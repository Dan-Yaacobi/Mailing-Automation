using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using PrintRequestApp.Core.Services.PageCounting;
using PrintRequestApp.ViewModels;

namespace PrintRequestApp.Views;

public partial class MainWindow : Window
{
    private static readonly string[] SupportedAttachmentExtensions = { ".pdf", ".doc", ".docx", ".ppt", ".pptx" };

    private readonly PageCounterFactory _pageCounterFactory = PageCounterFactory.CreateDefault();

    public ObservableCollection<AttachmentItemViewModel> Attachments { get; } = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
    }

    private void NumericTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !e.Text.All(char.IsDigit);
    }

    private async void AddFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "קבצי שכפול (*.pdf;*.doc;*.docx;*.ppt;*.pptx)|*.pdf;*.doc;*.docx;*.ppt;*.pptx"
        };

        if (dialog.ShowDialog(this) == true)
        {
            await AddFilesAsync(dialog.FileNames);
        }
    }

    private void Attachments_DragEnter(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private async void Attachments_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            await AddFilesAsync(paths);
        }
    }

    // Silently skips unsupported file types and exact-duplicate paths already attached.
    // Page counting runs on a dedicated STA thread per file (Office COM objects need
    // an STA apartment, and the thread pool is MTA) so the UI thread stays free to
    // keep the loading overlay's spinner animating and the window responsive instead
    // of looking frozen for the few seconds Word/PDF/PPTX counting can take (§3 of
    // docs/DESIGN.md).
    private async Task AddFilesAsync(IEnumerable<string> filePaths)
    {
        var pathsToProcess = filePaths
            .Where(path => SupportedAttachmentExtensions.Contains(Path.GetExtension(path).ToLowerInvariant()))
            .Where(path => !Attachments.Any(a => string.Equals(a.FilePath, path, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (pathsToProcess.Count == 0)
        {
            return;
        }

        LoadingOverlay.Visibility = Visibility.Visible;

        try
        {
            foreach (var filePath in pathsToProcess)
            {
                LoadingStatusText.Text = $"מזהה מספר עמודים: {Path.GetFileName(filePath)}";

                var fileKind = Core.Models.FileKindDetector.FromFilePath(filePath);
                var detectedPageCount = await CountPagesOnStaThreadAsync(fileKind, filePath);
                Attachments.Add(new AttachmentItemViewModel(filePath, fileKind, detectedPageCount));
            }
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }
    }

    private Task<int?> CountPagesOnStaThreadAsync(Core.Models.FileKind fileKind, string filePath)
    {
        var tcs = new TaskCompletionSource<int?>();

        var thread = new Thread(() =>
        {
            try
            {
                tcs.SetResult(_pageCounterFactory.TryCountPages(fileKind, filePath));
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return tcs.Task;
    }

    private void MarkAllColor_Click(object sender, RoutedEventArgs e)
    {
        foreach (var attachment in Attachments)
        {
            attachment.ColorMode = Core.Models.ColorMode.Color;
        }
    }

    private void MarkAllBlackAndWhite_Click(object sender, RoutedEventArgs e)
    {
        foreach (var attachment in Attachments)
        {
            attachment.ColorMode = Core.Models.ColorMode.BlackAndWhite;
        }
    }

    private void RemoveAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: AttachmentItemViewModel attachment })
        {
            Attachments.Remove(attachment);
        }
    }

    // No backend yet - this just collects and echoes what the form holds, so the
    // field set and control types can be sanity-checked before wiring anything real.
    private void Send_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(TxtCopiesCount.Text, out var copiesCount) || copiesCount < 1)
        {
            MessageBox.Show(
                this,
                "יש להזין מספר עותקים תקין (1 ומעלה)",
                "שגיאה - מספר עותקים חסר",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
                MessageBoxResult.OK,
                MessageBoxOptions.RtlReading);
            return;
        }

        if (Attachments.Count == 0)
        {
            MessageBox.Show(
                this,
                "לא ניתן לשלוח בקשה ללא קבצים מצורפים",
                "שגיאה - אין קבצים מצורפים",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
                MessageBoxResult.OK,
                MessageBoxOptions.RtlReading);
            return;
        }

        var missingPageCount = Attachments.Where(a => a.PageCount is null).ToList();
        if (missingPageCount.Count > 0)
        {
            var fileList = string.Join("\n", missingPageCount.Select(a => $"  - {a.FileName}"));
            MessageBox.Show(
                this,
                $"לא ניתן לשלוח: לא הוזן מספר עמודים עבור הקבצים הבאים:\n{fileList}",
                "שגיאה - חסר מספר עמודים",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
                MessageBoxResult.OK,
                MessageBoxOptions.RtlReading);
            return;
        }

        var pageTypeText = (string)((ComboBoxItem)CmbPageType.SelectedItem).Content;
        var request = new Core.Models.PrintRequest
        {
            SubmittedAt = DateTime.Now,
            ProgramName = TxtProgramName.Text,
            BudgetLine = TxtBudgetLine.Text,
            CopiesCount = copiesCount,
            HolePunch = IsYes(CmbHolePunch),
            DoubleSided = IsYes(CmbDoubleSided),
            Stapling = IsYes(CmbStapling),
            PageType = Enum.Parse<Core.Models.PaperSize>(pageTypeText),
            SlidesPerPage = int.TryParse(TxtSlidesPerPage.Text, out var slides) ? slides : null,
            Notes = string.IsNullOrWhiteSpace(TxtNotes.Text) ? null : TxtNotes.Text,
            Attachments = Attachments.Select(a => a.ToAttachmentItem()).ToList()
        };

        var confirmationDialog = new ConfirmationDialog(request) { Owner = this };
        if (confirmationDialog.ShowDialog() != true)
        {
            return;
        }

        Debug.WriteLine($"=== בקשה נשלחה (בדיקה בלבד): {request.ProgramName}, {request.Attachments.Count} קבצים ===");

        var colorPagesTotal = request.Attachments
            .Where(a => a.ColorMode == Core.Models.ColorMode.Color)
            .Sum(a => a.PageCount) * request.CopiesCount;
        var blackAndWhitePagesTotal = request.Attachments
            .Where(a => a.ColorMode == Core.Models.ColorMode.BlackAndWhite)
            .Sum(a => a.PageCount) * request.CopiesCount;

        Debug.WriteLine($"סה\"כ עמודים לשכפול (עמודים לקובץ × מספר עותקים) - צבעוני: {colorPagesTotal}, שחור-לבן: {blackAndWhitePagesTotal}");

        MessageBox.Show(
            this,
            "הבקשה נשלחה בהצלחה (בדיקה בלבד - טרם חובר backend אמיתי)",
            "נשלח",
            MessageBoxButton.OK,
            MessageBoxImage.Information,
            MessageBoxResult.OK,
            MessageBoxOptions.RtlReading);

        ResetForm();
    }

    // Ready for the next request: clears every field and every attached file.
    private void ResetForm()
    {
        TxtProgramName.Clear();
        TxtBudgetLine.Clear();
        TxtCopiesCount.Clear();
        CmbHolePunch.SelectedIndex = 0;
        CmbDoubleSided.SelectedIndex = 0;
        CmbStapling.SelectedIndex = 0;
        CmbPageType.SelectedIndex = 0;
        TxtSlidesPerPage.Clear();
        TxtNotes.Clear();
        Attachments.Clear();
    }

    private static bool IsYes(ComboBox comboBox) => comboBox.SelectedIndex == 1;
}
