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
using PrintRequestApp.Core.Services.Excel;
using PrintRequestApp.Core.Services.Outlook;
using PrintRequestApp.Core.Services.PageCounting;
using PrintRequestApp.ViewModels;

namespace PrintRequestApp.Views;

public partial class MainWindow : Window
{
    private static readonly string[] SupportedAttachmentExtensions = { ".pdf", ".doc", ".docx", ".ppt", ".pptx" };

    // Manual-testing wiring only, per the design doc's Excel phase: a dummy file on
    // the Desktop, not the real shared production log (that path doesn't exist yet -
    // still pending a walkthrough with whoever owns it).
    private static readonly string DummyExcelLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        "PrintRequestApp_DummyLog.xlsx");

    private readonly PageCounterFactory _pageCounterFactory = PageCounterFactory.CreateDefault();
    private readonly ExcelRequestWriter _excelWriter = new();
    private readonly IOutlookEmailService _outlookEmailService = new OutlookEmailService();

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
                var detectedPageCount = await RunOnStaThreadAsync(() => _pageCounterFactory.TryCountPages(fileKind, filePath));
                Attachments.Add(new AttachmentItemViewModel(filePath, fileKind, detectedPageCount));
            }
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }
    }

    // Office COM objects need an STA apartment (the thread pool is MTA) - shared by
    // page counting and Outlook sending, both of which also need the UI thread free
    // so the loading overlay's spinner can actually animate (§3 of docs/DESIGN.md).
    private static Task<T> RunOnStaThreadAsync<T>(Func<T> action)
    {
        var tcs = new TaskCompletionSource<T>();

        var thread = new Thread(() =>
        {
            try
            {
                tcs.SetResult(action());
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

    private async void Send_Click(object sender, RoutedEventArgs e)
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

        var recipientEmail = (string)((ComboBoxItem)CmbRecipient.SelectedItem).Content;

        LoadingStatusText.Text = "שולח אימייל...";
        LoadingOverlay.Visibility = Visibility.Visible;
        Core.Models.EmailSendResult emailResult;
        try
        {
            emailResult = await RunOnStaThreadAsync(() => _outlookEmailService.TrySendRequestEmail(request, recipientEmail));
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }

        // The email is the actual submission mechanism (§9.2 of docs/DESIGN.md) - a
        // failure here blocks, with a clear error, rather than silently continuing.
        if (!emailResult.Success)
        {
            MessageBox.Show(
                this,
                $"שליחת האימייל נכשלה:\n{emailResult.ErrorMessage}",
                "שגיאה - שליחת אימייל נכשלה",
                MessageBoxButton.OK,
                MessageBoxImage.Error,
                MessageBoxResult.OK,
                MessageBoxOptions.RtlReading);
            return;
        }

        Debug.WriteLine($"=== בקשה נשלחה: {request.ProgramName}, {request.Attachments.Count} קבצים, נשלח אל {recipientEmail} ===");

        var colorPagesTotal = request.Attachments
            .Where(a => a.ColorMode == Core.Models.ColorMode.Color)
            .Sum(a => a.PageCount) * request.CopiesCount;
        var blackAndWhitePagesTotal = request.Attachments
            .Where(a => a.ColorMode == Core.Models.ColorMode.BlackAndWhite)
            .Sum(a => a.PageCount) * request.CopiesCount;

        Debug.WriteLine($"סה\"כ עמודים לשכפול (עמודים לקובץ × מספר עותקים) - צבעוני: {colorPagesTotal}, שחור-לבן: {blackAndWhitePagesTotal}");

        // Manual-testing wiring only (§9.5 of docs/DESIGN.md) - writes to a dummy file
        // on the Desktop, not the real shared production log. Doesn't block the
        // success message below even if it fails - the email already went out.
        ExcelRequestWriter.EnsureDummyWorkbookExists(DummyExcelLogPath);
        var excelWriteSucceeded = _excelWriter.TryWrite(request, DummyExcelLogPath);
        var excelStatusText = excelWriteSucceeded
            ? $"נשמר ליומן הבדיקה (Excel):\n{DummyExcelLogPath}"
            : "הכתיבה ליומן הבדיקה (Excel) נכשלה - ראה Debug Output";

        Debug.WriteLine(excelWriteSucceeded
            ? $"נכתב ליומן האקסל הדמה: {DummyExcelLogPath}"
            : "כתיבה ליומן האקסל הדמה נכשלה");

        MessageBox.Show(
            this,
            $"הבקשה נשלחה בהצלחה בדוא\"ל אל {recipientEmail}\n\n{excelStatusText}",
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
        CmbRecipient.SelectedIndex = 0;
    }

    private static bool IsYes(ComboBox comboBox) => comboBox.SelectedIndex == 1;
}
