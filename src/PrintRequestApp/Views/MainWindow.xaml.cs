using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

    private void OpenConfirmationPreview_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ConfirmationDialog
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void NumericTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !e.Text.All(char.IsDigit);
    }

    private void AddFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "קבצי הדפסה (*.pdf;*.doc;*.docx;*.ppt;*.pptx)|*.pdf;*.doc;*.docx;*.ppt;*.pptx"
        };

        if (dialog.ShowDialog(this) == true)
        {
            AddFiles(dialog.FileNames);
        }
    }

    private void Attachments_DragEnter(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Attachments_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            AddFiles(paths);
        }
    }

    // Silently skips unsupported file types and exact-duplicate paths already attached.
    private void AddFiles(IEnumerable<string> filePaths)
    {
        foreach (var filePath in filePaths)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (!SupportedAttachmentExtensions.Contains(extension))
            {
                continue;
            }

            if (Attachments.Any(a => string.Equals(a.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var fileKind = Core.Models.FileKindDetector.FromFilePath(filePath);
            var detectedPageCount = _pageCounterFactory.TryCountPages(fileKind, filePath);
            Attachments.Add(new AttachmentItemViewModel(filePath, fileKind, detectedPageCount));
        }
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
        var programName = TxtProgramName.Text;
        var budgetLine = TxtBudgetLine.Text;
        int.TryParse(TxtCopiesCount.Text, out var copiesCount);
        var holePunch = IsYes(CmbHolePunch);
        var doubleSided = IsYes(CmbDoubleSided);
        var stapling = IsYes(CmbStapling);
        var pageType = (CmbPageType.SelectedItem as ComboBoxItem)?.Content as string;
        int? slidesPerPage = int.TryParse(TxtSlidesPerPage.Text, out var slides) ? slides : null;
        var notes = TxtNotes.Text;

        var attachmentsSummary = Attachments.Count == 0
            ? "אין קבצים מצורפים"
            : string.Join("\n", Attachments.Select(a =>
                $"  - {a.FileName}: {a.EffectivePageCount?.ToString() ?? "לא הוזן"} עמודים, " +
                (a.ColorMode == Core.Models.ColorMode.Color ? "צבעוני"
                    : a.ColorMode == Core.Models.ColorMode.BlackAndWhite ? "שחור-לבן"
                    : "לא נבחר")));

        var summary =
            $"שם התכנית: {programName}\n" +
            $"סעיף תקציבי: {budgetLine}\n" +
            $"מספר עותקים: {copiesCount}\n" +
            $"חירור: {YesNoText(holePunch)}\n" +
            $"דו\"צ: {YesNoText(doubleSided)}\n" +
            $"הידוק: {YesNoText(stapling)}\n" +
            $"סוג דף: {pageType}\n" +
            $"שקפים בעמוד: {(slidesPerPage?.ToString() ?? "לא הוזן")}\n" +
            $"הערות נוספות: {(string.IsNullOrWhiteSpace(notes) ? "אין" : notes)}\n\n" +
            $"קבצים מצורפים ({Attachments.Count}):\n{attachmentsSummary}";

        Debug.WriteLine("=== נתוני הבקשה שנאספו (טרם נשלח בפועל) ===");
        Debug.WriteLine(summary);

        MessageBox.Show(
            this,
            summary,
            "נתוני הבקשה שנאספו (בדיקה בלבד)",
            MessageBoxButton.OK,
            MessageBoxImage.Information,
            MessageBoxResult.OK,
            MessageBoxOptions.RtlReading);
    }

    private static bool IsYes(ComboBox comboBox) => comboBox.SelectedIndex == 1;

    private static string YesNoText(bool value) => value ? "כן" : "לא";
}
