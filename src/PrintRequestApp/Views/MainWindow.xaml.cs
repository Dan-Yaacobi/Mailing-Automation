using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PrintRequestApp.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
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

        var summary =
            $"שם התכנית: {programName}\n" +
            $"סעיף תקציבי: {budgetLine}\n" +
            $"מספר עותקים: {copiesCount}\n" +
            $"חירור: {YesNoText(holePunch)}\n" +
            $"דו\"צ: {YesNoText(doubleSided)}\n" +
            $"הידוק: {YesNoText(stapling)}\n" +
            $"סוג דף: {pageType}\n" +
            $"שקפים בעמוד: {(slidesPerPage?.ToString() ?? "לא הוזן")}\n" +
            $"הערות נוספות: {(string.IsNullOrWhiteSpace(notes) ? "אין" : notes)}";

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
