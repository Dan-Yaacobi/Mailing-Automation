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

    // Disables the blank-page field whenever "cover page" is answered "no" -
    // that field is only meaningful when there is a cover page (§4 of docs/DESIGN.md).
    private void CmbHasCoverPage_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CmbBlankPageBetweenCoverAndContent == null)
        {
            // Fires once during InitializeComponent, before this later-declared
            // control exists yet.
            return;
        }

        CmbBlankPageBetweenCoverAndContent.IsEnabled = CmbHasCoverPage.SelectedIndex == 1;
        if (!CmbBlankPageBetweenCoverAndContent.IsEnabled)
        {
            CmbBlankPageBetweenCoverAndContent.SelectedIndex = 0;
        }
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
        var hasCoverPage = IsYes(CmbHasCoverPage);
        var blankPageBetweenCoverAndContent = hasCoverPage && IsYes(CmbBlankPageBetweenCoverAndContent);
        int? slidesPerPage = int.TryParse(TxtSlidesPerPage.Text, out var slides) ? slides : null;

        var summary =
            $"שם התכנית: {programName}\n" +
            $"סעיף תקציבי: {budgetLine}\n" +
            $"מספר עותקים: {copiesCount}\n" +
            $"חירור: {YesNoText(holePunch)}\n" +
            $"דו צדדי: {YesNoText(doubleSided)}\n" +
            $"הידוק: {YesNoText(stapling)}\n" +
            $"יש או אין דף פתיח: {YesNoText(hasCoverPage)}\n" +
            $"עמוד ריק בין דף פתיחה לתוכן: {(hasCoverPage ? YesNoText(blankPageBetweenCoverAndContent) : "לא רלוונטי")}\n" +
            $"שקפים בעמוד: {(slidesPerPage?.ToString() ?? "לא הוזן")}";

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
