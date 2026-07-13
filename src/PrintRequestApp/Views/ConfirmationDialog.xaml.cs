using System.Linq;
using System.Windows;
using PrintRequestApp.Core.Models;
using PrintRequestApp.ViewModels;

namespace PrintRequestApp.Views;

public partial class ConfirmationDialog : Window
{
    public ConfirmationDialog(PrintRequest request)
    {
        InitializeComponent();

        TxtProgramNameValue.Text = request.ProgramName;
        TxtBudgetLineValue.Text = request.BudgetLine;
        TxtCopiesCountValue.Text = request.CopiesCount.ToString();
        TxtHolePunchValue.Text = YesNoText(request.HolePunch);
        TxtDoubleSidedValue.Text = YesNoText(request.DoubleSided);
        TxtStaplingValue.Text = YesNoText(request.Stapling);
        TxtPageTypeValue.Text = request.PageType.ToString();
        TxtSlidesPerPageValue.Text = request.SlidesPerPage?.ToString() ?? "לא רלוונטי";
        TxtNotesValue.Text = string.IsNullOrWhiteSpace(request.Notes) ? "אין" : request.Notes;

        AttachmentsList.ItemsSource = request.Attachments.Select(a => new AttachmentSummaryItem
        {
            FileName = a.FileName,
            PageCount = a.PageCount,
            ColorModeDisplay = a.ColorMode == ColorMode.Color ? "צבעוני" : "שחור-לבן"
        }).ToList();
    }

    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private static string YesNoText(bool value) => value ? "כן" : "לא";
}
