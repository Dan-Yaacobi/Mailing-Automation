using System.Windows;

namespace PrintRequestApp.Views;

public partial class ConfirmationDialog : Window
{
    public ConfirmationDialog(string summaryText)
    {
        InitializeComponent();
        TxtSummary.Text = summaryText;
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
}
