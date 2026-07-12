using System.Windows;

namespace PrintRequestApp.Views;

public partial class ConfirmationDialog : Window
{
    public ConfirmationDialog()
    {
        InitializeComponent();
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
