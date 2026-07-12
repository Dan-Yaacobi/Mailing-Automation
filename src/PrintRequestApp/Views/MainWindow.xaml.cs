using System.Windows;

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
}
