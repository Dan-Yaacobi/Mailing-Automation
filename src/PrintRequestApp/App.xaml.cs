using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Markup;

namespace PrintRequestApp;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        var hebrewCulture = new CultureInfo("he-IL");
        Thread.CurrentThread.CurrentCulture = hebrewCulture;
        Thread.CurrentThread.CurrentUICulture = hebrewCulture;
        CultureInfo.DefaultThreadCurrentCulture = hebrewCulture;
        CultureInfo.DefaultThreadCurrentUICulture = hebrewCulture;

        FrameworkElement.LanguageProperty.OverrideMetadata(
            typeof(FrameworkElement),
            new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(hebrewCulture.IetfLanguageTag)));

        base.OnStartup(e);
    }
}
