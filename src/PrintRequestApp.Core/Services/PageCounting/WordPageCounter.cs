using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

public sealed class WordPageCounter : IPageCounter
{
    public FileKind SupportedKind => FileKind.Word;

    public int? TryCountPages(string filePath)
    {
        Application? wordApp = null;
        Document? document = null;

        try
        {
            wordApp = new Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };

            document = wordApp.Documents.Open(
                FileName: filePath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            // Document.ComputedStatistics can return a stale cached page count for a
            // document that hasn't been repaginated since it was last edited -
            // Repaginate() first guarantees an accurate, live count (§6.2 of docs/DESIGN.md).
            document.Repaginate();

            // Late-bound: this PIA version doesn't statically expose ComputedStatistics,
            // but the real Word COM object supports it via IDispatch regardless.
            dynamic dynamicDocument = document;
            return (int)dynamicDocument.ComputedStatistics(WdStatistic.wdStatisticPages);
        }
        catch
        {
            // Covers Word not installed, corrupt/password-protected files, legacy
            // formats it refuses to open, etc. - falls back to manual entry like any
            // other detection failure.
            return null;
        }
        finally
        {
            if (document is not null)
            {
                document.Close(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
                Marshal.ReleaseComObject(document);
            }

            if (wordApp is not null)
            {
                wordApp.Quit(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}
