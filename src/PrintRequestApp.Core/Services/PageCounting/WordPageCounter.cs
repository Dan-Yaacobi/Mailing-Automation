using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

// Fully late-bound COM automation (no Microsoft.Office.Interop.Word PIA reference):
// referencing that PIA pulls in a specific Office-version-pinned "office" core
// assembly dependency that isn't guaranteed to be present/resolvable at runtime
// (observed as an unhandled FileNotFoundException for "office, Version=15.0.0.0"
// on a real machine). Going through Word's ProgID + dynamic/IDispatch instead
// works against whatever Word version is actually installed, with no compile-time
// assembly-version dependency at all.
public sealed class WordPageCounter : IPageCounter
{
    private const int WdAlertsNone = 0;
    private const int WdDoNotSaveChanges = 0;

    // WdInformation.wdNumberOfPagesInDocument - confirmed to actually exist via
    // IDispatch on a real machine, unlike Document.ComputedStatistics, which a real
    // test showed COM automation doesn't expose at all ("does not contain a
    // definition for 'ComputedStatistics'") despite being the documented VBA API.
    private const int WdNumberOfPagesInDocument = 4;

    public FileKind SupportedKind => FileKind.Word;

    public int? TryCountPages(string filePath)
    {
        dynamic? wordApp = null;
        dynamic? document = null;

        try
        {
            var wordApplicationType = Type.GetTypeFromProgID("Word.Application");
            if (wordApplicationType is null)
            {
                Debug.WriteLine("[WordPageCounter] Word.Application ProgID not found - Word not installed?");
                return null;
            }

            wordApp = Activator.CreateInstance(wordApplicationType);
            wordApp!.Visible = false;
            wordApp.DisplayAlerts = WdAlertsNone;

            // Deliberately not passing Visible:false here (only Application.Visible
            // above suppresses on-screen rendering) - passing it to Open too appears
            // to suppress creation of the document's own window entirely, leaving
            // Selection null (confirmed by a real "cannot perform runtime binding on
            // a null reference" failure with it set).
            document = wordApp.Documents.Open(
                filePath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false);

            // Repaginate() first guarantees an accurate, live count rather than a
            // possibly-stale cached one (§6.2 of docs/DESIGN.md), then read the page
            // count off this document's own window/selection - more robust than the
            // Application's "currently active" selection.
            document.Repaginate();
            return (int)document.ActiveWindow.Selection.Information(WdNumberOfPagesInDocument);
        }
        catch (Exception ex)
        {
            // Covers Word not installed, corrupt/password-protected files, legacy
            // formats it refuses to open, etc. - falls back to manual entry like any
            // other detection failure. Logged so a real failure can be diagnosed
            // instead of silently guessed at.
            Debug.WriteLine($"[WordPageCounter] Failed to count pages for '{filePath}': {ex}");
            return null;
        }
        finally
        {
            if (document is not null)
            {
                document.Close(SaveChanges: WdDoNotSaveChanges);
                Marshal.ReleaseComObject(document);
            }

            if (wordApp is not null)
            {
                wordApp.Quit(SaveChanges: WdDoNotSaveChanges);
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}
