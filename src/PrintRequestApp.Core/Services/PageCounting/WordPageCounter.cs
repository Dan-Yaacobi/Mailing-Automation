using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
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
    private const int WdStatisticPages = 2;

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

            document = wordApp.Documents.Open(
                filePath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            // Document.ComputedStatistics can return a stale cached page count for a
            // document that hasn't been repaginated since it was last edited -
            // Repaginate() first guarantees an accurate, live count (§6.2 of docs/DESIGN.md).
            document.Repaginate();

            try
            {
                // Method-call shape, passing both parameters explicitly since
                // late-bound/IDispatch calls don't reliably support omitting the
                // second (optional in VBA) one the way a compiled PIA would.
                return (int)document.ComputedStatistics(WdStatisticPages, false);
            }
            catch (RuntimeBinderException)
            {
                // Some late-bound automation surfaces this as an indexed property
                // instead of a method - try that shape before giving up entirely.
                return (int)document.ComputedStatistics[WdStatisticPages];
            }
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
