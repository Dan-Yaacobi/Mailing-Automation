using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.Outlook;

public interface IOutlookEmailService
{
    /// <summary>Composes and sends the request as an HTML email with the original
    /// files attached. Never throws - failures (Outlook not installed, send
    /// rejected, etc.) come back as a failed EmailSendResult, since this is the
    /// primary submission mechanism and the caller needs a clear error to show
    /// rather than a silent fallback (§9.2 of docs/DESIGN.md).</summary>
    EmailSendResult TrySendRequestEmail(PrintRequest request, string recipientEmail);
}
