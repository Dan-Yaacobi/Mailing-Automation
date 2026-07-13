using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.Outlook;

// Late-bound COM automation (Type.GetTypeFromProgID + dynamic), same approach as
// WordPageCounter and for the same reason: no compiled PIA reference means no
// Office-version-pinned assembly dependency that might not resolve on a given
// machine (§6.5 of docs/DESIGN.md). Creating a new Outlook.Application when Outlook
// is already running attaches to that existing instance automatically - Outlook's
// COM server does this itself, unlike Word, so there's no separate "attach to
// running instance" step needed here.
public sealed class OutlookEmailService : IOutlookEmailService
{
    private const int OlMailItem = 0;

    public EmailSendResult TrySendRequestEmail(PrintRequest request, string recipientEmail)
    {
        dynamic? outlookApp = null;
        dynamic? mailItem = null;

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType is null)
            {
                return EmailSendResult.Failed("Outlook אינו מותקן במחשב זה.");
            }

            outlookApp = Activator.CreateInstance(outlookType);
            mailItem = outlookApp!.CreateItem(OlMailItem);

            mailItem.To = recipientEmail;
            mailItem.Subject = $"בקשת שכפול - {request.ProgramName}";
            mailItem.HTMLBody = BuildHtmlBody(request);

            foreach (var attachment in request.Attachments)
            {
                mailItem.Attachments.Add(attachment.FilePath);
            }

            mailItem.Send();

            return EmailSendResult.Succeeded();
        }
        catch (Exception ex)
        {
            return EmailSendResult.Failed(ex.Message);
        }
        finally
        {
            if (mailItem is not null)
            {
                Marshal.ReleaseComObject(mailItem);
            }

            if (outlookApp is not null)
            {
                Marshal.ReleaseComObject(outlookApp);
            }
        }
    }

    // Page counts are deliberately excluded here - the person printing acts on file +
    // color/B&W + copies/finishing options, not page count, which is tracking data
    // for the Excel log rather than a printing instruction (per earlier discussion).
    private static string BuildHtmlBody(PrintRequest request)
    {
        var html = new StringBuilder();
        html.Append("<html dir=\"rtl\"><body style=\"font-family:'Segoe UI',sans-serif;font-size:14px;\">");
        html.Append($"<h2>בקשת שכפול - {WebUtility.HtmlEncode(request.ProgramName)}</h2>");
        html.Append("<table cellpadding=\"4\">");

        AppendRow(html, "שם התכנית", request.ProgramName);
        AppendRow(html, "סעיף תקציבי", request.BudgetLine);
        AppendRow(html, "מספר עותקים", request.CopiesCount.ToString());
        AppendRow(html, "חירור", request.HolePunch ? "כן" : "לא");
        AppendRow(html, "דו\"צ", request.DoubleSided ? "כן" : "לא");
        AppendRow(html, "הידוק", request.Stapling ? "כן" : "לא");
        AppendRow(html, "סוג דף", request.PageType.ToString());

        if (request.SlidesPerPage is not null)
        {
            AppendRow(html, "שקפים בעמוד", request.SlidesPerPage.Value.ToString());
        }

        if (!string.IsNullOrWhiteSpace(request.Notes))
        {
            AppendRow(html, "הערות נוספות", request.Notes);
        }

        html.Append("</table>");
        html.Append("<h3>קבצים מצורפים</h3><ul>");

        foreach (var attachment in request.Attachments)
        {
            var colorText = attachment.ColorMode == ColorMode.Color ? "צבעוני" : "שחור-לבן";
            html.Append($"<li>{WebUtility.HtmlEncode(attachment.FileName)} - {colorText}</li>");
        }

        html.Append("</ul></body></html>");
        return html.ToString();
    }

    private static void AppendRow(StringBuilder html, string label, string value)
    {
        html.Append("<tr><td style=\"font-weight:bold;\">");
        html.Append(WebUtility.HtmlEncode(label));
        html.Append("</td><td>");
        html.Append(WebUtility.HtmlEncode(value));
        html.Append("</td></tr>");
    }
}
