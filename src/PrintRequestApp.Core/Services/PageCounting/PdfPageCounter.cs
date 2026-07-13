using PrintRequestApp.Core.Models;
using UglyToad.PdfPig;

namespace PrintRequestApp.Core.Services.PageCounting;

public sealed class PdfPageCounter : IPageCounter
{
    public FileKind SupportedKind => FileKind.Pdf;

    public int? TryCountPages(string filePath)
    {
        try
        {
            using var document = PdfDocument.Open(filePath);
            return document.NumberOfPages;
        }
        catch
        {
            return null;
        }
    }
}
