using PdfSharp.Pdf.IO;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

public sealed class PdfPageCounter : IPageCounter
{
    public FileKind SupportedKind => FileKind.Pdf;

    public int? TryCountPages(string filePath)
    {
        try
        {
            // InformationOnly is documented by PdfSharp itself as not actually
            // implemented; Import is the library's own recommended replacement.
            using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
            return document.PageCount;
        }
        catch
        {
            return null;
        }
    }
}
