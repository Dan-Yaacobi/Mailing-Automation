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
            using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.InformationOnly);
            return document.PageCount;
        }
        catch
        {
            return null;
        }
    }
}
