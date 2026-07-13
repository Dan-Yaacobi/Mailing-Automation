using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

public sealed class PptxSlideCounter : IPageCounter
{
    public FileKind SupportedKind => FileKind.PowerPoint;

    public int? TryCountPages(string filePath)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, false);
            return presentation.PresentationPart?.SlideParts.Count();
        }
        catch
        {
            // Also covers legacy binary .ppt, which Open XML SDK can't open at all
            // (§6.3 of docs/DESIGN.md) - falls back to manual entry like any other
            // detection failure.
            return null;
        }
    }
}
