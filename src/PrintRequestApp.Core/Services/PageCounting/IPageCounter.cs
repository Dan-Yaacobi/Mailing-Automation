using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

public interface IPageCounter
{
    FileKind SupportedKind { get; }

    /// <summary>Returns the page/slide count, or null if it can't be determined (caller falls back to manual entry).</summary>
    int? TryCountPages(string filePath);
}
