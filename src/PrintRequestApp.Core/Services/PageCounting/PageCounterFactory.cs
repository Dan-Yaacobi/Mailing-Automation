using System.Collections.Generic;
using System.Linq;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.PageCounting;

public sealed class PageCounterFactory
{
    private readonly IReadOnlyList<IPageCounter> _counters;

    public PageCounterFactory(IEnumerable<IPageCounter> counters)
    {
        _counters = counters.ToList();
    }

    public static PageCounterFactory CreateDefault()
    {
        return new PageCounterFactory(new IPageCounter[]
        {
            new PdfPageCounter(),
            new PptxSlideCounter(),
            new WordPageCounter()
        });
    }

    public int? TryCountPages(FileKind fileKind, string filePath)
    {
        var counter = _counters.FirstOrDefault(c => c.SupportedKind == fileKind);
        return counter?.TryCountPages(filePath);
    }
}
