using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.Excel;

public interface IExcelRequestWriter
{
    /// <summary>Appends one row per attachment to the shared log. Returns false (never
    /// throws) on any failure - missing/unreachable file, locked file, etc. - so the
    /// caller can fall back to the reconciliation email instead (§9.4 of docs/DESIGN.md).</summary>
    bool TryWrite(PrintRequest request, string filePath);
}
