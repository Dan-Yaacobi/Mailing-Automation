using System.Collections.Generic;

namespace PrintRequestApp.Core.Models;

// One per submission - the request-level fields plus the attachments snapshot,
// built once client-side validation passes (§4 of docs/DESIGN.md).
public sealed class PrintRequest
{
    public required string ProgramName { get; init; }

    public required string BudgetLine { get; init; }

    public required int CopiesCount { get; init; }

    public bool HolePunch { get; init; }

    public bool DoubleSided { get; init; }

    public bool Stapling { get; init; }

    public required PaperSize PageType { get; init; }

    public int? SlidesPerPage { get; init; }

    public string? Notes { get; init; }

    public required IReadOnlyList<AttachmentItem> Attachments { get; init; }
}
