namespace PrintRequestApp.Core.Models;

// Immutable snapshot of one attached file at submit time - the WPF-side
// AttachmentItemViewModel is the mutable, bindable form used while the user is
// still editing (§4 of docs/DESIGN.md).
public sealed class AttachmentItem
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required FileKind FileKind { get; init; }

    public required int PageCount { get; init; }

    public required ColorMode ColorMode { get; init; }
}
