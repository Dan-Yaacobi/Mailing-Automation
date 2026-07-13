namespace PrintRequestApp.ViewModels;

// Display-ready row for the confirmation dialog's attachment list - keeps the
// Hebrew "צבעוני"/"שחור-לבן" display text out of Core (which stays UI-agnostic).
public sealed class AttachmentSummaryItem
{
    public required string FileName { get; init; }

    public required int PageCount { get; init; }

    public required string ColorModeDisplay { get; init; }
}
