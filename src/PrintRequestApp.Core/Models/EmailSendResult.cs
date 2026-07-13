namespace PrintRequestApp.Core.Models;

public sealed class EmailSendResult
{
    public required bool Success { get; init; }

    public string? ErrorMessage { get; init; }

    public static EmailSendResult Succeeded() => new() { Success = true };

    public static EmailSendResult Failed(string errorMessage) => new() { Success = false, ErrorMessage = errorMessage };
}
