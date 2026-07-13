using System.IO;

namespace PrintRequestApp.Core.Models;

public static class FileKindDetector
{
    public static FileKind FromFilePath(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".pdf" => FileKind.Pdf,
            ".doc" or ".docx" => FileKind.Word,
            ".ppt" or ".pptx" => FileKind.PowerPoint,
            _ => FileKind.Unsupported
        };
    }
}
