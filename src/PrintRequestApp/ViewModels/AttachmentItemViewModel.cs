using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.ViewModels;

public sealed class AttachmentItemViewModel : INotifyPropertyChanged
{
    private int? _pageCount;
    private ColorMode? _colorMode = Core.Models.ColorMode.BlackAndWhite;

    public AttachmentItemViewModel(string filePath, FileKind fileKind, int? detectedPageCount)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
        FileKind = fileKind;
        DetectedPageCount = detectedPageCount;
        _pageCount = detectedPageCount;
    }

    public string FilePath { get; }

    public string FileName { get; }

    public FileKind FileKind { get; }

    public int? DetectedPageCount { get; }

    // The one field the user actually looks at and edits - pre-filled from
    // detection when available, so there's a single number to reason about
    // instead of a separate "detected" display next to a separate "override" box.
    public int? PageCount
    {
        get => _pageCount;
        set
        {
            if (_pageCount == value)
            {
                return;
            }

            _pageCount = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(PageCountStatus));
        }
    }

    // Read-only companion label explaining where PageCount came from, or that
    // it's still missing - drives the red "needs attention" highlight too.
    public string PageCountStatus
    {
        get
        {
            if (PageCount is null)
            {
                return "⚠ יש להזין מספר עמודים";
            }

            return DetectedPageCount is not null && PageCount == DetectedPageCount
                ? "זוהה אוטומטית"
                : "הוזן ידנית";
        }
    }

    public ColorMode? ColorMode
    {
        get => _colorMode;
        set
        {
            if (_colorMode == value)
            {
                return;
            }

            _colorMode = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ColorModeIndex));
        }
    }

    // Bound to the color/B&W ComboBox's SelectedIndex: 0 = color, 1 = black & white
    // (default for every new attachment), -1 only if ColorMode is somehow cleared.
    public int ColorModeIndex
    {
        get => ColorMode switch
        {
            Core.Models.ColorMode.Color => 0,
            Core.Models.ColorMode.BlackAndWhite => 1,
            _ => -1
        };
        set
        {
            ColorMode = value switch
            {
                0 => Core.Models.ColorMode.Color,
                1 => Core.Models.ColorMode.BlackAndWhite,
                _ => (Core.Models.ColorMode?)null
            };
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
