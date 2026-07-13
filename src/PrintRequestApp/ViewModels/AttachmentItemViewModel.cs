using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.ViewModels;

public sealed class AttachmentItemViewModel : INotifyPropertyChanged
{
    private int? _manualPageCount;
    private ColorMode? _colorMode = Core.Models.ColorMode.BlackAndWhite;

    public AttachmentItemViewModel(string filePath, FileKind fileKind, int? detectedPageCount)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
        FileKind = fileKind;
        DetectedPageCount = detectedPageCount;
    }

    public string FilePath { get; }

    public string FileName { get; }

    public FileKind FileKind { get; }

    public int? DetectedPageCount { get; }

    public string DetectedPageCountDisplay => DetectedPageCount?.ToString() ?? "לא זוהה אוטומטית";

    public int? ManualPageCount
    {
        get => _manualPageCount;
        set
        {
            if (_manualPageCount == value)
            {
                return;
            }

            _manualPageCount = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(EffectivePageCount));
        }
    }

    public int? EffectivePageCount => ManualPageCount ?? DetectedPageCount;

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
