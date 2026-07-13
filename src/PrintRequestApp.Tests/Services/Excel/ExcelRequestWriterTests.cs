using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using PrintRequestApp.Core.Models;
using PrintRequestApp.Core.Services.Excel;
using Xunit;

namespace PrintRequestApp.Tests.Services.Excel;

// Runs against a fresh, empty workbook created per test (the "dummy file" the design
// doc's Excel phase asked to validate against before pointing anything at a real one),
// not a real production spreadsheet.
public sealed class ExcelRequestWriterTests : IDisposable
{
    private readonly string _dummyFilePath;

    public ExcelRequestWriterTests()
    {
        _dummyFilePath = Path.Combine(Path.GetTempPath(), $"PrintRequestApp_ExcelTest_{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("placeholder");
        workbook.SaveAs(_dummyFilePath);
    }

    public void Dispose()
    {
        if (File.Exists(_dummyFilePath))
        {
            File.Delete(_dummyFilePath);
        }
    }

    [Fact]
    public void TryWrite_CreatesMonthSheetWithHeaderAndRow()
    {
        var writer = new ExcelRequestWriter();
        var request = CreateRequest(submittedAt: new DateTime(2026, 7, 13));

        var succeeded = writer.TryWrite(request, _dummyFilePath);

        Assert.True(succeeded);

        using var workbook = new XLWorkbook(_dummyFilePath);
        var worksheet = workbook.Worksheet("07-2026");

        Assert.Equal("שם תכנית", worksheet.Cell(1, 2).GetString());
        Assert.Equal("תפעול הפורטל", worksheet.Cell(2, 2).GetString());
        Assert.Equal("222", worksheet.Cell(2, 3).GetString());
        Assert.Equal("document.docx", worksheet.Cell(2, 4).GetString());
        Assert.Equal("שחור-לבן", worksheet.Cell(2, 5).GetString());
        Assert.Equal(30, worksheet.Cell(2, 6).GetValue<int>()); // 5 pages x 6 copies
    }

    [Fact]
    public void TryWrite_DefaultsBlankProgramNameToKlali()
    {
        var writer = new ExcelRequestWriter();
        var request = CreateRequest(programName: "");

        writer.TryWrite(request, _dummyFilePath);

        using var workbook = new XLWorkbook(_dummyFilePath);
        var worksheet = workbook.Worksheet(request.SubmittedAt.ToString("MM-yyyy"));

        Assert.Equal("כללי", worksheet.Cell(2, 2).GetString());
    }

    [Fact]
    public void TryWrite_AppendsSecondRequestAsNewRowWithoutOverwriting()
    {
        var writer = new ExcelRequestWriter();
        var first = CreateRequest(submittedAt: new DateTime(2026, 7, 1));
        var second = CreateRequest(submittedAt: new DateTime(2026, 7, 20), programName: "תכנית אחרת");

        writer.TryWrite(first, _dummyFilePath);
        writer.TryWrite(second, _dummyFilePath);

        using var workbook = new XLWorkbook(_dummyFilePath);
        var worksheet = workbook.Worksheet("07-2026");

        Assert.Equal("תפעול הפורטל", worksheet.Cell(2, 2).GetString());
        Assert.Equal("תכנית אחרת", worksheet.Cell(3, 2).GetString());
    }

    [Fact]
    public void TryWrite_ReturnsFalse_WhenFileDoesNotExist()
    {
        var writer = new ExcelRequestWriter();
        var request = CreateRequest();
        var missingPath = Path.Combine(Path.GetTempPath(), $"does_not_exist_{Guid.NewGuid():N}.xlsx");

        Assert.False(writer.TryWrite(request, missingPath));
    }

    [Fact]
    public void TryWrite_ReturnsFalse_WhenFileIsLocked()
    {
        var writer = new ExcelRequestWriter();
        var request = CreateRequest();

        using var lockingHandle = new FileStream(_dummyFilePath, FileMode.Open, FileAccess.Read, FileShare.None);

        Assert.False(writer.TryWrite(request, _dummyFilePath));
    }

    private static PrintRequest CreateRequest(
        DateTime? submittedAt = null,
        string programName = "תפעול הפורטל")
    {
        return new PrintRequest
        {
            SubmittedAt = submittedAt ?? new DateTime(2026, 7, 13),
            ProgramName = programName,
            BudgetLine = "222",
            CopiesCount = 6,
            PageType = PaperSize.A4,
            Attachments = new List<AttachmentItem>
            {
                new()
                {
                    FilePath = @"C:\files\document.docx",
                    FileName = "document.docx",
                    FileKind = FileKind.Word,
                    PageCount = 5,
                    ColorMode = ColorMode.BlackAndWhite
                }
            }
        };
    }
}
