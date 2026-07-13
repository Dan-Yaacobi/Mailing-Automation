using System;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using PrintRequestApp.Core.Models;

namespace PrintRequestApp.Core.Services.Excel;

// Schema and behavior here are a provisional best-guess (§9.5 of docs/DESIGN.md) -
// pending a real walkthrough with whoever owns the production spreadsheet. One
// worksheet per calendar month/year (named "MM-yyyy"), one row per attachment:
// date, program name (defaults to "כללי" when blank - not every copy job is for a
// named program), budget line, file name, color mode, and total pages
// (page count x copies count for that request).
public sealed class ExcelRequestWriter : IExcelRequestWriter
{
    private const int WriteRetryAttempts = 3;
    private static readonly TimeSpan WriteRetryDelay = TimeSpan.FromMilliseconds(500);

    private static readonly string[] HeaderRow =
    {
        "תאריך",
        "שם תכנית",
        "סעיף תקציבי",
        "שם קובץ",
        "סוג צבע",
        "סה\"כ עמודים"
    };

    public bool TryWrite(PrintRequest request, string filePath)
    {
        if (!File.Exists(filePath))
        {
            return false;
        }

        for (var attempt = 1; attempt <= WriteRetryAttempts; attempt++)
        {
            try
            {
                WriteRows(request, filePath);
                return true;
            }
            catch (IOException) when (attempt < WriteRetryAttempts)
            {
                // Most likely the file is open/locked elsewhere (another submission
                // landing at the same moment, or someone has it open in Excel) -
                // worth a couple of short retries before giving up.
                Thread.Sleep(WriteRetryDelay);
            }
            catch
            {
                return false;
            }
        }

        return false;
    }

    private static void WriteRows(PrintRequest request, string filePath)
    {
        using var workbook = new XLWorkbook(filePath);

        var sheetName = request.SubmittedAt.ToString("MM-yyyy");
        var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
            ?? CreateMonthSheet(workbook, sheetName);

        var programName = string.IsNullOrWhiteSpace(request.ProgramName) ? "כללי" : request.ProgramName;
        var nextRow = (worksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

        foreach (var attachment in request.Attachments)
        {
            var totalPages = attachment.PageCount * request.CopiesCount;
            var colorText = attachment.ColorMode == ColorMode.Color ? "צבעוני" : "שחור-לבן";

            worksheet.Cell(nextRow, 1).Value = request.SubmittedAt;
            worksheet.Cell(nextRow, 1).Style.DateFormat.Format = "dd/MM/yyyy";
            worksheet.Cell(nextRow, 2).Value = programName;
            worksheet.Cell(nextRow, 3).Value = request.BudgetLine;
            worksheet.Cell(nextRow, 4).Value = attachment.FileName;
            worksheet.Cell(nextRow, 5).Value = colorText;
            worksheet.Cell(nextRow, 6).Value = totalPages;

            nextRow++;
        }

        workbook.Save();
    }

    private static IXLWorksheet CreateMonthSheet(XLWorkbook workbook, string sheetName)
    {
        var worksheet = workbook.Worksheets.Add(sheetName);
        worksheet.RightToLeft = true;

        for (var i = 0; i < HeaderRow.Length; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = HeaderRow[i];
            cell.Style.Font.Bold = true;
        }

        return worksheet;
    }
}
