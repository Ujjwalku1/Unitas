using System;
using System.Collections.Generic;
using System.Linq;
using GemBox.Spreadsheet;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

public class ExcelKeyValue
{
    public string KeyCell { get; set; }
    public string Key { get; set; }
    public string ValueCell { get; set; }
    public string Value { get; set; }
    public string Formula { get; set; }
}

public class ExcelServiceManager
{
    private readonly IConfiguration _config;
    private readonly ILogger<ExcelServiceManager> _logger;

    public ExcelServiceManager(IConfiguration config, ILogger<ExcelServiceManager> logger)
    {
        _config = config;
        _logger = logger;

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
    }

    public void UpdateAndRecalculate(string filePath, string sheetName, string cellRef, string newValue)
    {
        try
        {
            var workbook = ExcelFile.Load(filePath);
            var worksheet = workbook.Worksheets[sheetName];

            worksheet.Cells[cellRef].Value = newValue;

            workbook.Calculate();

            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogInformation("UpdateAndRecalculate Error: {Message} {StackTrace}", ex.Message, ex.StackTrace);
        }
    }

    public Dictionary<string, List<ExcelKeyValue>> ReadSectionData(string filePath)
    {
        var result = new Dictionary<string, List<ExcelKeyValue>>();

        try
        {
            var workbook = ExcelFile.Load(filePath);
            var worksheet = workbook.Worksheets[0];

            var excelSections = _config.GetSection("ExcelSections").GetChildren();

            foreach (var section in excelSections)
            {
                string sectionName = section.Key;
                string keyRange = section["KeyRange"];
                string valueRange = section["ValueRange"];

                var keyCells = GetCellsInRange(keyRange);
                var valueCells = GetCellsInRange(valueRange);

                var pairs = new List<ExcelKeyValue>();

                for (int i = 0; i < keyCells.Count && i < valueCells.Count; i++)
                {
                    string keyCellRef = keyCells[i];
                    string valueCellRef = valueCells[i];

                    var keyCell = worksheet.Cells[keyCellRef];
                    var valueCell = worksheet.Cells[valueCellRef];

                    string keyVal = keyCell.Value?.ToString().Trim() ?? string.Empty;
                    string formulaText = valueCell.Formula?.ToString() ?? string.Empty;
                    string valVal = GetFormattedValue(valueCell).Trim();

                    if (!string.IsNullOrWhiteSpace(keyVal))
                    {
                        pairs.Add(new ExcelKeyValue
                        {
                            KeyCell = keyCellRef,
                            Key = keyVal,
                            ValueCell = valueCellRef,
                            Value = valVal,
                            Formula = formulaText
                        });
                    }
                }

                result[sectionName] = pairs;
            }
        }
        catch (Exception ex)
        {
            _logger.LogInformation("ReadSectionData Error: {Message} {StackTrace}", ex.Message, ex.StackTrace);
        }

        return result;
    }

    private static string GetFormattedValue(ExcelCell cell)
    {
        if (cell == null || cell.Value == null)
            return string.Empty;

        // Try to use GetFormattedString if available (GemBox v38+)
        try
        {
            var method = typeof(ExcelCell).GetMethod("GetFormattedString");
            if (method != null)
            {
                return (string)method.Invoke(cell, null);
            }
        }
        catch
        {
            // ignore reflection errors and fallback
        }

        // Fallback: try to format manually based on cell style
        var format = cell.Style.NumberFormat.Format;

        if (!string.IsNullOrEmpty(format))
        {
            try
            {
                // This method formats a value according to the number format string
                return cell.GetFormattedValue(format);
            }
            catch
            {
                // fallback to plain value
            }
        }

        // Final fallback: just ToString()
        return cell.Value.ToString();
    }

    private static List<string> GetCellsInRange(string range)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException("Invalid range: " + range);

        var startAddress = new CellIndex(parts[0]);
        var endAddress = new CellIndex(parts[1]);

        var cells = new List<string>();

        for (int row = startAddress.Row; row <= endAddress.Row; row++)
        {
            for (int col = startAddress.Column; col <= endAddress.Column; col++)
            {
                var cellName = CellIndex.GetColumnLetter(col) + (row + 1);
                cells.Add(cellName);
            }
        }

        return cells;
    }

    private struct CellIndex
    {
        public int Row;
        public int Column;

        public CellIndex(string cellRef)
        {
            var letters = new string(cellRef.TakeWhile(char.IsLetter).ToArray());
            var digits = new string(cellRef.SkipWhile(char.IsLetter).ToArray());

            Column = GetColumnNumber(letters);
            Row = int.Parse(digits) - 1;
        }

        public static int GetColumnNumber(string colLetters)
        {
            int col = 0;
            foreach (char c in colLetters.ToUpper())
            {
                col = col * 26 + (c - 'A' + 1);
            }
            return col - 1;
        }

        public static string GetColumnLetter(int colNumber)
        {
            string colLetter = "";
            colNumber++;
            while (colNumber > 0)
            {
                int modulo = (colNumber - 1) % 26;
                colLetter = (char)('A' + modulo) + colLetter;
                colNumber = (colNumber - modulo) / 26;
            }
            return colLetter;
        }
    }
}
