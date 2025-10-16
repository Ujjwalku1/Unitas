using Aspose.Cells;
using DocumentFormat.OpenXml.Office2016.Excel;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using Unitas.Framework;


public class ExcelServiceManager
{
    private readonly IConfiguration _config;
    private readonly ILogger<ExcelServiceManager> _logger;
    public ExcelServiceManager(IConfiguration config, ILogger<ExcelServiceManager> logger)
    {
        _config = config;
        _logger = logger;
    }
    public void UpdateAndRecalculate(string filePath, string sheetName, string cellRef, string newValue)
    {
        try
        {
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Cells.lic");

            Workbook workbook = new Workbook(filePath);
            Worksheet sheet = workbook.Worksheets[sheetName];
            sheet.Cells[cellRef].PutValue(newValue);
            workbook.CalculateFormula();
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            _logger.LogInformation("UpdateAndRecalculate'{Message}:{StackTrace}'", ex.Message, ex.StackTrace);
        }
    }
    public Dictionary<string, List<ExcelKeyValue>> ReadSectionData_Aspose(string filePath)
    {
        var result = new Dictionary<string, List<ExcelKeyValue>>();

        var workbook = new Workbook(filePath);
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

                string keyVal = keyCell.StringValue?.Trim();
                string formulaText = valueCell.Formula ?? string.Empty;
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

        return result;
    }


    private static string GetFormattedValue(Cell cell)
    {
        if (cell == null) return string.Empty;

        // Apply number format (currency, %, etc.)
        Style style = cell.GetStyle();
        int numberFormat = style.Number;

        // Aspose will format the display value correctly if we call this:
        return cell.StringValue;
    }
    private static List<string> GetCellsInRange(string range)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException("Invalid range: " + range);

        string startCol = new string(parts[0].Where(char.IsLetter).ToArray());
        uint startRow = uint.Parse(new string(parts[0].Where(char.IsDigit).ToArray()));
        uint endRow = uint.Parse(new string(parts[1].Where(char.IsDigit).ToArray()));

        var cells = new List<string>();
        for (uint r = startRow; r <= endRow; r++)
            cells.Add(startCol + r);

        return cells;
    }

}
