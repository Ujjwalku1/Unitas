using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;  // for .xlsx
using NPOI.HSSF.UserModel;  // for .xls
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Configuration;


public class ExcelKeyValue
{
    public string KeyCell { get; set; }
    public string Key { get; set; }
    public string ValueCell { get; set; }
    public string Value { get; set; }
    public string Formula { get; set; }
}
public class ExcelNpoiService
{
    private readonly IConfiguration _config;

    public ExcelNpoiService(IConfiguration config)
    {
        _config = config;
    }

    public Dictionary<string, List<ExcelKeyValue>> ReadSectionData(string filePath)
    {
        var result = new Dictionary<string, List<ExcelKeyValue>>();

        IWorkbook workbook;
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = WorkbookFactory.Create(fs); // supports .xls and .xlsx
        }

        IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        ISheet sheet = workbook.GetSheetAt(0); // or use workbook.GetSheet("SheetName");

        var excelSections = _config.GetSection("ExcelSections").GetChildren();

        foreach (var section in excelSections)
        {
            string sectionName = section.Key;
            string keyRange = section["KeyRange"];     // e.g. "A2:A10"
            string valueRange = section["ValueRange"]; // e.g. "B2:B10"

            var keyCells = GetCellsInRange(keyRange);
            var valueCells = GetCellsInRange(valueRange);

            var pairs = new List<ExcelKeyValue>();

            for (int i = 0; i < keyCells.Count && i < valueCells.Count; i++)
            {
                string keyCellRef = keyCells[i];
                string valueCellRef = valueCells[i];

                ICell keyCell = GetCellByRef(sheet, keyCellRef);
                ICell valueCell = GetCellByRef(sheet, valueCellRef);

                string keyVal = GetCellValue(keyCell, evaluator);
                (string valVal, string formulaText) = GetCellValueAndFormula(valueCell, evaluator);

                if (!string.IsNullOrWhiteSpace(keyVal))
                {
                    pairs.Add(new ExcelKeyValue
                    {
                        KeyCell = keyCellRef,
                        Key = keyVal.Trim(),
                        ValueCell = valueCellRef,
                        Value = valVal?.Trim(),
                        Formula = formulaText
                    });
                }
            }

            result[sectionName] = pairs;
        }

        return result;
    }

    private ICell GetCellByRef(ISheet sheet, string cellRef)
    {
        var cr = new CellReference(cellRef);
        IRow row = sheet.GetRow(cr.Row);
        return row?.GetCell(cr.Col);
    }

    private string GetCellValue(ICell cell, IFormulaEvaluator evaluator)
    {
        if (cell == null) return string.Empty;

        if (cell.CellType == CellType.Formula)
        {
            CellValue eval = evaluator.Evaluate(cell);
            if (eval == null) return string.Empty;

            return eval.CellType switch
            {
                CellType.String => eval.StringValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                    ? cell.DateCellValue?.ToString("yyyy-MM-dd")
                    : $"{eval.NumberValue:0.######}",
                CellType.Boolean => eval.BooleanValue.ToString(),
                CellType.Error => $"#ERROR:{eval.ErrorValue}",
                _ => string.Empty
            };
        }

        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue,
            CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? cell.DateCellValue?.ToString("yyyy-MM-dd")
                : $"{cell.NumericCellValue:0.######}",
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Blank => string.Empty,
            _ => string.Empty
        };
    }

    private (string value, string formula) GetCellValueAndFormula(ICell cell, IFormulaEvaluator evaluator)
    {
        if (cell == null)
            return (string.Empty, string.Empty);

        string formulaText = cell.CellType == CellType.Formula ? cell.CellFormula : string.Empty;
        string value = GetCellValue(cell, evaluator);

        return (value, formulaText);
    }

    private List<string> GetCellsInRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            return new List<string>();

        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException("Invalid range: " + range);

        var start = new CellReference(parts[0]);
        var end = new CellReference(parts[1]);

        if (start.Col != end.Col)
            throw new NotSupportedException("Only vertical ranges (e.g., A2:A10) are supported.");

        List<string> cells = new List<string>();

        for (int row = start.Row; row <= end.Row; row++)
        {
            var cr = new CellReference(row, start.Col);
            cells.Add(cr.FormatAsString());
        }

        return cells;
    }
}