using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Configuration;

namespace Unitas.Framework
{
    public class ExcelService
    {
        private readonly IConfiguration _config;

        public ExcelService(IConfiguration config)
        {
            _config = config;
        }
        public void UpdateCell1(string filePath, string sheetName, string cellReference, string value)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
                if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));
                if (string.IsNullOrWhiteSpace(cellReference)) throw new ArgumentNullException(nameof(cellReference));

                using var document = SpreadsheetDocument.Open(filePath, true);

                var sheet = GetSheet(document, sheetName);
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                var (column, rowIndex) = ParseCellReference(cellReference);

                var row = GetOrCreateRow(sheetData, rowIndex);
                var cell = GetOrCreateCell(row, column, rowIndex);

                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                worksheetPart.Worksheet.Save();
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public void UpdateCell(string filePath, string sheetName, string cellReference, string newValue)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                    throw new Exception("WorkbookPart not found in file.");

                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found.");

                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

                // Get the cell reference (like "C5")
                Cell cell = worksheetPart.Worksheet.Descendants<Cell>()
                    .FirstOrDefault(c => c.CellReference == cellReference);

                if (cell == null)
                {
                    // Create cell if not exists
                    cell = CreateCell(worksheetPart, cellReference);
                }

                // Update cell value
                cell.CellValue = new CellValue(newValue);
                cell.DataType = CellValues.String;

                worksheetPart.Worksheet.Save();

                // 🔧 Force all formula recalculations
                foreach (var wsPart in workbookPart.WorksheetParts)
                {
                    foreach (var formulaCell in wsPart.Worksheet.Descendants<Cell>()
                                 .Where(c => c.CellFormula != null))
                    {
                        // Clear cached values
                        formulaCell.CellValue = null;
                        formulaCell.DataType = null;
                        formulaCell.CellFormula.CalculateCell = true;
                    }

                    wsPart.Worksheet.Save();
                }

                // ⚙ Ensure workbook-level recalculation properties
                var workbook = workbookPart.Workbook;
                var calcProps = workbook.GetFirstChild<CalculationProperties>();
                if (calcProps == null)
                {
                    calcProps = new CalculationProperties
                    {
                        CalculationId = 0,
                        FullCalculationOnLoad = true,
                        ForceFullCalculation = true
                    };
                    workbook.Append(calcProps);
                }
                else
                {
                    calcProps.FullCalculationOnLoad = true;
                    calcProps.ForceFullCalculation = true;
                }

                workbook.Save();
            }
        }


        public string ReadCellValue(string filePath, string sheetName, string cellReference)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));
            if (string.IsNullOrWhiteSpace(cellReference)) throw new ArgumentNullException(nameof(cellReference));

            using var document = SpreadsheetDocument.Open(filePath, false);

            var sheet = GetSheet(document, sheetName);
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

            if (cell == null) return string.Empty;

            var value = cell.InnerText;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                return document.WorkbookPart.SharedStringTablePart?
                           .SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText
                       ?? string.Empty;
            }

            return value;
        }

        private static Sheet GetSheet(SpreadsheetDocument document, string sheetName)
        {
            var sheet = document.WorkbookPart?.Workbook.Descendants<Sheet>()
                            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
                throw new InvalidOperationException($"Sheet '{sheetName}' not found.");

            return sheet;
        }

        private static (string column, uint rowIndex) ParseCellReference(string cellReference)
        {
            string column = new string(cellReference.Where(char.IsLetter).ToArray());
            uint rowIndex = uint.Parse(new string(cellReference.Where(char.IsDigit).ToArray()));

            return (column, rowIndex);
        }

        private static Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
        {
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);

            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            return row;
        }

        private static Cell GetOrCreateCell(Row row, string columnName, uint rowIndex)
        {
            var cellReference = columnName + rowIndex;

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
            if (cell == null)
            {
                cell = new Cell { CellReference = cellReference };
                row.Append(cell);
            }

            return cell;
        }

        private Cell CreateCell(WorksheetPart worksheetPart, string cellReference)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            string columnName = new string(cellReference.Where(Char.IsLetter).ToArray());
            uint rowIndex = uint.Parse(new string(cellReference.Where(Char.IsDigit).ToArray()));

            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            var cell = new Cell() { CellReference = cellReference };
            row.Append(cell);
            return cell;
        }


        #region Get Value
        public Dictionary<string, List<ExcelKeyValue>> ReadSectionData(string filePath)
        {
            var result = new Dictionary<string, List<ExcelKeyValue>>();

            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.OfType<Sheet>().First();
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            var excelSections = _config.GetSection("ExcelSections").GetChildren();

            foreach (var section in excelSections)
            {
                var sectionName = section.Key;
                var keyRange = section["KeyRange"];
                var valueRange = section["ValueRange"];

                var keyCells = GetCellsInRange(keyRange);
                var valueCells = GetCellsInRange(valueRange);

                var pairs = new List<ExcelKeyValue>();

                for (int i = 0; i < keyCells.Count && i < valueCells.Count; i++)
                {
                    string keyCellRef = keyCells[i];
                    string valueCellRef = valueCells[i];

                    // Read key text
                    string keyVal = GetCellValue(worksheetPart, workbookPart, keyCellRef);

                    // Read value + formula (formula may be blank)
                    var (valVal, formulaText) = GetCellValueAndFormula(worksheetPart, workbookPart, valueCellRef);

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
        private static string GetCellValue(WorksheetPart worksheetPart, WorkbookPart workbookPart, string cellRef)
        {
            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                            .FirstOrDefault(c => c.CellReference?.Value == cellRef);
            if (cell == null) return string.Empty;

            var value = cell.InnerText;
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var table = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int index))
                    return table.Elements<SharedStringItem>().ElementAt(index).InnerText;
            }
            return value;
        }

        private static (string value, string formula) GetCellValueAndFormula(WorksheetPart worksheetPart, WorkbookPart workbookPart, string cellRef)
        {
            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                            .FirstOrDefault(c => c.CellReference?.Value == cellRef);

            if (cell == null)
                return (string.Empty, string.Empty);

            string formulaText = cell.CellFormula?.Text ?? string.Empty;

      
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;

          
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStringTable != null && int.TryParse(rawValue, out int sharedStringIndex))
                {
                    rawValue = sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
                }
                return (rawValue, formulaText);
            }

         
            if (!double.TryParse(rawValue, out double numericVal))
                return (rawValue, formulaText);

           
            string formattedValue = rawValue;

            var styleIndex = cell.StyleIndex?.Value;
            if (styleIndex != null)
            {
                var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
                var cellFormats = stylesheet?.CellFormats;
                var cellFormat = cellFormats?.Elements<CellFormat>().ElementAt((int)styleIndex.Value);

                if (cellFormat != null && cellFormat.NumberFormatId != null)
                {
                    var formatId = cellFormat.NumberFormatId.Value;

                    switch (formatId)
                    {
                        case 9:
                            formattedValue = (numericVal * 100).ToString("0") + "%";
                            break;

                        case 10:
                            formattedValue = (numericVal * 100).ToString("0.00") + "%";
                            break;

                        case 164:
                            formattedValue = (numericVal * 100).ToString("0.00") + " %";
                            break;

                        case 4:
                        case 41:
                        case 42:
                            formattedValue = numericVal.ToString("C2");
                            break;

                        default:
                            formattedValue = numericVal.ToString("0.######");
                            break;
                    }
                }
                else
                {
                    formattedValue = numericVal.ToString("0.######");
                }
            }
            else
            {
                formattedValue = numericVal.ToString("0.######");
            }

            return (formattedValue, formulaText);
        }



        private static List<string> GetCellsInRange(string range)
        {
            var parts = range.Split(':');
            if (parts.Length != 2) throw new ArgumentException("Invalid range: " + range);

            string startCol = new string(parts[0].Where(char.IsLetter).ToArray());
            uint startRow = uint.Parse(new string(parts[0].Where(char.IsDigit).ToArray()));
            uint endRow = uint.Parse(new string(parts[1].Where(char.IsDigit).ToArray()));

            var cells = new List<string>();
            for (uint r = startRow; r <= endRow; r++)
                cells.Add(startCol + r);

            return cells;
        }



        #endregion

    }
}
