using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel; // For .xlsx files
using System;
using System.IO;

public class ExcelNPOIExample
{
    public static void UpdateAndRecalculate(string filePath, string sheetName, string cellRef, string newValue)
    {
        // Load the workbook
        IWorkbook workbook;
        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file); // For .xlsx files
        }

        // Get the sheet
        ISheet sheet = workbook.GetSheet(sheetName);
        if (sheet == null)
        {
            Console.WriteLine($"Sheet '{sheetName}' not found.");
            return;
        }

        // Convert cell reference (e.g., "B2") to row and column
        CellReference cr = new CellReference(cellRef);
        IRow row = sheet.GetRow(cr.Row) ?? sheet.CreateRow(cr.Row);
        ICell cell = row.GetCell(cr.Col) ?? row.CreateCell(cr.Col);

        // Update the cell value
        cell.SetCellValue(newValue);

        // Note: NPOI does not recalculate formulas automatically like Aspose
        // So we must tell Excel to do it by setting:
        sheet.ForceFormulaRecalculation = true;

        // Read value of a formula cell, e.g., "C5"
        CellReference depCr = new CellReference("C5");
        IRow depRow = sheet.GetRow(depCr.Row);
        if (depRow != null)
        {
            ICell depCell = depRow.GetCell(depCr.Col);
            if (depCell != null)
            {
                Console.WriteLine($"Formula in C5 (will update in Excel): {depCell.CellFormula}");
            }
        }

        // Save the updated workbook
        using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(outFile);
        }
    }
}
