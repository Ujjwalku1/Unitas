using Aspose.Cells;
using System;

public class ExcelAsposeExample
{
    public static void UpdateAndRecalculate(string filePath, string sheetName, string cellRef, string newValue)
    {
        // Load workbook
        Workbook workbook = new Workbook(filePath);

        // Access the worksheet
        Worksheet sheet = workbook.Worksheets[sheetName];

        // Update cell value (dropdown or normal)
        sheet.Cells[cellRef].PutValue(newValue);

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Example: Read dependent formula cell (say "C5" changes when dropdown "B2" changes)
        var dependentCell = sheet.Cells["C5"];
        Console.WriteLine($"Updated value in C5: {dependentCell.StringValue}");

        // Save updated file
        workbook.Save(filePath);
    }
}
