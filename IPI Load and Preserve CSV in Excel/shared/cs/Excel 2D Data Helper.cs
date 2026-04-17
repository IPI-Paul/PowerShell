using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

public class ExcelHelper 
{
    // If you ever want to paste numbers/dates without forcing text use object[] instead of string[]
    public static void FastPaste(object worksheetObj, object[] data, int startRow, int startCol) 
    {
        // Handle empty arrays safely
        if (data == null || data.Length == 0) return;

        Worksheet sheet = (Worksheet)worksheetObj;
        int rowCount = data.Length;

        // Create the 2D array inside C# where marshaling is stable
        object[,] data2D = new object[rowCount, 1];
        for (int i = 0; i < rowCount; i++)
        {
            data2D[i, 0] = data[i];
        }

        // Release COM objects (prevents "Excel stay running")
        Range startCell = null;
        Range targetRange = null;

        try {
            // Define the range and set Value2
            startCell = (Range)sheet.Cells[startRow, startCol];

            // Use get_Resize(), this usually costs fewer COM calls
            targetRange = startCell.get_Resize(rowCount, 1);

            targetRange.Value2 = data2D;
        } 
        finally
        {
            if (targetRange != null) Marshal.ReleaseComObject(targetRange);
            if (startCell != null) Marshal.ReleaseComObject(startCell);
        }
    }
}