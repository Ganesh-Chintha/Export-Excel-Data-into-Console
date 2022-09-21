using Aspose.Cells;
using System;

namespace NewExcel2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // File Path
            Workbook input = new Workbook("C:/Users/Sree Ganesh/Desktop/Excel_ To_CSV/Worldwide Rig Count Aug 2022.xlsx");

            WorksheetCollection collection = input.Worksheets;

            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {

                Worksheet file = collection[worksheetIndex];

                // Here is the Excel file Header
                Console.WriteLine(file.Name + "\n");

                int rows = file.Cells.MaxDataRow;
                int cols = file.Cells.MaxDataColumn;

                // Selected perticular Rows.
                // Before this rows of data is Titles or Unwanted data
                for (int i = 6; i < 35; i++)
                {

                    // Loop through each column in selected row
                    for (int j = 1; j < cols; j++)
                    {
                        // Print cell data
                        // Print camma (,) after the cell data.
                        
                        Console.Write(file.Cells[i, j].Value + " , ");
                    }
                    Console.WriteLine("");
                }
            }
        }
    }
}
