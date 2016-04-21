using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Bands
{
    class Program
    {
        public class Bands
        {
            public int position { get; set; }
            public string name { get; set; }
            public int numberofsongs { get; set; }
            public int time { get; set; }
           
        }

        static void Main(string[] args)
        {
            // Create a list of accounts.
            var band = new List<Bands> {
                  new Bands
                  {
                      position = 1,
                      name = "Queen",
                      numberofsongs = 10,
                      time = 30
                  }
            };

            // Display the list in an Excel spreadsheet.
            DisplayInExcel(band);

        }

        static void DisplayInExcel(IEnumerable<Bands> bands)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1] = "BANDS LIST";
            workSheet.Cells[2, "A"] = "Position";
            workSheet.Cells[2, "B"] = "Name";
            workSheet.Cells[2, "C"] = "Number of songs";
            workSheet.Cells[2, "D"] = "Time";

            var row = 2;
            foreach (var acct in bands)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.position;
                workSheet.Cells[row, "B"] = acct.name;
                workSheet.Cells[row, "C"] = acct.numberofsongs;
                workSheet.Cells[row, "D"] = acct.time;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
        }
    }
}
