#region Namespaces
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
//using System.Data;
//using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
#endregion

namespace xlScraper
{
    class Program
    {


        // parses list that is produced from 
        public static List<List<int>> ParseList (List<List<int>> ListName)
        {
            List<List<int>> outList = new List<List<int>>();

            int counter = -1;
            int bottom = 0;
            int top = 2000000;

            for (int i = 0; i < ListName.Count; i++)
            {
                int start;
                int end;


                // counter increments for each column
                counter++;

                if (bottom < ListName[i][2])
                {
                    bottom = ListName[i][2];
                }


                if (top > ListName[i][1])
                {
                    top = ListName[i][1];
                }

                try
                {
                    // identify last column of each table
                    if (ListName[i][0] + 1 != ListName[i + 1][0])
                    {

                        end = ListName[i][0];
                        start = ListName[i - counter][0];

                        outList.Add(new List<int>{start, end, top, bottom, ListName[i][3]});

                        //Console.WriteLine("{0}\t{1}: ({2}, {3}): {4}", i, counter, start, end, bottom);
                        counter = -1;
                        bottom = 0;
                    }

                }

                // if at the end of the list, set counter to zero
                catch (System.ArgumentOutOfRangeException)
                {
                    end = ListName[i][0];
                    start = ListName[i - counter][0];

                    outList.Add(new List<int> {start, end, top, bottom, ListName[i][3] });

                    //Console.WriteLine("{0}\t{1}: ({2}, {3}): {4}", i, counter, start, end, bottom);
                    counter = 0;
                    bottom = 0;
                }

                //Console.WriteLine("Start: {0}\tFinish: {1}\tCounter: {2}", start, end, counter);

            }


            return outList;
        }


        // inserts empty first row or column if doesn't exist
        public static void PrepareSheet (Excel.Worksheet sheet)
        {
                int firstFullRow;
                int firstFullColumn;
                               

                // check to see if first column and first row are completely blank. if either is not, create one.
                firstFullColumn = sheet.Cells[1, 1].End(Excel.XlDirection.xlDown).Row;
                firstFullRow = sheet.Cells[1, 1].End(Excel.XlDirection.xlToRight).Column;

                //Console.WriteLine("FirstFullColumn: {0}, ColCount: {1}, FirstFullRow: {2}, RowCount: {3}", firstFullColumn, sheet.Rows.Count, firstFullRow, sheet.Columns.Count);

                
                if (firstFullRow < sheet.Columns.Count)
                {
                    //insert 1st row
                    sheet.Rows[1].Insert();
                }


                if (firstFullColumn < sheet.Rows.Count)
                {
                    //insert 1st column
                    sheet.Columns[1].Insert();
                }

        }

        // Opens instance of Excel, defaults to visible
        public static Excel.Application OpenExcel (bool visible = true)
        {

            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = visible;
            return xlApp;
        }

        // closes Excel
        public static void CloseExcel (Excel.Application xlApp, Excel.Workbook wb)
        {
            // Close Excel
            object misValue = System.Reflection.Missing.Value;
            wb.Close(false, misValue, misValue);
            xlApp.Quit();
        }

        // gets data from excel file
        public static List<List<int>> GetExcelFile (Excel.Application xlApp, Excel.Workbook wb, Excel.Sheets ws, int colIndex = 16384)
        {

            List<List<int>> rangeSpec = new List<List<int>>();

            rangeSpec.Add(new List<int> {0,0,0});

            // Get sheet names
            foreach (Excel.Worksheet sheet in ws)
            {

                //Console.WriteLine("\n\n{0}\n", sheet.Name);

                int fullRow = sheet.Rows.Count; // gets count of all rows in excel spreadsheet. should be equal to 1,048,576 in excel 2007 and higher
                int lastRow;
                int firstRow;

                // inserts an empty first row or empty first column if does not exist
                PrepareSheet(sheet);


                // default value is 16384 and is extremely slow. use only what you need.
                for (int i = 1; i <= colIndex; i++)
                {
                    firstRow = sheet.Cells[1, i].End(Excel.XlDirection.xlDown).Row;
                    lastRow = sheet.Cells[fullRow, i].End(Excel.XlDirection.xlUp).Row;
                    
                    // if there is no data in the column, lastRow will return 1
                    if (lastRow == 1 && firstRow == fullRow)
                    {
                        continue;
                    }
                    else
                    {
                        rangeSpec.Add(new List<int> { i, firstRow, lastRow, sheet.Index });
                    }
                    
                    
                    //Console.WriteLine("Column: {0}, RowStart: {1}, RowEnd: {2}, Sheet: {3}", i, firstRow, lastRow, sheet.Name);
                }


            }

            rangeSpec.RemoveAt(0);
            return rangeSpec;

        }

        
        // creates file or appends to existing file
        public static void WriteTextFile (Excel.Application xlApp, Excel.Workbook wb, Excel.Sheets ws, List<List<int>> ListName, string outputDirectory)
        {


            // check if out folder exists. If not, create one.
            bool exists = System.IO.Directory.Exists(outputDirectory);

            if (!exists)
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            // loop through each line which represents
            foreach (List<int> row in ListName)
            {

                // row[4] contains the sheet index number
                int sheetIndex = row[4];
                Excel.Worksheet sht = wb.Sheets[sheetIndex];

                // this is a list of pipe-delimited strings that will be rows in the text file
                List<string> outText = new List<string>();
                int columnCount = row[1] - row[0];

                // get the address "$A$1:$B$2"cell range from r1c1 format
                string start = sht.Cells[row[2], row[0]].Address;
                string end = sht.Cells[row[3], row[1]].Address;

                string textFileName = outputDirectory;
                string delimiter = "|";

                string cell;

                cell = start.ToString() + ":" + end.ToString();

                Excel.Range rng = sht.get_Range(cell);

                int counter = 0;
                string line = "";
                string lineInput;

                foreach (Excel.Range thing in rng)
                {

                    if (thing.MergeCells)
                    {
                        textFileName += thing.Value2;
                        continue;
                    }

                    else if (thing.Value == null)
                    {
                        lineInput = "";
                    }
                    else
                    {
                        lineInput = (thing.Value).ToString();
                    }

                    if (counter < columnCount)
                    {
                        line += (lineInput + delimiter);
                    }
                    else if (counter == columnCount)
                    {
                        line += lineInput;
                        outText.Add(line);
                        line = "";
                        counter = -1;
                    }

                    counter++;
                }

                string dateTime = DateTime.Now.ToString("yyyyMMdd_Thhmmss");

                textFileName = textFileName.Replace(" ", "_");
                textFileName = textFileName.Replace("(", "_");
                textFileName = textFileName.Replace(")", "_");
                textFileName = textFileName.Replace("-", "_");
                textFileName = textFileName.Replace("__", "_");

                textFileName = "\\" + (textFileName + "_" + dateTime + ".txt").Replace("__", "_");




                // creates new file
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(textFileName, true))
                {
                    foreach (string strLine in outText)
                    {
                        file.WriteLine(strLine);
                    }
                            
                }

                Console.WriteLine(textFileName);
            }

         }





        static void Main(string[] args)
        {

            string inputFilePath = args[0];
            string outputFolderPath = args[1];

            string startTime = DateTime.Now.ToString("h:mm:ss tt");

            Excel.Application xlApp = OpenExcel(false);
            Excel.Workbook wb;
            Excel.Sheets ws;

            wb = xlApp.Workbooks.Open(inputFilePath);
            ws = wb.Sheets;

            // rangeSpec is used to find all 
            List<List<int>> rangeSpec = new List<List<int>>();
            List<List<int>> outList = new List<List<int>>();
            
            

            // limit to 78 columns (col BZ) for speed. 
            rangeSpec = GetExcelFile(xlApp, wb, ws, 78);

            Console.WriteLine("\n\n");
                      

            outList = ParseList(rangeSpec);


            WriteTextFile(xlApp, wb, ws, outList, outputFolderPath);


            Console.WriteLine("\n\nStart: {0}", startTime);

            Console.WriteLine("\n\nEnd: {0}", DateTime.Now.ToString("h:mm:ss tt"));

            // Pause
            Console.ReadKey();

            CloseExcel(xlApp, wb);

        }

    }
    
}
