using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;


namespace ReadExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //Instantiate the spreadsheet creation engine
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //Initialize application
                    IApplication app = excelEngine.Excel;

                    //Set default application version as Xlsx
                    app.DefaultVersion = ExcelVersion.Xlsx;

                    //Open existing Excel workbook from the specified location
                    string inputFileName = @"C:\Users\azureadmin\Documents\Book1.xlsx";
                    IWorkbook workbook = app.Workbooks.Open(inputFileName, ExcelOpenType.Automatic);

                    //for(int workbookCounts = 0; workbookCounts <= workbook.Worksheets.Count; workbookCounts++)
                    //{
                    //    Console.WriteLine(workbook.Worksheets[workbookCounts].Name);
                    //}
                    //Get worksheet by name
                    IWorksheet empDetailsSheet = workbook.Worksheets["Sheet1"];
                    //Access the first worksheet
                    IWorksheet worksheet = workbook.Worksheets[0];

                    var empDetailsTable = worksheet.ListObjects.Where(lb => lb.Name == "tblRegionalSettings").First();

                    //for (int empTableRowCount = 1; empTableRowCount < empDetailsTable.TotalsRowCount; empTableRowCount++)
                    //{
                    //    for (int empTableColumnsCount = 0; empTableColumnsCount < empDetailsTable.Columns.Count; empTableColumnsCount++)
                    //    {
                    //        Console.WriteLine(empDetailsTable.Columns[empTableColumnsCount]);
                    //        //Console.Write(worksheet[empDetailsTable.row, col].Value);
                    //        Console.Write("\t\t");
                    //    }
                    //}

                    //Access the used range of the Excel file
                    IRange usedRange = worksheet.UsedRange;
                    int lastRow = usedRange.LastRow;
                    int lastColumn = usedRange.LastColumn;


                    //Iterate the cells in the used range and print the cell values
                    for (int row = 1; row <= lastRow; row++)
                    {
                        //Get value using cell name
                        Console.WriteLine("CellValue with name: " + worksheet.Range["varFirstName"].Value);


                        for (int col = 1; col <= lastColumn; col++)
                        {
                            Console.Write(worksheet[row, col].Value);
                            Console.Write("\t\t");
                        }
                        Console.WriteLine("\n");
                    }
                    Console.WriteLine("\n\n");
                    //Iterate the cells in the used range and print the display text
                    for (int row = 1; row <= lastRow; row++)
                    {
                        for (int col = 1; col <= lastColumn; col++)
                        {
                            Console.Write(worksheet[row, col].DisplayText);
                            Console.Write("\t\t");
                        }
                        Console.WriteLine("\n");
                    }
                    Console.Read();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }


        }
    }
}
