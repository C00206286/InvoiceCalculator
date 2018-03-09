using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace WindowsFormsApp2
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)

        //public static void getExcelFile()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        public static void excelChange()
        {

            if (Form1.fileGot == true)
            {

                //string file = System.Reflection.Assembly.GetExecutingAssembly().Location;
                //string file2 = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                //file2 = file2.Replace("ExcelManager.exe", "Invoice");
                //file2 = file2.Replace("file:///", "");

                //Console.WriteLine(file2);

                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@Form1.fileSelected);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                if (xlRange.Cells[1, 1].Value2 != "ContactName")
                {
                    MessageBox.Show("It appears you have selected an invalid file.", "Incorrect File",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                { 
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                string invoiceNum = "start";
                string invoiceNum2 = "start";
                double totalAdd = 0;
                double totalTax = 0;
                double grandTotal = 0;
                bool last = false;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)

                //for (int i = 1; i <= 1; i++)

                {
                    for (int j = 1; j <= colCount; j++)
                    //for (int j = 1; j <= 1; j++)
                    {
                        //new line
                        if (j == 1)
                            //Console.Write("\r\n");

                            //write the value to the console
                            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {

                            }
                        //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                    }
                }
                for (int i = 2; i <= rowCount; i++)

                //for (int i = 1; i <= 1; i++)

                {

                    if (xlRange.Cells[i, 17] != null && xlRange.Cells[i, 17].Value2 != null)
                    {

                        double quantity = xlRange.Cells[i, 17].Value2;
                        double amount = xlRange.Cells[i, 18].Value2;
                        double total = quantity * amount;
                        xlRange.Cells[i, 19].Value2 = total;
                    }


                }
                    for (int i = 2; i <= rowCount; i++)

                    //for (int i = 1; i <= 1; i++)

                    {
                        if (last == false)
                        {
                            totalAdd = totalAdd + xlRange.Cells[i, 19].Value2;
                            totalTax = totalTax + xlRange.Cells[i, 22].Value2;
                        }
                        if (xlRange.Cells[i, 10].Value2 != xlRange.Cells[i + 1, 10].Value2)
                        {
                            last = true;
                        }
                        if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                        {
                            invoiceNum = xlRange.Cells[i, 10].Value2;
                        }
                        //if(invoiceNum == invoiceNum2 || invoiceNum2 == "start")


                        if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                        {
                            invoiceNum2 = xlRange.Cells[i, 10].Value2;
                        }
                        if (last == true)
                        {
                            xlRange.Cells[i, 20].Value2 = totalAdd;
                            xlRange.Cells[i, 23].Value2 = totalTax;
                            grandTotal = totalAdd + totalTax;
                            xlRange.Cells[i, 24].Value2 = grandTotal;

                            totalTax = 0;
                            totalAdd = 0;
                            grandTotal = 0;
                            last = false;
                        }

                    }
                }





                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                Console.ReadLine();
            }
        }

    }
}
