using System;
using System.Collections.Generic;
using Microsoft.Win32;
using AsyncTask;
using Microsoft.Office.Interop.Excel;

namespace ExcelComparator
{
    public class ApplicationController
    {
        private Application xlApp;
        private Workbook xlWorkbook1;
        private Workbook xlWorkbook2;

        public string CurrentColumn { get; private set; }

        internal string OpenDialogAndGetFileName()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            return openFileDialog.FileName;
        }

        internal bool VerifyFieldFile(string file1, string file2)
        {
            
            bool ok = !string.IsNullOrEmpty(file1) && !string.IsNullOrEmpty(file2);
            if (!ok)
            {
                Tools.Message.Show("Les fichiers ne sont pas saisis.");
            }
            else
            {
                ok = VerifyExcelExtensions(file1) && VerifyExcelExtensions(file2);
                if (!ok)
                {
                    Tools.Message.Show("Les fichiers ne sont pas des fichiers Excel.");
                }
                else
                {
                    ok = FileExist(file1) && FileExist(file2);
                    if (!ok)
                    {
                        Tools.Message.Show("Les fichiers n'existent pas.");
                    }
                }
            }
            return ok;
        }

        private bool FileExist(string file)
        {
            return System.IO.File.Exists(file);
        }


        private bool VerifyExcelExtensions(string file)
        {
            ICollection<string> extensions = new List<string>() { "xls", "xlsx", "xlsm" };
            bool isGoodFile = false;
            string[] fileSplit = file.Split('.');
            int length = fileSplit.Length;
            if (extensions.Contains(fileSplit[length-1]))
            {
                isGoodFile = true;
            }
            return isGoodFile;
        }



        private Dictionary<int, string> GetColumnsNumber(string columns, Worksheet sheet)
        {
            Dictionary<int, string> indexColumns = new Dictionary<int, string>();
            Range xlRange = sheet.UsedRange;
            int NumberColumns = xlRange.Columns.Count;

            if (string.IsNullOrEmpty(columns))
            {
                for (int j = 1; j <= NumberColumns; j++)
                {
                    indexColumns.Add(j, (string)(xlRange.Cells[1, j] as Range).Value);
                }
            }
            else
            {
                string[] split = columns.Split(';');
                foreach (string columnId in split)
                {
                    for (int j = 1; j <= NumberColumns; j++)
                    {
                        if (((string)(xlRange.Cells[1, j] as Range).Value) == columnId)
                        {
                            indexColumns.Add(j, columnId);
                        }
                    }
                }
            }
            
            return indexColumns;
        }

        private Worksheet GetExcelWorkSheetToFile(Workbook book, string sheetNameFile1)
        {
            Worksheet sheet = null;
            if (string.IsNullOrEmpty(sheetNameFile1))
            {
                object sheetTemp = book.ActiveSheet;
                if(sheetTemp != null)
                {
                    sheet = (Worksheet)sheetTemp;
                }
            }
            else
            {
                foreach (Worksheet worksheet in book.Sheets)
                {
                    if (worksheet.Name.Equals(sheetNameFile1))
                    {
                        sheet =worksheet;
                        break;
                    }
                }

            }

            return sheet;
        }

        internal List<Worksheet> GetExcelWorkSheetToFiles(string file1, string sheetNameFile1, string file2, string sheetNameFile2)
        {
            List<Worksheet> sheets = new List<Worksheet>();
            xlApp = new Application();
            xlWorkbook1 = xlApp.Workbooks.Open(file1);
            xlWorkbook2 = xlApp.Workbooks.Open(file2);
            sheets.Add(GetExcelWorkSheetToFile(xlWorkbook1, sheetNameFile1));
            sheets.Add(GetExcelWorkSheetToFile(xlWorkbook2, sheetNameFile2));
            return sheets;
        }

        internal void Compare(string file1, string sheet1, string file2, string sheet2,string columns)
        {
            try
            {
                xlApp = (Application) System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                Tools.Message.Show("Excel n'est pas disponible : " + ex.Message);
            }

            try
            {
                List<Worksheet> worksheets = GetExcelWorkSheetToFiles(file1, sheet1, file2, sheet2);
                Dictionary<int, string> columnsFile1 = GetColumnsNumber(columns, worksheets[0]);
                Dictionary<int, string> columnsFile2 = GetColumnsNumber(columns, worksheets[1]);
                new RunAsyncTask().Run(worksheets, columnsFile1, columnsFile2);
                FreeExcel();
            }
            catch (Exception excep)
            {
                try
                {
                    Tools.Message.Show(excep.Message);
                    if (xlWorkbook1 != null)
                    {
                        xlWorkbook1.Close(0);
                    }
                    if (xlWorkbook2 != null)
                    {
                        xlWorkbook2.Close(0);
                    }
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                    }
                }
                catch
                {
                    //
                }
            }
        }

        internal bool VerifyNumberRowsEqual(List<Worksheet> worksheets)
        {
            List<int> columnsNumber = new List<int>();
            List<int> rowsNumber = new List<int>();
            bool ok = true;
            foreach (Worksheet sheet in worksheets)
            {
                columnsNumber.Add(sheet.Columns.Count);
                rowsNumber.Add(sheet.Rows.Count);
            }

            if (columnsNumber[0] != columnsNumber[1])
            {
                Tools.Message.Show("Le nombre de colonne entre les deux fichiers est différent.");
                ok = false;
            }

            if (rowsNumber[0] != rowsNumber[1])
            {
                Tools.Message.Show("Le nombre de ligne entre les deux fichiers est différent.");
                ok = false;
            }
            return ok;
            
        }

        private void FreeExcel()
        {
            GC.WaitForPendingFinalizers();
                      //close and release
            xlWorkbook1.Close();
            xlWorkbook2.Close();

            //quit and release
            xlApp.Quit();
        }

        
    }
}
