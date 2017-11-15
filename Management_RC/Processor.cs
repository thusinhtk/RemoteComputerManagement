using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Management_RC
{
    class Processor
    {
        //Singleton 
        private static Processor instance = new Processor();

        private Processor()
        {
        }
        public static Processor getInstance()
        {
            return instance;
        }

        //Declare variables
        public const string _path = @"E:\\LEARNING\\Developer\\MR_Sequence22_Tools\\Management_RemoteComputer\\List.xls";

        public static string _computerName = Environment.MachineName;
        public static string _userName = Environment.UserName;
        
        public static Application excel = new Application();
        public static Workbook workbook = excel.Workbooks.Open(_path, ReadOnly: false, Editable: true);

        public void Run()
        {
            bool _isExistComputerNamInExcel = false;

            Worksheet currentSheet = GetCurrentSheet();

            _isExistComputerNamInExcel = IsExistComputer(currentSheet, _computerName);

            if (_isExistComputerNamInExcel)
            {
                //search index of cell contains value is computer name
                int numberUsedRow = 0;
                numberUsedRow = RowCount(currentSheet);

                for (int x = 2; x <= numberUsedRow; x++)
                {
                    //compare computer name of column 2
                    string excelComputerName = currentSheet.Rows.Cells[x, 2].Value;

                    if (excelComputerName.Equals(_computerName))
                    {
                        //update this cell for username and date time
                        currentSheet.Rows.Cells[x, 3].Value = _userName;
                        currentSheet.Rows.Cells[x, 4].Value = DateTime.Now;

                        SaveAndQuit();
                    }
                }
            }
            else
            {
                //write line in last off currentSheet

                int numberUsedRow = 0;
                numberUsedRow = RowCount(currentSheet);

                for (int x = 2; x <= numberUsedRow; x++)
                {
                    //compare computer name of column 2
                    string excelComputerName = currentSheet.Rows.Cells[x, 2].Value;
                    if (null == excelComputerName)
                    {
                        //insert this cell for computername, username and date time
                        currentSheet.Rows.Cells[x, 2].Value = _computerName;
                        currentSheet.Rows.Cells[x, 3].Value = _userName;
                        currentSheet.Rows.Cells[x, 4].Value = DateTime.Now;

                        SaveAndQuit();
                    }
                }
            }
        }

        private static void SaveAndQuit()
        {
            excel.Application.ActiveWorkbook.Save();

            excel.Application.Quit();
            excel.Quit();
            Console.WriteLine("Done");
            Environment.Exit(0);
        }

        private static bool IsExistComputer(Worksheet currentSheet, string computerName)
        {
            int numberUsedRow = 0;
            //int numberUsedColumn = 0;

            numberUsedRow = RowCount(currentSheet);
            //numberUsedColumn = ColumnCount(currentSheet);

            //from row 2 (except header of file excel)
            for (int x = 2; x <= numberUsedRow; x++)
            {
                //compare computer name of column 2
                string excelComputerName = currentSheet.Rows.Cells[x, 2].Value;
                if (excelComputerName == _computerName)
                {
                    return true;
                }
            }
            return false;
        }

        private static int ColumnCount(Worksheet currentSheet)
        {
            return currentSheet.Cells.Find("*", System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value,
                                                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns,
                                                Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                                false,
                                                System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value).Column;
        }

        private static int RowCount(Worksheet currentSheet)
        {
            return currentSheet.Cells.Find("*", System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value,
                                                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                                                Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                                false,
                                                System.Reflection.Missing.Value,
                                                System.Reflection.Missing.Value).Row;
        }

        private static Worksheet GetCurrentSheet()
        {
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            return worksheet;
        }
    }
}
