using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;

namespace AutoRebaringColumn
{
    public class ExcelFile
    {
        public Excel.Application Application { get; private set; }
        public Excel.Workbook Workbook { get; private set; }
        public Excel.Worksheet Worksheet { get; private set; }
        public ExcelFile(string path)
        {
            Excel.Application exApp = null;
            Excel.Workbook workBook = null;
            if (ExcelHandler.IsOpenedWB_ByPath(path))
            {
                workBook = ExcelHandler.GetOpenedWB_ByPath(path);
            }
            else if (File.Exists(path))
            {
                exApp = new Excel.Application();
                exApp.Visible = true;
                exApp.DisplayAlerts = false;
                workBook = exApp.Workbooks.Open(path);
            }
            else
            {
                throw new Exception("Excel file is not exist!");
            }
            Application = exApp; Workbook = workBook;
        }
        public void Save(string saveAsPath, TypeSave type)
        {
            switch (type)
            {
                case TypeSave.Save:
                    Workbook.Save();
                    break;
                case TypeSave.SaveAs:
                    Workbook.SaveAs(saveAsPath, XlFileFormat.xlWorkbookDefault);
                    break;
            }
        }
        public void Close()
        {
            Workbook.Close();
            Application.Quit();
        }
    }
    public enum TypeSave
    {
        Save, SaveAs
    }
    static class ExcelHandler
    {
        public static void OpenExcelFile(string path)
        {
            Excel.Application exApp = null;
            Excel.Workbook workBook = null;
            Excel.Worksheet workSheet = null;
            try
            {
                if (IsOpenedWB_ByPath(path))
                {
                    workBook = GetOpenedWB_ByPath(path);
                    workSheet = workBook.Worksheets.Add() as Excel.Worksheet;
                    workSheet = workBook.ActiveSheet;
                }
                else
                {
                    exApp = new Excel.Application();
                    exApp.Visible = true;
                    exApp.DisplayAlerts = false;
                    if (!File.Exists(path))
                    {
                        workBook = exApp.Workbooks.Add(Type.Missing);
                        workSheet = workBook.ActiveSheet;
                    }
                    else
                    {
                        throw new Exception("Excel file is not exist!");
                    }
                }
            }
            finally
            {
            }
        }
        public static void CreateExcelFile(string cPath, string cName)
        {
            Excel.Application exApp = null;
            Excel.Workbook workBook = null;
            Excel.Worksheet workSheet = null;
            DateTime date = DateTime.Now;
            bool isExisted = false;
            try
            {
                string path = Path.Combine(cPath, cName + ".xlsx");

                if (IsOpenedWB_ByPath(path))
                {
                    isExisted = true;
                    workBook = GetOpenedWB_ByPath(path);
                    workSheet = workBook.Worksheets.Add() as Excel.Worksheet;
                    workSheet.Activate();
                    //workSheet = workBook.ActiveSheet;
                    workSheet.Name = date.ToString(new CultureInfo("de-DE")).Replace(':', '.');
                }
                else
                {
                    exApp = new Excel.Application();
                    exApp.Visible = true;
                    exApp.DisplayAlerts = false;
                    if (!File.Exists(path))
                    {
                        isExisted = false;
                        workBook = exApp.Workbooks.Add(Type.Missing);
                        workSheet = workBook.ActiveSheet;
                        workSheet.Name = date.ToString(new CultureInfo("de-DE")).Replace(':', '.');
                    }
                    else
                    {
                        isExisted = true;
                        workBook = exApp.Workbooks.Open(path);
                        workSheet = workBook.Worksheets.Add() as Worksheet;
                        workSheet.Activate();
                        //workSheet = workBook.ActiveSheet;
                        workSheet.Name = date.ToString(new CultureInfo("de-DE")).Replace(':', '.');
                    }
                }

                //for (int i = 0; i < cDataTable.Columns.Count; i++)
                //{
                //    workSheet.Cells[1, i + 1] = cDataTable.Columns[i].ColumnName;
                //}

                //for (int i = 0; i < cDataTable.Rows.Count; i++)
                //{
                //    for (int j = 0; j < cDataTable.Columns.Count; j++)
                //    {
                //        workSheet.Cells[i + 2, j + 1] = cDataTable.Rows[i].ItemArray[j];
                //    }
                //}

                if (isExisted)
                    workBook.SaveAs(path, XlFileFormat.xlWorkbookDefault);
                else
                    workBook.SaveAs(path, XlFileFormat.xlWorkbookDefault);
                //workBook.Close();
                //exApp.Quit();
            }
            //catch (System.Exception ex)
            //{
            //    AcadApp.ShowAlertDialog(ex.Message);
            //}
            finally
            {
                //workSheet = null;
                //workBook = null;
                //exApp = null;
            }
        }
        public static bool IsOpenedWB_ByName(string wbName)
        {
            return (GetOpenedWB_ByName(wbName) != null);
        }

        public static bool IsOpenedWB_ByPath(string wbPath)
        {
            return (GetOpenedWB_ByPath(wbPath) != null);
        }

        public static Workbook GetOpenedWB_ByName(string wbName)
        {
            return (Workbook)GetRunningObjects().FirstOrDefault(x => (System.IO.Path.GetFileName(x.Path) == wbName) && (x.Obj is Workbook)).Obj;
        }

        public static Workbook GetOpenedWB_ByPath(string wbPath)
        {
            return (Workbook)GetRunningObjects().FirstOrDefault(x => (x.Path == wbPath) && (x.Obj is Workbook)).Obj;
        }
        public static List<RunningObject> GetRunningObjects()
        {
            // Get the table.
            List<RunningObject> roList = new List<RunningObject>();
            IBindCtx bc;
            CreateBindCtx(0, out bc);
            IRunningObjectTable runningObjectTable;
            bc.GetRunningObjectTable(out runningObjectTable);
            IEnumMoniker monikerEnumerator;
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            // Enumerate and fill list
            IMoniker[] monikers = new IMoniker[1];
            IntPtr numFetched = IntPtr.Zero;
            List<object> names = new List<object>();
            List<object> books = new List<object>();
            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                RunningObject running;
                monikers[0].GetDisplayName(bc, null, out running.Path);
                runningObjectTable.GetObject(monikers[0], out running.Obj);
                roList.Add(running);
            }
            return roList;
        }

        public struct RunningObject
        {
            public string Path;
            public object Obj;
        }

        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        static extern void CreateBindCtx(int a, out IBindCtx b);
    }
}
