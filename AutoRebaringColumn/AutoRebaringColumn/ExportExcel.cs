using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;
using AutoRebaringColumn;
namespace Export
{
    class ExportExcel
    {
        #region "convert unit"
        const double f2m = 0.3048;
        const double f2mm = f2m * 1000;
        const double mm2f = 1 / f2mm;
        #endregion
        public ExportExcel(string path, List<double> L0_bien,List<double> L1_bien, List<double> L2_bien, List<string> Comment_bien)
        {
            ExcelFile ex = new ExcelFile(path);
            Workbook wb = ex.Workbook;
            wb.Save();
            Worksheet sheet = ex.Workbook.Worksheets["ColumnRebar"];
            for (int i = 0; i < L1_bien.Count; i++)
            {
                sheet.Range["L0_"].Offset[i].Value = (Math.Round(L0_bien[i]*f2mm/25,0)*25).ToString();
                sheet.Range["L1_"].Offset[i].Value = (Math.Round(L1_bien[i] * f2mm/25,0)*25).ToString();
                sheet.Range["L2_"].Offset[i].Value = (Math.Round(L2_bien[i] * f2mm/25,0)*25).ToString();
                sheet.Range["comment_"].Offset[i].Value = Comment_bien[i].ToString();
            }
        }
    }
}
