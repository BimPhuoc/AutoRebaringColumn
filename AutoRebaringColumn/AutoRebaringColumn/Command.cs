#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using StandardRebar;
#endregion

namespace AutoRebaringColumn
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,ref string message,ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Access current selection
            string path = @"D:\LAP TRINH\Addin\AutoRebaringColumn\AutoRebaringColumn\ThepCot.xlsm";
            Selection sel = uidoc.Selection;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");
                StandardBarColumn st = new StandardBarColumn(commandData, ref message, elements, path);
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}
