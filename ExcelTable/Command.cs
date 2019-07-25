#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace ExcelTable
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;


            ElementId defaultTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

            string path = @"C:\Users\giovanni.brogiolo\Documents\190605 Anchor Schedule.xlsx";

            ExcelReader excelReader = new ExcelReader();

            excelReader.Read(path);

            //TaskDialog.Show("resutl", data.Count.ToString());

            double scaleWidth = 2;
            double scaleHeight = 0.3;

            using (Transaction t = new Transaction(doc, "Create table"))
                {
                    t.Start();

                    foreach (Cell cell in excelReader.data)
                    {
                    Rectangle(doc, new XYZ(ToFeet(cell.X) * scaleWidth, -ToFeet(cell.Y)*scaleHeight, 0), ToFeet(cell.CellWidth) * scaleWidth, ToFeet(cell.RowHeight) * scaleHeight);

                    

                    TextNote tn = TextNote.Create(doc, doc.ActiveView.Id, 
                        new XYZ( ToFeet(cell.X) * scaleWidth - ToFeet(cell.CellWidth) / 2 * scaleWidth, 
                        - ToFeet(cell.Y) * scaleHeight + ToFeet(cell.RowHeight)*scaleHeight*0.35, //+ ToFeet(cell.RowHeight) * scaleHeight / 2, 
                        0), 
                        ToFeet(cell.CellWidth) * scaleWidth, cell._value, defaultTypeId);
                    tn.HorizontalAlignment = HorizontalTextAlignment.Center;
                }
                    

                    t.Commit();
                }


            return Result.Succeeded;
        }

        private double ToFeet(double value)
        {
            return value / 304.8;
        }

        private void Rectangle(Document doc, XYZ center, double width, double height)
        {


            XYZ a = new XYZ(center.X - width / 2, center.Y - height / 2, 0);

            //PointCoords(a);

            XYZ b = new XYZ(a.X + width, a.Y, 0);
            //PointCoords(b);

            XYZ c = new XYZ(a.X + width, a.Y + height, 0);
            //PointCoords(c);

            XYZ d = new XYZ(a.X, a.Y + height, 0);
            //PointCoords(d);

            Line baseLine = Line.CreateBound(a, b);
            Line offsetHLine = Line.CreateBound(d, c);

            Line verticalLine = Line.CreateBound(a, d);
            Line offsetVlLine = Line.CreateBound(b, c);

            doc.Create.NewDetailCurve(doc.ActiveView, baseLine);
            doc.Create.NewDetailCurve(doc.ActiveView, verticalLine);
            doc.Create.NewDetailCurve(doc.ActiveView, offsetHLine);
            doc.Create.NewDetailCurve(doc.ActiveView, offsetVlLine);

        }

    }
}
