#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

            //string path = @"C:\Users\giovanni.brogiolo\Documents\190605 Anchor Schedule.xlsx";

            //TaskDialog.Show("resutl", data.Count.ToString());

            
            Dictionary<int, List<XYZ>> rowsPts = new Dictionary<int, List<XYZ>>();
            Dictionary<int, List<XYZ>> colPts = new Dictionary<int, List<XYZ>>();

            // TO BE ADDED: UPDATE TEXT ONLY (BASED ON TEXT POSITION)

            using (var form = new Form1())
            {

                form.ShowDialog();


                    string path = form.filePath;
                    double scaleWidth = form.widthFactor;
                    double scaleHeight = form.heightFactor;


                    //ExcelInteropReader excelReader = new ExcelInteropReader(); TOO SLOW

                    SpireExcelReader excelReader = new SpireExcelReader();
                    excelReader.Read(path);

                    using (Transaction t = new Transaction(doc, "Create table"))
                    {
                        t.Start();

                        StringBuilder sb = new StringBuilder();

                        foreach (Cell cell in excelReader.data)
                        {
                            //Rectangle(doc, new XYZ(ToFeet(cell.X) * scaleWidth, -ToFeet(cell.Y) * scaleHeight, 0), ToFeet(cell.CellWidth) * scaleWidth, ToFeet(cell.RowHeight) * scaleHeight);

                            #region Lines By Points
                            List<XYZ> corners = Corners(new XYZ(ToFeet(cell.X) * scaleWidth, -ToFeet(cell.Y) * scaleHeight, 0), ToFeet(cell.CellWidth) * scaleWidth, ToFeet(cell.RowHeight) * scaleHeight);

                            if (!rowsPts.ContainsKey(cell.Row))
                            {
                                rowsPts.Add(cell.Row, new List<XYZ> { corners[3], corners[2] });

                            }
                            else
                            {
                                rowsPts[cell.Row].Add(corners[3]);
                                rowsPts[cell.Row].Add(corners[2]);

                            }

                            if (!colPts.ContainsKey(cell.Column))
                            {
                                colPts.Add(cell.Column, new List<XYZ> { corners[3], corners[0] });
                            }
                            else
                            {
                                colPts[cell.Column].Add(corners[3]);
                                colPts[cell.Column].Add(corners[0]);
                            }
                            #endregion

                            try
                            {
                                string text = cell.Text;

                                if (cell.Text == "")
                                {
                                    text = "N/A";
                                }

                                TextNote tn = TextNote.Create(doc, doc.ActiveView.Id, new XYZ(ToFeet(cell.X) * scaleWidth, -ToFeet(cell.Y) * scaleHeight + ToFeet(cell.RowHeight) * scaleHeight * 0.35, ToFeet(cell.CellWidth) * scaleWidth), text, defaultTypeId);

                                tn.HorizontalAlignment = HorizontalTextAlignment.Center;

                            }
                            catch
                            {
                                sb.AppendLine($"{cell.Row} {cell.Column}");
                            }
                            
                        }

                        if (sb.Length > 0)
                        {
                            TaskDialog.Show("Error", $"Errors: {sb.ToString()}");
                        }
                     
                        //draw outer border
                        XYZ topLeftCorner = rowsPts.Values.First().First(); //0,0,0
                        XYZ rightX = rowsPts.Values.Last().Last();
                        XYZ bottomY = colPts.Values.Last().Last();

                        double tableWidth = rightX.X - topLeftCorner.X;
                        double tableHeight = bottomY.Y - topLeftCorner.Y;

                        XYZ midPoint = new XYZ(tableWidth / 2, tableHeight / 2, 0);

                        Rectangle(doc, midPoint, tableWidth, tableHeight);

                        //draw internal borders
                        #region Lines By Points creation

                        rowsPts.Remove(1);
                        colPts.Remove(1);

                        foreach (int item in rowsPts.Keys)
                        {
                            Line horLine = Line.CreateBound(rowsPts[item][0], rowsPts[item][rowsPts[item].Count - 1]);
                            doc.Create.NewDetailCurve(doc.ActiveView, horLine);

                        }
                        foreach (int item in colPts.Keys)
                        {
                            Line verLine = Line.CreateBound(colPts[item][0], colPts[item][colPts[item].Count - 1]);
                            doc.Create.NewDetailCurve(doc.ActiveView, verLine);

                        }
                        #endregion



                        t.Commit();
                    }//close transaction
                


            }

            ICollection<ElementId> fec = new FilteredElementCollector(doc, doc.ActiveView.Id).ToElementIds();

            uidoc.ShowElements(fec);

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

        private void PointCoords(XYZ point)
        {
            TaskDialog.Show("point", $"{point.X} {point.Y} {point.Z}");
        }


        private static List<XYZ> Corners(XYZ center, double width, double height)
        {
            List<XYZ> corners = new List<XYZ>();

            XYZ a = new XYZ(center.X - width / 2, center.Y - height / 2, 0);

            XYZ b = new XYZ(a.X + width, a.Y, 0);

            XYZ c = new XYZ(a.X + width, a.Y + height, 0);

            XYZ d = new XYZ(a.X, a.Y + height, 0);

            corners.Add(a);
            corners.Add(b);
            corners.Add(c);
            corners.Add(d);

            return corners;
        }
    }
}
