using System.Collections.Generic;
using System.Linq;
using Spire.Xls;

namespace ExcelTable
{
    class SpireExcelReader
    {
        
        public Dictionary<int, double> columnWidths { get; private set; }
        public Dictionary<int, double> rowHeights { get; private set; }

        public List<Cell> data { get; private set; }
        public List<string> dataValue = new List<string>();

        public SpireExcelReader()
        {
            data = new List<Cell>();

            dataValue = new List<string>();

            columnWidths = new Dictionary<int, double>();

            rowHeights = new Dictionary<int, double>();
        }

        public void Read(string path)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);

            Worksheet sheet = workbook.Worksheets[0];

            int columnCount = sheet.Columns.Count();
            int rowCount = sheet.Rows.Count();

            

            for (int row = 1; row < rowCount + 1; row++)
            {
                for (int col = 1; col < columnCount + 1; col++)
                {
                    CellRange objRange = sheet.Range[row, col];
                    double previousColumn;
                    columnWidths.TryGetValue(col - 1, out previousColumn);
                    double previousRow;
                    rowHeights.TryGetValue(row - 1, out previousRow);

                    if (!columnWidths.Keys.Contains(col))
                    {
                        columnWidths.Add(col, objRange.ColumnWidth + previousColumn);
                    }

                    if (!rowHeights.Keys.Contains(row))
                    {
                        rowHeights.Add(row, objRange.RowHeight + previousRow);
                    }
                }
            }


            for (int row = 1; row < rowCount + 1; row++)
            {
                for (int col = 1; col < columnCount + 1; col++)
                {
                    CellRange objRange = sheet.Range[row, col];

                    double currentColumnWidth = objRange.ColumnWidth;
                    double currentRowHeight = objRange.RowHeight;

                    double previousColumn;
                    columnWidths.TryGetValue(col - 1, out previousColumn);
                    double previousRow;
                    rowHeights.TryGetValue(row - 1, out previousRow);

                    //string cellText = objRange.Text;
                    string cellText = objRange.DisplayedText ?? "ERROR";

                    if (objRange.HasMerged)
                    {
                        if (!dataValue.Contains(cellText))
                        {
                            int rowsMerged = objRange.MergeArea.RowCount;
                            int colsMerged = objRange.MergeArea.ColumnCount;

                            double mergedColumnWidth;
                            columnWidths.TryGetValue(colsMerged, out mergedColumnWidth);
                            double mergedRowHeight;
                            rowHeights.TryGetValue(rowsMerged, out mergedRowHeight);



                            Cell mergedCell = new Cell(row, col, cellText, x: mergedColumnWidth / 2 + previousColumn,
                                                        y: mergedRowHeight / 2 + previousRow,
                                                        cellWidth: mergedColumnWidth,
                                                        rowHeight: mergedRowHeight);

                            data.Add(mergedCell);

                            dataValue.Add(cellText);
                        }

                    }//close merged cell
                    else
                    {
                        Cell singleCell = new Cell(row, col, cellText, x: currentColumnWidth / 2 + previousColumn,
                                                                y: currentRowHeight / 2 + previousRow,
                                                                cellWidth: currentColumnWidth, rowHeight: currentRowHeight);

                        data.Add(singleCell);

                    }
                }
            }

            
        }
    }//close class
}//close namespace
