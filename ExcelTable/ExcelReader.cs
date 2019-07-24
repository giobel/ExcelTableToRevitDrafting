using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTable
{
    class ExcelReader
    {
        public List<Cell> data { get; private set; }
        public List<string> dataValue { get; private set; }
        public Dictionary<int, double> columnWidths { get; private set; }
        public Dictionary<int, double> rowHeights { get; private set; }

        public ExcelReader()
        {
            data = new List<Cell>();

            dataValue = new List<string>();

            columnWidths = new Dictionary<int, double>();
        }

        public void Read(string path)
        {

            Application excelApplication = new Application();
            Workbook excelWorkBook = excelApplication.Workbooks.Open(path);

            Worksheet workSheet = excelWorkBook.Sheets[1];

            Sheets sheets = excelWorkBook.Sheets;

            int workSheetCounts = excelWorkBook.Worksheets.Count;

            int numberOfRows = sheets.Item[1].UsedRange.Cells.Rows.Count;

            int numberOfColumns = sheets.Item[1].UsedRange.Cells.Columns.Count;

            Range objRange = null;

            rowHeights = new Dictionary<int, double>();

            for (int row = 1; row < numberOfRows + 1; row++)
            {
                for (int col = 1; col < numberOfColumns + 1; col++)
                {
                    objRange = workSheet.Cells[row, col];

                    if (!columnWidths.Keys.Contains(col))
                    {
                        double previousColumn;
                        columnWidths.TryGetValue(col - 1, out previousColumn);
                        columnWidths.Add(col, objRange.ColumnWidth + previousColumn);
                    }

                    if (!rowHeights.Keys.Contains(row))
                    {
                        double previousRow;
                        rowHeights.TryGetValue(row - 1, out previousRow);
                        rowHeights.Add(row, objRange.RowHeight + previousRow);
                    }
                }
            }


            for (int row = 1; row < numberOfRows + 1; row++)
            {
                for (int col = 1; col < numberOfColumns + 1; col++)
                {
                    objRange = workSheet.Cells[row, col];

                    double currentColumnWidth = objRange.ColumnWidth;
                    double currentRowHeight = objRange.RowHeight;


                    if (objRange.MergeCells)
                    {
                        dynamic mergeAreaValue2 = objRange.MergeArea.Value2;

                        object[,] vals = mergeAreaValue2 as object[,];

                        double totalColumnWidth = objRange.MergeArea.Width;
                        double totalRowHeight = objRange.MergeArea.Height;

                        //NOT SURE WHY THE VALUES ARE WRONG: Excel's measurement system is not in pixels but in Points (1/72 of an inch). 
                        // https://social.msdn.microsoft.com/Forums/office/en-US/e0e176ab-fb54-4377-80f5-05473c81e8b8/excelrangewidth-wrong-pixel-size-also-topleftheight-c?forum=exceldev

                        string cellText = Convert.ToString(((Range)objRange.MergeArea[1, 1]).Text).Trim();

                        if (!dataValue.Contains(cellText))
                        {
                            int rowsMerged = vals.GetLength(0);
                            int colsMerged = vals.GetLength(1);

                            Console.WriteLine($"{vals.GetValue(1, 1)} {rowsMerged} {colsMerged}");

                            double previousColumn;
                            columnWidths.TryGetValue(col - 1, out previousColumn);
                            double previousRow;
                            rowHeights.TryGetValue(row - 1, out previousRow);

                            data.Add(new Cell(row, col, cellText, x: previousColumn + (columnWidths[col + (colsMerged - 1)] - previousColumn) / 2,
                                                                  y: previousRow + (rowHeights[row + (rowsMerged - 1)] - previousRow) / 2,
                                                                  cellWidth: totalColumnWidth / 5.625,
                                                                  rowHeight: totalRowHeight
                                                                  )
                                    );
                            dataValue.Add(cellText);
                        }
                    }
                    else
                    {
                        string cellText = Convert.ToString(objRange.Text).Trim();
                        double previousColumn;
                        columnWidths.TryGetValue(col - 1, out previousColumn);
                        double previousRow;
                        rowHeights.TryGetValue(row - 1, out previousRow);
                        data.Add(new Cell(row, col, cellText, x: currentColumnWidth / 2 + previousColumn,
                                        y: currentRowHeight / 2 + previousRow,
                                        cellWidth: currentColumnWidth, rowHeight: currentRowHeight));
                    }
                }

                

            }

            excelWorkBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            excelApplication.Workbooks.Close();
            excelApplication.Quit();

            //GC.GetTotalMemory(false);
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            //GC.Collect();
            //GC.GetTotalMemory(true);
        }

    }
}
