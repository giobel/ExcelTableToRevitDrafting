using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTable
{
    class Cell
    {
        private int _row { get; set; }
        private int _column { get; set; }
        public string _value { get; private set; }
        private double _x { get; set; }
        private double _y { get; set; }
        private double _cellWidth { get; set; }
        private double _rowHeight { get; set; }
        public double X
        {
            get { return _x; }
            set { _x = value; }
        }

        public double Y
        {
            get { return _y; }
            set { _y = value; }
        }

        public double CellWidth
        {
            get { return _cellWidth; }
            private set { _cellWidth = value; }
        }
        public double RowHeight
        {
            get { return _rowHeight; }
            private set { _rowHeight = value; }
        }

        public Cell(int row, int column, string value, double x, double y, double cellWidth, double rowHeight)
        {
            _row = row;
            _column = column;
            _value = value;
            _x = x;
            _y = y;
            _cellWidth = cellWidth;
            _rowHeight = rowHeight;
        }

        public override string ToString()
        {
            return $"row:{_row} column:{_column} text:{_value} x:{_x} y:{_y} cellWidth:{_cellWidth} rowHeight:{_rowHeight}";
        }


    }
}

