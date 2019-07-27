using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTable
{
    public partial class Form1 : Form
    {
        public string filePath { get; private set; }
        public double widthFactor { get; private set; }
        public double heightFactor { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void ButtonOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\Users\%username%\Documents",
                Title = "Browse Excel Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "excel files (*.xls)|*.xls",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxFilePath.Text = openFileDialog1.FileName;
                
            }

            filePath = textBoxFilePath.Text;

        }

        private void TextBoxFilePath_TextChanged(object sender, EventArgs e)
        {
            filePath = textBoxFilePath.Text;
        }

        private void TbWidthFactor_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void TbHeightFactor_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            widthFactor = Convert.ToDouble(tbWidthFactor.Text);
            heightFactor = Convert.ToDouble(tbHeightFactor.Text);
        }
    }
}
