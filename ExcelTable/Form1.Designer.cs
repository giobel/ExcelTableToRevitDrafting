namespace ExcelTable
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBoxFilePath = new System.Windows.Forms.TextBox();
            this.buttonOpen = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.tbWidthFactor = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbHeightFactor = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBoxFilePath
            // 
            this.textBoxFilePath.Location = new System.Drawing.Point(13, 15);
            this.textBoxFilePath.Name = "textBoxFilePath";
            this.textBoxFilePath.Size = new System.Drawing.Size(199, 20);
            this.textBoxFilePath.TabIndex = 0;
            this.textBoxFilePath.TextChanged += new System.EventHandler(this.TextBoxFilePath_TextChanged);
            // 
            // buttonOpen
            // 
            this.buttonOpen.Location = new System.Drawing.Point(226, 15);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(75, 20);
            this.buttonOpen.TabIndex = 1;
            this.buttonOpen.Text = "Browse";
            this.buttonOpen.UseVisualStyleBackColor = true;
            this.buttonOpen.Click += new System.EventHandler(this.ButtonOpen_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point(101, 84);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(111, 28);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.Button1_Click);
            // 
            // tbWidthFactor
            // 
            this.tbWidthFactor.Location = new System.Drawing.Point(13, 58);
            this.tbWidthFactor.Name = "tbWidthFactor";
            this.tbWidthFactor.Size = new System.Drawing.Size(60, 20);
            this.tbWidthFactor.TabIndex = 3;
            this.tbWidthFactor.Text = "2";
            this.tbWidthFactor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbWidthFactor.TextChanged += new System.EventHandler(this.TbWidthFactor_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Width factor";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(124, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Height factor";
            // 
            // tbHeightFactor
            // 
            this.tbHeightFactor.Location = new System.Drawing.Point(127, 58);
            this.tbHeightFactor.Name = "tbHeightFactor";
            this.tbHeightFactor.Size = new System.Drawing.Size(60, 20);
            this.tbHeightFactor.TabIndex = 5;
            this.tbHeightFactor.Text = "0.3";
            this.tbHeightFactor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbHeightFactor.TextChanged += new System.EventHandler(this.TbHeightFactor_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(313, 121);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbHeightFactor);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbWidthFactor);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonOpen);
            this.Controls.Add(this.textBoxFilePath);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Table from Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBoxFilePath;
        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.TextBox tbWidthFactor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbHeightFactor;
    }
}