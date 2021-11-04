namespace excel2json.GUI
{
    partial class DFExcelToolForm
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
            this.textBox_savePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_saveToFile = new System.Windows.Forms.Button();
            this.pictureBoxExcel = new System.Windows.Forms.PictureBox();
            this.labelExcelFile = new System.Windows.Forms.Label();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExcel)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox_savePath
            // 
            this.textBox_savePath.Location = new System.Drawing.Point(34, 52);
            this.textBox_savePath.Name = "textBox_savePath";
            this.textBox_savePath.Size = new System.Drawing.Size(290, 20);
            this.textBox_savePath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "导出路径[ Data文件夹]";
            // 
            // button_saveToFile
            // 
            this.button_saveToFile.Location = new System.Drawing.Point(34, 247);
            this.button_saveToFile.Name = "button_saveToFile";
            this.button_saveToFile.Size = new System.Drawing.Size(149, 50);
            this.button_saveToFile.TabIndex = 2;
            this.button_saveToFile.Text = "导出并保存json";
            this.button_saveToFile.UseVisualStyleBackColor = true;
            this.button_saveToFile.Click += new System.EventHandler(this.button_saveToFile_Click);
            // 
            // pictureBoxExcel
            // 
            this.pictureBoxExcel.Image = global::excel2json.Properties.Resources.excel_64;
            this.pictureBoxExcel.Location = new System.Drawing.Point(62, 92);
            this.pictureBoxExcel.Name = "pictureBoxExcel";
            this.pictureBoxExcel.Size = new System.Drawing.Size(262, 94);
            this.pictureBoxExcel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBoxExcel.TabIndex = 3;
            this.pictureBoxExcel.TabStop = false;
            // 
            // labelExcelFile
            // 
            this.labelExcelFile.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelExcelFile.Location = new System.Drawing.Point(62, 189);
            this.labelExcelFile.Name = "labelExcelFile";
            this.labelExcelFile.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.labelExcelFile.Size = new System.Drawing.Size(260, 38);
            this.labelExcelFile.TabIndex = 4;
            this.labelExcelFile.Text = "xlsx文件拖到这里";
            this.labelExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(58, 89);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(268, 140);
            this.flowLayoutPanel2.TabIndex = 5;
            this.flowLayoutPanel2.DragDrop += new System.Windows.Forms.DragEventHandler(this.flowLayoutPanel2_DragDrop);
            // 
            // DFExcelToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 352);
            this.Controls.Add(this.labelExcelFile);
            this.Controls.Add(this.pictureBoxExcel);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.button_saveToFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_savePath);
            this.Name = "DFExcelToolForm";
            this.Text = "DFExcelToolForm";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExcel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_savePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_saveToFile;
        private System.Windows.Forms.PictureBox pictureBoxExcel;
        private System.Windows.Forms.Label labelExcelFile;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
    }
}