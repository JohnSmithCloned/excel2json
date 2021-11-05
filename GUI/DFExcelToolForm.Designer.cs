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
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button_clearList = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_savePath
            // 
            this.textBox_savePath.Location = new System.Drawing.Point(34, 66);
            this.textBox_savePath.Name = "textBox_savePath";
            this.textBox_savePath.Size = new System.Drawing.Size(290, 20);
            this.textBox_savePath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "导出路径[ Data文件夹]";
            // 
            // button_saveToFile
            // 
            this.button_saveToFile.Location = new System.Drawing.Point(34, 279);
            this.button_saveToFile.Name = "button_saveToFile";
            this.button_saveToFile.Size = new System.Drawing.Size(149, 50);
            this.button_saveToFile.TabIndex = 2;
            this.button_saveToFile.Text = "导出并保存json";
            this.button_saveToFile.UseVisualStyleBackColor = true;
            this.button_saveToFile.Click += new System.EventHandler(this.button_saveToFile_Click);
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(34, 126);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(347, 108);
            this.listBox1.TabIndex = 3;
            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBox1_DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBox1_DragEnter);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Excel文件拖进来";
            // 
            // button_clearList
            // 
            this.button_clearList.Location = new System.Drawing.Point(156, 93);
            this.button_clearList.Name = "button_clearList";
            this.button_clearList.Size = new System.Drawing.Size(135, 27);
            this.button_clearList.TabIndex = 5;
            this.button_clearList.Text = "清空文件列表";
            this.button_clearList.UseVisualStyleBackColor = true;
            this.button_clearList.Click += new System.EventHandler(this.button_clearList_Click);
            // 
            // DFExcelToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 352);
            this.Controls.Add(this.button_clearList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button_saveToFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_savePath);
            this.Name = "DFExcelToolForm";
            this.Text = "DFExcelToolForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_savePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_saveToFile;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_clearList;
    }
}