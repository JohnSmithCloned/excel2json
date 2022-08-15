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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DFExcelToolForm));
            this.textBox_savePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button_saveToFile = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button_clearList = new System.Windows.Forms.Button();
            this.btn_export_protobuf = new System.Windows.Forms.Button();
            this.textBox_protoPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_csProtoPath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox_compiler = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_compiler = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox_savePath
            // 
            this.textBox_savePath.Location = new System.Drawing.Point(26, 71);
            this.textBox_savePath.Name = "textBox_savePath";
            this.textBox_savePath.Size = new System.Drawing.Size(290, 20);
            this.textBox_savePath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(147, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "数据导出路径 [Data文件夹]";
            // 
            // button_saveToFile
            // 
            this.button_saveToFile.Location = new System.Drawing.Point(22, 442);
            this.button_saveToFile.Name = "button_saveToFile";
            this.button_saveToFile.Size = new System.Drawing.Size(124, 40);
            this.button_saveToFile.TabIndex = 2;
            this.button_saveToFile.Text = "导出JSON";
            this.button_saveToFile.UseVisualStyleBackColor = true;
            this.button_saveToFile.Click += new System.EventHandler(this.button_saveToFile_Click);
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(22, 309);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(347, 108);
            this.listBox1.TabIndex = 3;
            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBox1_DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBox1_DragEnter);
            this.listBox1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.listBox1_KeyUp);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 281);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Excel文件拖进来";
            // 
            // button_clearList
            // 
            this.button_clearList.Location = new System.Drawing.Point(149, 273);
            this.button_clearList.Name = "button_clearList";
            this.button_clearList.Size = new System.Drawing.Size(135, 27);
            this.button_clearList.TabIndex = 5;
            this.button_clearList.Text = "清空文件列表";
            this.button_clearList.UseVisualStyleBackColor = true;
            this.button_clearList.Click += new System.EventHandler(this.button_clearList_Click);
            // 
            // btn_export_protobuf
            // 
            this.btn_export_protobuf.Location = new System.Drawing.Point(170, 442);
            this.btn_export_protobuf.Name = "btn_export_protobuf";
            this.btn_export_protobuf.Size = new System.Drawing.Size(114, 40);
            this.btn_export_protobuf.TabIndex = 7;
            this.btn_export_protobuf.Text = "导出Protobuf";
            this.btn_export_protobuf.UseVisualStyleBackColor = true;
            this.btn_export_protobuf.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox_protoPath
            // 
            this.textBox_protoPath.Location = new System.Drawing.Point(30, 128);
            this.textBox_protoPath.Name = "textBox_protoPath";
            this.textBox_protoPath.Size = new System.Drawing.Size(290, 20);
            this.textBox_protoPath.TabIndex = 8;
            this.textBox_protoPath.Text = "D:\\ProtoPath";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "proto文件保存路径";
            // 
            // textBox_csProtoPath
            // 
            this.textBox_csProtoPath.Location = new System.Drawing.Point(30, 181);
            this.textBox_csProtoPath.Name = "textBox_csProtoPath";
            this.textBox_csProtoPath.Size = new System.Drawing.Size(290, 20);
            this.textBox_csProtoPath.TabIndex = 10;
            this.textBox_csProtoPath.Text = "D:\\ProtoPath";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 165);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(159, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Bridge目录 \\Config\\ConfigProto";
            // 
            // textBox_compiler
            // 
            this.textBox_compiler.Location = new System.Drawing.Point(30, 247);
            this.textBox_compiler.Name = "textBox_compiler";
            this.textBox_compiler.Size = new System.Drawing.Size(141, 20);
            this.textBox_compiler.TabIndex = 13;
            this.textBox_compiler.Text = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe";
            this.textBox_compiler.TextChanged += new System.EventHandler(this.textBox_compiler_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(27, 231);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = ".Net编译器路径";
            // 
            // btn_compiler
            // 
            this.btn_compiler.Location = new System.Drawing.Point(177, 244);
            this.btn_compiler.Name = "btn_compiler";
            this.btn_compiler.Size = new System.Drawing.Size(75, 23);
            this.btn_compiler.TabIndex = 15;
            this.btn_compiler.Text = "选择";
            this.btn_compiler.UseVisualStyleBackColor = true;
            this.btn_compiler.Click += new System.EventHandler(this.btn_compiler_Click);
            // 
            // DFExcelToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(401, 560);
            this.Controls.Add(this.btn_compiler);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox_compiler);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox_csProtoPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox_protoPath);
            this.Controls.Add(this.btn_export_protobuf);
            this.Controls.Add(this.button_clearList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button_saveToFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_savePath);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DFExcelToolForm";
            this.Text = "转表神器-大禹工作室";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.DFExcelToolForm_FormClosed);
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
        private System.Windows.Forms.Button btn_export_protobuf;
        private System.Windows.Forms.TextBox textBox_protoPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_csProtoPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_compiler;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btn_compiler;
    }
}