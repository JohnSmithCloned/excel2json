using CCWin;
using excel2json.Properties;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace excel2json.GUI
{
    public partial class DFExcelToolForm : Skin_Color
    {
        public DFExcelToolForm()
        {
            InitializeComponent();

            this.KeyPreview = true;
            textBox_savePath.Text = Settings.Default.savePath;

            listBox1.Items.Clear();
            if (Settings.Default.lastFileList != null)
                foreach (var str in Settings.Default.lastFileList)
                {
                    listBox1.Items.Add(str);
                }

            if (!string.IsNullOrEmpty(Settings.Default.Compiler_Path))
                textBox_compiler.Text = Settings.Default.Compiler_Path;

            var st = Settings.Default;

            textBox_csProtoPath.Text = st.configProtoPath;
            textBox_protoPath.Text = st.protoPath;

            btn_export_protobuf.Enabled = false;//先不要用protobuf转表了 我想直接用鲁班
        }
        /// <summary>
        /// 点击导出并保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_saveToFile_Click(object sender, EventArgs e)
        {
            string savePath = textBox_savePath.Text;
            if (string.IsNullOrEmpty(savePath))
            {
                MessageBox.Show("请填写保存路径!");
                return;
            }
            //导出并保存
            foreach (string path in listBox1.Items)
            {
                //-- Load Excel
                ExcelLoader excel = new ExcelLoader(path, 3);
                DFJsonExporter.DebugMessage.fileName = Path.GetFileName(path);
                //一个excel可能导出多个文件额
                DFJsonExporter exporter = new DFJsonExporter(excel,
                    false, false, "yyyy/MM/dd", false, 3, "", false, false);
                //-- Export path
                string exportPath;
                exportPath = textBox_savePath.Text;
                exporter.SaveToFile(exportPath, new UTF8Encoding(false));
            }
            Settings.Default.savePath = textBox_savePath.Text;

            MessageBox.Show($"文件数量{listBox1.Items.Count}", "导出操作完成");
            //string output = DFJsonExporter.DebugMessage.GetAllSheetNameList();
            //MessageBox.Show($"report:{output}", "导出操作完成");
        }

        /// <summary>
        /// excel文件拖放
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {

            string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (dropData != null)
            {
                //this.loadExcelAsync(dropData[0]);
                int i;
                for (i = 0; i < dropData.Length; i++)
                {
                    string path = dropData[i];
                    if (path.EndsWith(".xlsx"))
                        listBox1.Items.Add(dropData[i]);
                }
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void button_clearList_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void listBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
        }

        private void DFExcelToolForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //界面关闭时

            Settings.Default.lastFileList = new System.Collections.Specialized.StringCollection();
            foreach (var item in listBox1.Items)
                Settings.Default.lastFileList.Add(item.ToString());

            var st = Settings.Default;
            st.Compiler_Path = textBox_compiler.Text;
            st.configProtoPath = textBox_csProtoPath.Text;
            st.protoPath = textBox_protoPath.Text;
            st.savePath = textBox_savePath.Text;
            st.Save();

        }

        /// <summary>
        /// 全套流程 excel to protobuf
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnExportProtobuf(object sender, EventArgs e)
        {
            string savePath = textBox_savePath.Text;
            if (string.IsNullOrEmpty(savePath))
            {
                MessageBox.Show("请填写保存路径!");
                return;
            }
            //导出并保存
            foreach (string path in listBox1.Items)
            {
                //-- Load Excel
                ExcelLoader excel = new ExcelLoader(path, 3);

                //一个excel可能导出多个文件额
                string protoPath = this.textBox_protoPath.Text;
                string datPath = this.textBox_savePath.Text;
                string csProtoPath = this.textBox_csProtoPath.Text;
                string excelFileName = Path.GetFileNameWithoutExtension(path);
                DFExcelReader exporter = new DFExcelReader(excel, protoPath, datPath, csProtoPath, excelFileName);
            }
            Settings.Default.savePath = textBox_savePath.Text;

            MessageBox.Show($"文件数量{listBox1.Items.Count}", "导出操作完成");
        }

        private void textBox_compiler_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.Compiler_Path = textBox_compiler.Text;
        }

        private void btn_compiler_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "csc(*.exe)|*.exe";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox_compiler.Text = dialog.FileName;
                Settings.Default.Compiler_Path = textBox_compiler.Text;
            }
        }

        private void btnOpenData_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(textBox_savePath.Text))
                Process.Start("explorer.exe", textBox_savePath.Text);
        }

        private void btnOpenProto_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(textBox_protoPath.Text))
                Process.Start("explorer.exe", textBox_protoPath.Text);
        }

        private void btnOpenConfig_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(textBox_csProtoPath.Text))
                Process.Start("explorer.exe", textBox_csProtoPath.Text);
        }
    }
}
