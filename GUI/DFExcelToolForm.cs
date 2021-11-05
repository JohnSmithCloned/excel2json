using excel2json.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CCWin;
using System.IO;
namespace excel2json.GUI
{
    public partial class DFExcelToolForm : Skin_Color
    {
        public DFExcelToolForm()
        {
            InitializeComponent();

            textBox_savePath.Text = Settings.Default.savePath;

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
                exporter.SaveToFile(exportPath, Encoding.UTF8);
            }
            Settings.Default.savePath = textBox_savePath.Text;
            Settings.Default.Save();

            MessageBox.Show($"文件数量{listBox1.Items.Count}", "导出操作完成");
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
    }
}
