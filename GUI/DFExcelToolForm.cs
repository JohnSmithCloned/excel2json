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

namespace excel2json.GUI
{
    public partial class DFExcelToolForm : Form
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
            string savePath = Settings.Default.savePath;
            if(string.IsNullOrEmpty(savePath))
            {
                MessageBox.Show("请填写保存路径!");
                return;
            }
            //todo 导出并保存
        }

        /// <summary>
        /// excel文件拖放
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void flowLayoutPanel2_DragDrop(object sender, DragEventArgs e)
        {
            string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (dropData != null)
            {
                //this.loadExcelAsync(dropData[0]);
            }
        }
    }
}
