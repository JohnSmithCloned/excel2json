using System;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;
using excel2json.Properties;

namespace excel2json
{
    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program
    {
        /// <summary>
        /// 应用程序入口
        /// </summary>
        /// <param name="args">命令行参数</param>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length <= 0)
            {
                //-- GUI MODE ----------------------------------------------------------
                Console.WriteLine("Launch excel2json GUI Mode...");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                //Application.Run(new GUI.MainForm());
                Application.Run(new GUI.DFExcelToolForm());
            }
            else
            {
                //-- COMMAND LINE MODE -------------------------------------------------
                Console.WriteLine(" cmd run excel to json");
                string defineExportPath="";
                string epHEAD = "export_path=";
                foreach (string arg in args)
                {
                    try
                    {
                        if(arg.StartsWith(epHEAD))
                        {
                            defineExportPath = arg.Substring(epHEAD.Length);
                            continue;
                        }

                        DateTime startTime = DateTime.Now;

                        //-- 程序计时
                        DateTime endTime = DateTime.Now;
                        TimeSpan dur = endTime - startTime;
                        //-- Load Excel
                        string path = arg;
                        ExcelLoader excel = new ExcelLoader(path, 3);
                        DFJsonExporter.DebugMessage.fileName = Path.GetFileName(path);
                        //一个excel可能导出多个文件额
                        DFJsonExporter exporter = new DFJsonExporter(excel,
                            false, false, "yyyy/MM/dd", false, 3, "", false, false);
                        //-- Export path
                        string exportPath;
                        if(!string.IsNullOrEmpty(defineExportPath))
                            exportPath = defineExportPath;
                        else
                            exportPath = Settings.Default.savePath;
                        exporter.SaveToFile(exportPath, new UTF8Encoding(false));

                        Console.WriteLine(
                            string.Format("[{0}]：\tConversion complete in [{1}ms].",
                            Path.GetFileName(path),
                            dur.TotalMilliseconds)
                            );
                    }
                    catch (Exception exp)
                    {
                        Console.WriteLine("Error: " + exp.Message);
                    }

                }
                Console.ReadKey();

            }// end of else
        }

    }
}
