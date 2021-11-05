using System;
using System.IO;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using LitJson;
using System.Windows.Forms;
namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class DFJsonExporter
    {
        string mContext = "";
        int mHeaderRows = 0;

        public string context
        {
            get
            {
                return mContext;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>
        public DFJsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, bool forceSheetName, int headerRows, string excludePrefix, bool cellJson, bool allString)
        {

            mHeaderRows = headerRows - 1;
            List<DataTable> validSheets = new List<DataTable>();
            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];

                // 过滤掉包含特定前缀的表单
                string sheetName = sheet.TableName;
                //名称 大小写字母开始  找到不符合的名称 则停止循环
                if (!Regex.Match(sheetName, "^[a-zA-z]").Success)
                    break;
                if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                    continue;

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheet);
            }

            if (!forceSheetName && validSheets.Count == 1)
            {   // single sheet
                //-- convert to object
                ConvertSheet(validSheets[0], exportArray, lowcase, excludePrefix, cellJson, allString);
            }
            else
            { // mutiple sheet
                foreach (var sheet in validSheets)
                {
                    ConvertSheet(sheet, exportArray, lowcase, excludePrefix, cellJson, allString);
                }
            }

        }

        private void ConvertSheet(DataTable sheet, bool exportArray, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
            convertSheetToDict(sheet, lowcase, excludePrefix, cellJson, allString);
        }

        /// <summary>
        /// key=sheet名称 value=json对象
        /// </summary>
        Dictionary<string, JsonData> jsonObjDict = new Dictionary<string, JsonData>();
        /// <summary>
        /// 根据sheet名称返回json对象
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        JsonData GetJsonDataBySheetName(string sheetName)
        {
            //sheetname 裁剪
            string trimedSheetName = Regex.Replace(sheetName, @"_.*", "");
            if (jsonObjDict.ContainsKey(trimedSheetName))
                return jsonObjDict[trimedSheetName];
            else
            {
                jsonObjDict.Add(trimedSheetName, new JsonData());
                return jsonObjDict[trimedSheetName];
            }
        }
        public static class DebugMessage
        {
            public static string sheetName;
            public static string fileName;
            public static int rowIndex;
            public static int colIndex;
            /// <summary>
            /// 字段名称
            /// </summary>
            public static string paramName;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
            DebugMessage.sheetName = sheet.TableName;
            Dictionary<string, object> importData =
                new Dictionary<string, object>();
            //key=第几列 value=字段数据类型 Integer
            Dictionary<int, string> headerDataType = new Dictionary<int, string>();
            //key=第几列 value=表头字段名 ID
            Dictionary<int, string> headerDic = new Dictionary<int, string>();
            int firstDataRow = mHeaderRows;
            DataRow dataTypeRow = sheet.Rows[0];//数据类型这一行

            DataRow sheetHead = sheet.Rows[1];//字段名这一行key key1 这一行
            //Key,Key1,Key2... 的列表
            List<string> KeyList = new List<string>();
            for (int j = 0; j < sheet.Columns.Count; j++)
            {
                string dataTypeName = dataTypeRow[sheet.Columns[j]].ToString();
                headerDataType[j] = dataTypeName.Trim();

                string paramName = sheetHead[sheet.Columns[j]].ToString();
                //Console.WriteLine($" 表头是 { paramName}");
                headerDic[j] = paramName.Trim();
                if (paramName.StartsWith("Key"))
                {
                    KeyList.Add(paramName);
                }
            }
            int kCount = KeyList.Count;
            //本sheet数据存储到这里
            JsonData outerJd = GetJsonDataBySheetName(sheet.TableName);

            //key=key层级:0,1,2  value=这个层级对应的最近使用的一个key
            Dictionary<int, string> KeyStrList = new Dictionary<int, string>();
            JsonData lastJD = null;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)//逐行读取数据
            {
                DebugMessage.rowIndex = i;
                DataRow row = sheet.Rows[i];
                bool foundKeyName = false;
                for (int j = 0; j < kCount; j++)
                {
                    string keyName = row[j].ToString();
                    bool keyNameFilled = !string.IsNullOrEmpty(keyName);
                    if (keyNameFilled)
                    {
                        if (j == 0) //第一层 
                        {
                            //有填写子key 则新建JsonData
                            KeyStrList[0] = keyName;
                            outerJd[keyName] = new JsonData();
                            lastJD = outerJd[keyName];
                        }
                        else if (j == 1) //第二层
                        {
                            KeyStrList[1] = keyName;
                            var fatherJd = outerJd[KeyStrList[0]];
                            fatherJd[keyName] = new JsonData();
                            lastJD = fatherJd[keyName];
                        }
                        else if (j == 2) //第三层
                        {
                            KeyStrList[2] = keyName;
                            var fatherJd1 = outerJd[KeyStrList[0]];
                            var fatherJd2 = fatherJd1[KeyStrList[1]];
                            fatherJd2[keyName] = new JsonData();
                            lastJD = fatherJd2[keyName];
                        }
                        foundKeyName = true;
                    }
                    //Console.WriteLine($"keyName {j} = {keyName}");
                }
                if (foundKeyName)
                    for (int m = kCount; m < sheet.Columns.Count; m++)
                    {
                        DebugMessage.colIndex = m;
                        string tileContent = row[m].ToString().Trim();
                        string paramName = headerDic[m]; //字段名 ItemId
                        string dataType = headerDataType[m]; //数据类型 String
                        DebugMessage.paramName = paramName;
                        if (!string.IsNullOrEmpty(tileContent) && !string.IsNullOrEmpty(paramName)
                            && !string.IsNullOrEmpty(dataType))
                        {
                            if (paramName == "Value")
                            {
                                string fatherKey = KeyStrList[0];
                                outerJd[fatherKey] = tileContent;
                            }
                            else
                            {
                                JsonData tempJd = ParseDataString(dataType, tileContent);
                                if (tempJd != null)
                                    lastJD[paramName] = tempJd;
                            }
                        }
                    }
            }
            //Console.WriteLine(outerJd.ToJson());
            return importData;
        }
        /// <summary>
        /// 给定数据类型字符串 和 数据字符串 返回Parse后的数据对象
        /// </summary>
        /// <param name="format"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        JsonData ParseDataString(string format, string data)
        {
            JsonData obj = null;
            try
            {
                switch (format.ToLower())
                {
                    case "string":
                        obj = new JsonData(data);
                        break;
                    case "bool":
                    case "boolean":
                        bool bData = data.ToLower().Equals("true");
                        obj = new JsonData(bData);
                        break;
                    case "integer":
                    case "int":
                        obj = new JsonData(int.Parse(data));
                        break;
                    case "float":
                        obj = new JsonData(float.Parse(data));
                        break;
                    case "array":
                        obj = ParseStringToJsonData(data);
                        break;
                }
            }
            catch (Exception e)
            {
                string colName = GetExcelColumnName(DebugMessage.colIndex + 1);
                string debugMsg = $"文件名 {DebugMessage.fileName} Sheet名 {DebugMessage.sheetName}\n字段名 {DebugMessage.paramName}\n坐标 {colName}{DebugMessage.rowIndex + 2}\n数据 {data}\n数据类型是 {format}";
                MessageBox.Show(debugMsg, "转表格出现问题");
            }
            return obj;
        }
        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// 数组的字符串转json
        /// </summary>
        /// <param name="rawString"></param>
        /// <returns></returns>
        JsonData ParseStringToJsonData(string rawString)
        {
            rawString.Replace('，', ',');
            if (!rawString.StartsWith("["))
                rawString = "[" + rawString;
            if (!rawString.EndsWith("]"))
                rawString += "]";
            return JsonMapper.ToObject(rawString);
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            foreach (var pair in jsonObjDict)
            {
                string sheetName = pair.Key;
                string savePath = Path.Combine(filePath, sheetName);
                savePath += ".json";
                JsonData jd = pair.Value;
                if (jd.Keys == null && jd.Keys.Count == 0)
                    continue;
                //-- 保存文件
                using (FileStream file = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                {
                    StringBuilder sb = new StringBuilder();
                    JsonWriter jw = new JsonWriter(sb)
                    {
                        PrettyPrint = true,
                        IndentValue = 4
                    };
                    JsonMapper.ToJson(jd, jw);
                    string jdString = Regex.Unescape(sb.ToString());
                    using (var writer = new StreamWriter(file, encoding))
                    {
                        writer.Write(jdString);
                    }
                }
            }

        }
    }
}
