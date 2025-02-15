﻿using System;
using System.IO;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using LitJson;
using System.Windows.Forms;
using excel2json.GUI;

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
            static Dictionary<string, HashSet<string>> excelDic = new Dictionary<string, HashSet<string>>();
            public static void AddSheetName(string name)
            {
                if (excelDic.ContainsKey(fileName) == false)
                {
                    excelDic.Add(fileName, new HashSet<string>());
                }
                excelDic[fileName].Add(name);
            }
            public static string GetAllSheetNameList()
            {
                string output = "";
                foreach (var pair in excelDic)
                {
                    string fileName = Path.GetFileNameWithoutExtension(pair.Key);
                    output += $"{fileName}: ";
                    foreach (var sheetName in pair.Value)
                    {
                        output += sheetName + ", ";
                    }
                    output += "\n";
                }
                return output;
            }
        }



        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
            DebugMessage.AddSheetName(sheet.TableName);
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
                try
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
                            DebugMessage.colIndex = j;
                            if (j == 0) //第一层 
                            {
                                //有填写子key 则新建JsonData
                                KeyStrList[0] = keyName;
                                if (outerJd.ContainsKey(keyName))
                                {
                                    ShowDublicateKeyError(j, keyName);
                                    return null;
                                }
                                outerJd[keyName] = new JsonData();
                                lastJD = outerJd[keyName];
                            }
                            else if (j == 1) //第二层
                            {
                                KeyStrList[1] = keyName;
                                var fatherJd = outerJd[KeyStrList[0]];
                                if (fatherJd.ContainsKey(keyName))
                                {
                                    ShowDublicateKeyError(j, keyName);
                                    return null;
                                }
                                fatherJd[keyName] = new JsonData();
                                lastJD = fatherJd[keyName];
                            }
                            else if (j == 2) //第三层
                            {
                                KeyStrList[2] = keyName;
                                var fatherJd1 = outerJd[KeyStrList[0]];
                                var fatherJd2 = fatherJd1[KeyStrList[1]];
                                if (fatherJd2.ContainsKey(keyName))
                                {
                                    ShowDublicateKeyError(j, keyName);
                                    return null;
                                }
                                fatherJd2[keyName] = new JsonData();
                                lastJD = fatherJd2[keyName];
                            }
                            foundKeyName = true;
                        }
                        //Console.WriteLine($"keyName {j} = {keyName}");
                    }
                    bool isGlobalSheet = sheet.TableName == "GlobalConfig";
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
                                if (isGlobalSheet)//常量表特殊处理
                                {
                                    //常量表 数字 字符串 数组
                                    string fatherKey = KeyStrList[0].Trim();
                                    object obj = ParseValueString(tileContent.Trim());
                                    switch (obj)
                                    {
                                        case bool _:
                                            outerJd[fatherKey] = (bool)obj;
                                            break;
                                        case string _:
                                            outerJd[fatherKey] = (string)obj;
                                            break;
                                        case long _:
                                            outerJd[fatherKey] = (long)obj;
                                            break;
                                        case float _:
                                            outerJd[fatherKey] = (float)obj;
                                            break;
                                        case JsonData _:
                                            outerJd[fatherKey] = (JsonData)obj;
                                            break;
                                    }
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
                catch (Exception e)
                {
                    string msg = $"行号:{i}有问题";
                    MessageBox.Show($"{msg} {e.ToString()}");
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
            string lowerStr = format.ToLower();
            if (lowerStr.EndsWith("[]"))
            {
                obj = ParseStringToJsonData(data);
                return obj;
            }
            try
            {
                switch (lowerStr)
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
                    case "int32":
                    case "int64":
                        obj = new JsonData(long.Parse(data));
                        break;
                    case "float":
                        decimal number = decimal.Parse(data);
                        obj = new JsonData((double)decimal.Round(number, 4));
                        break;
                    case "array":
                        //case "int32[]":
                        //case "string[]":
                        //case "float[]":
                        obj = ParseStringToJsonData(data);
                        break;
                }
            }
            catch (Exception e)
            {
                string colName = GetExcelColumnName(DebugMessage.colIndex + 1);
                string debugMsg = $"文件名 {DebugMessage.fileName} Sheet名 {DebugMessage.sheetName}\n字段名 {DebugMessage.paramName}\n坐标 {colName}{DebugMessage.rowIndex + 2}\n数据 {data}\n数据类型是 {format}";
                var result = MessageBox.Show(debugMsg, "转表格出现问题:无效字段值", MessageBoxButtons.AbortRetryIgnore);
                if (result == DialogResult.Abort)
                {
                    DFExcelToolForm.ActiveForm.Close();
                }

            }
            return obj;
        }
        void ShowDublicateKeyError(int colIdx, string keyName)
        {
            string colName = GetExcelColumnName(DebugMessage.colIndex + 1);
            string debugMsg = $"文件名 {DebugMessage.fileName} Sheet名 {DebugMessage.sheetName}\n 坐标{colName}{DebugMessage.rowIndex + 2} \n重复key值{keyName}";
            var result = MessageBox.Show(debugMsg, "转表格出现问题:重复Key值", MessageBoxButtons.AbortRetryIgnore);

            if (result == DialogResult.Abort)
            {
                DFExcelToolForm.ActiveForm.Close();
            }
        }

        /// <summary>
        /// 常量表的参数转对象 支持 整数 浮点数 bool 字符串 数组
        /// </summary>
        /// <param name="format"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        object ParseValueString(string data)
        {
            long number = 0;
            if (long.TryParse(data, out number))
                return number;
            float number2 = 0;
            if (float.TryParse(data, out number2))
                return number2;
            if (data.StartsWith("\""))
            {
                return data.Substring(1, data.Length - 2);
            }
            if (data.StartsWith("["))
            {
                return ParseStringToJsonData(data);
            }
            string lower = data.ToLower();
            if (lower.Equals("true") || lower.Equals("false"))
                return lower.Equals("true");
            return data;
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
            var match = Regex.Match(rawString, @"[A-Za-z]");
            if (match.Success)//包含字母 则自动加单引号
            {
                rawString = ArrayStringAddQuote(rawString);
            }

            if (!rawString.StartsWith("["))
                rawString = "[" + rawString;
            if (!rawString.EndsWith("]"))
                rawString += "]";

            return JsonMapper.ToObject(rawString);
        }
        string ArrayStringAddQuote(string rawString)
        {
            string output = "";
            string[] strArray = rawString.Split(',');
            foreach (string str in strArray)
            {
                string str1 = str.Trim();
                var match = Regex.Match(str1, "\"(.*)\"");
                var match2 = Regex.Match(str1, "'(.*)'");
                if (!match.Success && !match2.Success)
                {
                    str1 = "'" + str1 + "'";
                }
                output += str1 + ',';
            }
            return output.Substring(0, output.Length - 1);
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
                        PrettyPrint = true
                        //IndentValue = 0
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
