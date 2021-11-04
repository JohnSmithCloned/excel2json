using System;
using System.IO;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using LitJson;
using System.Windows.Forms;

namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter
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
        public JsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, bool forceSheetName, int headerRows, string excludePrefix, bool cellJson, bool allString)
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

            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            if (!forceSheetName && validSheets.Count == 1)
            {   // single sheet

                //-- convert to object
                convertSheet(validSheets[0], exportArray, lowcase, excludePrefix, cellJson, allString);
                //-- convert to json string
                //mContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
            }
            else
            { // mutiple sheet

                Dictionary<string, object> data = new Dictionary<string, object>();
                foreach (var sheet in validSheets)
                {
                    convertSheet(sheet, exportArray, lowcase, excludePrefix, cellJson, allString);
                    //data.Add(sheet.TableName, sheetValue);
                }

                //-- convert to json string
                //mContext = JsonConvert.SerializeObject(data, jsonSettings);
            }

        }

        private void convertSheet(DataTable sheet, bool exportArray, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
            if (exportArray)
                convertSheetToArray(sheet, lowcase, excludePrefix, cellJson, allString);
            else
                convertSheetToDict(sheet, lowcase, excludePrefix, cellJson, allString);
        }

        private object convertSheetToArray(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
            List<object> values = new List<object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson, allString)
                    );
            }

            return values;
        }

        /// <summary>
        /// key=sheet名称 value=json对象
        /// </summary>
        Dictionary<string, JsonData> jsonDataByName = new Dictionary<string, JsonData>();
        /// <summary>
        /// 根据sheet名称返回json对象
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        JsonData GetJsonDataBySheetName(string sheetName)
        {
            //sheetname 裁剪
            string trimedSheetName = Regex.Replace(sheetName, @"_.*", "");
            if (jsonDataByName.ContainsKey(trimedSheetName))
                return jsonDataByName[trimedSheetName];
            else
            {
                jsonDataByName.Add(trimedSheetName, new JsonData());
                return jsonDataByName[trimedSheetName];
            }
        }
        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString)
        {
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
                DataRow row = sheet.Rows[i];
                JsonData jd = new JsonData();
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
                    }
                    //Console.WriteLine($"keyName {j} = {keyName}");
                }
                for (int m = kCount; m < sheet.Columns.Count; m++)
                {
                    string tileContent = row[m].ToString().Trim();
                    string paramName = headerDic[m];
                    string dataType = headerDataType[m];
                    if (!string.IsNullOrEmpty(tileContent) && !string.IsNullOrEmpty(paramName)
                        && !string.IsNullOrEmpty(dataType))
                    {
                        JsonData tempJd = ParseDataString(dataType, tileContent);
                        if (tempJd != null)
                            lastJD[paramName] = tempJd;
                        else
                        {
                            ;
                        }
                    }
                }
                //string ID = row[sheet.Columns[2]].ToString();

                //todo 正则判断 是否是 Key Key1 等列名
                //Console.WriteLine($"rol {i} , col[2] is {ID}");
                //var rowObject = convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson, allString);
                // 多余的字段
                // rowObject[ID] = ID;
                //importData[ID] = rowObject;
            }
            Console.WriteLine(outerJd.ToJson());
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
                        bool bData = data.ToLower().Equals("true") ? true : false;
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
                MessageBox.Show(e.ToString());
            }
            return obj;
        }
        /// <summary>
        /// 数组的字符串转json
        /// </summary>
        /// <param name="rawString"></param>
        /// <returns></returns>
        JsonData ParseStringToJsonData(string rawString)
        {
            rawString.Replace('，', ',');
            return JsonMapper.ToObject(rawString);
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row, bool lowcase, int firstDataRow, string excludePrefix, bool cellJson, bool allString)
        {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            foreach (DataColumn column in sheet.Columns)
            {
                // 过滤掉包含指定前缀的列
                string columnName = column.ToString();
                if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
                    continue;

                object value = row[column];

                // 尝试将单元格字符串转换成 Json Array 或者 Json Object
                if (cellJson)
                {
                    string cellText = value.ToString().Trim();
                    if (cellText.StartsWith("[") || cellText.StartsWith("{"))
                    {
                        try
                        {
                            object cellJsonObj = JsonConvert.DeserializeObject(cellText);
                            if (cellJsonObj != null)
                                value = cellJsonObj;
                        }
                        catch (Exception exp)
                        {
                        }
                    }
                }

                if (value.GetType() == typeof(System.DBNull))
                {
                    value = getColumnDefault(sheet, column, firstDataRow);
                }
                else if (value.GetType() == typeof(double))
                { // 去掉数值字段的“.0”
                    double num = (double)value;
                    if ((int)num == num)
                        value = (int)num;
                }

                //全部转换为string
                //方便LitJson.JsonMapper.ToObject<List<Dictionary<string, string>>>(textAsset.text)等使用方式 之后根据自己的需求进行解析
                if (allString && !(value is string))
                {
                    value = value.ToString();
                }

                string fieldName = column.ToString();
                // 表头自动转换成小写
                if (lowcase)
                    fieldName = fieldName.ToLower();

                if (string.IsNullOrEmpty(fieldName))
                    fieldName = string.Format("col_{0}", col);

                rowData[fieldName] = value;
                col++;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull))
                {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }


        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            foreach (var pair in jsonDataByName)
            {
                string sheetName = pair.Key;
                JsonData jd = pair.Value;
                //-- 保存文件
                using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    using (TextWriter writer = new StreamWriter(file, encoding))
                        writer.Write(jd.ToString());
                }
            }

        }
    }
}
