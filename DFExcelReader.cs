using excel2json.Properties;
using Google.Protobuf;
using Google.Protobuf.Collections;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace excel2json
{
    /// <summary>
    /// Excel转proto
    /// </summary>
    class DFExcelReader
    {
        public const string NameSpace = "ConfigData";
        string debugInfo;
        /// <summary>
        /// 读取DLL
        /// </summary>
        private Assembly _assembly;
        /// <summary>
        /// key=sheet名 value=数据对象
        /// </summary>
        private Dictionary<string, object> _excelInsDic;
        /// <summary>
        /// Excel文件里面的Sheet
        /// </summary>
        List<DataTable> validSheets;
        /*
         string	double	float	int32	int64	uint32	uint64	bool	string	double[]	float[]	int32[]	int64[]	uint32[]	uint64[]	bool[]	string[]
*/
        /// <summary>
        /// 旧版变量类型转换
        /// </summary>
        Dictionary<string, string> oldTypeDic = new Dictionary<string, string>()
        {
            {"boolean","bool" },
            {"integer","int32" },
            {"int" ,"int32" },
            {"long" ,"int64" },
            {"array" ,"int32[]" },
        };
        HashSet<string> validTypeNames = new HashSet<string>()
        {
            "string","double","float","int32","int64","uint32","uint64",
            "sint32","sint64","fixed32","fixed64","sfixed32","sfixed64",
            "bool","string"
        };
        /// <summary>
        /// 变量类型名=>proto变量类型 example:int32[]=>repeated int32
        /// </summary>
        /// <param name="oldName"></param>
        /// <returns></returns>
        string ConvertTypeName2Proto(string oldName)
        {
            oldName = oldName.Trim().ToLower();
            if (oldTypeDic.ContainsKey(oldName))//旧版数据类型转换
                oldName = oldTypeDic[oldName];
            string newName = string.Empty;
            string prefix = string.Empty;
            if (oldName.EndsWith("[]"))//数组改成 repeated
            {
                prefix = "repeated ";
                //去掉[]保存
                newName = oldName.Substring(0, oldName.Length - 2);
            }
            else
                newName = oldName;
            if (validTypeNames.Contains(newName))
                return prefix + newName;
            else
            {
                //MessageBox.Show($"变量名不合法 {saveOldName} 如果是Array请改成 int32[] 或者 string[]");
                //throw new Exception();
                return null;
            }
        }

        /// <summary>
        /// 变量类型名=>C#变量类型
        /// </summary>
        /// <param name="oldName"></param>
        /// <returns></returns>
        string ConvertTypeName2CS(string oldName)
        {
            if (string.IsNullOrEmpty(oldName)) return null;
            string saveOldName = oldName;
            oldName = oldName.Trim().ToLower();
            string checkName = oldName;
            if (oldTypeDic.ContainsKey(checkName))//旧版数据类型转换
                checkName = oldTypeDic[checkName];
            string tail = "";
            if (checkName.EndsWith("[]"))
            {
                checkName = checkName.Substring(0, checkName.Length - 2);
                tail = "[]";
            }
            if (!validTypeNames.Contains(checkName))
            {
                debugInfo += $"不支持的变量类型 {saveOldName}\n";
                return null;
            }
            return checkName + tail;
        }
        /// <summary>
        /// Sheet名 去掉_后面的
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        string TrimSheetName(string sheetName)
        {
            return Regex.Replace(sheetName, @"_.*", "");
        }
        //key = sheet名  value = {key=变量名 value=数据类型}
        Dictionary<string, Dictionary<string, string>> headerDic;
        public DFExcelReader(ExcelLoader excel, string protoPath, string _datPath, string _csProtoPath, string _excelName)
        {
            this.datPath = _datPath;
            this.dfProtoPath = _csProtoPath;
            this.excelFileName = _excelName;

            //1个excel文件里面有多个sheet 同名合并
            //有效的sheet
            validSheets = new List<DataTable>();

            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];

                // 过滤掉包含特定前缀的表单
                string sheetName = sheet.TableName;
                //非大小写字母开头的sheet 忽略
                if (!Regex.Match(sheetName, "^[a-zA-z]").Success)
                    break;
                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheet);
            }
            if (validSheets.Count > 0)
            {
                string subDirName = TrimSheetName(validSheets[0].TableName);
                //配置表名是中文 只好改成找到第一个sheet名作为子目录名称
                this.protoPath = Path.Combine(protoPath, subDirName);
                if (!Directory.Exists(this.protoPath))//建立子目录
                {
                    Directory.CreateDirectory(this.protoPath);
                }
                CollectVariableNames();
                foreach (var sheetPair in headerDic)
                {
                    //sheet转对应的 *.proto
                    GenerateProtoFile(sheetPair.Key);
                }
                //调用protoc.exe proto文本 生成 c#proto代码文件 
                CMDGenerateCSProto();
                //所有proto.cs文件 生成一个dll
                CMDGenerateDLL();
                //读取dll 反射构造数据
                ReadDllToAssembly();
                //Excel表格读取 保存到构造数据里面
                GenerateProtoObj();
                //构造数据序列化保存到硬盘
                SerializeDatFile();
                if (!string.IsNullOrEmpty(debugInfo))
                    MessageBox.Show(debugInfo);
            }
        }
        /// <summary>
        /// 收集变量名
        /// </summary>
        void CollectVariableNames()
        {
            if (validSheets.Count > 0)
            {
                headerDic = new Dictionary<string, Dictionary<string, string>>();
                foreach (DataTable sheet in validSheets)
                {
                    string sheetName = TrimSheetName(sheet.TableName);
                    if (!headerDic.ContainsKey(sheetName))
                        headerDic[sheetName] = new Dictionary<string, string>();
                    var innerDic = headerDic[sheetName];
                    DataRow dataTypeRow = sheet.Rows[0];//数据类型这一行
                    DataRow dataNameRow = sheet.Rows[1];//变量名这一行
                    for (int j = 0; j < sheet.Columns.Count; j++)
                    {
                        DataColumn col = sheet.Columns[j];
                        //变量类型 去空格 变小写
                        string dataTypeName = ConvertTypeName2Proto(dataTypeRow[col].ToString());
                        //变量名去空格
                        string variableName = dataNameRow[col].ToString().Trim();
                        bool isEmptyName = string.IsNullOrEmpty(variableName);
                        bool isKeyName = Regex.Match(variableName, @"^Key[0-9]?").Success;
                        bool isEmptyType = string.IsNullOrEmpty(dataTypeName);
                        //if (isEmptyName && isEmptyType)
                        //{
                        //    //变量名和类型都没有写 则后面的内容跳过
                        //    break;
                        //}
                        //debugInfo += $"类型{dataTypeName} 变量名{variableName} \n";
                        if (isKeyName)
                        {
                            string _type = isEmptyType ? "int32" : dataTypeName;
                            //Key列 视为int32
                            if (!innerDic.ContainsKey(variableName))
                                innerDic.Add(variableName, _type);
                        }
                        if (!isKeyName && !isEmptyType)//普通的变量
                        {
                            if (!innerDic.ContainsKey(variableName))
                                innerDic.Add(variableName, dataTypeName);
                        }
                    }
                }
            }


        }
        #region 生成proto文件
        /// <summary>
        /// proto文件名
        /// </summary>
        string protoFilePath;
        /// <summary>
        /// proto文件保存路径
        /// </summary>
        string protoPath;
        /// <summary>
        /// DAT文件保存路径
        /// </summary>
        string datPath;
        /// <summary>
        /// df工程的proto文件路径
        /// </summary>
        string dfProtoPath;
        /// <summary>
        /// 当前Excel文件名称
        /// </summary>
        string excelFileName;
        /// <summary>
        /// 生成Proto文件
        /// </summary>
        void GenerateProtoFile(string sheetName)
        {
            string fileName = $"{sheetName}.proto";
            protoFilePath = Path.Combine(protoPath, fileName);

            ProcessHeader();
            ProcessVariables(sheetName);
            ProcessMap(sheetName);
        }
        void ProcessHeader()
        {
            if (File.Exists(protoFilePath))
                File.WriteAllText(protoFilePath, "");
            var header = @"
//this proto is auto generated By ExcelToProto
syntax = ""proto3"";
package ConfigTable;
option csharp_namespace = ""ConfigData""; 
";
            var sw = File.AppendText(protoFilePath);
            sw.WriteLine(header);
            sw.Close();
        }
        void ProcessVariables(string sheetName)
        {
            var varDic = headerDic[sheetName];
            int count = 1;
            string str = $"message {sheetName}" + "{\n";
            foreach (var pair in varDic)
            {
                str += $"    {pair.Value} {pair.Key} = {count};\n";
                count++;
            }
            str += "}";
            var sw = File.AppendText(protoFilePath);
            sw.WriteLine(str);
            sw.Close();
        }
        void ProcessMap(string sheetName)
        {
            string str = @"
message Excel_{0}
{{
    repeated {1} {2} = 1;
}}";
            str = string.Format(str, sheetName, sheetName, "Data");
            var sw = File.AppendText(protoFilePath);
            sw.WriteLine(str);
            sw.Close();
        }
        #endregion

        #region 数据对象生成dat文件
        void SerializeDatFile()
        {
            foreach (var pair in _excelInsDic)
            {
                var obj = pair.Value;
                var type = obj.GetType();
                var path = Path.Combine(datPath, $"{type.Name}.dat");
                using (var output = File.Create(path))
                {
                    MessageExtensions.WriteTo((Google.Protobuf.IMessage)obj, output);
                }
            }
        }
        #endregion

        #region Excel生成数据对象
        class ColInfo
        {
            /// <summary>
            /// 列号
            /// </summary>
            public int colIndex;
            /// <summary>
            /// 类型字符串
            /// </summary>
            public string type_string;
            /// <summary>
            /// 变量名
            /// </summary>
            public string variable_string;
            /// <summary>
            /// 表格Key
            /// </summary>
            public bool isKey;
            public ColInfo(int _col, string _type, string _var, bool _isKey)
            {
                this.colIndex = _col;
                this.type_string = _type;
                this.variable_string = _var;
                this.isKey = _isKey;
            }
        }

        /// <summary>
        /// 读取excel文件 数据填充
        /// </summary>
        /// <exception cref="Exception"></exception>
        void GenerateProtoObj()
        {
            //key=变量名 value=proto变量类型名
            //var varDic = headerDic[sheetName];

            if (validSheets.Count == 0) return;
            foreach (DataTable sheet in validSheets)//每个sheet
            {
                //obj准备
                string sheetName = sheet.TableName;
                string trimedSheetName = TrimSheetName(sheetName);
                if (!_excelInsDic.ContainsKey(trimedSheetName))
                    continue;
                //数据保存到_excelIns里
                object _excelIns = _excelInsDic[trimedSheetName];
                //反射找到.Data
                Type excel_Type = _excelIns.GetType();
                PropertyInfo dataProp = excel_Type.GetProperty("Data");
                object dataIns = dataProp.GetValue(_excelIns);
                //Data的type
                Type dataType = dataProp.PropertyType;

                int row_count = sheet.Rows.Count;//总共多少行
                DataRow dataTypeRow = sheet.Rows[0];//数据类型这一行
                DataRow dataNameRow = sheet.Rows[1];//变量名这一行
                //key=有效的列index 这里存好应该读取的列Index
                Dictionary<int, ColInfo> validColDic = new Dictionary<int, ColInfo>();
                //查询到第几列
                int maxValidColIndex = 0;
                for (int j = 0; j < sheet.Columns.Count; j++)//从0列开始按列查询 (col)
                {
                    DataColumn col = sheet.Columns[j];
                    //变量类型 去空格 变小写
                    string dataTypeName = ConvertTypeName2Proto(dataTypeRow[col].ToString());
                    //变量名去空格
                    string variableName = dataNameRow[col].ToString().Trim();
                    bool isEmptyName = string.IsNullOrEmpty(variableName);//变量名空
                    bool isKeyName = Regex.Match(variableName, @"^Key[0-9]?").Success;//KeyX
                    bool isEmptyType = string.IsNullOrEmpty(dataTypeName);//类型空
                    //if (isEmptyName && isEmptyType)
                    //{
                    //    //变量名和类型都没有写 则后面的内容跳过
                    //    break;
                    //}
                    maxValidColIndex = j;
                    if (isKeyName)//Key列必读
                    {
                        string _type = isEmptyType ? "int32" : dataTypeName;
                        //判断Key是数字还是字符串
                        validColDic.Add(j, new ColInfo(j, _type, variableName, true));
                    }
                    else if (!isEmptyName && !isEmptyType)//已配置变量名和类型
                    {
                        if (validColDic.ContainsKey(j))
                        {
                            debugInfo += $"重复变量名 {dataTypeName} \n";
                            throw new Exception();
                        }
                        else
                        {
                            string csTypeName = ConvertTypeName2CS(dataTypeRow[col].ToString());
                            validColDic.Add(j, new ColInfo(j, csTypeName, variableName, false));
                        }
                    }
                }
                //保留最近一次读取到的KeyX
                Dictionary<int, string> KeyCache = new Dictionary<int, string>();
                for (int k = 2; k < row_count; k++)//按行查询
                {
                    DataRow thisRow = sheet.Rows[k];//这一行
                    //一条数据 类型是 ConfigData.{sheetName}
                    object ins = _assembly.CreateInstance($"{NameSpace}.{trimedSheetName}");
                    Type insType = ins.GetType();
                    //反射找 Add 方法
                    MethodInfo addMethod = dataType.GetMethod("Add", new Type[] { insType });
                    //数据插入
                    addMethod.Invoke(dataIns, new[] { ins });
                    foreach (var pair in validColDic)
                    {
                        int col = pair.Key;//有效的列index
                        ColInfo _colInfo = pair.Value;//有效列变量名
                        string valueStr = thisRow[col].ToString().Trim();
                        //多层表 key列自动补全
                        if (_colInfo.isKey)
                        {
                            if (string.IsNullOrEmpty(valueStr))
                                valueStr = KeyCache[col];
                            else
                                KeyCache[col] = valueStr;
                        }
                        string convertedStr;
                        //逗号,全角逗号转换为 |
                        convertedStr = valueStr.Replace(',', '|');
                        convertedStr = convertedStr.Replace('，', '|');
                        //移除方括号
                        convertedStr = convertedStr.Trim(removeChars);
                        ///数据string转换obj
                        object valueObj = GetVariableValue(_colInfo.type_string, convertedStr);

                        string fieldName = FirstCharToLower(_colInfo.variable_string + "_");
                        FieldInfo insField = insType.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                        if (insField != null)
                        {
                            if (valueObj == null)
                            {
                                throw new Exception();
                            }
                            insField?.SetValue(ins, valueObj);
                        }
                    }
                }
            }

        }
        public string FirstCharToLower(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            string str = input.First().ToString().ToLower() + input.Substring(1);
            return str;
        }
        static Char[] removeChars = { '[', ']' };

        object GetVariableValue(string type, string value)
        {
            var isEmpty = false;
            if (string.IsNullOrEmpty(value))
            {
                isEmpty = true;
            }
            if (type == Common.double_)
                return isEmpty ? 0 : double.Parse(value);
            if (type == Common.float_)
                return isEmpty ? 0 : float.Parse(value);
            if (type == Common.int32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.int64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.uint32_)
                return isEmpty ? 0 : uint.Parse(value);
            if (type == Common.uint64_)
                return isEmpty ? 0 : ulong.Parse(value);
            if (type == Common.sint32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.sint64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.fixed32_)
                return isEmpty ? 0 : uint.Parse(value);
            if (type == Common.fixed64_)
                return isEmpty ? 0 : ulong.Parse(value);
            if (type == Common.sfixed32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.sfixed64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.bool_)
                return isEmpty ? false : (value == "1");
            if (type == Common.string_)
                return isEmpty ? string.Empty : value.ToString();
            if (type == Common.bytes_)
                return isEmpty ? ByteString.CopyFromUtf8(string.Empty) : ByteString.CopyFromUtf8(value.ToString());
            if (type == Common.double_s)
            {
                RepeatedField<double> newValue = new RepeatedField<double>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(double.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.float_s)
            {
                RepeatedField<float> newValue = new RepeatedField<float>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(float.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.int32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.int64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.uint32_s)
            {
                RepeatedField<uint> newValue = new RepeatedField<uint>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(uint.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.uint64_s)
            {
                RepeatedField<ulong> newValue = new RepeatedField<ulong>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(ulong.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sint32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sint64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.fixed32_s)
            {
                RepeatedField<uint> newValue = new RepeatedField<uint>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(uint.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.fixed64_s)
            {
                RepeatedField<ulong> newValue = new RepeatedField<ulong>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(ulong.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sfixed32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sfixed64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.bool_s)
            {
                RepeatedField<bool> newValue = new RepeatedField<bool>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(datas[i] == "1");
                    }
                }
                return newValue;
            }
            if (type == Common.string_s)
            {
                RepeatedField<string> newValue = new RepeatedField<string>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(datas[i]);
                    }
                }
                return newValue;
            }
            //Log.Error($"type: {type}  value: {value}");
            return null;
        }
        #endregion

        /// <summary>
        /// 生成 c#proto代码文件
        /// </summary>
        void CMDGenerateCSProto()
        {
            string startUpPath = Assembly.GetExecutingAssembly().Location;
            startUpPath = Path.GetDirectoryName(startUpPath);
            string protocName = @"protoc.exe";

            //protoc.exe 程序路径
            string protocPath = Path.Combine(startUpPath, protocName);
            //protoc.exe 程序路径  for ilr版本的代码给客户端用
            string protocILRPath = Path.Combine(startUpPath, "protocILRuntime.exe");

            string _protoPath = this.protoPath;
            string csProtoILRPath = this.dfProtoPath;
            string[] files = Directory.GetFiles(_protoPath, "*.proto", SearchOption.AllDirectories);
            List<string> fileList = files.ToList();
            foreach (string fileName in fileList)
            {
                string cParams = $"{protocPath} -I={_protoPath} --csharp_out={this.protoPath}   {fileName}";
                Common.Cmd(cParams);

                string cParams2 = $"{protocILRPath} -I={_protoPath} --csharp_out={csProtoILRPath}   {fileName}";
                Common.Cmd(cParams2);
            }
        }

        /// <summary>
        /// 编译C#proto为dll
        /// </summary>
        void CMDGenerateDLL()
        {
            string startUpPath = Assembly.GetExecutingAssembly().Location;
            startUpPath = Path.GetDirectoryName(startUpPath);
            string compilerPath = Settings.Default.Compiler_Path;
            string pbDllPath = Path.Combine(startUpPath, @"Google.Protobuf.dll");
            string _protoPath = this.protoPath;
            string saveDllPath = Path.Combine(_protoPath, @"excel_csharp.dll");
            string csFilePath = Path.Combine(protoPath, "*.cs");
            string cParams = $"{compilerPath} -target:library -out:{saveDllPath} -reference:{pbDllPath} -recurse:{csFilePath}";
            Common.Cmd(cParams);

        }

        /// <summary>
        /// 读取dll 每个sheet名生成格式数据
        /// </summary>
        void ReadDllToAssembly()
        {
            string _protoPath = this.protoPath;
            string saveDllPath = Path.Combine(_protoPath, @"excel_csharp.dll");
            _assembly = Assembly.LoadFrom(saveDllPath);
            _excelInsDic = new Dictionary<string, object>();
            foreach (DataTable sheet in validSheets)//每个sheet
            {
                string sheetName = sheet.TableName;
                if (!_excelInsDic.ContainsKey(sheetName))
                {
                    //sheetName 转换
                    string validSheetName = TrimSheetName(sheetName);
                    if (!_excelInsDic.ContainsKey(validSheetName))
                    {
                        string instName = $"ConfigData.Excel_{validSheetName}";
                        var obj = _assembly.CreateInstance(instName);
                        if (obj == null)
                        {
                            throw new Exception();
                        }
                        _excelInsDic.Add(validSheetName, obj);
                    }
                }
            }
        }
    }
    internal class Common
    {
        #region const types 
        public const string double_ = "double";
        public const string float_ = "float";
        public const string int32_ = "int32";
        public const string int64_ = "int64";
        public const string uint32_ = "uint32";
        public const string uint64_ = "uint64";
        public const string sint32_ = "sint32";
        public const string sint64_ = "sint64";
        public const string fixed32_ = "fixed32";
        public const string fixed64_ = "fixed64";
        public const string sfixed32_ = "sfixed32";
        public const string sfixed64_ = "sfixed64";
        public const string bool_ = "bool";
        public const string string_ = "string";
        public const string bytes_ = "bytes";

        public const string double_s = "double[]";
        public const string float_s = "float[]";
        public const string int32_s = "int32[]";
        public const string int64_s = "int64[]";
        public const string uint32_s = "uint32[]";
        public const string uint64_s = "uint64[]";
        public const string sint32_s = "sint32[]";
        public const string sint64_s = "sint64[]";
        public const string fixed32_s = "fixed32[]";
        public const string fixed64_s = "fixed64[]";
        public const string sfixed32_s = "sfixed32[]";
        public const string sfixed64_s = "sfixed64[]";
        public const string bool_s = "bool[]";
        public const string string_s = "string[]";
        public const string bytes_s = "bytes[]";
        #endregion

        internal static string Cmd(string str)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();

            process.StandardInput.WriteLine(str);
            process.StandardInput.AutoFlush = true;
            process.StandardInput.WriteLine("exit");

            StreamReader reader = process.StandardOutput;//截取输出流

            string output = reader.ReadLine();//每次读取一行

            while (!reader.EndOfStream)
            {
                output += reader.ReadLine();
            }

            process.WaitForExit();
            return output;
        }
    }
}
