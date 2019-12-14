using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;
using System.Configuration;
using  System.Diagnostics;

namespace XlsKeySearch.ExcelHandler
{
    public static class Util
    {
        public static bool IsEnableLog;

        public static ISheet ContentSheet;
        public static ISheet KeySheet;
        public static Dictionary<string, KeyContentObj> KeyDicts = new Dictionary<string, KeyContentObj>();
        public static List<string> MultiKeys = new List<string>();
        //关键字列表
        public static List<string> KeyWords = new List<string>();

        public static void InitLogInfo()
        {
            string logState = ConfigurationManager.AppSettings["IsWriteLog"];
            IsEnableLog = logState.Equals("1");

            Debug.Log("Log write is opened ...");
        }

        public static string GetFileExtension(string filePath)
        {
            if (filePath.Equals(string.Empty))
            {
                return string.Empty;
            }

           return Path.GetExtension(filePath);
        }

        public static string GetFileName(string filePath)
        {
            if (filePath.Equals(string.Empty))
            {
                return string.Empty;
            }

            return Path.GetFileNameWithoutExtension(filePath);
        }

        public static bool IsExcelExtension(string fileExtension)
        {
            return fileExtension.Equals(".xls") || fileExtension.Equals(".xlsx");
        }

        public static void ExcelHandler(string filePath, string outputPath, string searchColNumStr, string outputColIndexStr)
        {
            string fileExtension = GetFileExtension(filePath);
            if(fileExtension.Equals(string.Empty))
            {
                return;
            }

            bool isExcelFile = IsExcelExtension(fileExtension);
            if (!isExcelFile)
            {
                return;
            }

            Int32 searchColNum = GetIntFromString(searchColNumStr);
            Int32 outputColIndex = GetIntFromString(outputColIndexStr);

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook mWorkBook = null;
                if (fileExtension.Equals(".xls"))
                {
                    mWorkBook = new HSSFWorkbook(fs);
                }
                else if(fileExtension.Equals(".xlsx"))
                {
                    mWorkBook = new XSSFWorkbook(fs);
                }

                ContentSheet = mWorkBook.GetSheetAt(0);
                KeySheet = mWorkBook.GetSheetAt(1);
                InitKeyValue();
                InitContainedKeyList();

                for (int i = 0; i < ContentSheet.LastRowNum; i++)
                {
                    if (i >= 2)
                    {
                        IRow row = ContentSheet.GetRow(i);
                        if (row != null)
                        {
                            bool isFindKey = false;

                            for (int j = 0; j < searchColNum; j++)
                            {
                                if (isFindKey)
                                {
                                    break;
                                }
                                string cellValue = row.GetCell(j).ToString();
                                if (string.IsNullOrEmpty(cellValue))
                                {
                                    continue;
                                }
                                foreach (string key in KeyDicts.Keys)
                                {
                                    if (cellValue.Contains(key))
                                    {
                                        KeyContentObj kcObj = KeyDicts[key];
                                        string finalKey = GetFinalKey(kcObj, cellValue);
                                        if (!string.IsNullOrEmpty(finalKey))
                                        {
                                            string finalVaule = KeyDicts[finalKey].ContentStr;

                                            int cellNum = row.LastCellNum;
                                            ICell targetCell = row.GetCell(outputColIndex - 1);
                                            if (targetCell == null)
                                            {
                                                targetCell = row.CreateCell(outputColIndex - 1, CellType.String);
                                            }
                                            targetCell.SetCellValue(finalVaule);
                                            isFindKey = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                string fileName = GetFileName(filePath);
                string outputFileName = outputPath + "\\" + fileName + "_output" + fileExtension;
                FileStream fs2 = File.Create(outputFileName);
                mWorkBook.Write(fs2);

                fs2.Close();

                fs.Close();
                mWorkBook.Close();
                Debug.Log("Write file success!");
            }
        }

        public static void InitKeyValue()
        {
            if (KeySheet == null)
            {
                return;
            }

            for (int i = 0; i < KeySheet.LastRowNum; i++)
            {
                //第三行才开始
                if (i >= 2)
                {
                    IRow row = KeySheet.GetRow(i);
                    if (row != null)
                    {
                        //只取前两列
                        string cellKey = row.GetCell(0).ToString();
                        string cellValue = row.GetCell(1).ToString();
                        KeyContentObj kcObj = new KeyContentObj(cellKey, cellValue);
                        if (!string.IsNullOrEmpty(cellKey))
                        {
                            if (!KeyDicts.ContainsKey(cellKey))
                            {
                                KeyDicts[cellKey] = kcObj;
                                KeyWords.Add(cellKey);
                            }
                            else
                            {
                                if (!MultiKeys.Contains(cellKey))
                                {
                                    Debug.Log("Multi key name:{0}, rowIndex:{1}", cellKey, i);
                                    MultiKeys.Add(cellKey);
                                }
                            }
                        }
                    }
                }

            }
        }

        public static void InitContainedKeyList()
        {
            foreach (string key in KeyDicts.Keys)
            {
                KeyContentObj kcObj = KeyDicts[key];
                kcObj.InitContainedKeys(KeyWords);
            }
        }

        //由于key之间有包含关系，需要决策最终使用的key
        //比如“一班，十一班，一十一班”，如果key是“一班”，对于一句“他在一十一班上学”，需要决策的结果是key为“一十一班”
        public static string GetFinalKey(KeyContentObj kcObj, string originalStr)
        {
            if (kcObj == null)
            {
                return string.Empty;
            }

            //BeContainedKeys是已经排好序，从长度最大的字符串开始判断即可
            for (int i = 0; i < kcObj.BeContainedKeys.Count; i++)
            {
                string curKey = kcObj.BeContainedKeys[i];
                if (originalStr.Contains(curKey))
                {
                    return curKey;
                }
            }

            return kcObj.KeyStr;
        }

        public static bool CheckStringContentToIntValid(string str)
        {
            if (str.Equals(string.Empty))
            {
                return false;
            }

            Int32 intNumber = 0;
            Int32.TryParse(str, out intNumber);
            return intNumber > 0;
        }

        public static Int32 GetIntFromString(string str)
        {
            if (!CheckStringContentToIntValid(str))
            {
                return 0;
            }

            Int32 intNumber = 0;
            Int32.TryParse(str, out intNumber);
            return intNumber;
        }

        public static int GetLineNum()
        {
            StackTrace st = new StackTrace(true);
            StackFrame sf = st.GetFrame(3);
            return sf.GetFileLineNumber();
        }

        public static string GetFileName()
        {
            StackTrace st = new StackTrace(true);
            StackFrame sf = st.GetFrame(3);
            return sf.GetFileName();
        }

        public static string GetMethodName()
        {
            StackTrace st = new StackTrace(true);
            StackFrame sf = st.GetFrame(3);
            return sf.GetMethod().Name;
        }
    }
}
