using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsKeySearch.ExcelHandler
{
    public class KeyContentObj
    {
        public string KeyStr { get; set; }
        public string ContentStr { get; set; }

        public List<string> BeContainedKeys { get; set; }

        public KeyContentObj(string keyStr, string contentStr)
        {
            this.KeyStr = keyStr;
            this.ContentStr = contentStr;
            this.BeContainedKeys = new List<string>();
        }

        public void InitContainedKeys(List<string> keys)
        {
            if (keys == null)
            {
                return;
            }

            for (int i = 0; i < keys.Count; i++)
            {
                string curKey = keys[i];
                if (!this.KeyStr.Equals(curKey))
                {
                    if (curKey.Contains(this.KeyStr))
                    {
                        if (!BeContainedKeys.Contains(curKey))
                        {
                            BeContainedKeys.Add(curKey);
                        }
                    }
                }
            }

            BeContainedKeys.Sort((delegate(string LeftStr, string RightStr) {
                int result = RightStr.Length.CompareTo(LeftStr.Length);
                return result;
            }));
        }

    }
}
