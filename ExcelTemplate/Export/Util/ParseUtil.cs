using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExportTemplate.Export.Element;

namespace ExportTemplate.Export.Util
{
    public class ParseUtil
    {
        public static int ParseInt(string str, int valueIfNullOrError = 0)
        {
            return string.IsNullOrEmpty(str) || !IsNumeric(str) ? valueIfNullOrError : int.Parse(str);
        }

        public static bool ParseBoolean(string str, bool valueIfNull = false)
        {
            return string.IsNullOrEmpty(str) ? valueIfNull : bool.Parse(str);
        }

        //public static bool ParseBoolean(string str, bool valueIfNull = false, bool valueIfError = false)
        //{
        //    str = string.IsNullOrEmpty(str) ? valueIfNull.ToString() : str;
        //    bool.TryParse(str, out valueIfError);
        //    return valueIfError;
        //}

        public static string IfNullOrEmpty(string str, string valueIfNull = "")
        {
            return string.IsNullOrEmpty(str) ? valueIfNull : str;
        }
        
        public static int ParseColumnIndex(string str, int defValue = -1)
        {
            return string.IsNullOrEmpty(str) ? defValue : ParseUtil.IsNumeric(str) ? int.Parse(str) : ParseUtil.FromBase26(str);
        }

        public static T ParseEnum<T>(string str, T defValue)
        {
            return string.IsNullOrEmpty(str) ? defValue : (T)Enum.Parse(typeof(T), str, true);
        }

        /// <summary>
        /// 将26进制的字母行号转换成十进制数字
        /// </summary>
        /// <param name="str">行字母</param>
        /// <returns>十进制数字(索引从1开始)</returns>
        public static int FromBase26(string str)
        {
            str = str.ToUpper();
            int value = 0;
            for (int i = 0; i < str.Length; i++)
            {
                value = value * 26 + (str[i] - 'A' + 1);
            }
            return value;
        }

        /// <summary>
        /// 将整数转换成26进制的字母行号
        /// </summary>
        /// <param name="value">数字：从1开始</param>
        /// <returns>返回26进制数</returns>
        public static string ToBase26(int value)
        {
            List<char> charLsit = new List<char>();
            while (value > 0)
            {
                charLsit.Add((char)(value % 26 + 'A' - 1));
                value /= 26;
            }
            charLsit.Reverse();
            return string.Join("", charLsit);
        }

        public static int ToInteger(string value, int errorValue = -1)
        {
            int result = errorValue;
            if (int.TryParse(value, out result))
            {
                return result;
            }
            return errorValue;
        }

        public static bool Match(string pattern, string value, RegexOptions options = RegexOptions.None)
        {
            return new Regex(pattern, options).IsMatch(value);
        }

        /// <summary>
        /// 判断是否是数值类型
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>是否数据类型</returns>
        public static bool IsNumberType(Type type)
        {
            List<Type> NumberTypes = new List<Type>(new Type[] { 
                typeof(Int16), typeof(Int32), typeof(Int64), typeof(float), typeof(Double), typeof(Decimal),
                typeof(Int16?),typeof(Int32?),typeof(Int64?), typeof(float?), typeof(Double?), typeof(Decimal?)
            });
            return NumberTypes.Contains(type);
        }

        public static bool IsNumeric(string value)
        {
            return Match(@"^\d+$", value);
        }

        public static bool IsAlphabets(string value, bool ignoreCase = false)
        {
            return Match(@"^[A-Z]+$", value, ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }
        
        #region 测试
        
        public static void test()
        {
            string str = true.ToString();
            bool boolvalue = bool.Parse("TRUE");
            boolvalue = bool.Parse("True");
            boolvalue = bool.Parse("true");

            str = BaseElement.ELEMENT_NAME;
            str = SourceElement.ELEMENT_NAME;
            str = DataListElement.ELEMENT_NAME;
        }

        public static void test2()
        {
            string formula = "Sum(a1:b3) + AB12 + $A23 + A$29 + a1";
            formula = formula.ToUpper();
            int rowCount = 1, colCount = 1, tmpIndex = -1;
            string rowNum;
            string colNum;
            MatchCollection matched = new Regex(@"\$?[A-Z]+\$?\d+", RegexOptions.IgnoreCase).Matches(formula);
            foreach (Match match in matched)
            {
                string value = match.Value;
                for (int i = 0; i < value.Length; i++)
                {
                    if ('0' <= value[i] && value[i] <= '9')
                    {
                        tmpIndex = i;
                        break;
                    }
                }
                if (value[0] == '$' && value[tmpIndex - 1] == '$')
                    continue;
                tmpIndex = value[tmpIndex - 1] == '$' ? tmpIndex - 1 : tmpIndex;//有"$"在数字前，调整到$位置
                colNum = value.Substring(0, tmpIndex);
                rowNum = value.Substring(tmpIndex);
                if (value[0] != '$')
                {
                    colNum = ToBase26(FromBase26(colNum) + colCount);
                }
                if (value[tmpIndex] != '$')
                    rowNum = (int.Parse(rowNum) + rowCount).ToString();
                formula = formula.Replace(value, colNum + rowNum);
            }
        }

        public static void test3()
        {
            string str = "ad";
            int value = FromBase26(str);
            str = ToBase26(value);

            str = "abcd{field1}.{field2}";
            bool tr = Match(@"{\w+}", str);

            str = "A_{table}_B";
            tr = Match(@"(?:{)([\w\.]+)(?:})", str);

            MatchCollection matches = new Regex(@"(?:{)(\w+\.\w+)(?:})").Matches("aaa{Cell.Field}bbb");
            for (int i = 0; i < matches.Count; i++)
            {
                string match = matches[i].Groups[1].Value;
            }
        }
        #endregion 测试
    }
}
