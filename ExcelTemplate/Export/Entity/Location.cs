using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 位置方式
    /// </summary>
    public enum LocationPolicy
    {
        /// <summary>
        /// 未定义
        /// </summary>
        Undefined,
        /// <summary>
        /// 数据行首部
        /// </summary>
        Head,
        /// <summary>
        /// 数据行尾部
        /// </summary>
        Tail,
        /// <summary>
        /// Excel模板绝对行位置
        /// </summary>
        Absolute
    }

    /// <summary>
    /// 位置范围
    /// </summary>
    public class Location
    {
        public int RowIndex;
        public int ColIndex;
        public int RowCount = 0;    //RowCount < 0 无效数据
        public int ColCount = 0;    //ColCount < 0 无效数据

        public Location() { }
        public Location(string area)
        {
            parse(area);
        }
        public Location(Location location)
        {
            RowIndex = location.RowIndex;
            RowCount = location.RowCount;
            ColIndex = location.ColIndex;
            ColCount = location.ColCount;
        }

        /// <summary>
        /// 根据字符串解析对象
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool parse(string value)
        {
            RowIndex = 0;
            ColIndex = 0;
            RowCount = 0;
            ColCount = 0;

            if (string.IsNullOrEmpty(value))
            {
                return false;
            }
            try
            {
                int tmpIndex = -1, tmpColIndex = 0, tmpRowIndex = 0;
                string tmp;
                value = value.ToUpper();
                int index = value.IndexOf(':');
                //先处理后部分
                tmp = value.Substring(index + 1);
                for (int i = 0; i < tmp.Length; i++)
                {
                    if (!char.IsLetter(tmp[i]))
                    {
                        tmpIndex = i;
                        break;
                    }
                    tmpColIndex = tmpColIndex * 26 + (tmp[i] - 'A' + 1);      //从0开始
                }
                tmpColIndex -= 1;
                tmpRowIndex = int.Parse(tmp.Substring(tmpIndex)) - 1;  //从0开始
                ColIndex = tmpColIndex;
                RowIndex = tmpRowIndex;

                if (index > -1)
                {
                    tmpColIndex = 0;
                    tmpRowIndex = 0;
                    tmpIndex = -1;
                    tmp = value.Substring(0, index);
                    for (int i = 0; i < tmp.Length; i++)
                    {
                        if (!char.IsLetter(tmp[i]))
                        {
                            tmpIndex = i;
                            break;
                        }
                        tmpColIndex = tmpColIndex * 26 + (tmp[i] - 'A' + 1);//从0开始
                    }
                    tmpColIndex -= 1;
                    tmpRowIndex = int.Parse(tmp.Substring(tmpIndex)) - 1;//从0开始

                    ColCount = Math.Abs(tmpColIndex - ColIndex) + 1;
                    RowCount = Math.Abs(tmpRowIndex - RowIndex) + 1;
                    ColIndex = Math.Min(ColIndex, tmpColIndex);
                    RowIndex = Math.Min(RowIndex, tmpRowIndex);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
                throw new Exception(string.Format("{0}不能解析为Location对象", value), ex);
            }
        }

        public override bool Equals(object obj)
        {
            Location location = obj as Location;
            return this == obj || (obj != null && this.ColIndex == location.ColIndex && this.ColCount == location.ColCount
                && this.RowIndex == location.RowIndex && this.RowCount == location.RowCount);
        }

        public override int GetHashCode()
        {
            return string.Format("{0}.{1}.{2}.{3}", ColIndex, RowIndex, ColCount, RowCount).GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("{0}{1}{2}{3}", (char)(ColIndex + 'A'), RowIndex + 1,
                ColCount > 0 ? ":" + (char)(ColIndex + ColCount - 1 + 'A') : "", 
                RowCount > 0 ? (RowIndex + RowCount).ToString() : "");
        }

        public Location Clone()
        {
            return new Location()
            {
                ColCount = this.ColCount,
                ColIndex = this.ColIndex,
                RowCount = this.RowCount,
                RowIndex = this.RowIndex
            };
        }
    }
}
