using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export
{
    public static class MyExtention
    {
        public static List<T> Clone<T>(this IList<T> list) where T : System.ICloneable
        {
            List<T> newList = new List<T>();
            for (int i = 0; i < list.Count; i++)
            {
                newList.Add((T)list[i].Clone());
            }
            return newList;
        }

        /// <summary>
        /// 实体复制
        /// </summary>
        /// <typeparam name="T">实体类</typeparam>
        /// <param name="list">实体类列表</param>
        /// <param name="productRule">规则实体</param>
        /// <param name="container">(上级)容器实体</param>
        /// <returns></returns>
        public static List<T> Clone<T>(this IList<T> list, ProductRule productRule, BaseEntity container) where T : ICloneable<T>, IRuleEntity
        {
            List<T> newList = new List<T>();
            for (int i = 0; i < list.Count; i++)
            {
                newList.Add((T)list[i].CloneEntity(productRule, container));
            }
            return newList;
        }

        /// <summary>
        /// 字典复制
        /// </summary>
        /// <typeparam name="T">键</typeparam>
        /// <typeparam name="V">值：要实现ICloneable接口</typeparam>
        /// <param name="dict">返回克隆后的字典序列</param>
        /// <returns></returns>
        public static IDictionary<T, V> Clone<T, V>(this IDictionary<T, V> dict) where V : ICloneable
        {
            IDictionary<T, V> newDict = new Dictionary<T, V>();
            foreach (var item in dict)
            {
                newDict.Add(item.Key, (V)item.Value.Clone());
            }
            return newDict;
        }

        /// <summary>
        /// 复制System.Data.DataTable数据
        /// </summary>
        public static System.Data.DataTable CloneData(this System.Data.DataTable table)
        {
            System.Data.DataTable newTable = table.Clone();
            foreach (System.Data.DataRow row in table.Rows)
            {
                newTable.Rows.Add(row.ItemArray);
            }
            return newTable;
        }
    }
}
