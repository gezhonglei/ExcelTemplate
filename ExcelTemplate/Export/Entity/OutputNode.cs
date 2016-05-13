using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Writer.CellRender;
using ExportTemplate.Export.Writer.Convertor;

namespace ExportTemplate.Export.Entity
{

    /// <summary>
    /// 输出对象：区域或结点
    /// </summary>
    public class OutputNode
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex;
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex;
        /// <summary>
        /// 跨行数
        /// </summary>
        public int RowSpan;
        /// <summary>
        /// 跨列数
        /// </summary>
        public int ColumnSpan;
        /// <summary>
        /// 输出内容
        /// </summary>
        public object Content;
        /// <summary>
        /// Tag标签（其它用处：多行多列标题作为连接ID使用）
        /// </summary>
        public object Tag;
        /// <summary>
        /// 输出值时是否添加
        /// </summary>
        public bool ValueAppend = false;

        /// <summary>
        /// 指定单元格是否在此结点区域
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="colIndex">列索引</param>
        /// <returns>是否在区域中</returns>
        public bool IsInArea(int rowIndex, int colIndex)
        {
            return RowIndex <= rowIndex && rowIndex < RowIndex + RowSpan
                && ColumnIndex <= colIndex && colIndex < ColumnIndex + ColumnSpan;
        }

        public bool IsRegion { get { return ColumnSpan > 1 || RowSpan > 1; } }

        protected List<ICellRender> renders = new List<ICellRender>();
        public void AddRender(ICellRender render)
        {
            renders.Add(render);
        }
        public IList<ICellRender> GetRenders()
        {
            return renders;
        }

        public void OnRender(NPOI.SS.UserModel.ICell exCell)
        {
            foreach (var render in renders)
            {
                render.Render(exCell);
            }
        }

        public AbstractValueSetter Convertor;
    }

}
