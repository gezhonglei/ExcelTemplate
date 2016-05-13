using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;
using ExportTemplate.Export.Writer;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportTemplate.Export.Writer
{
    /// <summary>
    /// 外加结点输出
    /// </summary>
    public class NodeWriter : BaseWriter
    {
        IList<OutputNode> _nodes = null;
        public NodeWriter(IList<OutputNode> nodes)
            : base(null, null)
        {
            _nodes = nodes;
        }

        protected override void Writing(object sender, WriteEventArgs args)
        {
            WriteNodes(args.ExSheet, _nodes);
        }

        public static void WriteNodes(ISheet exSheet, IList<OutputNode> nodes)
        {
            if (nodes == null) return;
            foreach (var node in nodes)
            {
                int rowIndex = node.RowIndex;
                IRow exRow = exSheet.GetRow(rowIndex) ?? exSheet.CreateRow(rowIndex);
                ICell exCell = exRow.GetCell(node.ColumnIndex) ?? exRow.CreateCell(node.ColumnIndex);
                if (node.IsRegion)
                {
                    exSheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + node.RowSpan - 1, node.ColumnIndex, node.ColumnIndex + node.ColumnSpan - 1));
                }
                if (node.Convertor != null)
                {
                    node.Convertor.OnSetValue(exCell, node.Content);
                }
                else
                {
                    NPOIExcelUtil.SetCellValueByDataType(exCell, node.Content);
                }
                node.OnRender(exCell);
            }
        }
    }
}
