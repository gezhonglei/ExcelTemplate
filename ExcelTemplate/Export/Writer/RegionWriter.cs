using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Entity.Region;
using ExportTemplate.Export.Util;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportTemplate.Export.Writer
{
    public class RegionWriter : BaseWriter
    {
        public RegionWriter(RegionTable entity, BaseWriterContainer parent) : base(entity, parent) { }

        protected override void Writing(object sender, WriteEventArgs args)
        {
            ISheet exSheet = args.ExSheet;
            RegionTable headerTable = args.Entity as RegionTable;
            if (headerTable == null) return;

            int colIndex = this.ColIndex, rowIndex = this.RowIndex, colCount = this.ColCount, rowCount = this.RowCount;
            int colHeaderLevel = headerTable.ColumnHeaderLevel;
            int rowHeaderLevel = headerTable.RowHeaderLevel;
            //1、Excel格式准备
            ICell exCell = GetStandardCell(exSheet, RowIndex, colIndex, true);
            NPOIExcelUtil.CopyCell(exSheet, exCell, exCell.ColumnIndex + 1, colCount - 1);
            NPOIExcelUtil.CopyRow(exSheet, exCell.RowIndex, exCell.RowIndex + 1, rowHeaderLevel - 1);

            exCell = GetStandardCell(exSheet, rowIndex + rowHeaderLevel, colIndex);
            NPOIExcelUtil.CopyCell(exSheet, exCell, exCell.ColumnIndex + 1, colCount - 1);
            NPOIExcelUtil.CopyRow(exSheet, exCell.RowIndex, exCell.RowIndex + 1, rowCount - rowHeaderLevel - 1);

            //2、数据填充
            IList<OutputNode> nodes = this.GetNodes();
            foreach (var node in nodes)
            {
                IRow exRow = exSheet.GetRow(node.RowIndex) ?? exSheet.CreateRow(node.RowIndex);
                //if (exRow == null) continue;
                exCell = exRow.GetCell(node.ColumnIndex) ?? exRow.CreateCell(node.ColumnIndex);
                //if (exCell == null) continue;
                if (node.IsRegion)
                {
                    exSheet.AddMergedRegion(new CellRangeAddress(node.RowIndex, node.RowIndex + node.RowSpan - 1, node.ColumnIndex, node.ColumnIndex + node.ColumnSpan - 1));
                }
                NPOIExcelUtil.SetCellValueByDataType(exCell, node.Content);
            }

            //3、自适应宽度
            for (int i = 0; i < this.ColCount; i++)
            {
                int endRow = this.RowIndex + (i < colHeaderLevel || rowHeaderLevel == 0 ? this.RowCount : rowHeaderLevel) - 1;
                NPOIExcelUtil.AutoFitColumnWidth(exSheet, colIndex + i, rowIndex, endRow);
            }

            //4、锁定标题区域(只有列标题、只有行标题、多行多列标题三种情况锁定不一样）
            if (headerTable.Freeze)
            {
                NPOIExcelUtil.Freeze(exSheet, rowHeaderLevel > 0 ? rowIndex + rowHeaderLevel : 0, colHeaderLevel > 0 ? colIndex + colHeaderLevel : 0);
            }

            NPOIExcelUtil.SetAreaBorder(exSheet, rowIndex, colIndex, rowCount, colCount);
        }

        private ICell GetStandardCell(ISheet sheet, int rowIndex, int colIndex, bool isTitle = false)
        {
            //如存在直接返回
            IRow row = sheet.GetRow(rowIndex);
            if (row == null) row = sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(colIndex);
            if (cell != null) return cell;

            //如不存在创建
            cell = row.CreateCell(colIndex);
            cell.CellStyle = NPOIExcelUtil.CreateCellStyle(sheet.Workbook, 0, isTitle ? HSSFColor.LightTurquoise.Index : HSSFColor.White.Index, isTitle ? HorizontalAlignment.Center : HorizontalAlignment.Left);
            cell.CellStyle.SetFont(NPOIExcelUtil.CreateFont(sheet.Workbook, isTitle ? "微软雅黑" : "宋体"));
            return cell;
        }

        /// <summary>
        /// 填充模式（多行多列必须是复制填充模式的）
        /// </summary>
        public override bool CopyFill { get { return true; } }
        public override int TempleteRows { get { return 2; } }
        public override int RowCount
        {
            get
            {
                RegionTable entity = Entity as RegionTable;
                ExportTemplate.Export.Entity.Region.Region rowheader = entity.GetRegion(RegionType.RowHeader);
                ExportTemplate.Export.Entity.Region.Region bodyHeader = entity.GetRegion(RegionType.Body);
                return bodyHeader.RowCount + (rowheader != null ? rowheader.RowCount : 0);
            }
        }

        public override int ColCount
        {
            get
            {
                RegionTable entity = Entity as RegionTable;
                ExportTemplate.Export.Entity.Region.Region colheader = entity.GetRegion(RegionType.ColumnHeader);
                ExportTemplate.Export.Entity.Region.Region bodyHeader = entity.GetRegion(RegionType.Body);
                return bodyHeader.ColumnCount + (colheader != null ? colheader.ColumnCount : 0);
            }
        }

        public override IList<OutputNode> GetNodes()
        {
            RegionTable entity = Entity as RegionTable;
            List<OutputNode> nodes = new List<OutputNode>();
            ExportTemplate.Export.Entity.Region.Region[] regions = new ExportTemplate.Export.Entity.Region.Region[] { 
                entity.GetRegion(RegionType.Body),
                entity.GetRegion(RegionType.RowHeader),
                entity.GetRegion(RegionType.ColumnHeader),
                entity.GetRegion(RegionType.Corner)
            };
            foreach (var region in regions)
            {
                if (region != null)
                {
                    nodes.AddRange(region.GetNodes());
                }
            }
            return nodes;
        }

    }

}
