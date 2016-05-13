using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;
using NPOI.SS.UserModel;

namespace ExportTemplate.Export.Writer
{

    public class DynamicAreaWriter : BaseWriterContainer
    {
        private DynamicSource _dynamicObject = null;
        //private Source _source;
        public Source Source
        {
            get
            {
                if (_dynamicObject.Source == null)
                {
                    DynamicArea entity = Entity as DynamicArea;
                    _dynamicObject.Source = entity.ProductRule.GetSource(entity.SourceName);
                }
                return _dynamicObject.Source;
            }
        }

        public DynamicAreaWriter(DynamicArea entity, BaseWriterContainer parent)
            : base(entity, parent)
        {
            _dynamicObject = new DynamicSource(entity.ProductRule);
            _dynamicObject.SetParam("RowIndex", p => p);
            _dynamicObject.SetParam("RowNum", p => p + 1);
        }

        protected override void CreatingSubWriters()
        {
            DynamicArea dynamicArea = (DynamicArea)base.Entity;
            foreach (var container in this.GetAllContainers())
            {
                Table table = container as Table;

                if (table != null && table.CopyFill)
                {
                    Components.Add(new TableWriter(table as Table, this));
                }
                else if (container is Cell)
                {
                    Components.Add(new CellWriter(container as Cell, this));
                }
            }
        }

        public override void PreWrite(WriteEventArgs args)
        {
            base.PreWrite(args);
            ISheet exSheet = args.ExSheet;
            DynamicArea dynamicArea = args.Entity as DynamicArea;
            if (dynamicArea == null) return;

            int baseRow = this.RowIndex;
            for (int i = 1; i < this.Count; i++)
            {
                for (int j = 0; j < dynamicArea.Location.RowCount; j++)
                {
                    NPOIExcelUtil.CopyRow(exSheet, baseRow + j, baseRow + j + dynamicArea.Location.RowCount * i);
                }
            }
        }

        public int Count
        {
            get
            {
                return this.Source != null && this.Source.Table != null ? this.Source.Table.Rows.Count : 0;
            }
        }

        public override int TempleteRows { get { return Entity.Location.RowCount; } }
        public override bool CopyFill { get { return true; } }
        public override int ColCount { get { return Entity.Location.ColCount; } }

        public override int RowCount
        {
            get
            {
                return Count * Entity.Location.RowCount + this.IncreasedRows;
            }
        }

        public override int IncreasedRows
        {
            get
            {
                return Components.Where(p => p.CopyFill).Sum(p => p.IncreasedRows);
            }
        }

        private List<BaseEntity> allContainers = null;
        public List<BaseEntity> GetAllContainers(bool reset = false)
        {
            if (reset) allContainers = null;
            if (allContainers != null) return allContainers;

            DynamicArea entity = Entity as DynamicArea;
            allContainers = new List<BaseEntity>();
            DataTable dt = this.Source != null ? this.Source.Table : null;
            if (dt == null) return allContainers;
            for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                for (int i = 0; i < entity.Tables.Length; i++)
                {
                    Table tmpTable = entity.Tables[i];
                    Table table = (Table)tmpTable.Clone(tmpTable.ProductRule, tmpTable.Container);
                    //调整行位置(相对)
                    table.Location.RowIndex = entity.Location.RowCount * rowIndex + table.Location.RowIndex;
                    //获取数据源
                    table.SourceName = _dynamicObject.GetDynamicValue(table.SourceName, rowIndex);
                    //Source source = entity.GetDynamicSource(table.SourceName, rowIndex);
                    //if (source != null)
                    //{
                    //    table.Source = source;
                    //}
                    allContainers.Add(table);
                }

                for (int i = 0; i < entity.Cells.Length; i++)
                {
                    Cell tmpCell = entity.Cells[i];
                    Cell cell = (Cell)tmpCell.Clone(tmpCell.ProductRule, tmpCell.Container);
                    //调整行位置（相对）
                    cell.Location.RowIndex = entity.Location.RowCount * rowIndex + cell.Location.RowIndex;
                    if (!string.IsNullOrEmpty(cell.SourceName))
                    {
                        //Source source = entity.GetDynamicSource(cell.SourceName, rowIndex);
                        //if (source != null) cell.Source = source;
                        cell.SourceName = _dynamicObject.GetDynamicValue(cell.SourceName, rowIndex);
                    }
                    if (!string.IsNullOrEmpty(cell.Value))
                    {
                        cell.Value = _dynamicObject.GetDynamicValue(cell.Value, rowIndex);
                    }
                    allContainers.Add(cell);
                }
            }
            return allContainers;
        }

        public override IList<OutputNode> GetNodes()
        {
            DynamicArea entity = Entity as DynamicArea;
            List<OutputNode> nodes = new List<OutputNode>();
            //allContainers = GetAllContainers();
            for (int i = 0; i < Components.Count; i++)
            {
                nodes.AddRange(Components[i].GetNodes());
            }
            return nodes;
        }
    }

}
