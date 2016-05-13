using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity.Region
{
    /// <summary>
    /// Corner区域的合并类型
    /// </summary>
    public enum CornerSpanRule
    {
        /// <summary>
        /// 根据行标题合并
        /// </summary>
        BaseOnRowHeader,
        /// <summary>
        /// 根据列标题合并
        /// </summary>
        BaseOnColumnHeader,
        /// <summary>
        /// 合并成一个单元格
        /// </summary>
        AllInOne,
        /// <summary>
        /// 不合并
        /// </summary>
        None,
    }

    /// <summary>
    /// 左上角区域
    /// </summary>
    public class CornerRegion : Region
    {
        public CornerSpanRule SpanRule = CornerSpanRule.None;

        public CornerRegion(RegionTable table)
        {
            this._table = table;
        }

        protected override List<OutputNode> CalulateNodes()
        {
            List<OutputNode> nodes = new List<OutputNode>();
            if (Source.Table == null)
            {
                return nodes;
            }

            if (SpanRule == CornerSpanRule.AllInOne)
            {
                OutputNode node = new OutputNode()
                {
                    Content = Source.Table.Rows[0][0],
                    ColumnIndex = this.ColumnIndex,
                    RowIndex = this.RowIndex,
                    ColumnSpan = this.ColumnCount,
                    RowSpan = this.RowCount
                };
                nodes.Add(node);
            }
            else if (SpanRule == CornerSpanRule.BaseOnRowHeader)
            {
                string tmpField = !string.IsNullOrEmpty(Field) && Source.Table.Columns.Contains(Field) ? Field : Source.Table.Columns[0].ColumnName;
                int span = this.ColumnCount;
                int rows = Math.Min(this.ColumnCount, Source.Table.Rows.Count);
                for (int i = 0; i < rows; i++)
                {
                    OutputNode node = new OutputNode()
                    {
                        Content = Source.Table.Rows[i][tmpField],
                        ColumnIndex = this.ColumnIndex,
                        RowIndex = this.RowIndex + i,
                        ColumnSpan = span,
                        RowSpan = 1
                    };
                    nodes.Add(node);
                }
            }
            else if (SpanRule == CornerSpanRule.BaseOnColumnHeader)
            {
                string tmpField = !string.IsNullOrEmpty(Field) && Source.Table.Columns.Contains(Field) ? Field : Source.Table.Columns[0].ColumnName;
                int span = this.RowCount;
                int rows = Math.Min(this.ColumnCount, Source.Table.Rows.Count);
                for (int i = 0; i < rows; i++)
                {
                    OutputNode node = new OutputNode()
                    {
                        Content = Source.Table.Rows[i][tmpField],
                        ColumnIndex = this.ColumnIndex + i,
                        RowIndex = this.RowIndex,
                        ColumnSpan = 1,
                        RowSpan = span
                    };
                    nodes.Add(node);
                }
            }
            else
            {
                int rows = Math.Min(this.RowCount, Source.Table.Rows.Count);
                int columns = Math.Min(this.ColumnCount, Source.Table.Columns.Count);
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        OutputNode node = new OutputNode()
                        {
                            Content = Source.Table.Rows[i][j],
                            RowIndex = this.RowIndex + i,
                            ColumnIndex = this.ColumnIndex + j,
                            RowSpan = 1,
                            ColumnSpan = 1
                        };
                        nodes.Add(node);
                    }
                }
            }
            return nodes;
        }

        public override int RowIndex
        {
            get { return _table.Location.RowIndex; }
        }

        public override int ColumnIndex
        {
            get { return _table.Location.ColIndex; }
        }

        public override int RowCount
        {
            get
            {
                if (_location != null)
                {
                    return _location.RowCount;
                }
                RowHeaderRegion rowHeader = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                return rowHeader != null ? rowHeader.RowCount : 0;
            }
        }

        public override int ColumnCount
        {
            get
            {
                if (_location != null)
                {
                    return _location.ColCount;
                }
                ColumnHeaderRegion colHeader = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                return colHeader != null ? colHeader.RowCount : 0;
            }
        }

        public override string ToString()
        {
            return string.Format("<Region{0} />",
                " type=\"corner\"" +
                (Source != null ? string.Format(" source=\"{0}{1}\"", Source.Name, !string.IsNullOrEmpty(Field) ? "." + Field : string.Empty) : string.Empty) +
                (SpanRule != CornerSpanRule.None ? string.Format(" spanRule=\"{0}\"", SpanRule == CornerSpanRule.BaseOnRowHeader ? "row" : SpanRule == CornerSpanRule.BaseOnColumnHeader ? "column" : "one") : string.Empty) +
                (!string.IsNullOrEmpty(EmptyFill) ? string.Format(" emptyFill=\"{0}\"", EmptyFill) : string.Empty));
        }

        public override Region CloneEntity(ProductRule productRule, BaseEntity container)
        {
            CornerRegion newRegion = new CornerRegion(container as RegionTable)
            {
                _location = _location != null ? _location.Clone() : null,
                EmptyFill = EmptyFill,
                Field = Field,
                //Source = Source.Clone() as Source,
                SpanRule = SpanRule
            };
            if (Source != null)
            {
                newRegion.Source = productRule.GetSource(Source.Name);
            }
            return newRegion;
        }
    }

}
