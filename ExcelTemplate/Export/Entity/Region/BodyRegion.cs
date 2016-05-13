using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExportTemplate.Export.Entity.Region
{
    /// <summary>
    /// Body区域：数据区域
    /// </summary>
    public class BodyRegion : Region
    {
        public BodyRegion(RegionTable table)
        {
            this._table = table;
        }

        protected override List<OutputNode> CalulateNodes()
        {
            List<OutputNode> nodes = new List<OutputNode>();
            RowHeaderRegion rowheader = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
            ColumnHeaderRegion colheader = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
            if (rowheader != null && colheader != null)
            {
                IEnumerable<OutputNode> rowHeaderNodes = rowheader.GetNodes().Where(p => p.Tag != null);
                IEnumerable<OutputNode> colHeaderNodes = colheader.GetNodes().Where(p => p.Tag != null);
                string refRowField = rowheader.HeaderBodyRelation.ReferecedField;
                string refColField = colheader.HeaderBodyRelation.ReferecedField;
                IEnumerable<OutputNode> tmpRowNodes = null;
                IEnumerable<OutputNode> tmpColNodes = null;
                foreach (DataRow row in Source.Table.Rows)
                {
                    tmpRowNodes = rowHeaderNodes.Where(p => p.Tag.Equals(row[refRowField]));
                    tmpColNodes = colHeaderNodes.Where(p => p.Tag.Equals(row[refColField]));
                    if (tmpColNodes.Count() > 0 && tmpRowNodes.Count() > 0)
                    {
                        nodes.Add(new OutputNode()
                        {
                            Content = row[Field],
                            ColumnIndex = tmpColNodes.First().RowIndex,
                            RowIndex = tmpRowNodes.First().ColumnIndex,
                            ColumnSpan = 1,
                            RowSpan = 1
                        });
                    }
                }

                if (!string.IsNullOrEmpty(EmptyFill))
                {
                    int rIndex = this.RowIndex;
                    int cIndex = this.ColumnIndex;
                    for (int i = 0; i < this.RowCount; i++)
                    {
                        for (int j = 0; j < this.ColumnCount; j++)
                        {
                            if (!nodes.Exists(p => p.IsInArea(rIndex + i, cIndex + j)))
                            {
                                nodes.Add(new OutputNode()
                                {
                                    Content = EmptyFill,
                                    RowIndex = rIndex + i,
                                    ColumnIndex = cIndex + j,
                                    RowSpan = 1,
                                    ColumnSpan = 1
                                });
                            }
                        }
                    }
                }
            }
            else
            {
                HeaderRegion header = colheader != null ? (HeaderRegion)colheader : rowheader;
                bool isRowHeader = header is RowHeaderRegion;
                IEnumerable<OutputNode> headerNodes = header.GetNodes().Where(p => p.Tag != null);
                DataColumnCollection columns = Source.Table.Columns;
                DataRowCollection rows = Source.Table.Rows;
                string refField = header.HeaderBodyRelation.ReferecedField;
                if (!string.IsNullOrEmpty(refField))
                {
                    foreach (DataRow row in rows)
                    {
                        IEnumerable<OutputNode> tmpRowNodes = headerNodes.Where(p => p.Tag.Equals(row[refField]));
                        if (tmpRowNodes.Count() > 0)
                        {
                            OutputNode headerNode = tmpRowNodes.First();
                            for (int i = 0; i < columns.Count; i++)
                            {
                                if (/*columns[i].ColumnName != refField &&*/ row[i] != null && !string.IsNullOrEmpty(row[i].ToString()))
                                {
                                    nodes.Add(new OutputNode()
                                    {
                                        Content = row[i],
                                        ColumnIndex = isRowHeader ? headerNode.ColumnIndex : ColumnIndex + i,
                                        RowIndex = isRowHeader ? RowIndex + i : headerNode.RowIndex,
                                        ColumnSpan = 1,
                                        RowSpan = 1
                                    });
                                }
                            }
                        }
                    }
                }
                else
                {
                    for (int j = 0; j < columns.Count; j++)
                    {
                        IEnumerable<OutputNode> tmpRowNodes = headerNodes.Where(p => p.Tag.Equals(columns[j].ColumnName));
                        if (tmpRowNodes.Count() > 0)
                        {
                            OutputNode headerNode = tmpRowNodes.First();
                            for (int i = 0; i < rows.Count; i++)
                            {
                                if (rows[i][j] != null && !string.IsNullOrEmpty(rows[i][j].ToString()))
                                {
                                    nodes.Add(new OutputNode()
                                    {
                                        Content = rows[i][j],
                                        ColumnIndex = isRowHeader ? headerNode.ColumnIndex : ColumnIndex + i,
                                        RowIndex = isRowHeader ? RowIndex + i : headerNode.RowIndex,
                                        ColumnSpan = 1,
                                        RowSpan = 1
                                    });
                                }
                            }
                        }
                    }
                }
            }
            return nodes;
        }

        public override int RowIndex
        {
            get
            {
                if (_location != null)
                {
                    return _location.RowIndex;
                }
                HeaderRegion header = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                return _table.Location.RowIndex + (header != null ? header.RowCount : 0);
            }
        }

        public override int ColumnIndex
        {
            get
            {
                if (_location != null)
                {
                    return _location.ColIndex;
                }
                ColumnHeaderRegion header = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                return _table.Location.ColIndex + (header != null ? header.ColumnCount : 0);
            }
        }

        public override int RowCount
        {
            get
            {
                if (_location != null)
                {
                    return _location.RowCount;
                }
                ColumnHeaderRegion header = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                RowHeaderRegion rowheader = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                //三种情况：(1)存在行标题；(2)只存在列标题且指定与Body某字段关联；（3）只存在标题且与Body数据源字段名关联；（4）行列标题都不存在（不允许）
                return header != null ? header.RowCount :
                    rowheader != null ? (!string.IsNullOrEmpty(rowheader.HeaderBodyRelation.ReferecedField) ? Source.Table.Columns.Count : Source.Table.Rows.Count) : 0;
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
                RowHeaderRegion header = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                ColumnHeaderRegion colheader = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                return header != null ? header.ColumnCount :
                    colheader != null ? (!string.IsNullOrEmpty(colheader.HeaderBodyRelation.ReferecedField) ? Source.Table.Columns.Count : Source.Table.Rows.Count) : 0;
            }
        }

        public override string ToString()
        {
            return string.Format("<Region{0} />",
                " type=\"body\"" +
                (Source != null ? string.Format(" source=\"{0}.{1}\"", Source.Name, Field) : string.Empty) +
                (!string.IsNullOrEmpty(EmptyFill) ? string.Format(" emptyFill=\"{0}\"", EmptyFill) : string.Empty));
        }

        public override Region CloneEntity(ProductRule productRule, BaseEntity container)
        {
            BodyRegion newRegion = new BodyRegion(container as RegionTable)
            {
                _location = _location != null ? _location.Clone() : null,
                EmptyFill = EmptyFill,
                Field = Field,
                //Source = Source.Clone() as Source
            };
            if (Source != null)
            {
                newRegion.Source = productRule.GetSource(Source.Name);
            }
            return newRegion;
        }
    }

}
