using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExportTemplate.Export.Entity.Region
{
    /// <summary>
    /// 标题区域
    /// </summary>
    public abstract class HeaderRegion : Region
    {
        public int MaxLevel = -1;
        public bool ColSpannable;
        public bool RowSpannable;
        public bool IsBasedOn;
        public TreeSource TreeSource;
        /// <summary>
        /// 与Body区域之间的关联
        /// </summary>
        public SourceRelation HeaderBodyRelation;
        /// <summary>
        /// 与TreeSource区域之间的关联
        /// </summary>
        public SourceRelation HeaderTreeRelation;

        /// <summary>
        /// 不让实例化
        /// </summary>
        protected HeaderRegion() { }
        protected HeaderRegion(RegionTable table)
        {
            this._table = table;
        }

        public HeaderRegion NewHeaderRegion(RegionTable table, bool isRow)
        {
            return isRow ? (HeaderRegion)new RowHeaderRegion(table) : new ColumnHeaderRegion(table);
        }

        /// <summary>
        /// 获取与Body直接关联那一级标题
        /// </summary>
        /// <returns></returns>
        public object[] GetHeaderContent()
        {
            //如果不指定最后一级标题的输出，将从Body区域的数据取
            //return IsBasedOn && _maxLevelHeader != null ? _maxLevelHeader.GetReferecedData("", HeaderBodyRelation.Field) : HeaderBodyRelation.GetReferecedData("", HeaderBodyRelation.Field);
            return Source.GetValues(Field).ToArray();
        }

        /// <summary>
        /// 获取标题级数
        /// </summary>
        /// <returns></returns>
        public int GetHeaderLevel()
        {
            if (HeaderTreeRelation == null) return 1;

            int maxDept = 0;
            foreach (var row in HeaderTreeRelation.GetReferencedRows())
            {
                /**
                 * 可能存在两个问题：
                 * 1、递归的终止条件：用null表示可能不够
                 * 2、最后一级标题的指定源：必须是与Body数据源关联的字段（缺少灵活性）
                 */
                int count = 0;
                object tmpId = row[TreeSource.IdField];
                while (tmpId != null && !(tmpId is DBNull))//中止条件？
                {
                    count++;
                    tmpId = TreeSource.GetParentId(tmpId);
                }
                //只取最大值
                maxDept = count > maxDept ? count : maxDept;
            }
            return MaxLevel <= 0 ? maxDept : (maxDept > MaxLevel ? MaxLevel : maxDept);
        }

        //private string GetIdField()
        //{
        //    return !string.IsNullOrEmpty(TreeSource.IdField) ? TreeSource.IdField : Field;
        //}

        public override string ToString()
        {
            return string.Format("<Region{0} />",
                (string.Format(" type=\"{0}\"", this is RowHeaderRegion ? "rowheader" : "columnheader")) +
                (Source != null ? string.Format(" source=\"{0}.{1}\"", Source.Name, Field) : string.Empty) +
                (!string.IsNullOrEmpty(EmptyFill) ? string.Format(" emptyFill=\"{0}\"", EmptyFill) : string.Empty) +
                (HeaderBodyRelation != null ? " " + HeaderBodyRelation.ToString("headerBodyMapping") : string.Empty) +
                (TreeSource != null ? " " + TreeSource.ToString() : string.Empty) +
                (HeaderTreeRelation != null ? " " + HeaderTreeRelation.ToString("headerTreeMapping") : string.Empty) +
                (MaxLevel > 0 ? string.Format(" maxLevel=\"{0}\"", MaxLevel) : string.Empty) +
                (ColSpannable ? " colSpannable=\"true\"" : string.Empty) +
                (RowSpannable ? " rowSpannable=\"true\"" : string.Empty) +
                (IsBasedOn ? " basedSource=\"true\"" : string.Empty));
        }

        #region 抽象接口
        protected abstract override List<OutputNode> CalulateNodes();

        public abstract override int RowIndex { get; }

        public abstract override int ColumnIndex { get; }

        public abstract override int RowCount { get; }

        public abstract override int ColumnCount { get; }

        //public abstract override object Clone();
        public abstract override Region CloneEntity(ProductRule productRule, BaseEntity container);
        #endregion 抽象接口
    }

    /// <summary>
    /// 行标题区域
    /// </summary>
    public class ColumnHeaderRegion : HeaderRegion
    {
        public ColumnHeaderRegion(RegionTable table) : base(table) { }

        protected override List<OutputNode> CalulateNodes()
        {
            List<OutputNode> nodes = new List<OutputNode>();
            if (Source.Table != null)
            {
                //1、将数据组织到DataTable临时变量tmpDataTable中
                DataTable tmpDataTable = new DataTable();
                for (int i = 0; i <= this.ColumnCount; i++)
                {
                    tmpDataTable.Columns.Add("Column" + i);
                }
                for (int i = 0; i < this.RowCount; i++)
                {
                    tmpDataTable.Rows.Add(tmpDataTable.NewRow());
                }

                for (int i = 0; i < this.RowCount; i++)
                {
                    var headerRow = Source.Table.Rows[i];
                    tmpDataTable.Rows[i][0] = headerRow[Field];
                    //用于给Body确定列号使用
                    tmpDataTable.Rows[i][ColumnCount] = headerRow[HeaderBodyRelation.Field];

                    if (HeaderTreeRelation != null)
                    {
                        var tmpRow = TreeSource.GetRow(headerRow[HeaderTreeRelation.Field]);
                        for (int j = 0; j < this.ColumnCount; j++)
                        {
                            if (tmpRow == null) break;
                            tmpDataTable.Rows[i][j] = tmpRow[TreeSource.ContentField];
                            tmpRow = TreeSource.GetParent(tmpRow[TreeSource.ParentIdField]);
                        }
                    }
                }
                //TODO: 2、根据需要对数据进行排序，达到分组、合并的目的

                //3、根据DataTable计算出输出结点
                for (int j = 0; j < ColumnCount; j++)
                {
                    for (int i = 0; i < RowCount; )
                    {
                        if (tmpDataTable.Rows[i][j] is DBNull) { i++; continue; }
                        OutputNode node = new OutputNode() { Content = tmpDataTable.Rows[i][j] };
                        if (j == 0)
                        {
                            node.Tag = tmpDataTable.Rows[i][ColumnCount];
                        }
                        //node.ColumnIndex = ColumnIndex + ColumnCount - 1 - j;
                        node.RowIndex = RowIndex + i;
                        node.ColumnSpan = ColSpannable && j + 1 != ColumnCount && (tmpDataTable.Rows[i][j + 1] is DBNull) ?
                            ColumnCount - j : 1;
                        //如果需列合并,应该向外移位
                        node.ColumnIndex = ColumnIndex + ColumnCount - node.ColumnSpan - j;
                        node.RowSpan = 1;
                        if (RowSpannable)
                        {
                            while (++i < RowCount && node.Content.Equals(tmpDataTable.Rows[i][j]))
                            {
                                node.RowSpan++;
                            }
                        }
                        else
                        {
                            i++;
                        }
                        nodes.Add(node);
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
                RowHeaderRegion rowHeader = _table.GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                return _table.Location.RowIndex + (rowHeader != null ? rowHeader.RowCount : 0);
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
                return _table.Location.ColIndex;
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
                return Source != null && Source.Table != null ? Source.Table.Rows.Count : 0;
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
                return GetHeaderLevel();
            }
        }

        public override Region CloneEntity(ProductRule productRule, BaseEntity container)
        {
            ColumnHeaderRegion newRegion = new ColumnHeaderRegion(container as RegionTable)
            {
                _location = _location,
                ColSpannable = ColSpannable,
                RowSpannable = RowSpannable,
                EmptyFill = EmptyFill,
                Field = Field,
                IsBasedOn = IsBasedOn,
                MaxLevel = MaxLevel,
                //HeaderBodyRelation = HeaderBodyRelation.Clone() as SourceRelation,
                //HeaderTreeRelation = HeaderTreeRelation.Clone() as SourceRelation,
                //Source = Source.Clone() as Source,
                //TreeSource = TreeSource.Clone() as TreeSource
            };

            if (Source != null)
            {
                newRegion.Source = productRule.GetSource(Source.Name);
            }
            if (HeaderBodyRelation != null)
            {
                newRegion.HeaderBodyRelation = HeaderBodyRelation.Clone(productRule);
            }
            if (HeaderTreeRelation != null)
            {
                newRegion.HeaderTreeRelation = HeaderTreeRelation.Clone(productRule);
            }
            if (TreeSource != null)
            {
                newRegion.TreeSource = productRule.GetSource(TreeSource.Name) as TreeSource;
            }
            return newRegion;
        }
    }

    /// <summary>
    /// 列标题区域
    /// </summary>
    public class RowHeaderRegion : HeaderRegion
    {
        public RowHeaderRegion(RegionTable table) : base(table) { }

        protected override List<OutputNode> CalulateNodes()
        {
            List<OutputNode> nodes = new List<OutputNode>();
            if (Source.Table != null)
            {
                //1、将数据组织到DataTable临时变量tmpDataTable中
                DataTable tmpDataTable = new DataTable();
                for (int i = 0; i < this.ColumnCount; i++)
                {
                    tmpDataTable.Columns.Add("Column" + i);
                }
                for (int i = 0; i <= this.RowCount; i++)
                {
                    tmpDataTable.Rows.Add(tmpDataTable.NewRow());
                }

                for (int i = 0; i < this.ColumnCount; i++)
                {
                    var headerRow = Source.Table.Rows[i];
                    tmpDataTable.Rows[0][i] = headerRow[Field];
                    //用于给Body确定行号使用
                    tmpDataTable.Rows[RowCount][i] = headerRow[HeaderBodyRelation.Field];

                    if (HeaderTreeRelation != null)
                    {
                        var tmpRow = TreeSource.GetRow(headerRow[HeaderTreeRelation.Field]);
                        for (int j = 0; j < this.RowCount; j++)
                        {
                            if (tmpRow == null) break;
                            tmpDataTable.Rows[j][i] = tmpRow[TreeSource.ContentField];
                            tmpRow = TreeSource.GetParent(tmpRow[TreeSource.ParentIdField]);
                        }
                    }
                }
                //TODO: 2、根据需要对数据进行排序，达到分组、合并的目的

                //3、根据DataTable计算出输出结点
                for (int i = 0; i < RowCount; i++)
                {
                    for (int j = 0; j < ColumnCount; )
                    {
                        object content = tmpDataTable.Rows[i][j];
                        if (content is DBNull) { j++; continue; }
                        OutputNode node = new OutputNode() { Content = content };
                        if (i == 0)
                        {
                            node.Tag = tmpDataTable.Rows[RowCount][j];
                        }
                        node.ColumnIndex = ColumnIndex + j;
                        //node.RowIndex = RowIndex + RowCount - 1 - i;
                        //父级合并
                        node.RowSpan = RowSpannable && i + 1 != RowCount && (tmpDataTable.Rows[i + 1][j] is DBNull) ?
                            RowCount - i : 1;
                        node.RowIndex = RowIndex + RowCount - node.RowSpan - i;
                        node.ColumnSpan = 1;
                        if (ColSpannable)
                        {
                            //同一级的合并
                            while (++j < ColumnCount && content.Equals(tmpDataTable.Rows[i][j]))
                            {
                                node.ColumnSpan++;
                            }
                        }
                        else
                        {
                            j++;
                        }
                        nodes.Add(node);
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
                return _table.Location.RowIndex;
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
                ColumnHeaderRegion colHeader = _table.GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                return _table.Location.ColIndex + (colHeader != null ? colHeader.ColumnCount : 0);
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
                return GetHeaderLevel();
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
                return Source != null && Source.Table != null ? Source.Table.Rows.Count : 0;
            }
        }

        public override Region CloneEntity(ProductRule productRule, BaseEntity container)
        {
            RowHeaderRegion newRegion = new RowHeaderRegion(container as RegionTable)
            {
                _location = _location != null ? _location.Clone() : null,
                ColSpannable = ColSpannable,
                RowSpannable = RowSpannable,
                EmptyFill = EmptyFill,
                Field = Field,
                IsBasedOn = IsBasedOn,
                MaxLevel = MaxLevel,
                //HeaderBodyRelation = HeaderBodyRelation.Clone() as SourceRelation,
                //HeaderTreeRelation = HeaderTreeRelation.Clone() as SourceRelation,
                //Source = Source.Clone() as Source,
                //TreeSource = TreeSource.Clone() as TreeSource
            };
            if (Source != null)
            {
                newRegion.Source = productRule.GetSource(Source.Name);
            }
            if (HeaderBodyRelation != null)
            {
                newRegion.HeaderBodyRelation = HeaderBodyRelation.Clone(productRule);
            }
            if (HeaderTreeRelation != null)
            {
                newRegion.HeaderTreeRelation = HeaderTreeRelation.Clone(productRule);
            }
            if (TreeSource != null)
            {
                newRegion.TreeSource = productRule.GetSource(TreeSource.Name) as TreeSource;
            }
            return newRegion;
        }
    }

}
