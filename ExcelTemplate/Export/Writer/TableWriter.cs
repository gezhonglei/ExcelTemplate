using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;
using ExportTemplate.Export.Writer.CellRender;
using ExportTemplate.Export.Writer.Convertor;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportTemplate.Export.Writer
{
    public class TableWriter : BaseWriter
    {
        public Source Source
        {
            get
            {
                Table table = Entity as Table;
                return table.ProductRule.GetSource(table.SourceName);
            }
        }

        public TableWriter(Table entity, BaseWriterContainer parent) : base(entity, parent) { }

        protected override void Writing(object sender, WriteEventArgs args)
        {
            ISheet exSheet = args.ExSheet;
            Table table = args.Entity as Table;
            if (table == null) return;

            Source source = this.Source;
            if (source == null) return;
            DataTable dt = source.Table;
            if (dt == null) return;

            ////根据XML指定区域与数据源，计算实际可填充区域(table根据字段数与字段位置或指定区域自动获得区域）
            int rowCount = this.RowCount;//table.Location.RowCount == 0 ? dt.Rows.Count : table.Location.RowCount;
            int colCount = this.ColCount; //列数必须由XML决定（指定区域或根据字段[位置或数量]计算）
            int rowIndexBase = this.RowIndex;//table.Location.RowIndex + increasedRowCount;//XML是根据模板设置，要加上填充区域的基数
            int colIndexBase = this.ColIndex;

            IRow styleRow = exSheet.GetRow(rowIndexBase) ?? exSheet.CreateRow(rowIndexBase);
            //if (styleRow == null) return;

            //1、暂时移除区域中的合并单元格
            Dictionary<int, int> dict = ClearMergeRegion(styleRow, colIndexBase, colIndexBase + colCount - 1);
            if (table.CopyFill && rowCount > 1)
            {
                NPOIExcelUtil.CopyRow(exSheet, rowIndexBase, rowIndexBase + 1, rowCount - 1);
            }
            IList<OutputNode> nodes = this.GetNodes();
            //2、根据合并单元格调整结点区域
            for (int i = 0; i < nodes.Count; i++)
            {
                //在行区域范围内，有合并单元格且导出规则未指定合并时，需要以合并单元格为准
                if (rowIndexBase <= nodes[i].RowIndex && nodes[i].RowIndex < rowIndexBase + rowCount
                    && dict.ContainsKey(nodes[i].ColumnIndex) && nodes[i].ColumnSpan == 1)
                {
                    nodes[i].ColumnSpan = dict[nodes[i].ColumnIndex];
                }
            }
            //3、将数据写入Sheet
            NodeWriter.WriteNodes(exSheet, nodes);
        }
        /// <summary>
        /// 移除合并单元格并返回合并区域“列索引及列宽”
        /// </summary>
        /// <param name="row">行</param>
        /// <returns></returns>
        private Dictionary<int, int> ClearMergeRegion(IRow row, int fromCol, int endCol)
        {
            List<int> regions = new List<int>();
            Dictionary<int, int> dict = new Dictionary<int, int>();
            ISheet sheet = row.Sheet;
            int colCount = 0;
            for (int i = fromCol; i <= endCol; i++)
            {
                for (int j = 0; j < sheet.NumMergedRegions; j++)
                {
                    CellRangeAddress region = sheet.GetMergedRegion(j);
                    if (region.IsInRange(row.RowNum, i))
                    {
                        colCount = region.LastColumn - region.FirstColumn + 1;
                        regions.Add(j);
                        dict.Add(i, colCount);
                        i += colCount;
                        if (i > row.LastCellNum)
                            break;
                    }
                }
            }
            //避免删除过程中指针问题
            regions = regions.OrderByDescending(p => p).ToList();
            foreach (var index in regions)
            {
                sheet.RemoveMergedRegion(index);
            }
            return dict;
        }

        public override int TempleteRows { get { return (this.Entity as Table).TempleteRows; } }

        public override int RowCount
        {
            get
            {
                Table entity = Entity as Table;
                return entity.Location.RowCount == 0 && this.Source != null && this.Source.Table != null ? this.Source.Table.Rows.Count * this.TempleteRows : entity.Location.RowCount;
            }
        }
        public override int ColCount { get { return Entity.Location.ColCount; } }

        public override IList<OutputNode> GetNodes()
        {
            Table table = Entity as Table;
            //Source source = GetSource();
            int rowIndex = this.RowIndex;
            int rowCount = this.RowCount;
            List<OutputNode> nodes = new List<OutputNode>();
            //计算行号结点
            if (table.RowNumIndex > -1)
            {
                for (int i = 0; i < this.Source.Table.Rows.Count; i++)
                {
                    nodes.Add(new OutputNode()
                    {
                        Content = i + 1,
                        RowIndex = rowIndex + i,
                        ColumnIndex = table.RowNumIndex,
                        RowSpan = 1,
                        ColumnSpan = 1,
                        Convertor = new NumericConvertor(p => (int)p)
                    });
                }
            }
            //计算Field结点
            foreach (var field in table.Fields)
            {
                nodes.AddRange(GetNodes(table, field));
            }
            //计算汇总结点
            int sumRowIndex = GetSumRowIndex(table);
            if (sumRowIndex > -1)
            {
                foreach (var colIndex in table.SumColumns)
                {
                    nodes.Add(new OutputNode()
                    {
                        Content = string.Format("SUM({0}{1}:{0}{2})", ParseUtil.ToBase26(colIndex + 1), rowIndex + 1, rowIndex + rowCount),
                        RowIndex = sumRowIndex,
                        ColumnIndex = colIndex,
                        RowSpan = 1,
                        ColumnSpan = 1,
                        Convertor = new FormulaConvertor(p => p.ToString())
                    });
                }
            }
            return nodes;
        }

        /// <summary>
        /// 获取汇总行的行索引号
        /// <remarks>汇总行不会像数据填充那样增加Excel行</remarks>
        /// </summary>
        /// <returns>行号：-1表示无效值，即不存在汇总行</returns>
        public int GetSumRowIndex(Table table)
        {
            table.AdjustSumOffset();
            return table.SumLocation == LocationPolicy.Undefined ? -1 ://频率高
                table.SumLocation == LocationPolicy.Absolute ? table.SumOffset :
                table.SumLocation == LocationPolicy.Head ? this.RowIndex + (table.SumOffset < 0 ? table.SumOffset : -1) :
                table.SumLocation == LocationPolicy.Tail ?
                this.RowIndex + RowCount - 1 + (table.SumOffset > 0 ? table.SumOffset : 1) : -1;
        }

        public List<OutputNode> GetNodes(Table Table, Field field)
        {
            int rowIndex = this.RowIndex;
            int rowCount = this.RowCount;
            DropDownListRender dropDownRender = null;
            if (field.DropDownListSource != null)
            {
                string[] values = field.DropDownListSource.GetStringValues();
                if (values.Length > 0)
                {
                    dropDownRender = new DropDownListRender()
                    {
                        ValueList = values,
                        FillArea = new Location()
                        {
                            RowIndex = rowIndex,
                            RowCount = rowCount,
                            ColIndex = field.ColIndex,
                            ColCount = field.ColSpan
                        }
                    };
                }
            }
            List<OutputNode> nodes = new List<OutputNode>();
            Source source = this.Source;
            if (this.IsValid(field))
            {
                IDictionary<int, object> dict = this.GetGroupedValues(field);
                IList<int> indexes = dict.Keys.ToList();
                indexes.Add(rowCount);
                AbstractValueSetter converter = GetConvert(field);
                for (int i = 0; i < indexes.Count - 1; i++)
                {
                    OutputNode node = new OutputNode()
                    {
                        Content = dict[indexes[i]],
                        RowIndex = rowIndex + indexes[i],
                        ColumnIndex = field.ColIndex,
                        ColumnSpan = field.ColSpan,
                        RowSpan = indexes[i] < rowCount - 1 ? indexes[i + 1] - indexes[i] : 1,
                        Convertor = converter
                    };
                    nodes.Add(node);
                    //1、只需要第一行的结点设置
                    if (i == 0 && dropDownRender != null)
                    {
                        node.AddRender(dropDownRender);
                    }
                    //2、批注
                    if (!string.IsNullOrEmpty(field.CommentColumn))
                    {
                        string value = (this.GetValue(indexes[i], field.CommentColumn) ?? "").ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            CommentRender render = new CommentRender();
                            render.Author = field.Table.ProductRule.CommentAuthor;
                            render.Comment = value;
                            node.AddRender(render);
                        }
                    }
                    //3、链接
                    if (!string.IsNullOrEmpty(field.RefColumn))
                    {
                        string value = (this.GetValue(indexes[i], field.RefColumn) ?? "").ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            HyperLinkRender render = new HyperLinkRender();
                            render.Address = value;
                            render.LinkType = field.LinkType;
                            node.AddRender(render);
                        }
                    }
                }
            }
            return nodes;
        }

        private AbstractValueSetter GetConvert(Field field)
        {
            if (field.Convertor != null)
                return field.Convertor;
            else
            {
                Type type = this.GetFieldType(field.Name);
                if (type != null && type != typeof(DBNull))
                {
                    if (field.Type != FieldType.Unknown || field.Table.ProductRule.DataFirst)
                    {
                        return GetConvertorByType(type, field.Type, field.Format);
                    }
                }
            }
            return null;
        }

        private AbstractValueSetter GetConvertorByType(Type dataType, FieldType fieldType, string format = null)
        {
            if (dataType == typeof(string))
            {
                if (fieldType == FieldType.Text || fieldType == FieldType.Unknown)
                    return new TextConvertor(p => (string)p);
                else if (fieldType == FieldType.Formula)
                    return new FormulaConvertor(p => (string)p);
            }
            else if (ParseUtil.IsNumberType(dataType) && (fieldType == FieldType.Numeric || fieldType == FieldType.Unknown))
            {
                //非Double类型用强制转换会报错，因此先转换成字符串，再转换成Double
                return new NumericConvertor(p => double.Parse(p.ToString()), format);
            }
            else if ((dataType == typeof(DateTime) || dataType == typeof(DateTime?))
                && (fieldType == FieldType.Datetime || fieldType == FieldType.Unknown))
            {
                return new DateTimeConvertor(p => (DateTime)p, format);
            }
            else if ((dataType == typeof(bool) || dataType == typeof(bool?)) && (fieldType == FieldType.Boolean || fieldType == FieldType.Unknown))
            {
                return new BooleanConvertor(p => (bool)p);
            }
            else if ((dataType == typeof(byte) || dataType == typeof(byte?)) && fieldType == FieldType.Error)
            {
                return new ErrorConvertor(p => (byte)p);
            }
            else if (dataType == typeof(byte[]) && fieldType == FieldType.Picture)
            {
                return new PictureConvertor(p => (byte[])p);
            }
            return null;
        }


        public bool IsValid(Field field)
        {
            Table table = Entity as Table;
            return Source != null && Source.Contains(field.Name) &&
                //在有效范围内
                table.Location.ColIndex <= field.ColIndex && field.ColIndex < table.Location.ColIndex + table.Location.ColCount;
        }

        public object GetValue(int index, string field)
        {
            //Source source = GetSource();
            return Source != null && Source.Table != null && Source.Table.Columns.Contains(field) ? Source.Table.Rows[index][field] : null;
        }

        public IList<object> GetValues(string field, bool distinct = false)
        {
            //Source source = GetSource();
            IList<object> objects = new List<object>();
            if (Source != null)
            {
                objects = Source.GetValues(field);
            }
            if (distinct)
            {
                objects = objects.Distinct().ToList();
            }
            return objects;
        }

        /// <summary>
        /// 获取分组后的“索引-值”键值对
        /// </summary>
        /// <param name="field">字段</param>
        /// <returns>索引与值键值对</returns>
        public IDictionary<int, object> GetGroupedValues(Field field)
        {
            Table table = Entity as Table;
            //TODO: 增加分组功能时需要根据当前字段与上级字段分组关系调整分组结果
            Dictionary<int, object> dict = new Dictionary<int, object>();
            IList<object> values = GetValues(field.Name);
            IList<object> tokens = values;
            int index = table.Fields.IndexOf(field);
            if (index > -1 && index < table.GroupList.Length)
            {
                field.Spannable = false;//与分组冲突的属性
                string[] fields = new string[index + 1];
                Array.Copy(table.Fields.Select(p => p.Name).ToArray(), fields, index + 1);
                tokens = Source.GroupToken(fields);
            }
            if (values.Count > 0)
                dict.Add(0, values[0]);
            for (int i = 1; i < values.Count; i++)
            {
                //非分组也非Spannable，相邻值不相等两种满足条件
                if ((!field.Spannable && tokens == values) || !tokens[i].Equals(tokens[i - 1]))
                {
                    dict.Add(i, values[i]);
                }
            }
            return dict;
        }

        public Type GetFieldType(string field)
        {
            return Source != null && Source.Table != null && Source.Table.Columns.Contains(field) ? Source.Table.Columns[field].DataType : null;
        }

    }

}
