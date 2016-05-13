using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Util;
using ExportTemplate.Export.Writer.CellRender;
using ExportTemplate.Export.Writer.Convertor;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 填充区域的列规则定义
    /// </summary>
    public class Field : IRuleEntity,ICloneable<Field>
    {
        /// <summary>
        /// 字段名
        /// </summary>
        public string Name;
        /// <summary>
        /// Excel中列号
        /// </summary>
        public int ColIndex = -1;
        /// <summary>
        /// 合并行数(默认1）
        /// </summary>
        public int ColSpan = 1;
        /// <summary>
        /// 字段值渲染类型(输出类型)
        /// </summary>
        public FieldType Type;
        /// <summary>
        /// 格式：FieldType=Numeric|Datetime 可继续指定输出格式
        /// </summary>
        public string Format;
        /// <summary>
        /// 批注使用数据源字段
        /// </summary>
        public string CommentColumn;
        /// <summary>
        /// 链接类型：FieldType=l
        /// </summary>
        public string LinkType;
        /// <summary>
        /// 引用字段：链接使用的数据源字段名
        /// </summary>
        public string RefColumn;
        /// <summary>
        /// 下拉框需要的数据源字段
        /// </summary>
        public ListSource DropDownListSource;
        /// <summary>
        /// 本字段所在列是否需要根据数据合并单元格
        /// </summary>
        public bool Spannable = false;
        /// <summary>
        /// 是否汇总字段
        /// </summary>
        public bool SumField = false;
        /// <summary>
        /// 空值填充字符
        /// </summary>
        public string EmptyFill = string.Empty;

        /// <summary>
        /// 所属Table
        /// </summary>
        public Table Table = null;
        /// <summary>
        /// 给Field指定转换器
        /// </summary>
        public AbstractValueSetter Convertor = null;

        public Field(Table table, string name, FieldType fieldType = FieldType.Unknown)
        {
            this.Table = table;
            this.Name = name;
            this.Type = fieldType;
        }

        //private ICellValueSetter GetConvertorByType(FieldType fieldType)
        //{
        //    switch (fieldType)
        //    {
        //        case FieldType.Text:
        //            return new TextConvertor(p => (p ?? string.Empty).ToString());
        //        case FieldType.Numeric:
        //            return new NumericConvertor(p => double.Parse(p.ToString()), this.Format);
        //        case FieldType.Datetime:
        //            return new DateTimeConvertor(p => (DateTime)p, this.Format);
        //        case FieldType.Boolean:
        //            return new BooleanConvertor(p => (bool)p);
        //        case FieldType.Picture:
        //            return new PictureConvertor(p => (byte[])p);
        //        case FieldType.Formula:
        //            return new FormulaConvertor(p => (string)p);
        //        case FieldType.Error:
        //            return new ErrorConvertor(p => (byte)p);
        //        default:
        //            return new TextConvertor(p => (p ?? string.Empty).ToString());
        //    }
        //}

        public override string ToString()
        {
            return string.Format("<Field name=\"{0}\" colIndex=\"{1}\"{2} />", this.Name, this.ColIndex,
                (this.Type != FieldType.Unknown ? string.Format(" type=\"{0}\"", Type.ToString().ToLower()) : string.Empty) +
                (!string.IsNullOrEmpty(Format) ? string.Format(" format=\"{0}\"", Format) : string.Empty) +
                (!string.IsNullOrEmpty(CommentColumn) ? string.Format(" annnotationField=\"{0}\"", CommentColumn) : string.Empty) +
                (!string.IsNullOrEmpty(LinkType) ? string.Format(" linkType=\"{0}\"", LinkType) : string.Empty) +
                (!string.IsNullOrEmpty(RefColumn) ? string.Format(" refField=\"{0}\"", this.RefColumn) : string.Empty) +
                (DropDownListSource != null ? string.Format(" dropDownListSource=\"{0}\"", DropDownListSource.Name + "." + DropDownListSource.Field) : string.Empty) +
                (Spannable ? " spanable=\"true\"" : string.Empty) +
                (ColSpan > 1 ? string.Format(" colspan=\"{0}\"", ColSpan) : string.Empty) +
                (SumField ? " sumfield=\"true\"" : string.Empty) +
                (!string.IsNullOrEmpty(EmptyFill) ? string.Format(" emptyFill=\"{0}\"", EmptyFill) : string.Empty) +
                string.Empty);
        }

        public Field CloneEntity(ProductRule productRule, BaseEntity container)
        {
            Table table = container as Table;
            Field newObject = new Field(table, Name, Type)
            {
                Name = Name,
                ColIndex = ColIndex,
                ColSpan = ColSpan,
                Spannable = Spannable,
                EmptyFill = EmptyFill,
                Format = Format,
                CommentColumn = CommentColumn,
                LinkType = LinkType,
                RefColumn = RefColumn,
                DropDownListSource = DropDownListSource,
                Convertor = Convertor,
                SumField = SumField
            };
            if (DropDownListSource != null)
            {
                newObject.DropDownListSource = productRule.GetSource(DropDownListSource.Name) as ListSource;
            }
            return newObject;
        }
    }

    /// <summary>
    /// 输出类型
    /// </summary>
    public enum FieldType
    {
        /// <summary>
        /// 无类型（未指定类型）
        /// </summary>
        Unknown,
        /// <summary>
        /// 文本类型
        /// </summary>
        Text,
        /// <summary>
        /// 富文本
        /// </summary>
        //RichText,
        /// <summary>
        /// 数字类型：小数、整数等
        /// </summary>
        Numeric,
        /// <summary>
        /// 日期
        /// </summary>
        Datetime,
        /// <summary>
        /// Bool类型
        /// </summary>
        Boolean,
        /// <summary>
        /// 图片类型, 默认支持数据类型(byte[])
        /// </summary>
        Picture,
        /// <summary>
        /// 公式
        /// </summary>
        Formula,
        /// <summary>
        /// 错误代码，Excel的ErrorCode,一般不用
        /// </summary>
        Error
    }

}
