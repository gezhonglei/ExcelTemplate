using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Util;
using NPOI.SS.UserModel;

namespace ExportTemplate.Export.Writer.Convertor
{

    //**
    // * 1、输出类型优先级顺序：(1) 转换器优先级别最高；(2)XML配置；(3)数据源数据或Excel单元格类型
    // * 2、设计目标：
    // * (1) 程序能够连续处理单元格值设置(使用同一接口)
    // * (2) 在调用者确定输出类型时，①能够自定义转换类型逻辑；②必须转换成指定类型
    // */ 
    /// <summary>
    /// 单元格设值接口
    /// </summary>
    public abstract class AbstractValueSetter
    {
        protected abstract void SetValue(ICell cell, object source);

        public void OnSetValue(ICell cell, object source)
        {
            if (source != null && !(source is DBNull))
            {
                SetValue(cell, source);
            }
        }
    }

    /// <summary>
    /// 文本转换器
    /// </summary>
    public class TextConvertor : AbstractValueSetter
    {
        private Func<object, string> _func;
        public TextConvertor(Func<object, string> func)
        {
            this._func = func;
        }

        protected override void SetValue(ICell cell, object source)
        {
            cell.SetCellValue(_func(source));
        }
    }

    /// <summary>
    /// 暂不适用（使用时需引入NPOI）
    /// </summary>
    public class RichTextConvertor : AbstractValueSetter
    {
        private Func<object, IRichTextString> _func;
        public RichTextConvertor(Func<object, IRichTextString> func)
        {
            this._func = func;
        }
        protected override void SetValue(ICell cell, object source)
        {
            cell.SetCellValue(_func(source));
        }
    }

    /// <summary>
    /// 数字转换器
    /// </summary>
    public class NumericConvertor : AbstractValueSetter
    {
        private Func<object, double> _func;
        private string _format = null;

        /// <summary>
        /// 创建数字转换器
        /// </summary>
        /// <param name="func">转换接口</param>
        /// <param name="format">
        /// 数据格式,如货币“¥#,##0”、精确小数位“0.00”、科学计算法“0.00E+00”、百分数“0.00%”、
        /// 电话号码“000-00000000”、中文大写数字“[DbNum2][$-804]0 元”等等
        /// </param>
        public NumericConvertor(Func<object, double> func, string format = null)
        {
            this._func = func;
            this._format = format;
        }

        protected override void SetValue(ICell cell, object source)
        {
            if (!string.IsNullOrEmpty(_format))
            {
                ICellStyle cellStyle = cell.CellStyle ?? cell.Sheet.Workbook.CreateCellStyle();
                cellStyle.DataFormat = cell.Sheet.Workbook.CreateDataFormat().GetFormat(_format);
                cell.CellStyle = cellStyle;
            }
            cell.SetCellValue(_func(source));
        }
    }

    /// <summary>
    /// 时间转换器
    /// </summary>
    public class DateTimeConvertor : AbstractValueSetter
    {
        private Func<object, DateTime> _func;
        private string _format = null;

        /// <summary>
        /// 创建时间转换器
        /// </summary>
        /// <param name="func">转换接口</param>
        /// <param name="format">日期格式,如"yyyy/mm/dd hh:mm:ss","yyyy年mm月dd日"</param>
        public DateTimeConvertor(Func<object, DateTime> func, string format = "yyyy/mm/dd")
        {
            this._func = func;
            //在以单元格值类型优先输出时需要将format设置null
            this._format = /*string.IsNullOrEmpty(format) ? "yyyy/mm/dd" :*/ format;
        }
        protected override void SetValue(ICell cell, object source)
        {
            if (!string.IsNullOrEmpty(_format))
            {
                ICellStyle cellStyle = cell.CellStyle ?? cell.Sheet.Workbook.CreateCellStyle();
                cellStyle.DataFormat = cell.Sheet.Workbook.CreateDataFormat().GetFormat(_format);
                cell.CellStyle = cellStyle;
            }
            cell.SetCellValue(_func(source));
        }
    }

    /// <summary>
    /// Bool值转换器
    /// </summary>
    public class BooleanConvertor : AbstractValueSetter
    {
        private Func<object, bool> _func;
        public BooleanConvertor(Func<object, bool> func)
        {
            this._func = func;
        }

        protected override void SetValue(ICell cell, object source)
        {
            cell.SetCellValue(_func(source));
        }
    }

    /// <summary>
    /// 公式值转换器
    /// </summary>
    public class FormulaConvertor : AbstractValueSetter
    {
        private Func<object, string> _func;
        public FormulaConvertor(Func<object, string> func, bool keepFormula = true)
        {
            this._func = func;
        }
        protected override void SetValue(ICell cell, object source)
        {
            cell.SetCellFormula(_func(source));
        }
    }

    public class ErrorConvertor : AbstractValueSetter
    {
        private Func<object, byte> _func;
        public ErrorConvertor(Func<object, byte> func)
        {
            this._func = func;
        }
        protected override void SetValue(ICell cell, object source)
        {
            cell.SetCellErrorValue(_func(source));
        }
    }

    public class PictureConvertor : AbstractValueSetter
    {
        private Func<object, byte[]> _func;
        public PictureConvertor(Func<object, byte[]> func)
        {
            this._func = func;
        }
        protected override void SetValue(ICell cell, object source)
        {
            NPOI.SS.Util.CellRangeAddress region = NPOIExcelUtil.GetRange(cell);
            ISheet sheet = cell.Sheet;
            IDrawing draw = sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = region != null ?
                draw.CreateAnchor(20, 20, 0, 0, region.FirstColumn, region.FirstRow, region.LastColumn + 1, region.LastRow + 1) :
                draw.CreateAnchor(20, 20, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex + 1, cell.RowIndex + 1);
            draw.CreatePicture(anchor, sheet.Workbook.AddPicture(_func(source), PictureType.JPEG));//PNG、JPEG都没问题
        }
    }

}
