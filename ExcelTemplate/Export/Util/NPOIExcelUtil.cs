using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExportTemplate.Export.Entity;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExportTemplate.Export.Util
{

    public class NPOIExcelUtil
    {
        /// <summary>
        /// 默认的日期格式
        /// </summary>
        protected const string DEFAULT_DATEFORMAT = "yyyy/mm/dd";

        /// <summary>
        /// 添加批注
        /// </summary>
        /// <param name="exCell">单元格</param>
        /// <param name="commentAuthor">批注作者</param>
        /// <param name="commentContent">批注内容</param>
        public static void AddComment(ICell exCell, string commentAuthor, string commentContent)
        {
            IWorkbook book = exCell.Sheet.Workbook;
            //IDrawing draw = exCell.Sheet.DrawingPatriarch ?? exCell.Sheet.CreateDrawingPatriarch();
            IDrawing draw = exCell.Sheet.CreateDrawingPatriarch();
            IComment comment = draw.CreateCellComment(draw.CreateAnchor(0, 0, 255, 255, exCell.ColumnIndex + 1, exCell.RowIndex + 1,
                exCell.ColumnIndex + 3, exCell.RowIndex + 4));
            comment.Author = commentAuthor;
            comment.String = book.GetCreationHelper().CreateRichTextString(commentContent);
            //用于解决XSSF批注不显示的问题
            comment.String.ApplyFont(book.CreateFont());
            //comment.Column = exCell.ColumnIndex;
            //comment.Row = exCell.RowIndex;
            exCell.CellComment = comment;
        }

        /// <summary>
        /// 添加超链接
        /// </summary>
        /// <param name="exCell">单元格</param>
        /// <param name="field">字段规则</param>
        /// <param name="addr">地址</param>
        public static void AddHyperLink(ICell exCell, HyperlinkType linkType, string addr)
        {
            IWorkbook book = exCell.Sheet.Workbook;
            IHyperlink link = book.GetCreationHelper().CreateHyperlink(linkType);
            link.Address = addr;
            exCell.Hyperlink = link;

            IFont linkFont = book.CreateFont();
            linkFont.Underline = FontUnderlineType.Single;
            linkFont.Color = IndexedColors.Blue.Index;
            ICellStyle linkStyle = exCell.CellStyle ?? book.CreateCellStyle();
            linkStyle.SetFont(linkFont);
            exCell.CellStyle = linkStyle;
        }

        /// <summary>
        /// 添加下拉框
        /// </summary>
        /// <param name="exSheet">Excel表单对象</param>
        /// <param name="field">字段规则</param>
        /// <param name="rowIndex">起始行号</param>
        /// <param name="rowSpan">结束行号</param>
        public static void AddDropDownList(ISheet exSheet, string[] values, int colIndex, int rowIndex, int colSpan = 1, int rowSpan = 1)
        {
            if (values.Length > 0)
            {
                IDataValidationHelper helper = exSheet.GetDataValidationHelper();
                IDataValidationConstraint dvconstraint = helper.CreateExplicitListConstraint(values);
                CellRangeAddressList rangeList = new CellRangeAddressList(rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex + colSpan - 1);
                //DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
                //exSheet.AddValidationData(new HSSFDataValidation(rangeList, constraint)); 
                exSheet.AddValidationData(helper.CreateValidation(dvconstraint, rangeList));
            }
        }

        /// <summary>
        /// 按标准配置填值
        /// </summary>
        /// <param name="exCell">Excel单元格</param>
        /// <param name="field">字段类型</param>
        /// <param name="value"></param>
        /// <returns>操作状态:是否成功</returns>
        public static bool SetCellValue(ICell exCell, Field field, object value)
        {
            ICellStyle exCellStyle = null;
            Type dataType = value.GetType();
            IWorkbook book = exCell.Sheet.Workbook;
            if (dataType == typeof(string))
            {
                if (field.Type == FieldType.Formula)
                    exCell.SetCellFormula(value as string);
                else if (field.Type == FieldType.Text)
                    exCell.SetCellValue(value as string);
            }
            else if (ParseUtil.IsNumberType(dataType) && field.Type == FieldType.Numeric)
            {
                if (!string.IsNullOrEmpty(field.Format))
                {
                    exCellStyle = exCell.CellStyle ?? book.CreateCellStyle();
                    exCellStyle.DataFormat = book.CreateDataFormat().GetFormat(field.Format);
                    exCell.CellStyle = exCellStyle;
                }
                exCell.SetCellValue(double.Parse(value.ToString()));
            }
            else if ((dataType == typeof(DateTime) || dataType == typeof(DateTime?)) && field.Type == FieldType.Datetime)
            {
                //如不给日期指定格式会显示成数字,因此强制使用默认格式:"yyyy/mm/dd"
                string tmpStr = string.IsNullOrEmpty(field.Format) ? DEFAULT_DATEFORMAT : field.Format;
                exCellStyle = exCell.CellStyle ?? book.CreateCellStyle();
                exCellStyle.DataFormat = book.CreateDataFormat().GetFormat(field.Format);
                exCell.CellStyle = exCellStyle;
                exCell.SetCellValue((DateTime)value);
            }
            else if ((dataType == typeof(bool) || dataType == typeof(bool?)) && field.Type == FieldType.Boolean)
            {
                exCell.SetCellValue((bool)value);
            }
            else if (dataType == typeof(byte[]) && field.Type == FieldType.Picture)
            {
                //IDrawing draw = exCell.Sheet.DrawingPatriarch ?? exCell.Sheet.CreateDrawingPatriarch();//XSSF未实现DrawingPatriarch接口
                IDrawing draw = exCell.Sheet.CreateDrawingPatriarch();
                IPicture pic = draw.CreatePicture(draw.CreateAnchor(0, 0, 0, 0, exCell.ColumnIndex, exCell.RowIndex,
                        exCell.ColumnIndex + field.ColSpan, exCell.RowIndex + 1),
                        book.AddPicture((byte[])value, PictureType.JPEG));
            }
            else if ((dataType == typeof(byte) || dataType == typeof(byte?)) && field.Type == FieldType.Error)
            {
                exCell.SetCellErrorValue((byte)value);
            }
            else
            {
                /**
                 * (1) 如有转换器：想办法将DataType转换成FieldType与之匹配的数据类型
                 * (2) 否则，默认用字符串类型
                 */
                return false;
            }
            return true;
        }

        /// <summary>
        /// XML未配置FieldType：根据数据源字段类型自动设置
        /// </summary>
        /// <param name="exCell">单元格</param>
        /// <param name="tmpObject">数据</param>
        /// <returns>返回操作状态</returns>
        public static bool SetCellValueByDataType(ICell exCell, object tmpObject)
        {
            /**
             * (1) XML未配置FieldType，将不用考虑与类型相关的其它属性（format)
             * (2) 特殊的数据类型，必须指定显示指定FieldType（如Picture、Formula）；否则，将值全部当作String处理
             */
            Type tmpType = tmpObject.GetType();
            if (ParseUtil.IsNumberType(tmpType))
            {
                //int->double报错: (double)tmpObject
                exCell.SetCellValue(double.Parse(tmpObject.ToString()));
            }
            else if (tmpType == typeof(DateTime) || tmpType == typeof(DateTime?))
            {
                exCell.SetCellValue((DateTime)tmpObject);
            }
            else if (tmpType == typeof(bool) || tmpType == typeof(bool?))
            {
                exCell.SetCellValue((bool)tmpObject);
            }
            else
            {
                exCell.SetCellValue(tmpObject.ToString());
            }
            return true;
        }

        /// <summary>
        /// 根据Excel单元格类型确定输出类型
        /// </summary>
        /// <param name="exCell"></param>
        /// <param name="tmpObject"></param>
        /// <returns></returns>
        public static bool SetCellValueByCellType(ICell exCell, object tmpObject)
        {
            if (tmpObject == null) return false;

            CellType cellType = exCell.CellType;
            Type tmpType = tmpObject.GetType();
            if (cellType == CellType.Numeric && ParseUtil.IsNumberType(tmpType))
            {
                //int->double报错: (double)tmpObject
                exCell.SetCellValue(double.Parse(tmpObject.ToString()));
            }
            else if (cellType == CellType.Numeric && (tmpType == typeof(DateTime) || tmpType == typeof(DateTime?)))
            {
                exCell.SetCellValue((DateTime)tmpObject);
            }
            else if (cellType == CellType.Boolean && (tmpType == typeof(bool) || tmpType == typeof(bool?)))
            {
                exCell.SetCellValue((bool)tmpObject);
            }
            else
            {
                exCell.SetCellValue(tmpObject.ToString());
            }
            return true;
        }

        /// <summary>
        /// 根据数据合并某列的多行
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="field"></param>
        /// <param name="exSheet"></param>
        /// <param name="rowIndexBase"></param>
        /// <param name="rowCount"></param>
        /// <returns></returns>
        public static bool MergeRowsByData(DataTable dataTable, Field field, ISheet exSheet, int rowIndexBase, int rowCount)
        {
            //1、删除区域中所有的合并域（否则导出Excel需要修改）
            for (int col = field.ColIndex; col < field.ColIndex + field.ColSpan; col++)
            {
                for (int row = rowIndexBase; row < rowIndexBase + rowCount; row++)
                {
                    //性能问题会受影响：NPOI没有提供Remove(CellRangeAddress)或更多接口，只能通过索引删除
                    for (int i = 0; i < exSheet.NumMergedRegions; i++)
                    {
                        CellRangeAddress region = exSheet.GetMergedRegion(i);
                        if (region.IsInRange(row, col))
                        {
                            exSheet.RemoveMergedRegion(i);
                        }
                    }
                }
            }

            //2、找出行的不同值
            List<int> indexlist = new List<int>();
            indexlist.Add(rowIndexBase);
            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                //获取相邻列值不同的序号
                if (dataTable.Rows[i][field.Name] != dataTable.Rows[i-1][field.Name])
                {
                    indexlist.Add(rowIndexBase + i);
                }
            }
            indexlist.Add(rowIndexBase + rowCount);
            for (int i = 0; i < indexlist.Count - 1; i++)
            {
                if (indexlist[i + 1] - indexlist[i] > 1 || field.ColSpan > 1)
                {
                    exSheet.AddMergedRegion(new CellRangeAddress(indexlist[i], indexlist[i + 1] - 1, field.ColIndex, field.ColIndex + field.ColSpan - 1));
                }
            }
            return true;
        }

        /// <summary>
        /// 根据数据合并某列的多行
        /// </summary>
        /// <returns></returns>
        public static bool MergeRowsByCellData(ISheet exSheet, int rowIndex, int rowCount, int colIndex, int mergedColumns = 1)
        {
            //1、删除区域中所有的合并域（否则导出Excel需要修改）
            for (int col = colIndex; col < colIndex + mergedColumns; col++)
            {
                for (int row = rowIndex; row < rowIndex + rowCount; row++)
                {
                    //性能问题会受影响：NPOI没有提供Remove(CellRangeAddress)或更多接口，只能通过索引删除
                    for (int i = 0; i < exSheet.NumMergedRegions; i++)
                    {
                        CellRangeAddress region = exSheet.GetMergedRegion(i);
                        if (region.IsInRange(row, col))
                        {
                            exSheet.RemoveMergedRegion(i);
                        }
                    }

                }
            }

            //2、找出行的不同值
            List<int> indexlist = new List<int>();
            indexlist.Add(rowIndex);
            string tmpObj = exSheet.GetRow(rowIndex).GetCell(colIndex).ToString();
            for (int i = 1; i < rowCount; i++)
            {
                //获取相邻列值不同的序号
                ICell cell = exSheet.GetRow(i).GetCell(colIndex);
                if (cell.ToString() != tmpObj)
                {
                    indexlist.Add(rowIndex + i);
                    tmpObj = cell.ToString();
                }
            }
            indexlist.Add(rowIndex + rowCount);
            for (int i = 0; i < indexlist.Count - 1; i++)
            {
                if (indexlist[i + 1] - indexlist[i] > 1 || mergedColumns > 1)
                {
                    exSheet.AddMergedRegion(new CellRangeAddress(indexlist[i], indexlist[i + 1] - 1, colIndex, colIndex + mergedColumns - 1));
                }
            }
            return true;
        }

        private static IDictionary<int, CellRangeAddress> GetRangeBetweenRows(ISheet sheet, int sourceIndex, int targetIndex)
        {
            IDictionary<int, CellRangeAddress> rangeList = new Dictionary<int, CellRangeAddress>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (region.FirstRow <= sourceIndex && sourceIndex <= region.LastRow
                    && region.FirstRow <= targetIndex && targetIndex <= region.LastRow)
                {
                    rangeList.Add(i, region.Copy());//克隆解决版本兼容问题:CellRangeAddress在HSSF中会随Copy操作移动,而在XSSF中不会
                }
            }
            return rangeList;
        }

        /// <summary>
        /// 获取单元格所属合并区域
        /// </summary>
        /// <param name="cell">单元格对象</param>
        /// <returns>合并区域</returns>
        public static CellRangeAddress GetRange(ICell cell)
        {
            ISheet sheet = cell.Row.Sheet;
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (region.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return region;
                }
            }
            return null;
        }
        
        /// <summary>
        /// 复制单元格(解决xssf处理excel2007复制行=null的bug）
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <param name="sourceIndex">原始行</param>
        /// <param name="targetIndex">目标(起始)行</param>
        /// <param name="count">复制次数</param>
        /// <param name="dataIncluded">是否复制数据</param>
        /// <returns>复制行数</returns>
        public static int CopyRow(ISheet sheet, int sourceIndex, int targetIndex, int count = 1, bool dataIncluded = true)
        {
            if (count > 0)
            {
                IDictionary<int, CellRangeAddress> rowMergeDic = GetRangeBetweenRows(sheet, sourceIndex, targetIndex);

                IRow styleRow = sheet.GetRow(sourceIndex);
                //模板无效时直接退出
                if (styleRow == null) return 0;

                //NPOI的Copy操作默认是复制数据的，如不需要复制数据，用字符串清空
                if (!dataIncluded)
                {
                    foreach (var exCell in styleRow.Cells)
                    {
                        exCell.SetCellValue("");
                    }
                }
                /***
                 * ISheet.CopyRow 只复制不插入（字体、边框、样式、数据、单元格合并，不包括数据有效性）
                 */
                if (sheet is HSSFSheet)
                {
                    if (sheet.GetRow(targetIndex) == null)
                    {
                        sheet.CreateRow(targetIndex);//达到每次Copy都插入行的目标
                    }
                    for (int indexOffset = 0; indexOffset < count; indexOffset++)
                    {
                        //已经存在的行，自动插入新行；不存在的行，直接Copy行
                        sheet.CopyRow(sourceIndex, targetIndex);
                        //XSSF的此操作会导致targetIndex变成空行
                    }
                }
                else if (sheet is XSSFSheet)
                {
                    if (targetIndex <= sheet.LastRowNum)
                    {
                        sheet.ShiftRows(targetIndex, sheet.LastRowNum, count, true, false);//目的：让目标行始终是空的
                    }
                    for (int indexOffset = 0; indexOffset < count; indexOffset++)
                    {
                        //目标行：不存在时，直接Copy；存在时，有bug！执行后目标行反而变成了空行
                        sheet.CopyRow(sourceIndex, targetIndex + indexOffset);
                    }
                }
                CopyFormula(styleRow, targetIndex, count);

                CellRangeAddress tmpRange = null;
                foreach (var keyValue in rowMergeDic)
                {
                    tmpRange = keyValue.Value;
                    tmpRange.LastRow += count;
                    sheet.RemoveMergedRegion(keyValue.Key);
                    sheet.AddMergedRegion(tmpRange);
                }
                return count;
            }
            return 0;
        }

        /// <summary>
        /// 复制单元格(最原始的复制方法）
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <param name="sourceIndex">原始行</param>
        /// <param name="targetIndex">目标(起始)行</param>
        /// <param name="count">复制次数</param>
        /// <param name="dataIncluded">是否复制数据</param>
        /// <returns>复制行数</returns>
        public static int CopyRowByOriginalWay(ISheet sheet, int sourceIndex, int targetIndex, int count = 1, bool dataIncluded = false)
        {
            if (count > 0)
            {
                ICell styleCell, cell;
                IRow row = null;
                IRow styleRow = sheet.GetRow(sourceIndex);
                //模板无效时直接退出
                if (styleRow == null) return 0;
                //目标行超出范围时处理
                if (targetIndex <= sheet.LastRowNum)
                {
                    sheet.ShiftRows(targetIndex, sheet.LastRowNum, count);
                    //模板行在目标行之后：模板行也移动了
                    sourceIndex = sourceIndex < targetIndex ? sourceIndex : sourceIndex + count;
                    styleRow = sheet.GetRow(sourceIndex);
                }
                //取出模板行中的行合并（提高性能）
                List<CellRangeAddress> rangeList = new List<CellRangeAddress>(), tmpRangeList;
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    CellRangeAddress region = sheet.GetMergedRegion(i);
                    if (region.FirstRow == styleRow.RowNum)
                    {
                        rangeList.Add(region);
                    }
                }

                for (int indexOffset = 0; indexOffset < count; indexOffset++)
                {
                    row = sheet.CreateRow(targetIndex + indexOffset);
                    for (int colIndex = styleRow.FirstCellNum; colIndex <= styleRow.LastCellNum; colIndex++)
                    {
                        styleCell = styleRow.GetCell(colIndex);
                        if (styleCell != null)
                        {
                            cell = row.CreateCell(colIndex);
                            //（1）复制Style/Comment/Hyperlink等
                            if (styleCell.CellStyle != null)
                            {
                                cell.CellStyle = styleCell.CellStyle;//指定行的Style
                            }
                            if (styleCell.CellComment != null)
                            {
                                cell.CellComment = cell.CellComment;
                            }
                            if (styleCell.Hyperlink != null)
                            {
                                cell.Hyperlink = cell.Hyperlink;
                            }
                            //（2）复制单元格行合并
                            if (styleCell.IsMergedCell)
                            {
                                tmpRangeList = rangeList.Where(p => p.FirstColumn == cell.ColumnIndex).ToList();
                                if (tmpRangeList.Count > 0)
                                {
                                    sheet.AddMergedRegion(new CellRangeAddress(cell.RowIndex, cell.RowIndex,
                                        tmpRangeList[0].FirstColumn, tmpRangeList[0].LastColumn));
                                }
                            }
                            //(3) 复制数据
                            if (dataIncluded)
                            {
                                switch (styleCell.CellType)
                                {
                                    case CellType.String:
                                        cell.SetCellValue(styleCell.StringCellValue);
                                        break;
                                    case CellType.Boolean:
                                        cell.SetCellValue(styleCell.BooleanCellValue);
                                        break;
                                    case CellType.Numeric:
                                        //cell.SetCellValue(styleCell.NumericCellValue);
                                        if (DateUtil.IsCellDateFormatted(styleCell))
                                            cell.SetCellValue(styleCell.DateCellValue);
                                        else
                                            cell.SetCellValue(styleCell.NumericCellValue);
                                        break;
                                    case CellType.Formula:
                                        //cell.SetCellFormula(styleCell.CellFormula);
                                        if (!string.IsNullOrEmpty(styleCell.CellFormula))
                                            CopyFormula(styleCell, cell);
                                        break;
                                    case CellType.Error:
                                        cell.SetCellErrorValue(styleCell.ErrorCellValue);
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
                CopyFormula(styleRow, targetIndex, count);
                return count;
            }
            return 0;
        }

        /// <summary>
        /// 同一行复制单元格
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <param name="cell"></param>
        /// <param name="targetColIndex"></param>
        /// <param name="count"></param>
        public static void CopyCell(ISheet sheet, ICell cell, int targetColIndex, int count)
        {
            if (cell == null) return;
            IRow row = cell.Row;
            for (int i = targetColIndex; i < targetColIndex + count; i++)
            {
                ICell tmpCell = row.GetCell(i);
                if (tmpCell == null)
                {
                    tmpCell = row.CreateCell(i);
                    tmpCell.CellStyle = cell.CellStyle;
                }
            }
        }

        /// <summary>
        /// 公式复制（相对位置自动计算）：从模板行复制到目标多行
        /// </summary>
        /// <param name="sourceRow">模板行</param>
        /// <param name="targetIndex">目标行</param>
        /// <param name="count">行数</param>
        public static void CopyFormula(IRow sourceRow, int targetIndex, int count)
        {
            ICell cell;
            ISheet sheet = sourceRow.Sheet;

            //注意：IRow.GetCell与Cells取出结点是有区别的！
            foreach (var styleCell in sourceRow.Cells)
            {
                if (styleCell != null && styleCell.CellType == CellType.Formula && !string.IsNullOrEmpty(styleCell.CellFormula))
                {
                    for (int i = targetIndex; i < targetIndex + count; i++)
                    {
                        cell = sheet.GetRow(i).GetCell(styleCell.ColumnIndex);
                        if (cell != null)
                        {
                            cell.SetCellFormula(FormulaCopy(styleCell.CellFormula, i - styleCell.RowIndex, 0));
                        }
                    }
                }
            }
            //for (int j = sourceRow.FirstCellNum; j <= sourceRow.LastCellNum; j++)
            //{
            ////在XSSF下GetCell获取单元格总是CellType=Blank
            //styleCell = sourceRow.GetCell(j);
            //if (styleCell != null && styleCell.CellType == CellType.Formula && !string.IsNullOrEmpty(styleCell.CellFormula))
            //{
            //    for (int i = targetIndex; i < targetIndex + count; i++)
            //    {
            //        cell = sheet.GetRow(i).GetCell(j);
            //        cell.SetCellFormula(FormulaCopy(styleCell.CellFormula, i - styleCell.RowIndex, 0));
            //    }
            //}
            //}
        }

        /// <summary>
        /// 公式复制（相对位置自动计算）：将某单元格公式复制到另一单元格
        /// </summary>
        /// <param name="sourceCell">源单元格</param>
        /// <param name="targetCell">目标单元格</param>
        public static void CopyFormula(ICell sourceCell, ICell targetCell)
        {
            targetCell.SetCellFormula(FormulaCopy(sourceCell.CellFormula, targetCell.RowIndex - sourceCell.RowIndex, targetCell.ColumnIndex - sourceCell.ColumnIndex));
        }

        /// <summary>
        /// 自适用列宽
        /// </summary>
        /// <param name="sheet">表单Sheet对象</param>
        /// <param name="colIndex">列号</param>
        /// <param name="rowIndex">行号(可选),-1表示所有行</param>
        public static void AutoFitColumnWidth(ISheet sheet, int colIndex, int beginRow = -1, int endRow = -1)
        {
            //1、HSSF支持有点问题：(1)第一列过宽第二列过窄（较严重）(2)不能根据指定范围的内容调整宽度
            //sheet.AutoSizeColumn(colIndex);

            //2、对英文的支持不太好
            int width = sheet.GetColumnWidth(colIndex) / 256;
            beginRow = beginRow < 0 ? 0 : beginRow;
            endRow = endRow < 0 ? sheet.LastRowNum : endRow;
            for (int i = beginRow; i <= endRow; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                    {
                        int length = Encoding.Default.GetBytes(cell.ToString()).Length;
                        if (width < length)
                        {
                            width = length;
                        }
                    }
                }
            }
            sheet.SetColumnWidth(colIndex, width * 256);
        }

        /// <summary>
        /// 自适应行高
        /// </summary>
        /// <param name="sheet">Sheet对象</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="beginCol">开始列</param>
        /// <param name="endCol">结束列</param>
        public static void AutoFitRowHeight(ISheet sheet, int rowIndex, int beginCol = -1, int endCol = -1)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null) return;

            float height = sheet.GetRow(rowIndex).HeightInPoints;
            int begin = beginCol < 0 ? row.FirstCellNum : beginCol;
            int end = endCol < 0 ? row.LastCellNum : endCol;
            for (int i = begin; i <= end; i++)
            {
                ICell cell = row.GetCell(i);
                if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                {
                    string[] values = cell.ToString().Split('\n');
                    //float length = sheet.DefaultRowHeightInPoints * ((!cell.CellStyle.WrapText ? 0 : values.Max(p => Encoding.Default.GetBytes(p).Length) / 60) + values.Length);
                    float length = sheet.DefaultRowHeightInPoints * (values.Max(p => Encoding.Default.GetBytes(p).Length) / 60 + values.Length);
                    height = length > height ? length : height;
                }
            }
            row.HeightInPoints = height;
        }

        /// <summary>
        /// 根据内容识别链接类型
        /// </summary>
        /// <param name="linkAddr"></param>
        /// <returns></returns>
        public static HyperlinkType GetLinkTypeByData(string linkAddr)
        {
            linkAddr = linkAddr.Trim().ToLower();
            if (linkAddr.StartsWith("mailto:"))
                return HyperlinkType.Email;
            else if (linkAddr.StartsWith("file://"))
                return HyperlinkType.File;
            else if (Uri.IsWellFormedUriString(linkAddr, UriKind.Absolute))// && new Regex(@"^[a-zA-Z]+://(\w+(-\w+)*)(\.(\w+(-\w+)*))*(\?\S*)?$").IsMatch(linkAddr)
                return HyperlinkType.Url;
            else if (linkAddr.IndexOf('!') > 0)
                return HyperlinkType.Document;
            return HyperlinkType.Unknown;
        }

        /// <summary>
        /// 冻结窗口
        /// </summary>
        /// <param name="sheet">Sheet对象</param>
        /// <param name="rowIndex">行,默认值0表示不冻结行</param>
        /// <param name="colIndex">列,默认值0表示不冻结列</param>
        public static void Freeze(ISheet sheet, int rowIndex = 0, int colIndex = 0)
        {
            rowIndex = rowIndex < 0 ? 0 : rowIndex;
            colIndex = colIndex < 0 ? 0 : colIndex;

            if (rowIndex > 0 || colIndex > 0)
            {
                sheet.CreateFreezePane(colIndex, rowIndex);
            }
        }

        /// <summary>
        /// 对公式中相对位置移动的计算
        /// </summary>
        /// <param name="formula">公式</param>
        /// <param name="rowOffset">移动行数</param>
        /// <param name="colOffset">移动列数</param>
        /// <returns>移动后的公式</returns>
        public static string FormulaCopy(string formula, int rowOffset, int colOffset)
        {
            MatchCollection matched = new Regex(@"[A-Z]+\d+", RegexOptions.IgnoreCase).Matches(formula);
            int tmpIndex = -1;
            string colNum;
            string rowNum;
            foreach (Match match in matched)
            {
                string value = match.Value;
                for (int i = 0; i < value.Length; i++)
                {
                    if ('0' <= value[i] && value[i] <= '9')
                    {
                        tmpIndex = i;
                        break;
                    }
                }
                //绝对路径不作处理
                if (value[0] == '$' && value[tmpIndex - 1] == '$')
                    continue;
                //有"$"在数字前，调整到$位置
                tmpIndex = value[tmpIndex - 1] == '$' ? tmpIndex - 1 : tmpIndex;
                colNum = value.Substring(0, tmpIndex);
                rowNum = value.Substring(tmpIndex);
                if (value[0] != '$')
                {
                    colNum = ParseUtil.ToBase26(ParseUtil.FromBase26(colNum) + colOffset);
                }
                if (value[tmpIndex] != '$')
                    rowNum = (int.Parse(rowNum) + rowOffset).ToString();
                formula = formula.Replace(value, colNum + rowNum);
            }
            return formula;
        }

        #region NPOI样式相关

        /// <summary>
        /// 单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="backColor">背景色索引</param>
        /// <param name="foreColor">前景色索引</param>
        /// <param name="hAlign">横向对齐</param>
        /// <param name="vAlign">纵向对齐</param>
        /// <param name="boderStyle">边框样式</param>
        /// <returns>单元格样式</returns>
        public static ICellStyle CreateCellStyle(IWorkbook workbook, short backColor = 0, short foreColor = 0,
            HorizontalAlignment hAlign = HorizontalAlignment.Left, VerticalAlignment vAlign = VerticalAlignment.Center,
            BorderStyle boderStyle = BorderStyle.Thin)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = boderStyle;
            style.BorderLeft = boderStyle;
            style.BorderRight = boderStyle;
            style.BorderTop = boderStyle;
            style.Alignment = hAlign;
            style.VerticalAlignment = vAlign;
            style.FillPattern = FillPattern.SolidForeground;
            if (backColor > 0) style.FillBackgroundColor = backColor;//无用？
            if (foreColor > 0) style.FillForegroundColor = foreColor;
            return style;
        }

        /// <summary>
        /// 创建字体
        /// </summary>
        /// <param name="workbook">Excel对象</param>
        /// <param name="fontName">字体名</param>
        /// <param name="color">字体颜色</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="isItalic">斜体</param>
        /// <param name="isStrikeOut">粗体</param>
        /// <param name="underline">下划线</param>
        /// <returns></returns>
        public static IFont CreateFont(IWorkbook workbook, string fontName, short color = 0, short fontSize = 10, bool isItalic = false, bool isStrikeOut = false, FontUnderlineType underline = FontUnderlineType.None)
        {
            IFont font = workbook.CreateFont();
            if (!string.IsNullOrEmpty(fontName)) font.FontName = fontName;
            if (color > 0) font.Color = color;
            font.IsItalic = isItalic;
            font.IsStrikeout = isStrikeOut;
            //不加此判断：XSSF会有问题
            if (font.Underline != FontUnderlineType.None) font.Underline = underline;
            font.FontHeightInPoints = fontSize;
            //HSSF与HSSX对此属性的实现不一致
            //font.FontHeight = fontSize; 
            return font;
        }

        /// <summary>
        /// 设置区域粗边框(目标尚不可用)
        /// </summary>
        /// <param name="sheet">表单Sheet</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="colIndex">列索引</param>
        /// <param name="rowCount">区域行数</param>
        /// <param name="colCount">区域列数</param>
        public static void SetAreaBorder(ISheet sheet, int rowIndex, int colIndex, int rowCount, int colCount)
        {
            //在XSSF下修改一个单元格的样式将修改所有单元格样式
            BorderStyle borderStyle = BorderStyle.Medium;
            //short borderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            for (int i = rowIndex; i < rowIndex + rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) row = sheet.CreateRow(rowIndex);
                for (int j = colIndex; j < colIndex + colCount; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null) cell = row.CreateCell(j);
                    if (i == rowIndex || i == rowIndex + rowCount - 1 || j == colIndex || j == colIndex + colCount - 1)
                    {
                        ICellStyle style = sheet.Workbook.CreateCellStyle();
                        style.CloneStyleFrom(cell.CellStyle);
                        //上边框
                        if (i == rowIndex)
                        {
                            style.BorderTop = borderStyle;
                        }
                        //下边框
                        if (i == rowIndex + rowCount - 1)
                        {
                            style.BorderBottom = borderStyle;
                        }
                        //左边框
                        if (j == colIndex)
                        {
                            style.BorderLeft = borderStyle;
                        }
                        //右边框
                        if (j == colIndex + colCount - 1)
                        {
                            style.BorderRight = borderStyle;
                        }
                        cell.CellStyle = style;
                    }
                }
            }
        }

        #endregion NPOI样式相关

    }
}
