using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Entity.Region;
using ExportTemplate.Export.Util;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportTemplate.Export.Writer
{
    public class SheetWriter : BaseWriterContainer
    {
        private DynamicSource _dynamicObject;

        private ISheet _exSheet;
        /// <summary>
        /// 按规则写入开始之前添加输出结点（结点位置是相对位置，会随着前面区域填充而变化）
        /// </summary>
        public List<OutputNode> NodesAddBefore = new List<OutputNode>();
        /// <summary>
        /// 按规则写入结束时添加处理输出结点（结点位置是绝对位置）
        /// </summary>
        public List<OutputNode> NodesAddAfter = new List<OutputNode>();

        public SheetWriter(ISheet exSheet, Sheet sheet, BaseWriterContainer parent)
            : base(sheet, parent)
        {
            this._exSheet = exSheet;
            this.Entity = sheet;//(Sheet)sheet.Clone();//防止原始的Xml规则被修改[数据源未Clone]

            if (sheet.IsDynamic)
            {
                _dynamicObject = new DynamicSource(sheet.ProductRule, Entity.ProductRule.GetSource(sheet.SourceName));
                _dynamicObject.SetParam("SheetNum", p => p + 1);
                _dynamicObject.SetParam("SheetIndex", p => p);
                _dynamicObject.DistinctFunc = (dt, index) => { return parseExpression(sheet.NameRule, dt.Rows[index]); };
            }
        }
        public bool IsDynamic { get { return Sheet.IsDynamic; } }
        public Sheet Sheet
        {
            get { return (Sheet)this.Entity; }
            set { this.Entity = value; }
        }

        protected override void CreatingSubWriters()
        {
            //if (!IsDynamic)
            //{
                CreateComponent(_exSheet, Sheet);
            //}
            //else
            //{
            //    _exSheet.IsSelected = false;
            //    foreach (var dSheet in this.GetDynamics())
            //    {
            //        ISheet newSheet = _exSheet.CopySheet(dSheet.NameRule);
            //        CreateComponent(newSheet, dSheet);
            //    }
            //}
        }

        private void CreateComponent(ISheet exSheet, Sheet sheet)
        {
            if (NodesAddBefore != null && NodesAddBefore.Count > 0)
            {
                Components.Add(new NodeWriter(NodesAddBefore));
            }
            foreach (var table in sheet.Tables)
            {
                if (table is RegionTable)
                {
                    Components.Add(new RegionWriter(table as RegionTable, this));
                }
                else if (table is Table)
                {
                    Components.Add(new TableWriter(table as Table, this));
                }
                else if (table is DynamicArea)
                {
                    Components.Add(new DynamicAreaWriter(table as DynamicArea, this));
                }
            }
            //在Table之后处理的目的:让依赖于填充区域的Cell(图片)能够自由填充
            foreach (var cell in sheet.Cells)
            {
                Components.Add(new CellWriter(cell, this));
            }
            if (NodesAddAfter != null && NodesAddAfter.Count > 0)
            {
                Components.Add(new NodeWriter(NodesAddAfter));
            }
        }

        public override void PreWrite(WriteEventArgs args)
        {
            //WriteNodes(_exSheet, NodesAddAfter);
            args.ExSheet = _exSheet;
            base.PreWrite(args);
        }

        protected override void EndWrite(WriteEventArgs args)
        {
            ISheet exSheet = args.ExSheet;
            IWorkbook book = exSheet.Workbook;

            exSheet.ForceFormulaRecalculation = true;
            //删除模板
            if (IsDynamic && !Sheet.KeepTemplate)
            {
                int index = book.GetSheetIndex(exSheet);
                if (index > -1)
                    book.RemoveSheetAt(book.GetSheetIndex(exSheet));
            }
            //Sheet更名条件：非动态Sheet且指定了NameRule属性
            if (!IsDynamic && !string.IsNullOrEmpty(Sheet.NameRule))
            {
                book.SetSheetName(book.GetSheetIndex(exSheet), Sheet.GetExportName());
            }
            base.EndWrite(args);
            //WriteNodes(_exSheet, NodesAddAfter);
        }

        public override IList<OutputNode> GetNodes()
        {
            throw new NotImplementedException();
        }

        private string parseExpression(string expression, DataRow row)
        {
            foreach (DataColumn column in row.Table.Columns)
            {
                string strPart = string.Format("{{{0}}}", column.ColumnName);
                if (expression.Contains(strPart))
                {
                    expression = expression.Replace(strPart, row[column.ColumnName].ToString());
                }
            }
            return expression;
        }

        #region 遍历动态Sheet对象
        /// <summary>
        /// 获取可遍历的动态Sheet对象
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Sheet> GetDynamics()
        {
            return IsDynamic ? new SheetEnumerable(this) : null;
        }

        internal class SheetEnumerator : IEnumerator<Sheet>
        {
            protected Sheet _sheet;
            protected SheetWriter _writer;
            protected DynamicSource _dynamicObject;
            protected int _curIndex;

            public SheetEnumerator(SheetWriter writer)
            {
                _sheet = writer.Entity as Sheet;
                _writer = writer;
                _dynamicObject = writer._dynamicObject;
                _curIndex = -1;
            }

            public Sheet Current
            {
                get
                {
                    Sheet sheet = (Sheet)_sheet.Clone(_sheet.ProductRule, _sheet.Container);
                    sheet.IsDynamic = false;//已解析的Writer不是动态的
                    sheet.NameRule = _dynamicObject.GetDynamicValue(sheet.NameRule, _curIndex);

                    List<Cell> cells = new List<Cell>();
                    for (int i = 0; i < _sheet.Cells.Count; i++)
                    {
                        Cell tmpCell = _sheet.Cells[i];
                        Cell cell = tmpCell.CloneEntity(_sheet.ProductRule, _sheet);
                        if (!string.IsNullOrEmpty(cell.SourceName))
                        {
                            cell.SourceName = _dynamicObject.GetDynamicValue(cell.SourceName, _curIndex);
                            //Source source = _dynamicObject.GetDynamicSource(cell.SourceName, _curIndex);
                            //if (source != null) cell.Source = source;
                        }
                        if (!string.IsNullOrEmpty(cell.Value))
                        {
                            cell.Value = _dynamicObject.GetDynamicValue(cell.Value, _curIndex);
                        }
                        cells.Add(cell);
                    }
                    sheet.Cells = cells;

                    List<BaseEntity> tables = new List<BaseEntity>();
                    for (int i = 0; i < _sheet.Tables.Count; i++)
                    {
                        BaseEntity tmpContainer = _sheet.Tables[i];
                        BaseEntity container = tmpContainer.Clone(_sheet.ProductRule, _sheet);
                        if (container is Table)
                        {
                            Table tmpTable = (container as Table);
                            tmpTable.SourceName = _dynamicObject.GetDynamicValue(tmpTable.SourceName, _curIndex);
                            //Source tmpSource = _dynamicObject.GetDynamicSource(tmpTable.SourceName, _curIndex);
                            //if (tmpSource != null)
                            //{
                            //    tmpTable.Source = tmpSource;
                            //}
                            tables.Add(tmpTable);
                        }
                        //TODO:Sheet与DynamicArea、RegionTable不兼容
                    }
                    sheet.Tables = tables;
                    return sheet;
                }
            }

            public void Dispose() { _sheet = null; _dynamicObject = null; }

            object System.Collections.IEnumerator.Current
            {
                get { return Current; }
            }

            public bool MoveNext()
            {
                _curIndex += 1;
                return _curIndex < _dynamicObject.Count;
            }

            public void Reset()
            {
                _curIndex = -1;
            }
        }

        internal class SheetEnumerable : IEnumerable<Sheet>
        {
            protected SheetWriter _writer;
            public SheetEnumerable(SheetWriter writer)
            {
                _writer = writer;
            }

            public IEnumerator<Sheet> GetEnumerator()
            {
                return new SheetEnumerator(_writer);
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
        #endregion 遍历动态Sheet对象
    }




}
