using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Element;
using ExportTemplate.Export.Entity;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExportTemplate.Export.Writer
{
    public class ProductWriter : BaseWriterContainer
    {
        public ProductRule ProductRule { get { return base.Entity as ProductRule; } }
        private DataSet _dataSet;
        private string[] _allTempleteSheets;

        private Stream _stream;
        private IWorkbook _book;
        private bool _isXSSF = false;

        /// <summary>
        /// Excel是否2007版本
        /// 当导出规则中指定Excel模板无效时设置此参数才有效。
        /// </summary>
        public bool IsExcel2007
        {
            get { return (_book != null && _book is XSSFWorkbook) || _isXSSF; }
            set
            {
                //_book在初始化就加载Excel模板，如_book = null表示模板无效
                //_book != null表示已加载模板，或（不论有无模板）已经完成导出，都不允许修改
                if (_book == null)
                {
                    _isXSSF = value;
                }
            }
        }

        /// <summary>
        /// 构建产出物
        /// </summary>
        /// <param name="productType">产出物导出规则</param>
        public ProductWriter(ProductRule productType) : base(productType, null) { }

        #region 实例构建接口
        protected static ProductWriter CreateInstance(ProductRule productRule, bool excel2007 = false, DataSet datas = null, Action<IWorkbook, ProductRule> action = null)
        {
            ProductWriter writer = null;
            if (productRule != null)
            {
                writer = new ProductWriter(productRule);
                writer.IsExcel2007 = excel2007;
                if (action != null)
                {
                    writer.BeforeLoadAction = action;
                }
                if (datas != null)
                {
                    writer.LoadData(datas);
                }
            }
            return writer;
        }
        /// <summary>
        /// 构建ProductWriter实例
        /// </summary>
        /// <param name="productRule">Excel规则实体</param>
        /// <param name="excel2007">指定导出Excel格式（模板为空时有效）</param>
        /// <returns>ProductWriter实例</returns>
        public static ProductWriter NewInstance(ProductRule productRule, bool excel2007 = false)
        {
            return CreateInstance(productRule, excel2007);
        }
        public static ProductWriter NewInstance(ProductRule productRule, DataSet datas, Action<IWorkbook, ProductRule> action = null)
        {
            return CreateInstance(productRule, false, datas, action);
        }
        #endregion 实例构建接口

        /// <summary>
        /// 加载数据源
        /// 加载数据源的同时，实例化IWorkbook对象（确定版本类型），并根据数据源构建Writer（动态对象依赖于数据源的构建)
        /// </summary>
        /// <param name="datas">数据集</param>
        public void LoadData(DataSet datas)
        {
            this.Init();
            if (BeforeLoadAction != null)
            {
                BeforeLoadAction(_book, ProductRule);
            }

            this._dataSet = datas;
            this.ProductRule.LoadData(datas);
            if (datas != null)
            {
                Components.Clear();
                //动态对象的创建依赖于数据源
                base.CreateAllSubWriters();
            }
        }

        /// <summary>
        /// 加载数据之前的操作
        /// </summary>
        public Action<IWorkbook, ProductRule> BeforeLoadAction;

        /// <summary>
        /// 获取SheetWriter
        /// </summary>
        /// <param name="sheetName">Sheet模板名称（非动态Sheet或动态Sheet产生的模板Sheet）</param>
        /// <returns></returns>
        public SheetWriter GetSheetWriter(string sheetName)
        {
            SheetWriter tmpWriter = null;
            foreach (var writer in Components)
            {
                tmpWriter = writer as SheetWriter;
                if (tmpWriter != null && tmpWriter.Entity is Sheet)
                {
                    Sheet sheet = tmpWriter.Entity as Sheet;
                    if (sheet != null && (sheet.Name ?? "").Equals(sheetName))
                    {
                        break;
                    }
                }
            }
            return tmpWriter;
        }

        /// <summary>
        /// 初始化IWorkbook对象
        /// </summary>
        private void Init()
        {
            if (File.Exists(this.ProductRule.Template))
            {
                try
                {
                    using (Stream fStream = new FileStream(this.ProductRule.Template, FileMode.Open, FileAccess.Read))
                    {
                        //IWorkbook book = isOld ? (IWorkbook)new HSSFWorkbook(fStream) : new XSSFWorkbook(fStream);
                        _book = WorkbookFactory.Create(fStream);
                        _isXSSF = _book is XSSFWorkbook;
                    }
                }
                catch { }
            }
        }

        protected override void CreatingSubWriters()
        {
            //如未指定模板Excel或模板Excel
            if (_book == null)
            {
                _book = _isXSSF ? (IWorkbook)new XSSFWorkbook() : new HSSFWorkbook(); // WorkbookFactory.Create(new MemoryStream());
            }
            //记录模板的Sheet
            List<string> sheetnames = new List<string>();
            for (int i = 0; i < _book.NumberOfSheets; i++)
            {
                sheetnames.Add(_book.GetSheetName(i));
            }
            _allTempleteSheets = sheetnames.ToArray();

            foreach (var sheet in ProductRule.Sheets)
            {
                ISheet exSheet = _book.GetSheet(sheet.Name) ?? _book.CreateSheet(sheet.Name);
                SheetWriter writer = new SheetWriter(exSheet, sheet, this);
                if (!sheet.IsDynamic)
                {
                    Components.Add(writer);
                }
                else
                {
                    exSheet.IsSelected = false;
                    foreach (var dSheet in writer.GetDynamics())
                    {
                        ISheet newSheet = exSheet.CopySheet(dSheet.NameRule);
                        Components.Add(new SheetWriter(newSheet, dSheet, this));
                    }
                }
            }
        }

        protected override void EndWrite(WriteEventArgs args)
        {
            if (_book == null) return;

            RemoveDynamicSheet(_book);
            ShrinkSheet(_book);
            _book.Write(_stream);
        }

        /// <summary>
        /// 根据配置移除无用的Sheet
        /// </summary>
        /// <param name="book">WorkBook对象</param>
        protected void ShrinkSheet(IWorkbook book)
        {
            if (ProductRule.ShrinkSheet)
            {
                string[] usedlessSheets = GetUnusedSheets();
                if (ProductRule.ShrinkExSheets != null)
                {
                    usedlessSheets = usedlessSheets.Except(this.ProductRule.ShrinkExSheets).ToArray();
                }
                for (int i = 0; i < usedlessSheets.Length; i++)
                {
                    int sheetIndex = book.GetSheetIndex(usedlessSheets[i]);
                    if (sheetIndex > -1)
                    {
                        book.RemoveSheetAt(sheetIndex);
                    }
                }
            }
        }

        /// <summary>
        /// 获取未使用的Sheet名称
        /// </summary>
        /// <returns></returns>
        protected string[] GetUnusedSheets()
        {
            string[] usedSheets = ProductRule.Sheets.Select(p => p.Name).ToArray();
            return _allTempleteSheets != null ? _allTempleteSheets.Except(usedSheets).ToArray() : new string[0];
        }

        protected void RemoveDynamicSheet(IWorkbook book)
        {
            string[] usedlessDynamicSheets = ProductRule.Sheets.Where(p => p.IsDynamic && !p.KeepTemplate).Select(p => p.Name).ToArray();
            for (int i = 0; i < usedlessDynamicSheets.Length; i++)
            {
                int sheetIndex = book.GetSheetIndex(usedlessDynamicSheets[i]);
                if (sheetIndex > -1)
                {
                    book.RemoveSheetAt(sheetIndex);
                }
            }
        }

        public void Export(Stream stream)
        {
            this._stream = stream;
            WriteEventArgs args = new WriteEventArgs() { ExSheet = null, Entity = Entity };
            this.OnWrite(this, args);
        }

        public void Export(string filename)
        {
            using (var stream = new FileStream(filename, FileMode.Create))
            {
                Export(stream);
            }
        }
    }

}
