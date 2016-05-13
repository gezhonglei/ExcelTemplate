using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using NPOI.SS.UserModel;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// Excel表单规则
    /// <!-- 说明 -->
    /// 1、继承自BaseContainer原因？它用作其它部件的父容器
    /// 2、Tables属性为什么用BaseContainer？它要兼容Table、RegionTable、DynamicArea等
    /// </summary>
    public class Sheet : BaseEntity, ICloneable<Sheet>
    {
        private List<Cell> _cells = new List<Cell>();
        private List<BaseEntity> _tables = new List<BaseEntity>();
        private string _name;
        private string _nameRule;
        private string _sourceName;
        private bool _keepTemplate = false;
        private bool _isDynamic;

        public Sheet(ProductRule productRule) : base(productRule, null, new Location("A1")) { }

        /// <summary>
        /// Sheet模板名称
        /// </summary>
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        /// <summary>
        /// 动态Sheet
        /// </summary>
        public string SourceName
        {
            get { return _sourceName; }
            set
            {
                _sourceName = value;
                if (_sourceName != null)
                {
                    this.ProductRule.RegistSource(_sourceName);
                }
            }
        }

        /// <summary>
        /// Sheet输出名称
        /// 规则(tableName.fieldname)
        /// </summary>
        public string NameRule
        {
            get { return _nameRule; }
            set
            {
                _nameRule = value;
                MatchCollection matches = new Regex(@"(?:{)(\w+\.\w+)(?:})").Matches(_nameRule);
                for (int i = 0; i < matches.Count; i++)
                {
                    string tableField = matches[i].Groups[1].Value;
                    if (tableField.Contains('.'))
                    {
                        string[] strArr = tableField.Split('.');
                        _productRule.RegistSource(strArr[0]);
                    }
                }
            }
        }

        /// <summary>
        /// 动态生成表单时是否保留模板表单（不保留在使用完时将会删除）
        /// </summary>
        public bool KeepTemplate
        {
            get { return _keepTemplate; }
            set { _keepTemplate = value; }
        }

        /// <summary>
        /// 是否动态生成
        /// </summary>
        public bool IsDynamic
        {
            get { return _isDynamic; }
            set { _isDynamic = value; }
        }

        public List<Cell> Cells { get { return _cells; } set { _cells = value; } }
        public List<BaseEntity> Tables { get { return _tables; } set { _tables = value; } }

        /// <summary>
        /// 添加Table元素
        /// </summary>
        /// <param name="tables"></param>
        public void AddTables(IEnumerable<BaseEntity> tables)
        {
            _tables.AddRange(tables);

            //保证Table顺序：填充区域扩展行会对后面的填充区域有影响，必须将Tables根据行位置排序
            _tables.Sort(new Comparison<BaseEntity>((t1, t2) =>
            {
                return t1.Location.RowIndex - t2.Location.RowIndex;
            }));
        }

        /// <summary>
        /// 添加Cell元素集
        /// </summary>
        /// <param name="cells"></param>
        public void AddCells(IEnumerable<Cell> cells)
        {
            _cells.AddRange(cells);
        }

        /// <summary>
        /// 获取非动态Sheet导出名称
        /// </summary>
        /// <returns>解析后Sheet名称</returns>
        public string GetExportName()
        {
            if (IsDynamic || string.IsNullOrEmpty(_nameRule)) return Name;

            string sheetName = this._nameRule;
            MatchCollection matches = new Regex(@"(?:{)(\w+\.\w+)(?:})").Matches(sheetName);
            for (int i = 0; i < matches.Count; i++)
            {
                string tableField = matches[i].Groups[1].Value;
                if (tableField.Contains('.'))
                {
                    string[] strArr = tableField.Split('.');
                    Source source = _productRule.GetSource(strArr[0]);
                    if (source != null && source.Contains(strArr[1]) && source.Table.Rows.Count > 0)
                    {
                        sheetName = sheetName.Replace(string.Format("{{{0}}}", tableField),
                            source.Table.Rows[0][strArr[1]].ToString());
                    }
                }
            }
            return sheetName;
        }

        public override string ToString()
        {
            return string.Format("<Sheet name=\"{0}\"{3}>\n<Cells>\n{1}\n</Cells>\n<Tables>\n{2}\n</Tables>\n</Sheet>", _name, string.Join("\n", _cells), string.Join("\n", _tables),
                (IsDynamic ? " dynamic=\"true\"" : string.Empty)
                + (IsDynamic ? string.Format(" source=\"{0}\"", /*_dynamicObject.Source.Name*/_sourceName ?? "") : string.Empty)
                + (IsDynamic ? string.Format(" nameRule=\"{0}\"", _nameRule) : string.Empty));
        }

        public override BaseEntity Clone(ProductRule productRule, BaseEntity container)
        {
            Sheet newObject = new Sheet(productRule)
            {
                _location = _location.Clone(),
                //_cells = cells,
                //_tables = tables,
                _container = _container,
                //_dynamicObject = _dynamicObject,
                //TempleteRows = TempleteRows,
                _nameRule = _nameRule,
                _sourceName = _sourceName,
                _isDynamic = _isDynamic,
                _name = _name,
                _copyFill = _copyFill,
                _keepTemplate = _keepTemplate
            };
            //if (_dynamicObject != null && _dynamicObject.Source != null)
            //{
            //    newObject.SetDynamicSource(productRule.GetSource(_dynamicObject.Source.Name));
            //}
            newObject._cells = _cells.Clone(productRule, newObject);
            newObject._tables = BaseEntity.Clone(_tables, productRule, newObject);
            return newObject;
        }

        public Sheet CloneEntity(ProductRule productRule, BaseEntity container)
        {
            return this.Clone(productRule, container) as Sheet;
        }
    }

}
