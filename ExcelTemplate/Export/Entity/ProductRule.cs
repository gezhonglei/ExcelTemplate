using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 产出物规则定义
    /// </summary>
    public class ProductRule : BaseEntity
    {
        private string _name;
        private string _description;
        private string _template;
        private string _export;
        private bool _shrinkSheet = false;
        private string[] _shrinkExSheets = new string[0];
                
        protected IDictionary<string, Source> _sourceDict = new Dictionary<string, Source>();
        protected List<Sheet> _sheets = new List<Sheet>();

        private string _commentAuthor = string.Empty;
        private bool _dataFirst = true;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }
        public string Template
        {
            get { return _template; }
            set { _template = value; }
        }
        public string Export
        {
            get { return _export; }
            set { _export = value; }
        }
        /// <summary>
        /// 移除无用的Sheet
        /// </summary>
        public bool ShrinkSheet
        {
            get { return _shrinkSheet; }
            set { _shrinkSheet = value; }
        }
        /// <summary>
        /// Shrink操作保留的Sheet名称列表
        /// </summary>
        public string[] ShrinkExSheets
        {
            get { return _shrinkExSheets; }
            set { _shrinkExSheets = value; }
        }

        /// <summary>
        /// 批注作者
        /// </summary>
        public string CommentAuthor
        {
            get { return _commentAuthor; }
            set { _commentAuthor = value; }
        }
        /// <summary>
        /// 填充类型以数据优先还是单元格优先（在XML未配置FieldType情况下生效）
        /// </summary>
        public bool DataFirst
        {
            get { return _dataFirst; }
            set { _dataFirst = value; }
        }

        public IList<Sheet> Sheets
        {
            get { return _sheets; }
        }

        public ProductRule() : base(null, null, default(Location)) { }

        /// <summary>
        /// 根据名称获取数据源
        /// </summary>
        /// <param name="name">数据源名称</param>
        /// <returns>数据源对象</returns>
        public Source GetSource(string name)
        {
            if (_sourceDict.ContainsKey(name))
            {
                return _sourceDict[name];
            }
            return null;
        }

        /// <summary>
        /// 添加或替换原有数据源
        /// </summary>
        /// <param name="name">数据源名称</param>
        /// <param name="source"></param>
        public void RegistSource(Source source)
        {
            if (source == null) return;
            if (_sourceDict.ContainsKey(source.Name))
            {
                _sourceDict[source.Name] = source;
            }
            else
            {
                _sourceDict.Add(source.Name, source);
            }
        }

        public void AddSheet(Sheet sheet)
        {
            if (sheet == null) return;
            _sheets.Add(sheet);
        }

        public void AddSheets(IEnumerable<Sheet> sheets)
        {
            _sheets.AddRange(sheets);
        }

        /// <summary>
        /// 注册数据源：如存在，直接返回数据源；否则，创建数据源、注册并返回数据源
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Source RegistSource(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            if (_sourceDict.ContainsKey(name))
            {
                return _sourceDict[name];
            }
            Source source = new Source(name);
            _sourceDict.Add(name, source);
            return source;
        }

        /// <summary>
        /// 添加数据源列表
        /// </summary>
        /// <param name="list">数据源列表</param>
        public void AddSources(IEnumerable<Source> list)
        {
            foreach (var source in list)
            {
                RegistSource(source);
            }
        }

        /// <summary>
        /// 加载数据源
        /// </summary>
        /// <param name="dataSet">数据集</param>
        public void LoadData(DataSet dataSet)
        {
            foreach (DataTable dataTable in dataSet.Tables)
            {
                Source source = GetSource(dataTable.TableName);
                if (source != null)
                {
                    source.Table = dataTable;
                }
                else
                {
                    RegistSource(new Source()
                    {
                        Name = dataTable.TableName,
                        Table = dataTable
                    });
                }
            }
        }

        public override string ToString()
        {
            return string.Format("<ExportProduct name=\"{0}\" templete=\"{1}\"{2}>\n<DataSource>\n{3}\n</DataSource>\n<Sheets>\n{4}\n</Sheets>\n</ExportProduct>",
                Name, Template, (ShrinkSheet ? "shrinkSheet=\"true\"" : string.Empty) +
                (ShrinkExSheets != null && ShrinkExSheets.Length > 0 ? string.Format("shrinkExSheets=\"{0}\"", string.Join(",", ShrinkExSheets)) : string.Empty),
                string.Join("\n", _sourceDict.Values.Select(p => p is ListSource ? string.Format("<DataList name=\"{0}\" value=\"{1}\" />", p.Name, string.Join(",", (p as ListSource).GetStringValues())) : string.Format("<TableSource name=\"{0}\"/>", p.Name))),
                string.Join("\n", Sheets));
        }

        public ProductRule Clone()
        {
            ProductRule newObject = new ProductRule()
            {
                Name = Name,
                Description = Description,
                Template = Template,
                Export = Export,
                _commentAuthor = _commentAuthor,
                _dataFirst = _dataFirst,
                ShrinkSheet = ShrinkSheet,
                ShrinkExSheets = ShrinkExSheets.Clone<string>().ToArray()
            };
            //1、数据源的复制
            newObject._sourceDict = _sourceDict.Clone();
            //2、父级引用是否重置？
            newObject._sheets = _sheets.Clone(newObject, null); ;

            return newObject;
        }

        public override BaseEntity Clone(ProductRule productRule, BaseEntity container)
        {
            return this.Clone();
        }
    }
}
