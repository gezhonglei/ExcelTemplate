using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 动态区域
    /// </summary>
    public class DynamicArea : BaseEntity//, ICloneable<DynamicArea>
    {
        private string _sourceName;
        protected List<Cell> _cells = new List<Cell>();
        protected List<Table> _tables = new List<Table>();

        public DynamicArea(ProductRule productRule, BaseEntity container, Location location)
            : base(productRule, container, location) { }

        /// <summary>
        /// 数据源规则：预留作后期实现动态Sheet与动态区域兼容
        /// </summary>
        public string SourceName
        {
            get { return _sourceName; }
            set { _sourceName = value; }
        }
        //public Source Source
        //{
        //    get { return _source; }
        //    set
        //    {
        //        _source = value;
        //        _dynamicObject.Source = value;
        //    }
        //}

        public Cell[] Cells
        {
            get
            {
                return _cells.ToArray();
            }
        }

        public Table[] Tables
        {
            get
            {
                return _tables.ToArray();
            }
        }

        public void AddCells(IEnumerable<Cell> cells)
        {
            _cells.AddRange(cells);
        }
        public void AddTables(IEnumerable<Table> tables)
        {
            _tables.AddRange(tables);
            _tables.Sort((t1, t2) => t1.Location.RowIndex - t2.Location.RowIndex);
        }

        public override string ToString()
        {
            return string.Format("<DynamicArea location=\"{0}\" source=\"{1}\">\n<Cells>\n{2}\n</Cells>\n<Tables>\n{3}\n</Tables>\n</DynamicArea>",
                _location, _sourceName, string.Join("\n", _cells), string.Join("", _tables));
        }

        public override BaseEntity Clone(ProductRule productRule, BaseEntity container)
        {
            DynamicArea newObject = new DynamicArea(productRule, container, _location.Clone())
            {
                //CopyFill = CopyFill,
                //_source = (Source)_source.Clone(),
                //_tables = tables,
                //_cells = cells,
                //_dynamicObject = _dynamicObject.Clone(productRule),
                SourceName = SourceName
            };
            //if (_source != null)
            //{
            //    newObject.Source = productRule.GetSource(_source.Name);
            //}
            newObject._tables = _tables.Clone(productRule, newObject);
            newObject._cells = _cells.Clone(productRule, newObject);
            return newObject;
        }
    }
}
