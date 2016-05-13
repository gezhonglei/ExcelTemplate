using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Entity.Region;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    public class SheetElement : Element<Sheet>, IProductRuleGetter
    {
        public new const string ELEMENT_NAME = "Sheet";
        private ProductRuleElement _productRuleElement;
        public ProductRuleElement ProductRuleElement
        {
            get { return _productRuleElement; }
        }
        public BaseEntity CurrentEntity
        {
            get { return base.Entity; }
        }

        public SheetElement(XmlElement data, ProductRuleElement productRuleElement)
        {
            this._data = data;
            this._productRuleElement = productRuleElement;
        }

        protected override Sheet CreateEntity()
        {
            _entity = new Sheet(_productRuleElement.Entity);
            _entity.Name = _data.GetAttribute("name");
            _entity.IsDynamic = ParseUtil.ParseBoolean(_data.GetAttribute("dynamic"), false);
            _entity.NameRule = _data.GetAttribute("nameRule");
            string sourceName = _data.GetAttribute("source");
            if (!string.IsNullOrEmpty(sourceName))
            {
                //_entity.SetDynamicSource(_entity.ProductRule.RegistSource(source));
                _entity.ProductRule.RegistSource(sourceName);
                _entity.SourceName = sourceName;
            }
            //_entity.AddCells(parseCell(_productRuleElement.QuerySubNodes("Cells/Cell", _data)));
            //_entity.AddTables(parseTable(_productRuleElement.QuerySubNodes("Tables/Table", _data)));
            //_entity.AddTables(parseHeaderTable(_productRuleElement.QuerySubNodes("Tables/HeaderTable", _data)));
            //_entity.AddTables(parseDynamicArea(_productRuleElement.QuerySubNodes("DynamicAreas/DynamicArea", _data)));
            _entity.AddCells(parseSubElement<CellElement, Cell>(_productRuleElement.QuerySubNodes("Cells/Cell", _data)));
            _entity.AddTables(parseSubElement<TableElement, Table>(_productRuleElement.QuerySubNodes("Tables/Table", _data)));
            _entity.AddTables(parseSubElement<RegionTableElement, RegionTable>(_productRuleElement.QuerySubNodes("Tables/HeaderTable", _data)));
            _entity.AddTables(parseSubElement<DynamicAreaElement, DynamicArea>(_productRuleElement.QuerySubNodes("DynamicAreas/DynamicArea", _data)));
            return _entity;
        }

        private IEnumerable<Cell> parseCell(XmlNodeList nodes)
        {
            List<Cell> cells = new List<Cell>();
            if (nodes == null) return cells;
            foreach (XmlElement node in nodes)
            {
                cells.Add(new CellElement(node, this).Entity);
            }
            return cells;
        }

        private IEnumerable<BaseEntity> parseTable(XmlNodeList nodes)
        {
            List<BaseEntity> entities = new List<BaseEntity>();
            if (nodes == null) return entities;
            foreach (XmlElement node in nodes)
            {
                entities.Add(new TableElement(node, this).Entity);
            }
            return entities;
        }

        private IEnumerable<BaseEntity> parseDynamicArea(XmlNodeList nodes)
        {
            List<BaseEntity> entities = new List<BaseEntity>();
            if (nodes == null) return entities;
            foreach (XmlElement node in nodes)
            {
                entities.Add(new DynamicAreaElement(node, this).Entity);
            }
            return entities;
        }

        private IEnumerable<BaseEntity> parseHeaderTable(XmlNodeList nodes)
        {
            List<RegionTable> tables = new List<RegionTable>();
            foreach (XmlElement node in nodes)
            {
                RegionTable table = new RegionTableElement(node, this).Entity;
                tables.Add(table);
            }
            return tables;
        }
    }

}
