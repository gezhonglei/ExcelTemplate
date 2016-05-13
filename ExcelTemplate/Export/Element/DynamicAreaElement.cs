using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export.Element
{
    public class DynamicAreaElement : Element<DynamicArea>, IProductRuleGetter
    {
        public new const string ELEMENT_NAME = "DynamicArea";
        public SheetElement _sheetElement;

        public DynamicAreaElement(XmlElement data, SheetElement sheetElement)
        {
            _data = data;
            _sheetElement = sheetElement;
        }

        protected override DynamicArea CreateEntity()
        {
            Sheet sheet = _sheetElement.Entity;
            ProductRule prodRule = sheet.ProductRule;

            Location location = new Location(_data.GetAttribute("location"));

            _entity = new DynamicArea(prodRule, sheet, location);
            string tmpStr = _data.GetAttribute("source");
            _entity.SourceName = tmpStr;
            sheet.ProductRule.RegistSource(tmpStr);
            //_entity.Source = sheet.ProductRule.RegistSource(tmpStr);
            //_entity.AddCells(parseCell(_sheetElement.ProductRuleElement.QuerySubNodes("Cells/Cell", _data)));
            //_entity.AddTables(parseTable(_sheetElement.ProductRuleElement.QuerySubNodes("Tables/Table", _data)));
            _entity.AddCells(parseSubElement<CellElement, Cell>(_sheetElement.ProductRuleElement.QuerySubNodes("Cells/Cell", _data)));
            _entity.AddTables(parseSubElement<TableElement, Table>(_sheetElement.ProductRuleElement.QuerySubNodes("Tables/Table", _data)));
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

        private IEnumerable<Table> parseTable(XmlNodeList nodes)
        {
            List<Table> entities = new List<Table>();
            if (nodes == null) return entities;
            foreach (XmlElement node in nodes)
            {
                entities.Add(new TableElement(node, this).Entity);
            }
            return entities;
        }

        public ProductRuleElement ProductRuleElement
        {
            get { return _sheetElement.ProductRuleElement; }
        }

        public BaseEntity CurrentEntity
        {
            get { return base.Entity; }
        }
    }

}
