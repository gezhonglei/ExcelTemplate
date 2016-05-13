using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    public class CellElement : Element<Cell>
    {
        public new const string ELEMENT_NAME = "Cell";
        public IProductRuleGetter _upperElement;

        public CellElement(XmlElement data, IProductRuleGetter sheetElement)
        {
            _data = data;
            _upperElement = sheetElement;
        }

        protected override Cell CreateEntity()
        {
            BaseEntity container = _upperElement.CurrentEntity;
            ProductRule prodRule = container.ProductRule;

            _entity = new Cell(prodRule, container, new Location(_data.GetAttribute("location")));
            parseSource(_data.GetAttribute("source"));
            _entity.DataIndex = ParseUtil.ParseInt(_data.GetAttribute("index"), 0);
            _entity.ValueAppend = ParseUtil.ParseBoolean(_data.GetAttribute("valueAppend"), false);
            _entity.Value = ParseUtil.IfNullOrEmpty(_data.GetAttribute("value"));
            _entity.FillType = ParseUtil.ParseEnum<FillType>(_data.GetAttribute("fill"), FillType.Origin);
            return _entity;
        }

        public void parseSource(string str)
        {
            if (str != null)
            {
                _entity.SourceName = str;
                if (str.Contains('.'))
                {
                    _entity.Field = str.Split('.')[1];
                    str = str.Split('.')[0];
                    _entity.SourceName = str;
                }
                _entity.ProductRule.RegistSource(str);
                //_entity.Source = _entity.ProductRule.RegistSource(str);
            }
        }
    }

}
