using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export.Element
{
    public class SourceElement : Element<Source>
    {
        public new const string ELEMENT_NAME = "TableSource";
        public SourceElement(XmlElement data, ProductRuleElement productRuleElement)
        {
            this._data = data;
        }

        protected override Source CreateEntity()
        {
            if (_data == null) return null;
            _entity = new Source(_data.GetAttribute("name"));
            return _entity;
        }
    }

    public class DataListElement : Element<ListSource>
    {
        public new const string ELEMENT_NAME = "DataList";
        public DataListElement(XmlElement data, ProductRuleElement productRuleElement)
        {
            this._data = data;
        }

        protected override ListSource CreateEntity()
        {
            if (_data == null) return null;
            _entity = new ListSource(_data.GetAttribute("name"));
            _entity.Field = _data.GetAttribute("field");
            string tmpStr = _data.GetAttribute("value");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                _entity.LoadData(tmpStr.Split(','));
            }
            return _entity;
        }
    }
}
