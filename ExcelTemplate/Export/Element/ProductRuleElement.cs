using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    public class ProductRuleElement : Element<ProductRule>
    {
        public new const string ELEMENT_NAME = "ExportProduct";
        ExportConfigElement _configElement;
        protected XmlValidator _validator;

        public override ProductRule NewEntity
        {
            get
            {
                lock (this)
                {
                    if (_tplEntity == null)
                    {
                        _tplEntity = CreateEntity();
                    }
                }
                _entity = _tplEntity.Clone();
                return _entity;
            }
        }

        public ProductRuleElement(XmlElement data, ExportConfigElement configElement, XmlValidator validator = null)
        {
            this._data = data;
            this._configElement = configElement;
            this._validator = validator ?? (configElement != null ? configElement.Validator : null);
        }

        protected override ProductRule CreateEntity()
        {
            //属性
            _entity = new ProductRule()
            {
                Name = _data.GetAttribute("name"),
                Description = _data.GetAttribute("description"),
                Template = _data.GetAttribute("template"),
                Export = _data.GetAttribute("export")
            };
            _entity.ShrinkSheet = ParseUtil.ParseBoolean(_data.GetAttribute("shrinkSheet"), false);
            string tmpStr = _data.GetAttribute("shrinkExSheets");
            _entity.ShrinkExSheets = string.IsNullOrEmpty(tmpStr) ? new string[0] : tmpStr.Split(',');

            //子结点:DataList
            //_entity.AddSources(parseDataList(_validator.QuerySubNodes("DataSource/DataList", _data)));
            //子结点:DataTable
            //_entity.AddSources(parseSource(_validator.QuerySubNodes("DataSource/TableSource", _data)));
            //子结点：Sheet
            //_entity.AddSheets(parseSheet(_validator.QuerySubNodes("Sheets/*", _data)));
            _entity.AddSources(parseSubElement<DataListElement, ListSource>(_validator.QuerySubNodes("DataSource/DataList", _data)));
            _entity.AddSources(parseSubElement<SourceElement, Source>(_validator.QuerySubNodes("DataSource/TableSource", _data)));
            _entity.AddSheets(parseSubElement<SheetElement, Sheet>(_validator.QuerySubNodes("Sheets/*", _data)));
            return _entity;
        }

        protected IList<Source> parseDataList(XmlNodeList nodes)
        {
            List<Source> list = new List<Source>();
            if (nodes == null) return list;

            foreach (XmlElement node in nodes)
            {
                Source source = new DataListElement(node, this).Entity;
                if (source != null)
                {
                    list.Add(source);
                }
            }
            return list;
        }

        protected IList<Source> parseSource(XmlNodeList nodes)
        {
            List<Source> list = new List<Source>();
            if (nodes == null) return list;

            foreach (XmlElement node in nodes)
            {
                Source source = new SourceElement(node, this).Entity;
                if (source != null)
                {
                    list.Add(source);
                }
            }
            return list;
        }

        protected IList<Sheet> parseSheet(XmlNodeList nodes)
        {
            List<Sheet> list = new List<Sheet>();
            if (nodes == null) return list;

            foreach (XmlElement node in nodes)
            {
                Sheet sheet = new SheetElement(node, this).Entity;
                if (sheet != null)
                {
                    list.Add(sheet);
                }
            }
            return list;
        }

        internal XmlNodeList QuerySubNodes(string path, XmlElement data = null)
        {
            return _validator.QuerySubNodes(path, data);
        }
    }

}
