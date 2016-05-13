using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    public class FieldElement : Element<Field>
    {
        public new const string ELEMENT_NAME = "Field";
        private TableElement _tableElement;

        public FieldElement(XmlElement data, TableElement tableElement)
        {
            this._data = data;
            this._tableElement = tableElement;
        }

        protected override Field CreateEntity()
        {
            Table table = _tableElement.Entity;
            ProductRule prod = table.ProductRule;

            _entity = new Field(table, _data.GetAttribute("name"));
            _entity.ColIndex = ParseUtil.ParseColumnIndex(_data.GetAttribute("colIndex"), -1);
            _entity.Type = ParseUtil.ParseEnum<FieldType>(_data.GetAttribute("type"), FieldType.Unknown);
            _entity.Format = _data.GetAttribute("format");
            _entity.CommentColumn = _data.GetAttribute("annnotationField");
            _entity.RefColumn = _data.GetAttribute("refField");
            _entity.LinkType = _data.GetAttribute("linkType");
            string tmpStr = _data.GetAttribute("dropDownListSource");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                Source tmpSource = prod.GetSource(tmpStr);
                /**
                 * dropDownListSource数据源要么引用预定义的DataList，要么指定了DataTable.Field；否则将被忽略
                 */
                if (tmpSource != null && tmpSource is ListSource)
                {
                    _entity.DropDownListSource = tmpSource as ListSource;
                }
                else if (tmpSource == null && tmpStr.Contains('.'))
                {
                    tmpSource = new ListSource(tmpStr.Split('.')[0], tmpStr.Split('.')[1]);
                    _entity.DropDownListSource = tmpSource as ListSource;
                    prod.RegistSource(tmpSource);
                }
            }
            _entity.Spannable = ParseUtil.ParseBoolean(_data.GetAttribute("spannable"), false);
            _entity.ColSpan = ParseUtil.ParseInt(_data.GetAttribute("colspan"), 1);
            _entity.SumField = ParseUtil.ParseBoolean(_data.GetAttribute("sumfield"), false);
            _entity.EmptyFill = _data.GetAttribute("emptyFill");
            return _entity;
        }
    }

}
