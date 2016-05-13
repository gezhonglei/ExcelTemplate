using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Entity.Region;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    public class RegionElement : Element<Region>
    {
        public new const string ELEMENT_NAME = "Region";
        private RegionTableElement _regionTableElement;

        public RegionElement(XmlElement data, RegionTableElement regionTableElement)
        {
            this._data = data;
            this._regionTableElement = regionTableElement;
        }

        protected override Region CreateEntity()
        {
            RegionTable table = _regionTableElement.Entity;
            ProductRule prodRule = table.ProductRule;

            string tmpStr = _data.GetAttribute("type");
            if (string.IsNullOrEmpty(tmpStr))
            {
                return null;
            }
            _entity = "corner".Equals(tmpStr.ToLower()) ? (Region)new CornerRegion(table) :
                "rowheader".Equals(tmpStr.ToLower()) ? new RowHeaderRegion(table) :
                "columnheader".Equals(tmpStr.ToLower()) ? (Region)new ColumnHeaderRegion(table) :
                new BodyRegion(table);
            tmpStr = _data.GetAttribute("source");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                string[] values = tmpStr.Split('.');
                if (values.Length > 1)
                {
                    _entity.Field = values[1];
                }
                _entity.Source = prodRule.RegistSource(values[0]);
            }
            _entity.EmptyFill = _data.GetAttribute("emptyFill");

            if (_entity is BodyRegion)
            {
                //暂无逻辑
            }
            else if (_entity is HeaderRegion)
            {
                HeaderRegion header = _entity as HeaderRegion;
                tmpStr = _data.GetAttribute("headerBodyMaping");
                header.HeaderBodyRelation = parseRelation(header.Source, tmpStr, prodRule);
                tmpStr = _data.GetAttribute("treeSource");
                header.TreeSource = parseTreeSource(tmpStr, _data.GetAttribute("treeInnerMapping"), prodRule);
                //header.IdField = element.GetAttribute("IdField");
                //header.ParentField = element.GetAttribute("parentField");
                tmpStr = _data.GetAttribute("headerTreeMapping");
                header.HeaderTreeRelation = parseRelation(header.Source, tmpStr, prodRule);
                if (header.HeaderTreeRelation != null)
                {
                    header.HeaderTreeRelation.ReferecedSource = header.TreeSource;
                    header.MaxLevel = ParseUtil.ParseInt(_data.GetAttribute("maxLevel"), -1);
                    header.ColSpannable = ParseUtil.ParseBoolean(_data.GetAttribute("colSpannable"));
                    header.RowSpannable = ParseUtil.ParseBoolean(_data.GetAttribute("rowSpannable"));
                    header.IsBasedOn = ParseUtil.ParseBoolean(_data.GetAttribute("basedSource"));
                }
            }
            else if (_entity is CornerRegion)
            {
                tmpStr = _data.GetAttribute("spanRule");
                if (!string.IsNullOrEmpty(tmpStr))
                {
                    CornerSpanRule spanRule = CornerSpanRule.None;
                    if (!Enum.TryParse(tmpStr, true, out spanRule))
                    {
                        spanRule = "row".Equals(tmpStr.ToLower()) ? CornerSpanRule.BaseOnRowHeader :
                            "column".Equals(tmpStr.ToLower()) ? CornerSpanRule.BaseOnColumnHeader :
                            "one".Equals(tmpStr.ToLower()) ? CornerSpanRule.AllInOne : CornerSpanRule.None;
                    }
                    (_entity as CornerRegion).SpanRule = spanRule;
                }
            }
            return _entity;
        }

        private TreeSource parseTreeSource(string source, string mapping, ProductRule prod)
        {
            TreeSource treeSource = null;
            if (!string.IsNullOrEmpty(source) && source.Contains('.')
                && !string.IsNullOrEmpty(mapping) && mapping.Contains(':'))
            {
                string name = source.Split('.')[0];
                Source tmpSource = prod.GetSource(name);
                //treeSource = tmpSource != null ? tmpSource as TreeSource : new TreeSource(name);
                if (tmpSource == null)
                {
                    treeSource = new TreeSource(name);
                    prod.RegistSource(treeSource);
                }
                else if (tmpSource is TreeSource)
                {
                    treeSource = tmpSource as TreeSource;
                }
                else
                {
                    //替换原有Source
                    treeSource = new TreeSource(name);
                    prod.RegistSource(treeSource);
                }
                treeSource.ContentField = source.Split('.')[1];
                treeSource.IdField = mapping.Split(':')[0];
                treeSource.ParentIdField = mapping.Split(':')[1];
            }
            return treeSource;
        }

        private SourceRelation parseRelation(Source source, string tmpStr, ProductRule prod)
        {
            if (!string.IsNullOrEmpty(tmpStr))
            {
                string[] values = tmpStr.Split(':');
                SourceRelation relation = new SourceRelation();
                relation.Source = source;
                relation.Field = values[0];
                if (values.Length > 0)
                {
                    tmpStr = values[1];
                    if (tmpStr.Contains('.'))
                    {
                        relation.ReferecedField = tmpStr.Split('.')[1];
                        tmpStr = tmpStr.Split('.')[0];
                        relation.ReferecedSource = prod.RegistSource(tmpStr);
                    }
                    else
                    {
                        relation.ReferecedField = tmpStr;
                    }
                }
                return relation;
            }
            return null;
        }

    }
}
