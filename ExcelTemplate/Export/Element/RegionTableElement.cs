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
    public class RegionTableElement : Element<RegionTable>
    {
        public new const string ELEMENT_NAME = "HeaderTable";
        public SheetElement _sheetElement;

        public RegionTableElement(XmlElement data, SheetElement sheetElement)
        {
            _data = data;
            _sheetElement = sheetElement;
        }

        protected override RegionTable CreateEntity()
        {
            Sheet sheet = _sheetElement.Entity;
            ProductRule prodRule = sheet.ProductRule;

            Location tmpLocation = new Location(_data.GetAttribute("location"));
            _entity = new RegionTable(prodRule, sheet, tmpLocation);
            _entity.Freeze = ParseUtil.ParseBoolean(_data.GetAttribute("freeze"), false);

            XmlNodeList regionNodes = _sheetElement.ProductRuleElement.QuerySubNodes("Region", _data);
            foreach (XmlElement regionNode in regionNodes)
            {
                Region region = new RegionElement(regionNode, this).Entity;
                if (region != null)
                {
                    _entity.AddRegion(region);
                }
            }
            _entity.LinkRegionSource();
            return _entity;
        }
    }

}
