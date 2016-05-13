using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity.Region
{
    /// <summary>
    /// 区域Table
    /// </summary>
    public class RegionTable : BaseEntity//,ICloneable<RegionTable>
    {
        //public HeaderTable(Sheet sheet, Location location) : base(sheet, location) { }
        public RegionTable(ProductRule productRule, BaseEntity container, Location location) : base(productRule, container, location) { }

        private List<Region> _regions = new List<Region>();
        /// <summary>
        /// 根据Body区域的位置冻结窗口
        /// </summary>
        public bool Freeze = false;

        public void AddRegion(Region region)
        {
            _regions.Add(region);
        }

        public Region GetRegion(RegionType type)
        {
            IEnumerable<Region> tmpRegions = _regions.Where(p => type == RegionType.Body ? p is BodyRegion :
                type == RegionType.RowHeader ? p is RowHeaderRegion :
                type == RegionType.ColumnHeader ? p is ColumnHeaderRegion :
                type == RegionType.Corner ? p is CornerRegion : false);
            if (tmpRegions.Count() > 0)
            {
                return tmpRegions.First();
            }
            return null;
        }

        /// <summary>
        /// 建立主体区域与标题区域之间数据引用关系(只在解析时调用）
        /// </summary>
        internal void LinkRegionSource()
        {
            BodyRegion body = GetRegion(RegionType.Body) as BodyRegion;
            if (body != null)
            {
                HeaderRegion header = GetRegion(RegionType.ColumnHeader) as ColumnHeaderRegion;
                if (header != null)
                {
                    header.HeaderBodyRelation.ReferecedSource = body.Source;
                }
                header = GetRegion(RegionType.RowHeader) as RowHeaderRegion;
                if (header != null)
                {
                    header.HeaderBodyRelation.ReferecedSource = body.Source;
                }
            }
        }

        public int RowHeaderLevel
        {
            get
            {
                Region region = GetRegion(RegionType.RowHeader);
                return region != null ? region.RowCount : 0;
            }
        }

        public int ColumnHeaderLevel
        {
            get
            {
                Region region = GetRegion(RegionType.ColumnHeader);
                return region != null ? region.ColumnCount : 0;
            }
        }


        //public override object Clone()
        //{
        //    return new RegionTable(_productRule, _container, (Location)_location.Clone())
        //    {
        //        _regions = new List<Region>(_regions),
        //        Freeze = Freeze,
        //        TempleteRows = TempleteRows
        //    };
        //}

        public override string ToString()
        {
            return string.Format("<HeaderTable location=\"{0}\"{1}>\n{2}\n</HeaderTable>",
                _location, Freeze ? " freeze=\"true\"" : string.Empty, string.Join("\n", _regions));
        }

        public override BaseEntity Clone(ProductRule productRule, BaseEntity container)
        {
            RegionTable newRegionTable = new RegionTable(productRule, container, (Location)_location.Clone())
            {
                Freeze = Freeze,
                //TempleteRows = TempleteRows
            };
            newRegionTable._regions = _regions.Clone(productRule, newRegionTable);
            return newRegionTable;
        }
    }
}
