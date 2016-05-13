using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity.Region
{
    /// <summary>
    /// 区域类型
    /// </summary>
    public enum RegionType
    {
        /// <summary>
        /// 主体区域
        /// </summary>
        Body,
        /// <summary>
        /// 行标题区域
        /// </summary>
        RowHeader,
        /// <summary>
        /// 列标题区域
        /// </summary>
        ColumnHeader,
        /// <summary>
        /// 左上角区域
        /// </summary>
        Corner
    }

    /// <summary>
    /// 填充区域
    /// Table划分四个区域：Corner、RowHeader、ColumnHeader、Body
    /// </summary>
    public abstract class Region : IRuleEntity, ICloneable<Region>
    {
        protected RegionTable _table;
        protected Location _location = null;
        protected List<OutputNode> _nodes = null;

        public Source Source;
        public string Field;
        public string EmptyFill = string.Empty;

        /// <summary>
        /// 获取位置
        /// </summary>
        /// <param name="reset">重新计算参数：true，需要重新计算；false不重新计算</param>
        public Location GetLocation(bool reset = false)
        {
            if (reset)
            {
                _location = null;
            }
            if (_location == null)
            {
                _location = new Location()
                {
                    ColIndex = ColumnIndex,
                    RowIndex = RowIndex,
                    ColCount = ColumnCount,
                    RowCount = RowCount
                };
            }
            return new Location(_location);
        }

        /// <summary>
        /// 获取输出结点列表
        /// </summary>
        /// <param name="reset">重新计算参数：true，需要重新计算；false不重新计算</param>
        /// <returns>结点列表</returns>
        public List<OutputNode> GetNodes(bool reset = false)
        {
            GetLocation(reset);
            if (reset)
            {
                _nodes = null;
            }
            if (_nodes == null)
            {
                _nodes = CalulateNodes();
            }
            return _nodes;
        }

        /// <summary>
        /// 计算输出结点列表
        /// </summary>
        /// <returns>结点列表</returns>
        protected abstract List<OutputNode> CalulateNodes();
        /// <summary>
        /// 获取行索引
        /// </summary>
        public abstract int RowIndex { get; }
        /// <summary>
        /// 获取列索引
        /// </summary>
        public abstract int ColumnIndex { get; }
        /// <summary>
        /// 获取区域行数
        /// </summary>
        public abstract int RowCount { get; }
        /// <summary>
        /// 获取区域列数
        /// </summary>
        public abstract int ColumnCount { get; }

        //public abstract object Clone();

        public abstract Region CloneEntity(ProductRule productRule, BaseEntity container);
    }

}
