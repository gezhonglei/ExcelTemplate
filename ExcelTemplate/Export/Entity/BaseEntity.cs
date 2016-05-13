using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 抽象容器
    /// 目的：抽提出规则实体模型必需元素，包括实体必需元素（大小、填充、模析行数）、实体间关系的必需元素（根实体、上级实体）
    /// </summary>
    public abstract class BaseEntity : IRuleEntity
    {
        protected ProductRule _productRule;
        protected BaseEntity _container;
        protected Location _location;
        protected bool _copyFill = true;
        protected int _templeteRows = 1;

        public BaseEntity(ProductRule productRule, BaseEntity container, Location location)
        {
            _productRule = productRule;
            _container = container;
            _location = location;
        }

        /// <summary>
        /// 规则根实体
        /// </summary>
        public ProductRule ProductRule
        {
            get { return _productRule; }
            set { _productRule = value; }
        }

        /// <summary>
        /// 上级实体
        /// </summary>
        public BaseEntity Container
        {
            get { return _container; }
            set { this._container = value; }
        }

        /// <summary>
        /// 填充位置或区域(规则值或原始值)
        /// </summary>
        public Location Location
        {
            get { return _location; }
            set { value = _location; }
        }

        /// <summary>
        /// 是否以Copy模式填充
        /// </summary>
        public virtual bool CopyFill
        {
            get { return _copyFill; }
            set { _copyFill = value; }
        }

        ///// <summary>
        ///// 用作模板的行数
        ///// </summary>
        //public virtual int TempleteRows { get { return _templeteRows; } set { _templeteRows = value; } }

        ///// <summary>
        ///// 实际行索引
        ///// </summary>
        //public int RowIndex
        //{
        //    get
        //    {
        //        //return _location.RowIndex + _sheet.Tables.Where(p => p.CopyFill && p._location.RowIndex < _location.RowIndex).Sum(p => p.RowCount > p.TempleteRows ? p.RowCount - p.TempleteRows : 0);
        //        return _location.RowIndex + AllIncreasedBefore;
        //    }
        //}

        ///// <summary>
        ///// 实际列索引
        ///// </summary>
        //public int ColIndex
        //{
        //    get { return _location.ColIndex; }
        //}

        ///// <summary>
        ///// 之前增长行数
        ///// </summary>
        //public int AllIncreasedBefore { get { return _container == null ? 0 : _container.IncreasedBefore(this) + _container.AllIncreasedBefore; } }

        ///// <summary>
        ///// (当前)增长行数
        ///// 注：对于容器对象需要覆盖此方法
        ///// </summary>
        //public virtual int IncreasedRows { get { return CopyFill && RowCount > TempleteRows ? RowCount - TempleteRows : 0; } }

        ///// <summary>
        ///// 同一父容器内前面对象增长行数
        ///// 注：对于容器对象需要覆盖此方法
        ///// </summary>
        ///// <param name="subObject"></param>
        ///// <returns></returns>
        //public virtual int IncreasedBefore(BaseContainer subObject) { return 0; }

        ///// <summary>
        ///// 列数
        ///// </summary>
        //public abstract int ColCount { get; }

        ///// <summary>
        ///// 行数
        ///// </summary>
        //public abstract int RowCount { get; }

        ///// <summary>
        ///// 获取输出结点
        ///// </summary>
        ///// <returns></returns>
        //public abstract IList<OutputNode> GetNodes();

        /// <summary>
        /// 获取克隆对象
        /// </summary>
        /// <returns></returns>
        public abstract BaseEntity Clone(ProductRule productRule, BaseEntity container);

        /// <summary>
        /// 批量复制BaseContainer列表对象
        /// </summary>
        /// <param name="list">列表</param>
        /// <param name="productRule">规则实体</param>
        /// <param name="container">上级实体</param>
        /// <returns></returns>
        public static List<BaseEntity> Clone(IList<BaseEntity> list, ProductRule productRule, BaseEntity container)
        {
            List<BaseEntity> newList = new List<BaseEntity>();
            for (int i = 0; i < list.Count; i++)
            {
                newList.Add(list[i].Clone(productRule, container));
            }
            return newList;
        }
    }
}
