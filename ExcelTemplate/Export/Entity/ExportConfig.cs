using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity
{
    public class ExportConfig : IRuleEntity
    {
        private IList<ProductRule> _productTypes;

        public ExportConfig()
        {
            _productTypes = new List<ProductRule>();
        }

        /// <summary>
        /// 获取所有产出物规则
        /// </summary>
        public IList<ProductRule> ProductTypes
        {
            get { return _productTypes; }
            set { _productTypes = value; }
        }

        /// <summary>
        /// 加载数据源
        /// </summary>
        /// <param name="productType">产出物类型名称</param>
        /// <param name="dataSet">数据源</param>
        /// <returns>产出物</returns>
        public ProductRule LoadData(string productType, DataSet dataSet)
        {
            ProductRule prod = GetProductRule(productType);
            if (prod != null)
            {
                prod.LoadData(dataSet);
            }
            return prod;
        }

        /// <summary>
        /// 加载数据源（DataSet的DataSetName作为产出物规则名称）
        /// </summary>
        /// <param name="prod">产出物规则</param>
        /// <param name="dataSet">数据源</param>
        public ProductRule LoadData(DataSet dataSet)
        {
            return LoadData(dataSet.DataSetName, dataSet);
        }

        /// <summary>
        /// 根据名称获取产出物规则对象
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ProductRule GetProductRule(string name)
        {
            IEnumerable<ProductRule> prods = _productTypes.Where(p => p.Name == name);
            if (prods.Count() > 0)
                return prods.First();
            return null;
        }

        public override string ToString()
        {
            return string.Format("<ExportConfig>\n{0}\n</ExportConfig>", string.Join("\n", _productTypes));
        }
    }
}
