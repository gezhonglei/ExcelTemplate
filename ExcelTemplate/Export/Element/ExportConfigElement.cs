using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export.Element
{
    public class ExportConfigElement : Element<ExportConfig>
    {
        public new const string ELEMENT_NAME = "ExportConfig";
        private XmlDocument _xmlDoc;
        private XmlValidator _validator;
        private Dictionary<string, ProductRuleElement> _productRuleDict = new Dictionary<string, ProductRuleElement>();

        public Dictionary<string, ProductRuleElement> ProductRules
        {
            get { return _productRuleDict; }
            set { _productRuleDict = value; }
        }

        public XmlValidator Validator
        {
            get { return _validator; }
            set { _validator = value; }
        }

        public ExportConfigElement()
        {
            _xmlDoc = new XmlDocument();
            _validator = new XmlValidator(_xmlDoc);
        }

        /// <summary>
        /// 加载导出规则XML配置
        /// </summary>
        /// <param name="xmlReader">xmlReader</param>
        /// <param name="xsdfile">Schema路径</param>
        /// <param name="parseErrorExit">是否解析异常退出</param>
        public void Load(XmlReader xmlReader, string xsdfile = null, bool parseErrorExit = true)
        {
            _xmlDoc.Load(xmlReader);
            _validator.SchemaUri = xsdfile;
            _validator.ParseErrorExit = parseErrorExit;
        }
        /// <summary>
        /// 加载导出规则XML配置
        /// </summary>
        /// <param name="filename">文件路径</param>
        /// <param name="xsdfile">Schema路径</param>
        /// <param name="parseErrorExit">是否解析异常退出</param>
        public void Load(string filename, string xsdfile = null, bool parseErrorExit = true)
        {
            _xmlDoc.Load(filename);
            _validator.SchemaUri = xsdfile;
            _validator.ParseErrorExit = parseErrorExit;
        }
        /// <summary>
        /// 加载导出规则XML配置
        /// </summary>
        /// <param name="xml">xml字符串</param>
        /// <param name="xsdfile">Schema路径</param>
        /// <param name="parseErrorExit">是否解析异常退出</param>
        public void LoadXml(string xml, string xsdfile = null, bool parseErrorExit = true)
        {
            _xmlDoc.LoadXml(xml);
            _validator.SchemaUri = xsdfile;
            _validator.ParseErrorExit = parseErrorExit;
        }

        protected override ExportConfig CreateEntity()
        {
            _entity = new ExportConfig();
            XmlNodeList nodes = _validator.QuerySubNodes("//ExportProduct");
            foreach (XmlElement node in nodes)
            {
                ProductRuleElement ruleElement = new ProductRuleElement(node, this);
                _entity.ProductTypes.Add(ruleElement.Entity);
                _productRuleDict.Add(ruleElement.Entity.Name, ruleElement);
            }
            return _entity;
        }

        public ProductRule GetRules(string name,bool createNew = false)
        {
            if (_productRuleDict.ContainsKey(name))
            {
                ProductRuleElement ruleElement = _productRuleDict[name];
                /**
                 * 如果发现性能较差时，可考虑对ProductRule实现深度Clone，然后修改NewEntity的实现，
                 * 用临时变量存储创建的实体对象，每次从OriginalEntity属性中返回其Clone对象
                 */
                return createNew ? ruleElement.NewEntity : ruleElement.Entity;
            }
            return null;
        }
    }
}
