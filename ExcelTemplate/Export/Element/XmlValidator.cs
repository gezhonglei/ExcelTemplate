using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Schema;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{
    /// <summary>
    /// 验证器
    /// <para>Schema无法定义属性与属性（依存关系）、属性与元素（决定与一致性）之间的约束关系，因此需要通过程序定义这种深层关系。</para>
    /// </summary>
    public interface IValidator
    {
        /// <summary>
        /// 用于验证接口
        /// </summary>
        /// <returns>正确</returns>
        bool Validate();
    }

    /// <summary>
    /// 使用Schema验证XML
    /// </summary>
    public class XmlValidator
    {
        private string _schemaUri = null;
        private bool _parseErrorExit;
        private bool? _schamaUsed;
        private XmlDocument _xmlDoc = new XmlDocument();
        private XmlNamespaceManager _xnameManager = null;
        private const string XMLNS_PREFIX = "szly:";   /* 可与xml文件上指定命名空间前缀无关 */

        public XmlValidator(XmlDocument xmlDoc, string xsdFile = null, bool parseErrorExit = true)
        {
            this._xmlDoc = xmlDoc;
            this._schemaUri = xsdFile;
            this._parseErrorExit = parseErrorExit;
        }

        public bool SchemaUsed
        {
            get
            {
                if (!_schamaUsed.HasValue)
                {
                    checkSchema();
                }
                return _schamaUsed.Value;
            }
        }
        public string SchemaUri
        {
            get { return _schemaUri; }
            set { _schemaUri = value; }
        }
        /// <summary>
        /// 解析出错直接退出
        /// </summary>
        public bool ParseErrorExit
        {
            get { return _parseErrorExit; }
            set { _parseErrorExit = value; }
        }

        public XmlNodeList QuerySubNodes(string query, XmlElement element = null)
        {
            string xpath = SchemaUsed ? getXPath(query, XMLNS_PREFIX) : query; //string.Format(query, SchemaUsed ? XMLNS_PREFIX : "");
            if (!SchemaUsed)
            {
                if (element == null) return _xmlDoc.SelectNodes(xpath);
                return element.SelectNodes(xpath);
            }
            else
            {
                if (element == null) return _xmlDoc.SelectNodes(xpath, _xnameManager);
                return element.SelectNodes(xpath, _xnameManager);
            }
        }

        private static string getXPath(string query, string prefix)
        {
            StringBuilder sbStr = new StringBuilder(query);
            MatchCollection matches = new Regex("^[a-zA-Z]+|(?:/)([a-zA-Z]+)").Matches(query);
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                Match match = matches[i];
                string str = match.Value[0] != '/' ? match.Value : match.Groups[1].Value;
                int index = match.Index + (match.Value[0] != '/' ? 0 : 1);
                sbStr.Remove(index, str.Length).Insert(index, prefix + str);
            }
            //foreach (Match match in matches)
            //{
            //    string str = match.Groups.Count == 1 ? match.Value : match.Groups[1].Value;
            //    int index = sbStr.ToString().IndexOf(str);
            //    sbStr.Remove(index, str.Length).Insert(index, prefix + str);
            //}
            return sbStr.ToString();
        }

        /// <summary>
        /// 检查xml是否引用XSD
        /// </summary>
        private bool checkSchema()
        {
            XmlNode root = null;
            foreach (XmlNode node in _xmlDoc.ChildNodes)
            {
                if (node.NodeType == XmlNodeType.Element)
                {
                    root = node;
                    break;
                }
            }
            if (root != null && !string.IsNullOrEmpty(root.NamespaceURI))
            {
                _schamaUsed = true;
                //xmlnsUri = root.NamespaceURI;
                _xnameManager = new XmlNamespaceManager(_xmlDoc.NameTable);
                _xnameManager.AddNamespace(XMLNS_PREFIX.Split(':')[0], root.NamespaceURI);
                return true;
            }
            return false;
        }

        /// <summary>
        /// 使用Schema检查XML
        /// </summary>
        public void ValidateXml()
        {
            if (!File.Exists(_schemaUri))
            {
                Console.WriteLine("SchemaFile '{0}' not found!");
                return;
            }
            if (SchemaUsed)
            {
                string targetNamespace = ParseUtil.IfNullOrEmpty(SchemaTargetNamespace(), "http://tempuri.org/schema.xsd");
                XmlSchemaSet schemas = new XmlSchemaSet();
                schemas.Add(targetNamespace, _schemaUri);

                _xmlDoc.Schemas = schemas;
                StringBuilder sbStr = new StringBuilder();
                _xmlDoc.Validate(new ValidationEventHandler((obj, e) =>
                {
                    if (e.Severity == XmlSeverityType.Error)
                    {
                        sbStr.AppendLine(e.Message);
                    }
                    else
                    {
                        sbStr.AppendLine(e.Message);
                    }
                }));

                if (_parseErrorExit && sbStr.Length > 0)
                    throw new ParseErrorException(sbStr.ToString());
            }
        }

        /// <summary>
        /// 从Schema中提取命名空间
        /// </summary>
        /// <returns></returns>
        public string SchemaTargetNamespace()
        {
            string targetNamespace = string.Empty;
            if (!string.IsNullOrEmpty(_schemaUri))
            {
                XmlDocument schemaDoc = new XmlDocument();
                schemaDoc.Load(_schemaUri);
                foreach (XmlNode node in schemaDoc.ChildNodes)
                {
                    if (node.LocalName == "schema")
                    {
                        targetNamespace = node.Attributes["targetNamespace"].Value;
                        break;
                    }
                }
            }
            return targetNamespace;
        }

        public static void test()
        {
            string _schemaFile = @"E:\DotNet\MyTestPlatform\Winform\ExcelTemplate\ExportConfig.xsd";
            string targetNamespace = "http://tempuri.org/schema.xsd";
            XmlReader xsdReader = XmlReader.Create(_schemaFile);
            XmlDocument schemaDoc = new XmlDocument();
            //schemaDoc.Load(_schemaFile);
            schemaDoc.Load(xsdReader);
            foreach (XmlNode node in schemaDoc.ChildNodes)
            {
                if (node.LocalName == "schema")
                {
                    targetNamespace = node.Attributes["targetNamespace"].Value;
                    break;
                }
            }

            XmlSchemaSet schemas = new XmlSchemaSet();
            //与前面共用Reader会报错:System.Xml.Schema.XmlSchemaException: W3C XML 架构的根元素应为 <schema>，命名空间应为“http://www.w3.org/2001/XMLSchema”。
            schemas.Add(targetNamespace, xsdReader);
            XmlDocument _xmlDoc = new XmlDocument();
            _xmlDoc.Load(@"E:\DotNet\MyTestPlatform\Winform\ExcelTemplate\ExportConfig.xml");
            _xmlDoc.Schemas = schemas;
        }

        public static void test2()
        {
            string test = getXPath("//ExportProduct", XMLNS_PREFIX);
            Console.WriteLine(test);
            test = getXPath("DataSource/*", XMLNS_PREFIX);
            Console.WriteLine(test);
            test = getXPath("Tables/HeaderTable", XMLNS_PREFIX);
            Console.WriteLine(test); 
            test = getXPath("Tables/Table", XMLNS_PREFIX);
            Console.WriteLine(test); 
            test = getXPath("Field", XMLNS_PREFIX);
            Console.WriteLine(test);
            test = getXPath("Cells/Cell", XMLNS_PREFIX);
            Console.WriteLine(test);
        }
    }

}
