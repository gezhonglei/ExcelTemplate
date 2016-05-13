using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Configuration;
using ExportTemplate.Export.Element;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Writer;

namespace ExportTemplate.Export
{
    /// <summary>
    /// 导出工具入口
    /// </summary>
    public class ExportMain
    {
        private static string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportConfig.xml");
        private static string configSchema = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExportConfig.xsd");
        private static ExportConfigElement _config = null;
        static ExportMain()
        {
            _config = new ExportConfigElement();
            _config.Load(configFile, configSchema);
        }

        //private static ExportConfig GetExportConfig(bool reload = false)
        //{
        //    if (reload || _config == null)
        //    {
        //        _config = new ExportConfigElement();
        //        _config.Load(configFile, configSchema);
        //    }
        //    return _config.Entity;
        //}
        public static ProductRule GetProductRule(string ruleName, bool createNew = false)
        {
            if (_config == null)
            {
                _config = new ExportConfigElement();
                _config.Load(configFile, configSchema);
            }
            return _config.GetRules(ruleName, createNew);
        }

        /// <summary>
        /// 调整Excel版本：
        /// 检查导出文件与模板文件的Excel版本一致：（用后缀名表示区分版本）
        /// ①如匹配，不处理；
        /// ②如不匹配，优先满足“导出文件路径”中指定的Excel版本，调整模板使用的Excel后缀；
        ///   如果此版本对应的Excel模板不存在，调整导出文件的后缀与模板一致。
        /// </summary>
        /// <param name="productRule">导出规则对象</param>
        /// <param name="exportfile">导出文件路径</param>
        private static void MatchExcelVersion(ProductRule productRule, ref string exportfile)
        {
            string tempFile = productRule.Template;
            string tempExt = new FileInfo(tempFile).Extension;
            string exportExt = new FileInfo(exportfile).Extension;
            if (!tempExt.ToLower().Equals(exportExt.ToLower()))
            {
                tempFile = tempFile.Replace(tempExt, exportExt);
                if (File.Exists(tempFile))
                {
                    productRule.Template = tempFile;
                }
                else
                {
                    exportfile = exportfile.Replace(exportExt, tempExt);
                }
            }
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dataSet">数据集,DataSetName="产出物名称"</param>
        /// <param name="exportfile">绝对路径;导出文件名</param>
        public static void Export(DataSet dataSet, string exportfile)
        {
            ProductRule productRule = _config.GetRules(dataSet.DataSetName, true);
            MatchExcelVersion(productRule, ref exportfile);
            ProductWriter writer = ProductWriter.NewInstance(productRule, dataSet);
            writer.Export(exportfile);
        }

        /// <summary>
        /// 导出Excel（用于Web）
        /// </summary>
        /// <param name="dataSet">数据集,DataSetName="产出物名称"</param>
        /// <param name="stream">流</param>
        public static void ExportForBS(DataSet dataSet, Stream stream)
        {
            ProductRule productRule = GetProductRule(dataSet.DataSetName, true);
            ProductWriter writer = ProductWriter.NewInstance(productRule, dataSet);
            writer.Export(stream);
        }

        /// <summary>
        /// 根据XML文件重新加载导出规则实体
        /// </summary>
        public void ReloadConfig()
        {
            _config.Load(configFile, configSchema);
        }

        /// <summary>
        /// 获取所有产出物名称列表
        /// </summary>
        /// <returns></returns>
        public static List<string> GetProductList()
        {
            return _config.Entity.ProductTypes.Select(p => p.Name).ToList();
        }

        /// <summary>
        /// 获取产出物名称与描述信息
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetProductDescription()
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (var prod in _config.Entity.ProductTypes)
            {
                dict.Add(prod.Name, prod.Description);
            }
            return dict;
        }

        public static void test()
        {
            //检查XML解析测试
            string xmlString = _config.Entity.ToString();
            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            try
            {
                xmldoc.LoadXml(xmlString);
            }
            catch (Exception e)
            {
                Console.Write(e.StackTrace);
            }
            //格式化XML
            StringBuilder sbStr = new StringBuilder();
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(new StringWriter(sbStr));
            writer.Formatting = System.Xml.Formatting.Indented;
            writer.IndentChar = '\t';
            writer.Indentation = 1;
            xmldoc.WriteTo(writer);
            Console.Write(sbStr.ToString());
        }
    }
}
