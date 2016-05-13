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
using ExportTemplate.Export;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Writer;

namespace ExportTemplate.Export
{
    /**
     * 兼容问题：
     * 【批注】
     * 1、在Office下批注不显示Author, 但在WPS下可以显示
     * 2、在Office 2007下批注随机开关, 但在WPS下不会（升级到NPOI-2.1.3.1就好了)
     * 【公式】
     * 1、HSSF与XSSF对公式处理：
     * （1）HSSF与XSSF，都是机械式地复制公式，相对引用未随复制自动变化；
     * （2）SetValue(value)不会将CellType=Formula修改其它类型,但当value=null或""时，HSSF会修改CellType=Blank
     * （3）触发计算时(打开文档)会覆盖原单元格的值
     * （4）公式取值：Workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateInCell(cell)
     */
    public class ExportTest
    {
        static string path = AppDomain.CurrentDomain.BaseDirectory;
        //public static void Test(string exportfile, string productTypeName)
        //{
        //    productTypeName = !string.IsNullOrEmpty(productTypeName) ? productTypeName : "Template";
        //    DataSet dataSet = GetDataSet(productTypeName);//GetTestData(path);
        //    dataSet.DataSetName = productTypeName;
        //    ExportRuleConfig config = ExportRuleConfig.NewInstance(Path.Combine(path, "ExcelTemplate\\ExportConfig.xml"));
        //    config.BasePath = path;
        //    ProductRule prodType = config.GetProductRule(dataSet.DataSetName);

        //    if (!string.IsNullOrEmpty(exportfile))
        //    {
        //        string name = prodType.Template;
        //        prodType.Template = name.Replace(new FileInfo(name).Extension, new FileInfo(exportfile).Extension);
        //    }

        //    string templatefile = Path.Combine(path, prodType.Template);
        //    exportfile = string.IsNullOrEmpty(exportfile) ? Path.Combine(path, "test" + new FileInfo(templatefile).Extension) : exportfile;
        //    NPOIExcelProduct product = new NPOIExcelProduct(prodType, dataSet);
        //    product.Export(exportfile);
        //}

        public static void Test2(string productTypeName, string exportfile)
        {
            DataSet dataSet = GetDataSet(productTypeName);
            dataSet.DataSetName = productTypeName;
            ExportMain.Export(dataSet, exportfile);
        }

        private static DataSet GetDataSet(string type)
        {
            if (type == "FunctionList")
                return GetFunctionList();
            else if (type == "BusinessMainFlow")
                return GetBusinessMainFlow();
            else if (type == "BusinessCategory")
                return GetBusinessCategory();
            else if (type == "Organization")
                return GetOrganization();
            else if (type == "Report")
                return GetReport();
            else if (type == "3ConfigReq")
                return Get3ConfigReq();
            else if (type == "4Line3Config")
                return Get4Line3Config();
            return GetTestData();
        }

        private static DataSet GetTestData()
        {
            DataSet dataSet = new DataSet("Form");
            DataTable dt = new DataTable("Cells");
            dt.Columns.Add("Title");
            dt.Columns.Add("FormName");
            dt.Columns.Add("FormCode");
            dt.Columns.Add("Remark");
            dt.Columns.Add("Constrain");
            dt.Columns.Add("FormStyle", typeof(byte[]));
            dt.Rows.Add("来店登记统计表", "来店登记", "from_0001", "(特殊情况说明)", "(约束说明)");

            using (FileStream filestream = new FileStream(Path.Combine(path, "ExcelTemplate\\FormStyle.jpg"), FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[(int)filestream.Length];
                filestream.Read(bytes, 0, bytes.Length);
                dt.Rows[0]["FormStyle"] = bytes;
            }
            //IImageBss test = ContainerContext.Container.Resolve<IImageBss>();
            //DataTable imgTable = test.GetFormImage("Form_20150430170001558384");
            //byte[] bytes = imgTable == null || imgTable.Rows.Count == 0 ? new byte[0] : (byte[])imgTable.Rows[0]["IMG_CONTENT"];
            //dt.Rows.Add(bytes);
            dataSet.Tables.Add(dt);

            dt = new DataTable("FromRule");
            dt.Columns.Add("Category1");
            dt.Columns.Add("Category2");
            dt.Columns.Add("Item");
            dt.Columns.Add("Save");
            dt.Columns.Add("Way");
            dt.Columns.Add("Source");
            dt.Columns.Add("Rule");
            dt.Columns.Add("Algorithm");
            dt.Columns.Add("Annotation");//用于批注
            dt.Columns.Add("LinkAddr");//用于链接
            dt.Rows.Add("来店管理", "来店登记", "日期", "√", null, null, null, null, "填写日期", "http://baidu.com");
            dt.Rows.Add("来店管理", "来店登记", "销售顾问", "√", "手工录入", null, null, 1, "填写销售顾问名称", "file:///E:/DotNetWorkingSpace/MyTestPlatform/Console");
            dt.Rows.Add("来店管理", "来店登记", "类型", "√", null, null, null, "2", "填写业务类型", "Sheet2!A1");
            dt.Rows.Add("来店管理", "来店登记", "来店人数", "√", "手工录入", null, null, "3", "填写来店人数", "http://baidu.com");
            dt.Rows.Add("来店管理", "来店登记", "车系", "√", "手工录入", "a", null, "4", "填写车系", "http://baidu.com");
            dt.Rows.Add("来店管理", "其它", "来店时间", "√", "手工录入", "b", null, "5", "填写来店时间", "mailto:test@szlanyou.com");
            dt.Rows.Add("来店管理", "其它", "离店时间", "√", "手工录入", "c", null, "6", "填写离店时间", "mailto:gezhonglei@szlanyou.com");
            dt.Rows.Add("来店管理", "其它", "备注", "√", "手工录入", "d", null, "7.465", "填写备注信息", "http://lydoa.szlanyou.com/Login.aspx");
            dataSet.Tables.Add(dt);

            dt = new DataTable("SDPChange");
            dt.Columns.Add("ColName");
            dt.Columns.Add("ColCode");
            dt.Columns.Add("DataType");
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("IsKey", typeof(bool));
            dt.Columns.Add("Required", typeof(bool));
            dt.Columns.Add("Constrain");
            dt.Columns.Add("Remark");
            dt.Columns.Add("ChangeType");
            dt.Rows.Add("来店时间", "InTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "新增");
            dt.Rows.Add("离店时间", "LeaveTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "修改");
            dt.Rows.Add("备注", "InTime", "DateTime", 20, false, false, "between 5am and 18pm", "", "删除");
            dataSet.Tables.Add(dt);

            dt = new DataTable("BillList");
            dt.Columns.Add("TableName");
            dt.Columns.Add("FormName");
            dt.Columns.Add("FormCode");
            dt.Columns.Add("Remark");
            dt.Columns.Add("Constrain");
            dt.Rows.Add("SDPChange1", "来店管理1", "bill_001", "特殊说明1", "(约束说明1)");
            dt.Rows.Add("SDPChange2", "来店管理2", "bill_002", "特殊说明2", "(约束说明2)");
            dt.Rows.Add("SDPChange3", "来店管理3", "bill_003", "特殊说明3", "(约束说明3)");
            dataSet.Tables.Add(dt);

            dt = new DataTable("SDPChange1");
            dt.Columns.Add("ColName");
            dt.Columns.Add("ColCode");
            dt.Columns.Add("DataType");
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("IsKey", typeof(bool));
            dt.Columns.Add("Required", typeof(bool));
            dt.Columns.Add("Constrain");
            dt.Columns.Add("Remark");
            dt.Columns.Add("ChangeType");
            dt.Rows.Add("来店时间1", "InTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "新增");
            dt.Rows.Add("离店时间1", "LeaveTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "修改");
            dt.Rows.Add("备注1", "InTime", "DateTime", 20, false, false, "between 5am and 18pm", "", "删除");
            dataSet.Tables.Add(dt);

            dt = new DataTable("SDPChange2");
            dt.Columns.Add("ColName");
            dt.Columns.Add("ColCode");
            dt.Columns.Add("DataType");
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("IsKey", typeof(bool));
            dt.Columns.Add("Required", typeof(bool));
            dt.Columns.Add("Constrain");
            dt.Columns.Add("Remark");
            dt.Columns.Add("ChangeType");
            dt.Rows.Add("来店时间2", "InTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "新增");
            dt.Rows.Add("离店时间2", "LeaveTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "修改");
            dt.Rows.Add("备注2", "InTime", "DateTime", 20, false, false, "between 5am and 18pm", "", "删除");
            dataSet.Tables.Add(dt);

            dt = new DataTable("SDPChange3");
            dt.Columns.Add("ColName");
            dt.Columns.Add("ColCode");
            dt.Columns.Add("DataType");
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("IsKey", typeof(bool));
            dt.Columns.Add("Required", typeof(bool));
            dt.Columns.Add("Constrain");
            dt.Columns.Add("Remark");
            dt.Columns.Add("ChangeType");
            dt.Rows.Add("来店时间3", "InTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "新增");
            dt.Rows.Add("离店时间3", "LeaveTime", "DateTime", 20, false, true, "between 5am and 18pm", "", "修改");
            dt.Rows.Add("备注3", "InTime", "DateTime", 20, false, false, "between 5am and 18pm", "", "删除");
            dataSet.Tables.Add(dt);

            return dataSet;
        }

        private static DataSet GetFunctionList()
        {
            DataSet dataSet = new DataSet("FunctionList");
            DataTable dt = new DataTable("ValueTypeSource");
            dt.Columns.Add("Name");
            dt.Rows.Add("文本");
            dt.Rows.Add("数字");
            dt.Rows.Add("日期");
            dataSet.Tables.Add(dt);
            dt = new DataTable("FuncTypeSource");
            dt.Columns.Add("Value");
            dt.Rows.Add("表单");
            dt.Rows.Add("报表");
            dt.Rows.Add("接口");
            dt.Rows.Add("后台");
            dataSet.Tables.Add(dt);

            dt = new DataTable("ModuleFuncListSource");
            dt.Columns.Add("Module");
            dt.Columns.Add("SubModule");
            dt.Columns.Add("FuncCode");
            dt.Columns.Add("FuncName");
            dt.Columns.Add("FuncType");
            dt.Columns.Add("Description");
            dt.Columns.Add("UseOrg");
            dt.Columns.Add("AppConfig");
            dt.Columns.Add("ReptConfig");
            dt.Columns.Add("Remark");
            dt.Rows.Add("整车采购", "需求计划", "reqplan001", "采购需求计划", "表单", "(功能说明)", "销售部", "√", null, null);
            dt.Rows.Add("整车采购", "需求计划", "reqplan002", "生产需求计划", "表单", "(功能说明)", "生产部", "√", "√", null);
            dt.Rows.Add("整车采购", "采购管理", "purchase001", "一般采购", "接口", "(功能说明)", "应收科", "", "", null);
            dt.Rows.Add("整车采购", "采购管理", "purchase002", "采购分配流程", "接口", "(功能说明)", "应付科", "√", "", null);
            dt.Rows.Add("整车采购", "采购管理", "purchase003", "大宗采购", "接口", "(功能说明)", "销售部", "", "√", null);
            dt.Rows.Add("整车采购", "采购管理", "purchase004", "特殊采购", "报表", "(功能说明)", "进口科", "", "√", null);
            dt.Rows.Add("整车采购", "采购管理", "purchase005", "进口车采购", "报表", "(功能说明)", "进口科", "", "", null);
            dt.Rows.Add("整车库存", "专营车库存", "starage001", "专营店库存管理", "后台", "(功能说明)", "仓管", "√", "√", null);
            dt.Rows.Add("整车库存", "主机厂库存", "starage002", "主机厂库存管理", "后台", "(功能说明)", "仓管", "√", "", null);
            dataSet.Tables.Add(dt);

            dt = new DataTable("ModuleTable");
            dt.Columns.Add("Module");
            List<string> modulelist = dataSet.Tables["ModuleFuncListSource"].AsEnumerable().Select(p => p["Module"].ToString()).ToList();
            modulelist = modulelist.Distinct().ToList();
            foreach (var item in modulelist)
            {
                dt.Rows.Add(item);
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("InterfaceList");
            dt.Columns.Add("InterfaceCode");
            dt.Columns.Add("InterfaceName");
            dt.Columns.Add("TransWay");
            dt.Columns.Add("TransTime", typeof(DateTime));
            dt.Columns.Add("TransFreq", typeof(int));
            dt.Columns.Add("TransDirection");
            dt.Columns.Add("Description");
            dt.Columns.Add("TableName");
            dt.Rows.Add("JK001", "一般采购接口", "DCS", DateTime.Now.AddHours(1), 1, "MDM -> DMS", "(接口描述)", "Table1");
            dt.Rows.Add("JK002", "采购分流程接口", "DCS", DateTime.Now, 2, "MDM -> DMS", "(接口描述)", "Table2");
            dt.Rows.Add("JK003", "大宗采购接口", "DCS", DateTime.Now.AddHours(2), 3, "MDM -> DMS", "(接口描述)", "Table3");
            dataSet.Tables.Add(dt);

            dt = new DataTable("Table1");
            dt.Columns.Add("ItemName");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("ValueType");
            dt.Columns.Add("Length", typeof(int));
            dt.Columns.Add("Key");
            dt.Columns.Add("BizRule");
            dt.Columns.Add("DefaultValue");
            dt.Columns.Add("Remark");
            dt.Rows.Add("采购Id", "ID", "数字", null, "主键", "", "", "");
            dt.Rows.Add("采购人", "PurchaserId", "文本", 100, "外键", "", "", "");
            dt.Rows.Add("操作时间", "OperateTime", "日期", null, "", "GetDate()", "", "");
            dataSet.Tables.Add(dt);

            dt = dt.Copy();
            dt.TableName = "Table2";
            dt.Rows.Add("货品类型", "ProductType", "数字", null, "", "", "", "");
            dataSet.Tables.Add(dt);
            dt = dt.Copy();
            dt.TableName = "Table3";
            dt.Rows.Add("采购数量", "Quality", "数字", null, "", "", "", "");
            dataSet.Tables.Add(dt);

            return dataSet;
        }

        private static DataSet GetBusinessMainFlow()
        {
            DataSet dataSet = new DataSet("BusinessMainFlow");
            DataTable dt = new DataTable("Cells");
            dt.Columns.Add("FlowChart", typeof(byte[]));
            using (FileStream filestream = new FileStream(Path.Combine(path, "ExcelTemplate\\FormStyle.jpg"), FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[(int)filestream.Length];
                filestream.Read(bytes, 0, bytes.Length);
                dt.Rows.Add(bytes);
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("NodeInfo");
            dt.Columns.Add("NodeName");
            dt.Columns.Add("OrgName");
            dt.Columns.Add("BizCategory");
            dt.Rows.Add("需求计划", "主机厂专营店物流", "整车采购");
            dt.Rows.Add("采购管理", "主机厂专营店物流", "整车采购");
            dt.Rows.Add("采购物流", "主机厂专营店物流", "整车采购");
            dt.Rows.Add("整车销退", "主机厂专营店", "整车销退");
            dt.Rows.Add("主机厂财务", "主机厂专营店", "整车财务");
            dt.Rows.Add("专营店财务", "主机厂专营店", "整车财务");
            dataSet.Tables.Add(dt);

            return dataSet;
        }

        private static DataSet GetBusinessCategory()
        {
            DataSet dataSet = new DataSet("BusinessCategory");
            DataTable dt = new DataTable("BodySource");
            dt.Columns.Add("Description");
            dt.Columns.Add("Organization");
            dt.Columns.Add("BusinessCategory");
            dt.Rows.Add("根据销售情况制定计划", "销售部经理", "采购需求计划");
            dt.Rows.Add("选择已经第二次转库的车辆", "销售顾问", "物流运输流程");
            dt.Rows.Add("确认销售顾问制作的采购单", "销售部经理", "订单余量转采购分配流程");
            dt.Rows.Add("根据实际需求，制作采购单", "销售顾问", "验收入库流程");
            dataSet.Tables.Add(dt);

            dt = new DataTable("RowHeaderSource");
            dt.Columns.Add("Organization");
            List<string> rowHeader = new List<string>() { 
                "销售部经理", "销售顾问", "进口车销售专员", "整车仓库管理员", "牌证员", "二网销售主管", 
                "大客户专员", "展厅主管", "试乘试驾专员", "财务主管", "收银员", "销售管理科员", 
                "销售计划科员", "大宗客户科员", "车辆业务科员", "大区督导", "应收科员", "应付科员", 
                "启辰销售科员", "进口车销售科员", "RowHeaderTest"
            };
            for (int i = 0; i < rowHeader.Count; i++)
            {
                dt.Rows.Add(rowHeader[i]);
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("RowTreeSource");
            dt.Columns.Add("Org");
            dt.Columns.Add("ParentOrg");
            for (int i = 0; i < rowHeader.Count; i++)
            {
                dt.Rows.Add(rowHeader[i], i <= 10 ? "专营店" : i < 20 ? "主机厂" : null);
            }
            dt.Rows.Add("专营店", "业务说明");
            dt.Rows.Add("主机厂", "业务说明");
            dt.Rows.Add("业务说明", null);
            dataSet.Tables.Add(dt);

            dt = new DataTable("ColumnHeaderSource");
            dt.Columns.Add("BizCategory");
            List<string> colHeader4 = new List<string>() { 
                "采购需求计划","一般采购分配流程","订单余量转采购分配流程","订单生产车型采购分配流程",
                "大宗采购分配流程","特殊采购分配流程","进口车采购分配流程","物流运输流程","DFDLC二次转库流程","搬入地变更流程",
                "验收入库流程","专营店库存管理1","主机厂库存管理1","ColumnTest"
            };
            for (int i = 0; i < colHeader4.Count; i++)
            {
                dt.Rows.Add(colHeader4[i]);
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("ColumnTreeSource");
            dt.Columns.Add("Category");
            dt.Columns.Add("ParentCategory");
            for (int i = 0; i < colHeader4.Count; i++)
            {
                dt.Rows.Add(colHeader4[i], i < 1 ? "需求计划" :
                    i < 7 ? "采购管理" :
                    i < 11 ? "采购物流" :
                    i < 12 ? "专营店库存管理" :
                    i < 13 ? "主机厂库存管理" : null);//3->4
            }
            dt.Rows.Add("需求计划", "整车采购");//2->3
            dt.Rows.Add("采购管理", "整车采购");
            dt.Rows.Add("采购物流", "整车采购");
            dt.Rows.Add("专营店库存管理", "整车库存");
            dt.Rows.Add("主机厂库存管理", "整车库存");
            dt.Rows.Add("整车采购", "整车");//1->2
            dt.Rows.Add("整车库存", "整车");
            dt.Rows.Add("整车", null);
            dataSet.Tables.Add(dt);

            dt = new DataTable("LeftUpperSource");
            dt.Columns.Add("Name");
            dt.Rows.Add("类别");
            dt.Rows.Add("业务大类");
            dt.Rows.Add("业务中类");
            dt.Rows.Add("业务小类");
            dataSet.Tables.Add(dt);

            #region 只行标题数据源
            dt = new DataTable("BodySource2");
            for (int i = 0; i < rowHeader.Count; i++)
            {
                dt.Columns.Add(rowHeader[i]);
            }
            dt.Rows.Add("", "", "", "", "Text", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            dt.Rows.Add("", "", "", "", "", "Test", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            dt.Rows.Add("", "", "", "", "", "", "", "根据销量制定销售计划", "", "", "", "", "", "", "", "", "", "", "", "");
            dt.Rows.Add("", "", "选择已经第二次转库的车辆", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            dt.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", "", "根据数据", "", "", "", "", "", "");
            dt.Rows.Add("", "", "", "", "", "", "", "", "确认销售顾问制作的采购单", "", "", "", "", "", "", "", "", "", "", "");
            dataSet.Tables.Add(dt);

            dt = new DataTable("BodySource3");
            for (int i = 0; i < colHeader4.Count; i++)
            {
                dt.Columns.Add(colHeader4[i]);
            }
            dt.Rows.Add("A", "", "", "", "", "", "", "", "I", "", "", "", "", "");
            dt.Rows.Add("", "B", "", "", "", "", "", "H", "", "J", "", "", "", "");
            dt.Rows.Add("", "", "C", "", "", "", "G", "", "", "", "K", "", "", "");
            dt.Rows.Add("", "", "", "D", "", "F", "", "", "", "", "", "L", "", "N");
            dt.Rows.Add("", "", "", "", "E", "", "", "", "", "", "", "", "M", "");
            dataSet.Tables.Add(dt);

            dt = new DataTable("BodySource4");
            dt.Columns.Add("OrgName");
            for (int i = 0; i < 10; i++)
            {
                dt.Columns.Add("Column" + i);
            }
            for (int i = 0; i < rowHeader.Count; i++)
            {
                dt.Rows.Add(rowHeader[i], "A" + i, "B" + i, "C" + i, "D" + i, "E" + i,
                    "F" + i, "G" + i, "H" + i, "I" + i, "J" + i);
            }
            dataSet.Tables.Add(dt);
            #endregion 只有标题数据源

            return dataSet;
        }

        private static DataSet GetOrganization()
        {
            DataSet dataSet = new DataSet("Organization");
            DataTable dt = new DataTable("Cells");
            dt.Columns.Add("OrgChart", typeof(byte[]));
            using (FileStream filestream = new FileStream(Path.Combine(path, "ExcelTemplate\\FormStyle.jpg"), FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[(int)filestream.Length];
                filestream.Read(bytes, 0, bytes.Length);
                dt.Rows.Add(bytes);
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("OrganizationInfo");
            dt.Columns.Add("PostName");
            dt.Columns.Add("PostReq");
            dt.Columns.Add("PostReqDescription");
            dt.Columns.Add("Remark");
            dt.Rows.Add("销售经理", "完成销售目标", "1、制定销售计划aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("销售顾问", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("进口车销售专员", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("整车仓库管理员", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("牌证员", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("二网销售主管", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("大客户专员", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dt.Rows.Add("展厅主管", "完成销售目标", "1、制定销售计划\n2、定期汇报销售业绩\n3、策划销售方向", "。。。");
            dataSet.Tables.Add(dt);

            return dataSet;
        }

        private static DataSet GetReport()
        {
            DataSet dataSet = new DataSet("Report");
            DataTable dt = new DataTable("Cells");
            dt.Columns.Add("Title");
            dt.Columns.Add("ReportStyle", typeof(byte[]));
            dt.Rows.Add("标准来店量分日统计报表");
            using (FileStream filestream = new FileStream(Path.Combine(path, "ExcelTemplate\\ReportStyle.bmp"), FileMode.Open, FileAccess.Read))
            {
                byte[] bytes = new byte[(int)filestream.Length];
                filestream.Read(bytes, 0, bytes.Length);
                dt.Rows[0]["ReportStyle"] = bytes;
            }
            dataSet.Tables.Add(dt);

            dt = new DataTable("ReportRule");
            dt.Columns.Add("ReportName");
            dt.Columns.Add("Item");
            dt.Columns.Add("DataType");
            dt.Columns.Add("Query");
            dt.Columns.Add("Source");
            dt.Columns.Add("SourceItem");
            dt.Columns.Add("Algorithm");
            dt.Columns.Add("Remark");
            dt.Rows.Add("来店统计表-来店看板", "网点名称", "文本", "√", "来店登记", "网点ID", "", "");
            dt.Rows.Add("来店统计表-来店看板", "统计开始日期", "日期", "√", "", "", "", "");
            dt.Rows.Add("来店统计表-来店看板", "统计结束日期", "日期", "√", "", "", "", "");
            dt.Rows.Add("来店统计表-来店看板", "销售顾问", "文本", "", "来店登记", "销售顾问ID", "", "");
            dt.Rows.Add("来店统计表-来店看板", "日期", "日期", "", "来店登记", "登记日期", "", "");
            dt.Rows.Add("来店统计表-来店看板", "合计", "数字", "", "", "网点ID", "=Σ来店数合计", "");
            dt.Rows.Add("来店统计表-来店看板", "留资料", "数字", "", "", "网点ID", "=Σ有效来店数合计", "");
            dt.Rows.Add("来店统计表-来店看板", "留档率真", "数字", "", "", "网点ID", "=留资料/合计", "");
            dt.Rows.Add("Text", "留档率真", "数字", "", "", "网点ID", "=留资料/合计", "");
            dataSet.Tables.Add(dt);

            return dataSet;
        }

        private static DataSet Get3ConfigReq()
        {
            DataSet dataSet = new DataSet("3ConfigReq");
            DataTable dt = new DataTable("AppConfigSource");
            dt.Columns.Add("Platform");
            dt.Columns.Add("Module");
            dt.Columns.Add("FlowName");
            dt.Columns.Add("ConfigReq");
            dt.Columns.Add("BizScene");
            dt.Columns.Add("ConfigMethod");
            dt.Columns.Add("ConfigFunc");
            dt.Columns.Add("EffectedFunc");
            dt.Columns.Add("EffectedRegion");
            dt.Columns.Add("Remark");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-", "-", "-", "-");
            dataSet.Tables.Add(dt);

            dt = new DataTable("FlowConfigSource");
            dt.Columns.Add("Platform");
            dt.Columns.Add("Module");
            dt.Columns.Add("FlowName");
            dt.Columns.Add("ConfigReq");
            dt.Columns.Add("BizScene");
            dt.Columns.Add("ConfigMethod");
            dt.Columns.Add("Remark");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dataSet.Tables.Add(dt);

            dt = new DataTable("ReportConfigSource");
            dt.Columns.Add("Platform");
            dt.Columns.Add("Module");
            dt.Columns.Add("ReportName");
            dt.Columns.Add("Description");
            dt.Columns.Add("ConfigReq");
            dt.Columns.Add("ConfigMethod");
            dt.Columns.Add("Remark");
            dataSet.Tables.Add(dt);
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-");
            return dataSet;
        }

        private static DataSet Get4Line3Config()
        {
            DataSet dataSet = new DataSet("4Line3Config");
            DataTable dt = new DataTable("Cells");
            dt.Columns.Add("Title");
            dt.Rows.Add("来店管理流程_四线三配置分析");
            dataSet.Tables.Add(dt);

            dt = new DataTable("4Line3ConfigSource");
            string[] columns = "FlowName,WorkFlow,Node,Function,Organization,Input,Process,Rule,Output,OnOffLine,ReqBasicData,Standard,Monitor,MonitorRule,MonitorWay,UpdateRule,UpdateWay,MonitorReport,ScheduleReport,QualityReport,Rate,GeneralReport,Shared,ShareWay,SharedObject,AppConfig,FlowConfig,ReportName,ReportConfig".Split(',');
            foreach (var column in columns)
            {
                dt.Columns.Add(column);
            }
            Random random = new Random();
            for (int i = 0; i < 10; i++)
            {
                DataRow row = dt.NewRow();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    row[j] = (char)('A' + random.Next(100) % 26);
                }
                dt.Rows.Add(row);
            }
            dataSet.Tables.Add(dt);
            return dataSet;
        }
    }
}
