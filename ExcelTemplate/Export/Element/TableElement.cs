using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using ExportTemplate.Export.Entity;
using ExportTemplate.Export.Util;

namespace ExportTemplate.Export.Element
{

    public class TableElement : Element<Table>
    {
        public new const string ELEMENT_NAME = "Table";
        public IProductRuleGetter _upperElement;

        public TableElement(XmlElement data, IProductRuleGetter sheetElement)
        {
            _data = data;
            _upperElement = sheetElement;
        }

        protected override Table CreateEntity()
        {
            BaseEntity container = _upperElement.CurrentEntity;
            ProductRule prodRule = container.ProductRule;

            Location tmpLocation = new Location(_data.GetAttribute("location"));
            _entity = new Table(prodRule, container, tmpLocation);
            string tmpStr = _data.GetAttribute("source");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                _entity.SourceName = tmpStr;
                //只要不是动态解析的数据源，Source不能为空
                if ((container is Sheet && !(container as Sheet).IsDynamic) || !DynamicSource.NeedDynamicParse(tmpStr))
                {
                    //_entity.Source = prodRule.RegistSource(tmpStr);
                    prodRule.RegistSource(tmpStr);
                }
            }
            _entity.RowNumIndex = ParseUtil.ParseColumnIndex(_data.GetAttribute("rowNumIndex"), -1);
            _entity.CopyFill = ParseUtil.ParseBoolean(_data.GetAttribute("copyFill"), true);
            _entity.SumLocation = ParseUtil.ParseEnum<LocationPolicy>(_data.GetAttribute("sumLocation"), LocationPolicy.Undefined);
            _entity.SumOffset = ParseUtil.ParseInt(_data.GetAttribute("sumOffset"), _entity.SumLocation == LocationPolicy.Undefined || _entity.SumLocation == LocationPolicy.Absolute ? -1 : 0);
            _entity.AutoFitHeight = ParseUtil.ParseBoolean(_data.GetAttribute("autoFitHeight"), false);
            int groupLevel = ParseUtil.ParseInt(_data.GetAttribute("groupLevel"), 0);
            if (groupLevel > 0)
            {
                tmpStr = _data.GetAttribute("groupNumShow");
                if (!string.IsNullOrEmpty(tmpStr))
                {
                    bool[] shows = new bool[groupLevel];
                    string[] bools = tmpStr.Split(',');

                    for (int i = 0; i < bools.Length && i < shows.Length; i++)
                    {
                        if (bool.Parse(bools[i]))
                        {
                            shows[i] = true;
                        }
                    }
                    _entity.SetGroup(shows);
                }
            }

            _entity.AdjustSumOffset();

            #region Field重复出现的处理
            List<Field> fieldlist = new List<Field>();
            tmpStr = _data.GetAttribute("fields");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                string[] fieldArray = tmpStr.Split(',');
                foreach (var fieldname in fieldArray)
                {
                    fieldlist.Add(new Field(_entity, fieldname));
                }
            }
            _entity.AddFields(fieldlist);

            //子结点
            parseField(_upperElement.ProductRuleElement.QuerySubNodes("Field", _data));

            #endregion Field重复出现的处理

            // 行号处理
            parseRowNum(_upperElement.ProductRuleElement.QuerySubNodes("RowNum", _data));

            //区域计算：计算行号、字段的列位置，以及Table填充列范围
            _entity.CalculateArea();

            //汇总列：在确定字段位置之后处理
            tmpStr = _data.GetAttribute("sumColumns");
            if (!string.IsNullOrEmpty(tmpStr))
            {
                string[] columns = tmpStr.Split(',');
                _entity.AddSumColumns(columns);
            }
            return _entity;
        }

        private void parseField(XmlNodeList nodes)
        {
            List<Field> fields = new List<Field>();
            if (nodes.Count > 0)
            {
                foreach (XmlElement fieldNode in nodes)
                {
                    Field field = new FieldElement(fieldNode, this).Entity;
                    if (field != null)
                    {
                        fields.Add(field);
                    }
                }
            }
            if (fields.Count > 0)
            {
                _entity.AddFields(fields);
            }
        }

        private void parseRowNum(XmlNodeList nodes)
        {
            if (nodes.Count > 0 && nodes[0] is XmlElement)
            {
                string tmpStr = (nodes[0] as XmlElement).GetAttribute("index");
                _entity.RowNumIndex = string.IsNullOrEmpty(tmpStr) ? _entity.Location.ColIndex : int.Parse(tmpStr);
                //如设置行号所在列索引小于指定区域，将被调整为起点列
                if (_entity.RowNumIndex < _entity.Location.ColIndex)
                {
                    _entity.RowNumIndex = _entity.Location.ColIndex;
                }
            }
        }
    }

}
