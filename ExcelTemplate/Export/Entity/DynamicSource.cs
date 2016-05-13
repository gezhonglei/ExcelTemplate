using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 动态对象
    /// </summary>
    public class DynamicSource
    {
        protected List<int> _sourceIndexes = null;
        protected Source _source;
        protected ProductRule _productRule;
        protected Dictionary<string, Func<int, int>> _expressionBasedIndex = new Dictionary<string, Func<int, int>>();

        public DynamicSource(ProductRule productRule)
        {
            this._productRule = productRule;
        }

        /// <summary>
        /// 数据源Distinct规则
        /// 注：如DistinctFunc = null,将默认每行的Distinct值不同
        /// </summary>
        public Func<DataTable, int, object> DistinctFunc = null;

        /// <summary>
        /// 动态数据源
        /// </summary>
        public Source Source
        {
            get { return _source; }
            set { _source = value; }
        }

        /// <summary>
        /// 动态对象个数
        /// </summary>
        public int Count
        {
            get
            {
                GetSourceIndexes();
                return _sourceIndexes.Count;
            }
        }

        public DynamicSource(ProductRule productRule, Source source)
            : this(productRule)
        {
            this._source = source;
        }

        /// <summary>
        /// 基于索引设置内置参数
        /// </summary>
        /// <param name="name">参数名称：不要与动态数据源字段重名，否则会被数据源覆盖</param>
        /// <param name="func">自定义lamda表达式函数，输入参数当前索引号</param>
        public void SetParam(string name, Func<int, int> func)
        {
            if (_expressionBasedIndex.ContainsKey(name))
            {
                _expressionBasedIndex[name] = func;
            }
            else
            {
                _expressionBasedIndex.Add(name, func);
            }
        }

        private List<int> GetSourceIndexes()
        {
            if (_sourceIndexes != null) return _sourceIndexes;

            if (_source != null && _source.Table != null)
            {
                if (DistinctFunc == null)
                {
                    DistinctFunc = (dt, rowNum) => rowNum;
                }
                //用于处理重名表单名称：只取重复的第一条数据
                Dictionary<object, int> dict = new Dictionary<object, int>();
                int count = _source.Table.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    object distinctObject = DistinctFunc(_source.Table, i);
                    if (!dict.ContainsKey(distinctObject))
                    {
                        dict.Add(distinctObject, i);
                    }
                }
                _sourceIndexes = dict.Values.ToList();
            }
            return _sourceIndexes;
        }

        /// <summary>
        /// 获取动态Sheet中Cell元素值
        /// </summary>
        /// <param name="expression">表达式</param>
        /// <param name="index">索引号</param>
        /// <param name="preExParams">前置扩展参数（优先级低）</param>
        /// <param name="postExParams">后置扩展参数（优先级高）</param>
        /// <returns>返回值</returns>
        public string GetDynamicValue(string expression, int index,
            IDictionary<string, string> preExParams = null, IDictionary<string, string> postExParams = null)
        {
            GetSourceIndexes();
            Dictionary<string, object> dict = new Dictionary<string, object>();
            if (preExParams != null)
            {
                foreach (var item in preExParams)
                {
                    dict.Add(item.Key, item.Value);
                }
            }
            foreach (var item in _expressionBasedIndex)
            {
                dict.Add(item.Key, item.Value(index));
            }
            DataRow row = _source.Table.Rows[_sourceIndexes[index]];
            foreach (DataColumn column in _source.Table.Columns)
            {
                if (!dict.ContainsKey(column.ColumnName))
                    dict.Add(column.ColumnName, row[column.ColumnName]);
                else
                    dict[column.ColumnName] = row[column.ColumnName];//如果重复则替换

                if (!dict.ContainsKey(column.ColumnName))
                    dict.Add(string.Format("{0}.{1}", _source.Name, column.ColumnName), row[column.ColumnName]);
            }
            if (postExParams != null)
            {
                foreach (var item in postExParams)
                {
                    dict.Add(item.Key, item.Value);
                }
            }

            string[] fields = GetFields(expression);
            foreach (var fieldname in fields)
            {
                if (dict.ContainsKey(fieldname))
                {
                    expression = expression.Replace(string.Format("{{{0}}}", fieldname), (dict[fieldname] ?? "").ToString());
                }
            }
            return expression;
        }

        protected static string[] GetFields(string expression)
        {
            MatchCollection matches = new Regex(@"(?:{)([\w\.]+)(?:})").Matches(expression);
            string[] fields = new string[matches.Count];
            for (int i = 0; i < matches.Count; i++)
            {
                fields[i] = matches[i].Groups[1].Value;
            }
            return fields;
        }

        /// <summary>
        /// 在动态Sheet中获取Table或Cell元素的数据源
        /// </summary>
        /// <param name="expression">表达式</param>
        /// <param name="index">索引号</param>
        /// <returns>数据源</returns>
        public Source GetDynamicSource(string expression, int index,
            IDictionary<string, string> preExParams = null, IDictionary<string, string> postExParams = null)
        {
            Source tmpSource = null;
            string sourceName = GetDynamicValue(expression, index, preExParams, postExParams);
            if (!string.IsNullOrEmpty(sourceName))
            {
                tmpSource = _productRule.GetSource(sourceName);
            }
            //如果表达式包含".",可能是"TableName.FieldName"组成，解析出前TableName部分
            if (tmpSource == null && sourceName.Contains("."))
            {
                sourceName = sourceName.Substring(0, sourceName.IndexOf('.'));
                tmpSource = _productRule.GetSource(sourceName);
            }
            return tmpSource;
        }

        /// <summary>
        /// 是否需要解析
        /// 注：只要包含一对大括号的字符串就是需要解析的
        /// </summary>
        /// <param name="expresssion">表达式</param>
        /// <returns></returns>
        public static bool NeedDynamicParse(string expresssion)
        {
            return new Regex(@"(?:{)([\w\.]+)(?:})").IsMatch(expresssion);
        }

        public DynamicSource Clone(ProductRule productRule)
        {
            return new DynamicSource(productRule)
            {
                _source = _source.Clone() as Source,
                DistinctFunc = DistinctFunc.Clone() as Func<DataTable, int, object>
            };
        }
    }

}
