using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExportTemplate.Export.Entity
{
    /// <summary>
    /// 数据接口(未使用)
    /// 目的：
    /// (1) 实现ListSource与TreeSource之间的兼容
    /// (2) 将DataTable隐藏（暂无法达到，使用频度太高且对外联系太紧密）
    /// </summary>
    public abstract class BaseSource
    {
        public string Name { get; set; }
        public DataTable Table { get; set; }
    }

    /// <summary>
    /// 数据源
    /// </summary>
    public class Source : IRuleEntity, ICloneable
    {
        /// <summary>
        /// 数据源名称标识
        /// </summary>
        public string Name;
        public DataTable Table;

        public Source() { }
        public Source(string name) { Name = name; }

        public bool IsEmpty()
        {
            return Table == null || Table.Rows.Count == 0;
        }

        public bool Contains(string field)
        {
            return Table != null && Table.Columns.Contains(field);
        }

        public Type GetField(string field)
        {
            return Table.Columns[field].DataType;
        }

        public IList<object> GetValues(string field)
        {
            return Table != null ? Table.AsEnumerable().Select(p => p[field]).ToList() : new List<object>();
        }

        public IList<T> GetValues<T>(string field, Func<object, T> func)
        {
            return Table != null ? Table.AsEnumerable().Select(p => func(p[field])).ToList() : new List<T>();
        }
        public IList<T> GetValues<T>(int colIndex, Func<object, T> func)
        {
            return Table != null ? Table.AsEnumerable().Select(p => func(p[colIndex])).ToList() : new List<T>();
        }

        public IDictionary<string, object> Get(int index)
        {
            Dictionary<string, object> dict = new Dictionary<string, object>();
            if (index < 0 || index >= Table.Rows.Count) return dict;
            foreach (DataColumn column in Table.Columns)
            {
                dict.Add(column.ColumnName, Table.Rows[index][column.ColumnName]);
            }
            return dict;
        }

        public IList<object> GroupToken(string[] fields)
        {
            IList<object> values = new object[Table.Rows.Count];
            for (int i = 0; i < values.Count; i++)
            {
                StringBuilder sbStr = new StringBuilder();
                foreach (var field in fields)
                {
                    sbStr.Append(Table.Rows[i][field] + "$");
                }
                values[i] = sbStr.ToString();
            }
            return values;
        }

        public override string ToString()
        {
            return Name;
        }

        public virtual object Clone()
        {
            return new Source() { Name = Name, Table = Table != null ? Table.Clone() : null };
        }
    }

    /// <summary>
    /// 数据列表（单列数据）
    /// </summary>
    public class ListSource : Source, IRuleEntity
    {
        public string Field = "Column1";
        public bool IsInner = false;

        internal ListSource(string name) : base(name) { IsInner = true; }

        public ListSource(string name, string fieldname)
        {
            Name = name;
            Field = fieldname;
        }

        public void LoadData<T>(T[] array)
        {
            if (Table == null)
            {
                Table = new DataTable();
            }
            if (!Contains(Field))
            {
                Table.Columns.Add(Field);
            }
            Table.Clear();
            foreach (var item in array)
            {
                Table.Rows.Add(item);
            }
        }

        public string[] GetStringValues()
        {
            return GetValues(p => p.ToString());
        }

        public object[] GetValues()
        {
            return GetValues(p => p);
        }

        public T[] GetValues<T>(Func<object, T> func)
        {
            return Table == null || Table.Columns.Count == 0 ? new T[0] :
                Contains(Field) ? GetValues<T>(Field, func).ToArray() :
                GetValues<T>(0, func).ToArray();
        }

        public override string ToString()
        {
            return string.Format("{0}.{1}", Name, Field);
        }

        public override object Clone()
        {
            return new ListSource(Name) { Table = Table != null ? IsInner ? Table.CloneData() : Table.Clone() : null, Field = Field };
        }
    }

    /// <summary>
    /// 树结构数据源（用于多级标题）
    /// 输出字段不一样时，不能与ListSource兼容
    /// </summary>
    public class TreeSource : Source
    {
        public string ContentField;
        public string IdField;
        public string ParentIdField;

        public TreeSource(string name) : base(name) { }
        public TreeSource(string name, string content, string idField, string pIdField)
            : base(name)
        {
            ContentField = content;
            IdField = idField;
            ParentIdField = pIdField;
        }

        public DataRow GetParent(object id)
        {
            if (id == null || id is DBNull) return null;
            IEnumerable<DataRow> rows = Table.AsEnumerable().Where(p => p[IdField].Equals(id));
            if (rows.Count() > 0) return rows.First();
            return null;
        }
        public object GetParentId(object id)
        {
            DataRow row = GetParent(id);
            if (row != null) return row[ParentIdField];
            return null;
        }

        public DataRow GetRow(object id)
        {
            IEnumerable<DataRow> rows = Table.AsEnumerable().Where(p => p[IdField].Equals(id));
            if (rows.Count() > 0) return rows.First();
            return null;
        }

        public override string ToString()
        {
            return string.Format("treeSource=\"{0}.{1}\" treeInnerMapping=\"{2}:{3}\"", Name, ContentField, IdField, ParentIdField);
        }

        public override object Clone()
        {
            return new TreeSource(Name) { Table = Table != null ? Table.Clone() : null, IdField = IdField, ContentField = ContentField, ParentIdField = ParentIdField };
        }
    }

    /// <summary>
    /// 数据源之间的关系
    /// </summary>
    public class SourceRelation
    {
        public Source Source;
        public string Field;
        public Source ReferecedSource;
        public string ReferecedField;

        public IEnumerable<DataRow> GetReferencedRows()
        {
            object[] datas = Source.Table.AsEnumerable().Select(p => p[Field]).ToArray();
            return ReferecedSource.Table.AsEnumerable().Where(p => datas.Contains(p[ReferecedField]));
        }

        public object[] GetReferencedData()
        {
            return GetReferencedRows().Select(p => p[ReferecedField]).ToArray();
        }

        public string ToString(string attribute)
        {
            return string.Format("{0}=\"{1}.{2}:{3}.{4}\"", attribute, Source.Name, Field, ReferecedSource != null ? ReferecedSource.Name : "", ReferecedField);
        }

        public SourceRelation Clone(ProductRule productRule)
        {
            SourceRelation newSourceRel = new SourceRelation()
            {
                //Source = Source.Clone() as Source, 
                //ReferecedSource = ReferecedSource.Clone() as Source, 
                Field = Field,
                ReferecedField = ReferecedField
            };
            if (Source != null)
            {
                newSourceRel.Source = productRule.GetSource(Source.Name);
            }
            if (ReferecedSource != null)
            {
                newSourceRel.ReferecedSource = productRule.GetSource(ReferecedSource.Name);
            }
            return newSourceRel;
        }
    }

}
