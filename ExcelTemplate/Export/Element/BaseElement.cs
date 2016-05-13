using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export.Element
{
    //[Obsolete]
    //public class Element
    //{
    //    public string Name;
    //    public string Content;
    //    public List<Element> Children;
    //    public IDictionary<string, string> Attributes;
    //}

    //[Obsolete]
    //public class ElementAttribute
    //{
    //    public string Name;
    //    public string Value;
    //    public string Regex;
    //    public bool Required;
    //    public string Default;
    //}

    /// <summary>
    /// XML序列化接口
    /// </summary>
    public interface IXmlSerialize
    {
        /// <summary>
        /// 将实体转换成XML字符串
        /// </summary>
        /// <param name="param">可选参数</param>
        /// <returns></returns>
        string ToString(string param);
    }

    /// <summary>
    /// 实现实体接口
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    public interface IEntityGetter<T>
    {
        /// <summary>
        /// 当前实体对象
        /// </summary>
        T Entity { get; }
        /// <summary>
        /// 原始实体对象（规则）
        /// </summary>
        T NewEntity { get; }
    }

    /// <summary>
    /// 元素基本
    /// </summary>
    public abstract class BaseElement
    {
        public const string ELEMENT_NAME = "Unkown_Element";
        /// <summary>
        /// 元素名称
        /// </summary>
        public abstract string ElementName { get; }
        /// <summary>
        /// 获取规则实体
        /// </summary>
        /// <returns></returns>
        public abstract IRuleEntity GetEntity();
    }

    /// <summary>
    /// 泛型元素类
    /// 抽象IEntityGetter接口原因：强类型约束
    /// 抽提BaseElemnt原因：Element泛型类型不便于多态调用
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public abstract class Element<T> : BaseElement, IEntityGetter<T> where T : IRuleEntity
    {
        protected XmlElement _data;
        protected T _entity;
        protected T _tplEntity;

        /// <summary>
        /// 原始实体对象（规则）
        /// </summary>
        public virtual T NewEntity
        {
            get
            {
                T tmpEntity = CreateEntity();
                if (_entity == null) _entity = tmpEntity;
                return tmpEntity;
            }
        }

        /// <summary>
        /// 当前实体对象
        /// </summary>
        public T Entity
        {
            get { return _entity != null ? _entity : this.NewEntity; }
        }

        public override IRuleEntity GetEntity()
        {
            return this.Entity;
        }

        /// <summary>
        /// 元素名称
        /// 注意：子类要覆盖常量ELEMENT_NAME
        /// </summary>
        public override string ElementName { get { return ELEMENT_NAME; } }

        /// <summary>
        /// 解析子结点
        /// 注意：使用此方法，要求P类构造函数是两个参数：第一个是XmlElement，第二个是V类型
        /// </summary>
        /// <typeparam name="P">Element子类且泛型参数为V</typeparam>
        /// <typeparam name="V">规则实体类</typeparam>
        /// <param name="nodes">xml元素结点</param>
        /// <returns>实体列表</returns>
        protected List<V> parseSubElement<P, V>(XmlNodeList nodes)
            where P : Element<V>
            where V : IRuleEntity
        {
            List<V> entities = new List<V>();
            foreach (XmlElement node in nodes)
            {
                P element = (P)typeof(P).GetConstructor(new Type[] { node.GetType(), this.GetType() }).Invoke(new object[] { node, this });
                //P element = (P)typeof(P).GetConstructors()[0].Invoke(new object[] { node, this });
                //V entity = (V)typeof(P).GetProperty("Entity").GetValue(element, null);
                entities.Add(element.Entity);
            }
            return entities;
        }

        #region 抽象方法

        /// <summary>
        /// （受保护）子类实现实体的产生
        /// </summary>
        /// <returns></returns>
        protected abstract T CreateEntity();

        #endregion 抽象方法
    }
}
