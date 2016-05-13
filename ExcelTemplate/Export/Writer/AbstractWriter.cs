using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;
using NPOI.SS.UserModel;

namespace ExportTemplate.Export.Writer
{
    /// <summary>
    /// 修改目的：
    /// ①解耦，单一职责并可扩展；
    /// ②添加事件控制Render定制；
    /// ③抽提，实现减少对NPOI的依赖，快速切换到其它对Excel操作的API（难实现）
    /// </summary>

    public class CommonEventArgs<T, E> : EventArgs
    {
        protected T ExObject;
        public E Entity;
    }

    public class WriteEventArgs : CommonEventArgs<ISheet, IRuleEntity>
    {
        public ISheet ExSheet { get { return base.ExObject; } set { base.ExObject = value; } }
    }
    //public class ProductWriteEventArgs:CommonEventArgs<IWorkbook,IRuleEntity>
    //{
    //    public IWorkbook Workbook { get { return base.ExObject; } set { base.ExObject = value; } }
    //}
    //public delegate void ProductBeforeHandler(object sender, ProductWriteEventArgs args);
    public delegate void WriteHandler(object sender, WriteEventArgs args);
    public abstract class BaseWriter
    {
        private BaseEntity _entity;
        private BaseWriterContainer _parentWriter;

        public BaseWriterContainer ParentWriter
        {
            get { return _parentWriter; }
            set { _parentWriter = value; }
        }

        public BaseEntity Entity
        {
            get { return _entity; }
            set { _entity = value; }
        }


        public BaseWriter(BaseEntity entity, BaseWriterContainer parent) { this._entity = entity; this._parentWriter = parent; }

        public virtual void OnWrite(object sender, WriteEventArgs args)
        {
            if (BeforeWrite != null)
                BeforeWrite(sender, args);
            Writing(sender, args);
            if (AfterWrite != null)
                AfterWrite(sender, args);
        }
        public event WriteHandler BeforeWrite;
        public event WriteHandler AfterWrite;
        protected abstract void Writing(object sender, WriteEventArgs args);

        #region 计算模型

        protected bool _copyFill = true;

        /// <summary>
        /// 实际行索引
        /// </summary>
        public int RowIndex { get { return _entity.Location.RowIndex + AllIncreasedBefore; } }

        /// <summary>
        /// 实际列索引
        /// </summary>
        public int ColIndex { get { return _entity.Location.ColIndex; } }

        /// <summary>
        /// 列数
        /// </summary>
        public virtual int ColCount { get { return 0; } }

        /// <summary>
        /// 行数
        /// </summary>
        public virtual int RowCount { get { return 0; } }

        /// <summary>
        /// 是否以Copy模式填充
        /// </summary>
        public virtual bool CopyFill { get { return _copyFill; } set { _copyFill = value; } }

        /// <summary>
        /// 用作模板的行数
        /// </summary>
        public virtual int TempleteRows { get { return 1; } }

        /// <summary>
        /// 之前增长行数
        /// </summary>
        public int AllIncreasedBefore
        {
            get
            {
                return _parentWriter == null ? 0 : this.IncreasedInParent() + _parentWriter.AllIncreasedBefore;
                //return _parentWriter == null ? 0 : _parentWriter.IncreasedBefore(this) + _parentWriter.AllIncreasedBefore; 
            }
        }

        /// <summary>
        /// (当前)增长行数
        /// 注：对于容器对象需要覆盖此方法
        /// </summary>
        public virtual int IncreasedRows { get { return CopyFill && RowCount > TempleteRows ? RowCount - TempleteRows : 0; } }

        /// <summary>
        /// 同一父容器内前面对象增长行数
        /// 注：对于容器对象需要覆盖此方法
        /// </summary>
        /// <param name="subObject"></param>
        /// <returns></returns>
        public virtual int IncreasedInParent()
        {
            return _parentWriter == null || !_parentWriter.Components.Contains(this) ? 0 :
                _parentWriter.Components.Where(p => p.CopyFill && p.Entity.Location.RowIndex < this.Entity.Location.RowIndex).Sum(p => p.IncreasedRows);
        }

        /// <summary>
        /// 获取输出结点
        /// </summary>
        /// <returns></returns>
        public virtual IList<OutputNode> GetNodes() { return new OutputNode[0]; }
        #endregion 计算模型
    }

    public abstract class BaseWriterContainer : BaseWriter
    {
        private List<BaseWriter> _components = new List<BaseWriter>();
        public List<BaseWriter> Components { get { return _components; } }

        public BaseWriterContainer(BaseEntity entity, BaseWriterContainer parentWriter) : base(entity, parentWriter) { }

        public override void OnWrite(object sender, WriteEventArgs args)
        {
            PreWrite(args);
            base.OnWrite(this, args);
            EndWrite(args);
        }

        /// <summary>
        /// 初始化（第一步处理）
        /// </summary>
        /// <param name="args"></param>
        public virtual void PreWrite(WriteEventArgs args) { }
        /// <summary>
        /// 渲染后的结束工作（最后一步处理）
        /// </summary>
        protected virtual void EndWrite(WriteEventArgs args) { }

        protected override void Writing(object sender, WriteEventArgs args)
        {
            foreach (var item in _components)
            {
                WriteEventArgs newArgs = new WriteEventArgs() { ExSheet = args.ExSheet, Entity = item.Entity };
                item.OnWrite(item, newArgs);
            }
        }

        protected abstract void CreatingSubWriters();

        /// <summary>
        /// 构建子孙Writer
        /// </summary>
        public void CreateAllSubWriters()
        {
            _components.Clear();
            CreatingSubWriters();
            //构建孙子结点Writer
            foreach (var item in _components)
            {
                if (item is BaseWriterContainer)
                {
                    (item as BaseWriterContainer).CreateAllSubWriters();
                }
            }
        }
    }

}
