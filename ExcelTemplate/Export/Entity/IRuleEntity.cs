using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Entity
{
    public interface IRuleEntity
    {
        //T Clone<T>(ProductRule productRule, BaseContainer container) where T : IRuleEntity;
    }

    /// <summary>
    /// 用于实体规则的复制
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface ICloneable<T>
    {
        T CloneEntity(ProductRule productRule, BaseEntity container);
    }

}
