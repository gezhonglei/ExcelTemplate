using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExportTemplate.Export.Entity;

namespace ExportTemplate.Export.Element
{
    public interface IProductRuleGetter
    {
        ProductRuleElement ProductRuleElement { get; }
        BaseEntity CurrentEntity { get; }
    }
}
