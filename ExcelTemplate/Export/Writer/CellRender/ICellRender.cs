using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace ExportTemplate.Export.Writer.CellRender
{
    public interface ICellRender
    {
        void Render(ICell cell);
    }
}
