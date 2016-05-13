using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportTemplate.Export.Element
{
    /// <summary>
    /// 解析异常
    /// </summary>
    [Serializable]
    public class ParseErrorException : Exception
    {
        public ParseErrorException() : base() { }
        public ParseErrorException(string message) : base(message) { }
    }
}
