using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTempl
{[Serializable]
    public class Config
    {
       public Dictionary<Point, TemplateRange> dic;
       public object[] header;
        public Config(Dictionary<Point, TemplateRange> dic, object[] header)
        {
            this.dic = dic;
            this.header = header;
        }
    }
}
