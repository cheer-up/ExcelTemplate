using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ExcelTempl
{
    [Serializable]
    public class DicXML
    {
            [XmlAttribute]
            public DataGridViewCell Cell;
            [XmlAttribute]
            public TemplateRange Range;
        public DicXML(DataGridViewCell Cell, TemplateRange Range)
        {
            this.Cell = Cell;
            this.Range = Range;
        }
        public DicXML()
        {
        }
    }
}
