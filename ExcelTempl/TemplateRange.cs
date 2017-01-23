using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTempl
{[Serializable]
    public class TemplateRange
    {
        public Point FirstCell = new Point(0, 0);
        public Point EndCell = new Point(0, 0);
      
        public Point EndCell1
        {
            get
            {
                return EndCell;
            }

            set
            {
                EndCell = value;
            }
        }
        public Point FirstCell1
        {
            get
            {
                return FirstCell;
            }

            set
            {
                FirstCell = value;
            }
        }
        public TemplateRange(Point FirstCell, Point EndCell)
        {
            this.EndCell = EndCell;
            this.FirstCell = FirstCell;
        }
        public TemplateRange(Point FirstCell)
        {
            this.EndCell = Point.Empty;
            this.FirstCell = FirstCell;
        }
        public TemplateRange()
        {

        }
        public string EndCellString()
        {
            if (EndCell.IsEmpty)
            { return "INF"; }
            return ((char)('A' + (char)this.EndCell.Y)).ToString() + this.EndCell.X.ToString();
        }
        public string FirstCellString()
        {

            return ((char)('A' + (char)this.FirstCell.Y)).ToString() + this.FirstCell.X.ToString();
        }
    }
}
