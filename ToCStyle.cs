using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTProgressMaker
{
    public class ToCStyle
    {
        public ToCStyle()
        {
            Type = Types.Horizontal;
            Style = Styles.Gradient;
            RTL = false;
            FirstSlide = false;
        }

        public enum Types { Horizontal, Vertical };
        public enum Styles { Solid, Gradient };

        public Types Type { get; set; }
        public Styles Style { get; set; }
        public bool RTL { get; set;  }
        public bool FirstSlide { get; set; }
        public bool SlideNumbers { get; set; }
        public bool IgnoreLastSection { get; set;  }
    }
}
