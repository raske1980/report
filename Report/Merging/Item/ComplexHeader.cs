using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Report.Base;

namespace Report.Merging.Item
{
    public class ComplexHeader : List<ComplexHeaderCell>, IElement
    {
        public Style Style { get; set; }
    }
}
