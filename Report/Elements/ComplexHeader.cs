using System.Collections.Generic;

namespace Report.Base
{
    public class ComplexHeader : List<ComplexHeaderCell>, IElement
    {
        public Style Style { get; set; }
    }
}
