using System;
using System.Collections.Generic;

namespace Report.Base
{
    public class Report : List<IElement>
    {
        public DateTime TimeStamp { get; set; }

        public Report()
        {
        }
    }
}
