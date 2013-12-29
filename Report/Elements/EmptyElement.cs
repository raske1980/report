using System;

namespace Report.Base
{
    public class EmptyElement : IElement
    {
        public EmptyElement()
        {
        }

        public Style Style { get; set; }
    }
}
