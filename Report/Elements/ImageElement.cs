using System;
using System.Drawing;

namespace Report.Base
{
    public class ImageElement : IElement
    {
        public Image Image { get; set; }
        public Uri Url { get; set; }

        public ImageElement()
        {
        }

        public ImageElement(string url)
        {
            Url = new Uri(url);
        }

        public ImageElement(System.Drawing.Image image)
        {
            Image = image;
        }

        public Style Style { get; private set; }
    }
}
