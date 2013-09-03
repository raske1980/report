using System;
using System.Drawing;

namespace Report.Base
{
    public class Style
    {
        public String FontName { get; set; }
        public int FontSize { get; set; }
        public TextStyle FontType { get; set; }
        public Color FontColor { get; set; }
        public Aligment Aligment { get; set; }

        public DocumentTitle DocumentTitle { get; set; }

        public virtual string Name
        {
            get
            {
                //Внтуреннее имя стиля - используется для генерации классов css
                var result = String.Format("st{0}{1}{2}{3}{4}", FontName, FontSize, FontType, FontColor.Name, Aligment)
                    .Replace(" ", "_").GetHashCode().ToString().Replace("-", "x");

                return "s_" + result;
            }
        }

        public Style()
        {
            FontName = "Arial";
            FontSize = 11;
            FontType = TextStyle.Normal;
            FontColor = Color.Black;
            Aligment = new Aligment();
        }
    }

    /// <summary>
    /// пока не используется в реализации отчетов
    /// </summary>
    public class Aligment
    {
        public VerticalAligment VerticalAligment { get; set; }
        public HorizontalAligment HorizontalAligment { get; set; }
        public Boolean WrapText { get; set; }

        public Aligment()
        {
            VerticalAligment = VerticalAligment.Center;
            HorizontalAligment = HorizontalAligment.Left;
            WrapText = false;
        }

        public override string ToString()
        {
            return String.Format("{0} {1} {2}", VerticalAligment, HorizontalAligment, WrapText);
        }
    }
}
