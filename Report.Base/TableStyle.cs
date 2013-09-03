using System;
using System.Drawing;

namespace Report.Base
{
    public class TableStyle : Style
    {
        public Color BorderColor { get; set; }
        public BoderLine BorderLine { get; set; }

        public Color Foreground { get; set; }
        public Color Background { get; set; }
        public PatternType PatternType { get; set; }

        public TableStyle()
        {
            PatternType = PatternType.None;

            BorderColor = Color.Black;
            BorderLine = BoderLine.None;
            Background = Color.White;
            Foreground = Color.White;
        }

        public override string Name
        {
            get
            {
                //Внтуреннее имя стиля - используется для генерации классов css
                var result = String.Format("t{0}{1}{2}{3}{4}{5}", base.Name, BorderColor, BorderLine,
                                           Foreground.Name, Background.Name, PatternType).GetHashCode().ToString().Replace("-", "x");
                return "st_" + result;
            }
        }
    }
}
