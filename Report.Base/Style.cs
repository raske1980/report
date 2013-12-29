using System;
using System.Drawing;
using System.Collections.Generic;
using System.Text;

namespace Report.Base
{
    public class Style
    {
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = (FontName != null ? FontName.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ FontSize;
                //hashCode = (hashCode * 397) ^ (int)TextStyle;
                hashCode = (hashCode * 397) ^ FontColor.GetHashCode();
                hashCode = (hashCode * 397) ^ (Aligment != null ? Aligment.GetHashCode() : 0);
                return hashCode;
            }
        }

        public String FontName { get; set; }
        public int FontSize { get; set; }

        private List<TextStyleType> _textStyle;

        public List<TextStyleType> TextStyle
        {
            get { return _textStyle ?? (_textStyle = new List<TextStyleType>()); }
            set
            { _textStyle = value; }
        }

        private string TextStyleString
        {
            get
            {
                var sb = new StringBuilder();
                foreach (var item in TextStyle)
                    sb.Append(item);
                return sb.ToString();
            }
        }
        public Color FontColor { get; set; }
        public Aligment Aligment { get; set; }

        public DocumentTitle DocumentTitle { get; set; }

        public virtual string Name
        {
            get
            {
                //Внтуреннее имя стиля - используется для генерации классов css
                var result = String.Format("st{0}{1}{2}{3}{4}{5}", FontName, FontSize, TextStyleString, FontColor.Name, Aligment, DocumentTitle)
                    .Replace(" ", "_").GetHashCode().ToString().Replace("-", "x");

                return "s_" + result;
            }
        }

        public Style()
        {
            FontName = "Arial";
            FontSize = 11;
            TextStyle.Add(TextStyleType.Normal);
            FontColor = Color.Black;
            Aligment = new Aligment();
            DocumentTitle = DocumentTitle.None;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Style)obj);
        }

        protected bool Equals(Style other)
        {
            return string.Equals(FontName, other.FontName) && FontSize == other.FontSize && TextStyle == other.TextStyle && FontColor.Equals(other.FontColor) && Equals(Aligment, other.Aligment);
        }
    }

    /// <summary>
    /// пока не используется в реализации отчетов
    /// </summary>
    public class Aligment
    {
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = (int)VerticalAligment;
                hashCode = (hashCode * 397) ^ (int)HorizontalAligment;
                hashCode = (hashCode * 397) ^ WrapText.GetHashCode();
                return hashCode;
            }
        }

        public VerticalAligment VerticalAligment { get; set; }
        public HorizontalAligment HorizontalAligment { get; set; }
        public Boolean WrapText { get; set; }

        public Aligment()
        {
            VerticalAligment = VerticalAligment.Center;
            HorizontalAligment = HorizontalAligment.Left;
            WrapText = false;
        }

        public uint Rotation { get; set; }

        public override string ToString()
        {
            return String.Format("{0} {1} {2} {3}", VerticalAligment, HorizontalAligment, WrapText, Rotation);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Aligment)obj);
        }

        protected bool Equals(Aligment other)
        {
            return VerticalAligment == other.VerticalAligment && HorizontalAligment == other.HorizontalAligment && WrapText.Equals(other.WrapText);
        }
    }
}
