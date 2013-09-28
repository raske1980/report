namespace Report.Base
{
    public enum PatternType
    {
        None,
        Solid,
    }

    public enum DocumentTitle
    {
        None = 0,
        Title = 1,
        Heading1 = 2,
        Heading2 = 3,
        Heading3 = 4,
        Heading4 = 5,
        SubTitle = 6
    }

    public enum BoderLine
    {
        None,
        Thin,
        Medium,
        Thick,
    }

    public enum TextStyle
    {
        Normal,
        Bold,
        Italic,
        Underline
    }

    public enum VerticalAligment
    {
        Top,
        Center,
        Bottom,
        Justify
    }

    public enum HorizontalAligment
    {
        General,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed
    }
}
