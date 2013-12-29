using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Report.Base;

namespace Report.Utils.HorizontalAlignmentMapper
{
    public static class AlignmentMapper
    {
        public static EnumValue<HorizontalAlignmentValues> MapHorizontalAligment(HorizontalAligment alignToMap)
        {
            if (alignToMap != null)
            {
                switch (alignToMap)
                {
                    case HorizontalAligment.General:
                        return HorizontalAlignmentValues.General;
                    case HorizontalAligment.Left:
                        return HorizontalAlignmentValues.Left;
                    case HorizontalAligment.Center:
                        return HorizontalAlignmentValues.Center;
                    case HorizontalAligment.Right:
                        return HorizontalAlignmentValues.Right;
                    case HorizontalAligment.Fill:
                        return HorizontalAlignmentValues.Fill;
                    case HorizontalAligment.Justify:
                        return HorizontalAlignmentValues.Justify;
                    case HorizontalAligment.CenterContinuous:
                        return HorizontalAlignmentValues.CenterContinuous;
                    case HorizontalAligment.Distributed:
                        return HorizontalAlignmentValues.Distributed;
                    default:
                        break;
                }
            }

            return HorizontalAlignmentValues.Center;
        }

        public static EnumValue<VerticalAlignmentValues> MapVerticalAligment(VerticalAligment alignToMap)
        {
            if (alignToMap != null)
            {
                switch (alignToMap)
                {
                    case VerticalAligment.Top:
                        return VerticalAlignmentValues.Top;
                    case VerticalAligment.Center:
                        return VerticalAlignmentValues.Center;
                    case VerticalAligment.Bottom:
                        return VerticalAlignmentValues.Bottom;
                    case VerticalAligment.Justify:
                        return VerticalAlignmentValues.Justify;
                    case VerticalAligment.Distributed:
                        return VerticalAlignmentValues.Distributed;
                    default:
                        return VerticalAlignmentValues.Center;
                }
            }

            return VerticalAlignmentValues.Center;
        }

    }
}
