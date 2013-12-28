using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Report.Base;

namespace Report.Utils.HorizontalAlignmentMapper
{
    public static class ExcelHorizontalAlignmentMapper
    {
        public static EnumValue<HorizontalAlignmentValues> Map(HorizontalAligment alignToMap)
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
    }
}
