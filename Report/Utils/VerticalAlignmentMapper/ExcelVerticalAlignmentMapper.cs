using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Report.Base;

namespace Report.Utils.VerticalAlignmentMapper
{
    public static class ExcelVerticalAlignmentMapper
    {
        public static EnumValue<VerticalAlignmentValues> Map(VerticalAligment alignToMap)
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
