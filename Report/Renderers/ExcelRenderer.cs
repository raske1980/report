using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Report.Base;
using Report.Utils.HorizontalAlignmentMapper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TableStyle = Report.Base.TableStyle;

namespace Report.Renderers
{
    public class ExcelRenderer : RendererBase
    {
        public override byte[] Render(Base.Report report)
        {
            byte[] buffer;
            using (var outputMemoryStream = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(outputMemoryStream, SpreadsheetDocumentType.Workbook, true))
                {
                    // Add a WorkbookPart to the document.
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart.
                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    // Add Sheets to the Workbook.
                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                    // Append a new worksheet and associate it with the workbook.
                    var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    };

                    sheets.Append(sheet);

                    // Add a WorkbookStylesPart to the WorkbookPart
                    var stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    var shData = worksheetPart.Worksheet.GetFirstChild<SheetData>();



                    stylesPart.Stylesheet = InitializeDefaultStyles(stylesPart.Stylesheet);

                    uint styleId = 2; //обязательно 2 потому что уже додано два стиля по умолчанию с индексами 0 и 1;

                    var styles = report.Select(o => o.Style).ToList();

                    //нельзя перенести в отдельный метод :(
                    foreach (var s in styles)
                    {
                        if (s is TableStyle)
                        {
                            var style = s as TableStyle;

                            var border = new DocumentFormat.OpenXml.Spreadsheet.Border();

                            var leftBorder = new DocumentFormat.OpenXml.Spreadsheet.LeftBorder();
                            var rightBorder = new DocumentFormat.OpenXml.Spreadsheet.RightBorder();
                            var topBorder = new DocumentFormat.OpenXml.Spreadsheet.TopBorder();
                            var bottomBorder = new DocumentFormat.OpenXml.Spreadsheet.BottomBorder();

                            leftBorder.Style = SetLineType(style.BorderLine);
                            rightBorder.Style = SetLineType(style.BorderLine);
                            topBorder.Style = SetLineType(style.BorderLine);
                            bottomBorder.Style = SetLineType(style.BorderLine);

                            leftBorder.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            rightBorder.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            topBorder.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            bottomBorder.Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };

                            border.LeftBorder = leftBorder;
                            border.RightBorder = rightBorder;
                            border.TopBorder = topBorder;
                            border.BottomBorder = bottomBorder;

                            stylesPart.Stylesheet.Borders.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Border>(border);



                            var patternFill = new DocumentFormat.OpenXml.Spreadsheet.PatternFill
                            {
                                BackgroundColor = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor { Rgb = ToHexBinaryValue(style.Background) },
                                ForegroundColor = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = ToHexBinaryValue(style.Foreground) },
                                PatternType = SetPatternType(style.PatternType)
                            };

                            var fill = new DocumentFormat.OpenXml.Spreadsheet.Fill { PatternFill = patternFill };

                            stylesPart.Stylesheet.Fills.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Fill>(fill);

                            DocumentFormat.OpenXml.Spreadsheet.Font font = new DocumentFormat.OpenXml.Spreadsheet.Font()
                            {
                                FontName = new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = new StringValue { Value = s.FontName } },
                                FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = s.FontSize },
                                Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(s.FontColor) }
                            };
                            stylesPart.Stylesheet.Fonts.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Font>(font);
                        }
                        else
                        {

                            stylesPart.Stylesheet.Fills.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Fill>(new DocumentFormat.OpenXml.Spreadsheet.Fill());
                            stylesPart.Stylesheet.Borders.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Border>(new DocumentFormat.OpenXml.Spreadsheet.Border());
                            DocumentFormat.OpenXml.Spreadsheet.Font font;
                            if (s != null)
                            {
                                font = new DocumentFormat.OpenXml.Spreadsheet.Font
                                {
                                    FontName = new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = new StringValue { Value = s.FontName } },
                                    FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = s.FontSize },
                                    Color = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = ToHexBinaryValue(s.FontColor) }
                                };


                                foreach (var item in s.TextStyle)
                                {
                                    switch (item)
                                    {
                                        case TextStyleType.Bold: font.Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold { Val = true }; break;
                                        case TextStyleType.Italic: font.Italic = new DocumentFormat.OpenXml.Spreadsheet.Italic { Val = true }; break;
                                        case TextStyleType.Underline: font.Underline = new DocumentFormat.OpenXml.Spreadsheet.Underline { Val = DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Single }; break;
                                        case TextStyleType.Normal: break;
                                    }
                                }
                            }
                            else
                                font = new DocumentFormat.OpenXml.Spreadsheet.Font();

                            stylesPart.Stylesheet.Fonts.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Font>(font);
                        }


                        if (s != null)
                        {
                            var cellFormat = new DocumentFormat.OpenXml.Spreadsheet.CellFormat
                            {
                                BorderId = UInt32Value.ToUInt32(styleId),
                                FillId = UInt32Value.ToUInt32(styleId),
                                FontId = UInt32Value.ToUInt32(styleId)
                            };

                            if (s.Aligment != null)
                            {

                                var align = new Alignment()
                                {
                                    TextRotation = new UInt32Value(s.Aligment.Rotation),
                                    Horizontal = AlignmentMapper.MapHorizontalAligment(s.Aligment.HorizontalAligment),
                                    Vertical = AlignmentMapper.MapVerticalAligment(s.Aligment.VerticalAligment)
                                };

                                cellFormat.Append(align);
                            }

                            stylesPart.Stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(cellFormat);
                        }

                        styleId = styleId + 1;
                    }

                    foreach (var x in report)
                    {
                        if (x is ComplexHeaderCell)
                        {
                            var element = x as ComplexHeaderCell;

                            if (Render(element, styles, shData, worksheetPart, workbookpart))
                            {
                                continue;
                            }
                        }

                        if (x is TextElement)
                        {
                            var element = x as TextElement;

                            if (Render(element, styles, shData, workbookpart))
                            {
                                continue;
                            }
                        }

                        if (x is TableElement)
                        {
                            var element = x as TableElement;
                            if (Render(element, styles, shData, workbookpart))
                            {
                                continue;
                            }
                        }

                        if (x is ImageElement)
                        {
                            var element = x as ImageElement;

                            if (Render(element, styles, shData, workbookpart))
                            {
                                continue;
                            }
                        }
                    }

                    //var dateLine = shData.AppendChild(new Spreadsheet.Row());

                    //dateLine.AppendChild(new Spreadsheet.Cell()
                    //{
                    //    CellValue = new Spreadsheet.CellValue(_report.TimeStamp.Date.ToString()),
                    //    DataType = Spreadsheet.CellValues.String,
                    //    StyleIndex = 0
                    //});

                    workbookpart.Workbook.Save();
                }

                buffer = new byte[outputMemoryStream.Position];
                outputMemoryStream.Position = 0;
                outputMemoryStream.Read(buffer, 0, buffer.Length);
            }
            return buffer;
        }

        #region Render

        private static bool Render(ImageElement element, ICollection<Style> styles, SheetData shData, WorkbookPart workbookpart)
        {
#warning need to implement for ImageElement
            workbookpart.Workbook.Save();
            return false;
        }

        private static bool Render(ComplexHeaderCell element, ICollection<Style> styles, SheetData shData, WorksheetPart worksheetPart, WorkbookPart workbookpart)
        {
            uint styleId;
            styleId = GetStyleId(element.Style, styles);
            if (styleId > 0)
                element.Render(worksheetPart.Worksheet, styleId + 1);
            else
                element.Render(worksheetPart.Worksheet);
            workbookpart.Workbook.Save();

            return false;
        }

        private static bool Render(TableElement element, ICollection<Style> styles, SheetData shData, WorkbookPart workbookpart)
        {
            uint styleId;
            if (element.Table.Rows.Count == 0)
            {
                //throw new Exception("The table is empty");
                return true;
            }

            if (element.Headers.Count != 0)
            {
                styleId = GetStyleId(element.HeaderStyle, styles);

                var row = shData.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Row());

                foreach (var head in element.Headers)
                {
                    var cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                    {
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(head),
                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                        StyleIndex = UInt32Value.FromUInt32(styleId + 1)
                    };

                    row.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Cell>(cell);
                }
                workbookpart.Workbook.Save();
            }


            styleId = GetStyleId(element.TableStyle, styles);

            for (int i = 0; i < element.Table.Rows.Count; i++)
            {
                var row = shData.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Row());

                for (int j = 0; j < element.Table.Columns.Count; j++)
                {
                    var cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
                    {
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(element.Table.Rows[i][j].ToString()),
                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                        StyleIndex = UInt32Value.FromUInt32(styleId + 1)
                    };

                    row.AppendChild(cell);
                }
            }

            workbookpart.Workbook.Save();
            return false;
        }

        private static bool Render(TextElement element, ICollection<Style> styles, SheetData shData, WorkbookPart workbookpart)
        {
            uint styleId;
            var row = shData.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Row>(new DocumentFormat.OpenXml.Spreadsheet.Row());

            styleId = GetStyleId(element.Style, styles);

            var value = element.Value;

            if (element.Value == InternalContants.NewLine)
            {
                value = String.Empty;
            }

            if (element.Value == InternalContants.NewSection)
            {
                value = String.Empty;
            }

            var cell = new DocumentFormat.OpenXml.Spreadsheet.Cell
            {
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                StyleIndex = UInt32Value.FromUInt32(styleId + 1)
            };

            row.AppendChild(cell);

            workbookpart.Workbook.Save();

            return false;
        }

        #endregion

        private static BorderStyleValues SetLineType(BoderLine lineType)
        {
            switch (lineType)
            {
                case BoderLine.Thin: return DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin;
                case BoderLine.Medium: return DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium;
                case BoderLine.Thick: return DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick;

                default: return DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.None;
            }
        }

        private static Stylesheet InitializeDefaultStyles(Stylesheet stylesheet)
        {
            //!!!!!!!!!!!!!!!!!!!!!!!!!
            //Просьба не переделывать этот метод, это связано с непонятным багом в библиотеке Office XML SDK
            //!!!!!!!!!!!!!!!!!!!!!!!!!
            stylesheet = new DocumentFormat.OpenXml.Spreadsheet.Stylesheet();
            stylesheet.Borders = new DocumentFormat.OpenXml.Spreadsheet.Borders();
            stylesheet.Borders.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Border>(new DocumentFormat.OpenXml.Spreadsheet.Border());
            stylesheet.Borders.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Border>(new DocumentFormat.OpenXml.Spreadsheet.Border());

            stylesheet.Fills = new DocumentFormat.OpenXml.Spreadsheet.Fills();
            stylesheet.Fills.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Fill>(new DocumentFormat.OpenXml.Spreadsheet.Fill());
            stylesheet.Fills.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Fill>(new DocumentFormat.OpenXml.Spreadsheet.Fill());

            stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts();
            stylesheet.Fonts.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Font>(new DocumentFormat.OpenXml.Spreadsheet.Font());
            stylesheet.Fonts.AppendChild<DocumentFormat.OpenXml.Spreadsheet.Font>(new DocumentFormat.OpenXml.Spreadsheet.Font());

            stylesheet.CellFormats = new DocumentFormat.OpenXml.Spreadsheet.CellFormats();
            stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(new DocumentFormat.OpenXml.Spreadsheet.CellFormat());
            stylesheet.CellFormats.AppendChild<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(new DocumentFormat.OpenXml.Spreadsheet.CellFormat());

            return stylesheet;
        }

        private static PatternValues SetPatternType(PatternType patternType)
        {
            switch (patternType)
            {
                case PatternType.Solid: return DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid;

                default: return DocumentFormat.OpenXml.Spreadsheet.PatternValues.None;
            }
        }

        private static uint GetStyleId(Style style, ICollection<Style> styles)
        {
            uint styleId = 1;
            //Учитывая особенность построения списка стилей по умолчанию, 
            //а также особенность реализации и уточнения возможных ошибок,
            //удобно начинать индексацию с 1

            foreach (var s in styles)
            {
                if (s == null || style == null)
                    continue;

                if (style.Name == s.Name)
                {
                    return styleId;
                }

                if (styleId == styles.Count || style == null)
                {
                    styleId = 0;
                    return styleId;
                }
                styleId = styleId + 1;
            }

            return styleId = 0;
        }

        private static HexBinaryValue ToHexBinaryValue(System.Drawing.Color color)
        {
            return HexBinaryValue.FromString(ColorToRgb(color));
        }
    }
}
