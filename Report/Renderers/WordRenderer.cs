using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Report.Base;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Style = Report.Base.Style;
using TableStyle = Report.Base.TableStyle;

namespace Report.Renderers
{
    /// <summary>
    /// 
    /// </summary>
    public class WordRenderer : RendererBase
    {
        public override byte[] Render(Base.Report report)
        {

            #warning Need to extend text style support

            var memoryStream = new MemoryStream();

            using (var document = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                document.AddMainDocumentPart();
                document.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                var stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                var body = document.MainDocumentPart.Document.AppendChild<DocumentFormat.OpenXml.Wordprocessing.Body>(new DocumentFormat.OpenXml.Wordprocessing.Body());


                var styles = report.Select(o => o.Style).ToList();
                stylesPart.Styles = WordStyles(styles);
                stylesPart.Styles.Save();

                foreach (var x in report)
                {
                    if (x is TextElement)
                    {
                        var element = x as TextElement;

                        if (Render(element, document, body))
                        {
                            continue;
                        }
                    }

                    if (x is TableElement)
                    {
                        var element = x as TableElement;

                        if (Render(element, document, body))
                        {
                            continue;
                        }
                    }

                    if (x is ImageElement)
                    {
                        var element = x as ImageElement;

                        if (Render(element, document, body))
                        {
                            continue;
                        }
                    }

                    if (x is ComplexHeaderCell)
                    {
                        var element = x as ComplexHeaderCell;

                        if (Render(element, document, body))
                        {
                            continue;
                        }
                    }

                    if (x is ComplexHeader)
                    {
                        var element = x as ComplexHeader;

                        if (Render(element, document, body))
                        {
                            continue;
                        }
                    }

                    document.MainDocumentPart.Document.Save();
                }

            }

            return memoryStream.ToArray();
        }

        #region Render

        private static bool Render(ComplexHeaderCell element, WordprocessingDocument document, Body body)
        {
            return Render(new TextElement()
            {
                Style = new Style(),
                Value = "Rended for ComplexHeaderCell not implemented."
            }, document, body);
        }

        private static bool Render(ComplexHeader element, WordprocessingDocument document, Body body)
        {
            return Render(new TextElement()
            {
                Style = new Style(),
                Value = "Rended for ComplexHeader not implemented."
            }, document, body);
        }

        private static bool Render(ImageElement element, WordprocessingDocument document, Body body)
        {

#warning Image element not implemented
            //body.AppendChild<Wordprocessing.Paragraph>();
            document.MainDocumentPart.Document.Save();
            return false;
        }

        private static bool Render(TextElement element, WordprocessingDocument document, Body body)
        {
            if (element.Value == InternalContants.NewLine || element.Value == InternalContants.NewSection)
            {
                body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                return true;
            }

            var text = new DocumentFormat.OpenXml.Wordprocessing.Text(element.Value);
            var run = new DocumentFormat.OpenXml.Wordprocessing.Run(text);
            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

            paragraph.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
            {
                ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId()
                {
                    Val = WordStyleIdFromStyleName(document, element.Style.Name)
                },
                WordWrap = new DocumentFormat.OpenXml.Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                TextAlignment =
                    new DocumentFormat.OpenXml.Wordprocessing.TextAlignment
                    {
                        Val = DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignmentValues.Auto
                    },
            };

            body.Append(paragraph);
            return false;
        }

        private static bool Render(TableElement element, WordprocessingDocument document, Body body)
        {
            if (element.Table.Rows.Count == 0)
            {
                return true;

                //throw new Exception("The table is empty");
            }

            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();

            if (element.Headers.Count != 0)
            {
                var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                foreach (var head in element.Headers)
                {
                    var text = new DocumentFormat.OpenXml.Wordprocessing.Text(head.ToString());
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(text);
                    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

                    paragraph.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                    {
                        ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId()
                        {
                            Val = WordStyleIdFromStyleName(document, element.HeaderStyle.Name)
                        },
                        WordWrap = new DocumentFormat.OpenXml.Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                        TextAlignment =
                            new DocumentFormat.OpenXml.Wordprocessing.TextAlignment
                            {
                                Val = DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignmentValues.Auto
                            },
                    };

                    var cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(paragraph);

                    DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellprop = WordCellProperties(element.HeaderStyle);
                    cell.Append(cellprop);

                    row.Append(cell);
                }
                table.Append(row);
            }


            for (int i = 0; i < element.Table.Rows.Count; i++)
            {
                var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                for (int j = 0; j < element.Table.Columns.Count; j++)
                {
                    var text = new DocumentFormat.OpenXml.Wordprocessing.Text(element.Table.Rows[i][j].ToString());
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(text);
                    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

                    paragraph.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
                    {
                        ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId()
                        {
                            Val = WordStyleIdFromStyleName(document, element.TableStyle.Name)
                        },
                        WordWrap = new DocumentFormat.OpenXml.Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                        TextAlignment =
                            new DocumentFormat.OpenXml.Wordprocessing.TextAlignment
                            {
                                Val = DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignmentValues.Auto
                            },
                    };

                    var cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(paragraph);

                    var cellprop = WordCellProperties(element.TableStyle);

                    cell.Append(cellprop);

                    row.Append(cell);
                }
                table.Append(row);
            }
            body.AppendChild<DocumentFormat.OpenXml.Wordprocessing.Table>(table);
            return false;
        }

        #endregion

        private static Styles WordStyles(IEnumerable<Style> stylesList)
        {
            var styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            int countId = 1;

            foreach (var s in stylesList)
            {
                var paragraphStyle = new DocumentFormat.OpenXml.Wordprocessing.Style()
                {
                    Type = DocumentFormat.OpenXml.Wordprocessing.StyleValues.Paragraph,
                    StyleId = "p" + countId,
                    CustomStyle = true,
                    Default = false,
                    StyleName = new DocumentFormat.OpenXml.Wordprocessing.StyleName { Val = s.Name }
                };

                //if (s.DocumentTitle == DocumentTitle.None)
                //{
                var styleRunProperties = new DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties
                {
                    Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = ColorToRgb(s.FontColor) },
                    RunFonts = new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = s.FontName },
                    FontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = (s.FontSize * 2).ToString() },
                    Bold = s.TextStyle.Contains(TextStyleType.Bold) ? new DocumentFormat.OpenXml.Wordprocessing.Bold() : null,
                    Italic = s.TextStyle.Contains(TextStyleType.Italic) ? new DocumentFormat.OpenXml.Wordprocessing.Italic() : null,
                    Underline = s.TextStyle.Contains(TextStyleType.Underline) ? new DocumentFormat.OpenXml.Wordprocessing.Underline() : null
                };

                var helpStyleRunProperties = WordFontStyle(s.TextStyle);

                paragraphStyle.Append(styleRunProperties);
                paragraphStyle.Append(helpStyleRunProperties);
                //}
                //else
                //{
                //    var styleRunProperties = new Wordprocessing.StyleRunProperties
                //        {
                //            RunFonts = new Wordprocessing.RunFonts() {Ascii = "Arial"},
                //            FontSize = new Wordprocessing.FontSize() {Val = GetFontSizeFromDocumentTitle(s.DocumentTitle)},
                //        };

                //    paragraphStyle.Append(styleRunProperties);
                //}

                styles.Append(paragraphStyle);

                countId++;
            }

            return styles;
        }

        private static StyleRunProperties WordFontStyle(IEnumerable<TextStyleType> textStyle)
        {
            var property = new DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties();
            foreach (var item in textStyle)
            {
                switch (item)
                {
                    case TextStyleType.Bold: property.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold(); break;
                    case TextStyleType.Italic: property.Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic(); break;
                    case TextStyleType.Underline: property.Underline = new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single }; break;
                    default: break;
                }
            }

            return property;
        }

        private static TableCellProperties WordCellProperties(TableStyle style)
        {
            var leftBorder = new DocumentFormat.OpenXml.Wordprocessing.LeftBorder();
            var rightBorder = new DocumentFormat.OpenXml.Wordprocessing.RightBorder();
            var topBorder = new DocumentFormat.OpenXml.Wordprocessing.TopBorder();
            var bottomBorder = new DocumentFormat.OpenXml.Wordprocessing.BottomBorder();
            var insideHorizontalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder();
            var insideVerticalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder();

            leftBorder.Val = WordLineType(style.BorderLine);
            rightBorder.Val = WordLineType(style.BorderLine);
            topBorder.Val = WordLineType(style.BorderLine);
            bottomBorder.Val = WordLineType(style.BorderLine);
            insideHorizontalBorder.Val = WordLineType(style.BorderLine);
            insideVerticalBorder.Val = WordLineType(style.BorderLine);

            leftBorder.Color = ColorToRgb(style.BorderColor);
            rightBorder.Color = ColorToRgb(style.BorderColor);
            topBorder.Color = ColorToRgb(style.BorderColor);
            bottomBorder.Color = ColorToRgb(style.BorderColor);
            insideHorizontalBorder.Color = ColorToRgb(style.BorderColor);
            insideVerticalBorder.Color = ColorToRgb(style.BorderColor);

            leftBorder.Size = WordBorderWaight(style.BorderLine);
            rightBorder.Size = WordBorderWaight(style.BorderLine);
            topBorder.Size = WordBorderWaight(style.BorderLine);
            bottomBorder.Size = WordBorderWaight(style.BorderLine);
            insideHorizontalBorder.Size = WordBorderWaight(style.BorderLine);
            insideVerticalBorder.Size = WordBorderWaight(style.BorderLine);


            DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellprop = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();
            DocumentFormat.OpenXml.Wordprocessing.Shading shading = new DocumentFormat.OpenXml.Wordprocessing.Shading();
            shading.Val = DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues.Clear;
            shading.Fill = ColorToRgb(style.Foreground);

            DocumentFormat.OpenXml.Wordprocessing.TableCellBorders tableCellBorders = new DocumentFormat.OpenXml.Wordprocessing.TableCellBorders();
            tableCellBorders.LeftBorder = leftBorder;
            tableCellBorders.RightBorder = rightBorder;
            tableCellBorders.TopBorder = topBorder;
            tableCellBorders.BottomBorder = bottomBorder;
            tableCellBorders.InsideHorizontalBorder = insideHorizontalBorder;
            tableCellBorders.InsideVerticalBorder = insideVerticalBorder;

            cellprop.Shading = shading;
            cellprop.TableCellBorders = tableCellBorders;

            return cellprop;
        }

        private static string WordStyleIdFromStyleName(WordprocessingDocument document, string styleName)
        {
            var stylePart = document.MainDocumentPart.StyleDefinitionsPart;

            //string styleId = stylePart.Styles.Descendants<Wordprocessing.StyleName>()
            //    .Where(s => s.Val.Value.Equals(styleName) &&
            //        (((Wordprocessing.Style)s.Parent).Type == Wordprocessing.StyleValues.Paragraph))
            //    .Select(n => ((Wordprocessing.Style)n.Parent).StyleId).FirstOrDefault();

            string styleId = (from o in stylePart.Styles.Descendants<DocumentFormat.OpenXml.Wordprocessing.StyleName>()
                              where o.Val.Value == styleName && (((DocumentFormat.OpenXml.Wordprocessing.Style)o.Parent).Type == DocumentFormat.OpenXml.Wordprocessing.StyleValues.Paragraph)
                              select ((DocumentFormat.OpenXml.Wordprocessing.Style)o.Parent).StyleId).FirstOrDefault();

            return styleId;
        }

        private static EnumValue<BorderValues> WordLineType(BoderLine borderLine)
        {
            switch (borderLine)
            {
                case BoderLine.Thin: return DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single;
                case BoderLine.Medium: return DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single;
                case BoderLine.Thick: return DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single;
                default: return DocumentFormat.OpenXml.Wordprocessing.BorderValues.None;
            }
        }

        private static UInt32Value WordBorderWaight(BoderLine borderLine)
        {
            switch (borderLine)
            {
                case BoderLine.Thin: return (UInt32Value)4U;
                case BoderLine.Medium: return (UInt32Value)16U;
                case BoderLine.Thick: return (UInt32Value)24U;
                default: return (UInt32Value)0U;
            }
        }

    }
}
