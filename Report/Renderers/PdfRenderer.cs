using iTextSharp.text;
using iTextSharp.text.pdf;
using Report.Base;
using System;
using System.Collections.Generic;
using System.IO;

namespace Report.Renderers
{
    /// <summary>
    /// 
    /// </summary>
    public class PdfRenderer : RendererBase
    {
        public override byte[] Render(Base.Report report)
        {
            byte[] buffer;

            using (var outputMemoryStream = new MemoryStream())
            {
                var pdfDocument = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 25, 15, 10);

                var arialuniTff = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                FontFactory.Register(arialuniTff);

                //FontFactory.Register(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));

                var pdfWriter = PdfWriter.GetInstance(pdfDocument, outputMemoryStream);

                pdfWriter.CloseStream = false;
                pdfDocument.Open();

                //pdfDocument.NewPage();
                pdfDocument.Add(new Paragraph(new Phrase(" ")));

                foreach (var x in report)
                {
                    if (x is TextElement)
                    {
                        var element = x as TextElement;

                        if (Render(element, pdfDocument, arialuniTff))
                        {
                            continue;
                        }
                    }

                    if (x is TableElement)
                    {
                        var element = x as TableElement;
                        if (Render(element, pdfDocument, arialuniTff))
                        {
                            continue;
                        }
                    }

                    if (x is ImageElement)
                    {
                        var element = x as ImageElement;

                        if (Render(element, pdfDocument))
                        {
                            continue;
                        }
                    }

                    if (x is ComplexHeaderCell)
                    {
                        var element = x as ComplexHeaderCell;

                        if (Render(element, pdfDocument))
                        {
                            continue;
                        }
                    }

                    if (x is ComplexHeader)
                    {
                        var element = x as ComplexHeader;

                        if (Render(element, pdfDocument))
                        {
                            continue;
                        }
                    }
                }

                pdfDocument.Close();
                pdfDocument.CloseDocument();

                buffer = new byte[outputMemoryStream.Position];
                outputMemoryStream.Position = 0;
                outputMemoryStream.Read(buffer, 0, buffer.Length);
            }

            return buffer;
        }

        private bool Render(ComplexHeaderCell element, Document pdfDocument)
        {
            return Render(new TextElement()
            {
                Style = new Style
                {
                    FontColor = System.Drawing.Color.Red,
                    TextStyle = new List<TextStyleType> { TextStyleType.Bold }
                },
                Value = "Rended for ComplexHeaderCell not implemented."
            }, pdfDocument, "arialuniTff");
        }

        #region Render

        private bool Render(ImageElement element, Document pdfDocument)
        {
            if ((element.Url == null || String.IsNullOrEmpty(element.Url.ToString())) && element.Image == null)
            {
                return true;
            }

            var image = iTextSharp.text.Image.GetInstance(element.Image, element.Image.RawFormat);
            pdfDocument.Add(image);
            return false;
        }

        private bool Render(TableElement element, Document pdfDocument, string fontName)
        {
            if (element.Table.Rows.Count == 0)
            {
                return true;
                //throw new Exception("The table is empty");
            }

            var table = new PdfPTable(element.Table.Columns.Count);

            if (element.Headers.Count != 0)
            {
                var headerFont = FontFactory.GetFont(fontName, System.Text.Encoding.GetEncoding(855).BodyName, BaseFont.EMBEDDED);
                headerFont = PdfFontStyle(headerFont, element.Style.TextStyle);
                headerFont.Size = (float)element.Style.FontSize;
                headerFont.Color = new BaseColor(element.Style.FontColor);

                foreach (var head in element.Headers)
                {
                    var cell = new PdfPCell
                    {
                        Phrase = new Phrase(head, headerFont),
                        BackgroundColor = new BaseColor(element.HeaderStyle.Foreground),
                        BorderColor = new BaseColor(element.HeaderStyle.BorderColor),
                        BorderWidth = PdfBorderWidth(element.HeaderStyle.BorderLine),
                        NoWrap = element.HeaderStyle.Aligment.WrapText
                    };

                    table.AddCell(cell);
                }
            }


            var tableFont = FontFactory.GetFont(fontName, System.Text.Encoding.GetEncoding(855).BodyName, BaseFont.EMBEDDED);
            tableFont = PdfFontStyle(tableFont, element.Style.TextStyle);
            tableFont.Size = (float)element.Style.FontSize;
            tableFont.Color = new BaseColor(element.Style.FontColor);

            for (int i = 0; i < element.Table.Rows.Count; i++)
            {
                for (int j = 0; j < element.Table.Columns.Count; j++)
                {
                    var cell = new PdfPCell
                    {
                        Phrase = new Phrase(element.Table.Rows[i][j].ToString(), tableFont),
                        BackgroundColor = new BaseColor(element.TableStyle.Foreground),
                        BorderColor = new BaseColor(element.TableStyle.BorderColor),
                        BorderWidth = PdfBorderWidth(element.TableStyle.BorderLine),
                        NoWrap = element.TableStyle.Aligment.WrapText,
                        HorizontalAlignment = PdfHorizontalAlignment(element.TableStyle.Aligment.HorizontalAligment),
                        VerticalAlignment = PdfHorizontaVerticallAlignment(element.TableStyle.Aligment.VerticalAligment)
                    };

                    table.AddCell(cell);
                }
            }

            pdfDocument.Add(table);
            return false;
        }

        private bool Render(TextElement element, Document pdfDocument, string fontName)
        {
            if (element.Value == InternalContants.NewLine || element.Value == InternalContants.NewSection)
            {
                pdfDocument.Add(new Paragraph(new Phrase(" "))); //так должно быть!!!
                return true;
            }

            var font = FontFactory.GetFont(fontName, System.Text.Encoding.GetEncoding(855).BodyName, BaseFont.EMBEDDED);
            font = PdfFontStyle(font, element.Style.TextStyle);
            font.Size = (float)element.Style.FontSize;
            font.Color = new BaseColor(element.Style.FontColor);


            var paragraph = new Paragraph(element.Value, font);
            paragraph.Alignment = PdfHorizontalAlignment(element.Style.Aligment.HorizontalAligment);
            pdfDocument.Add(paragraph);
            return false;
        }

        private bool Render(ComplexHeader element, Document pdfDocument)
        {
            return Render(new TextElement()
            {
                Style = new Style
                {
                    FontColor = System.Drawing.Color.Red,
                    TextStyle = new List<TextStyleType> { TextStyleType.Bold }
                },
                Value = "Rended for ComplexHeader not implemented."
            }, pdfDocument, "arialuniTff");
        }

        #endregion

        private static int PdfHorizontaVerticallAlignment(VerticalAligment verticalAligment)
        {
            switch (verticalAligment)
            {
                case VerticalAligment.Top: return Element.ALIGN_TOP;
                case VerticalAligment.Center: return Element.ALIGN_CENTER;
                case VerticalAligment.Justify: return Element.ALIGN_JUSTIFIED;
                default: return Element.ALIGN_BOTTOM;
            }
        }

        private static int PdfHorizontalAlignment(HorizontalAligment horizontalAligment)
        {
            switch (horizontalAligment)
            {
#warning Не все варианты перечислены (надо экспериментировать)
                case HorizontalAligment.Right: return Element.ALIGN_RIGHT;
                case HorizontalAligment.Center: return Element.ALIGN_CENTER;
                default: return Element.ALIGN_LEFT;
            }
        }

        private static float PdfBorderWidth(BoderLine boderLine)
        {
            switch (boderLine)
            {
                case BoderLine.Thin: return 0.5f;
                case BoderLine.Medium: return 1f;
                case BoderLine.Thick: return 1.5f;
                default: return 0f;
            }

        }

        private static Font PdfFontStyle(iTextSharp.text.Font font, IEnumerable<TextStyleType> textStyle)
        {
            if (textStyle != null)
                foreach (var item in textStyle)
                {
                    switch (item)
                    {
                        case TextStyleType.Bold: font.SetStyle(iTextSharp.text.Font.BOLD); return font;
                        case TextStyleType.Italic: font.SetStyle(iTextSharp.text.Font.ITALIC); return font;
                        case TextStyleType.Underline: font.SetStyle(iTextSharp.text.Font.UNDERLINE); return font;
                        default: font.SetStyle(iTextSharp.text.Font.NORMAL); return font;
                    }
                }
            return font;
        }

    }
}
