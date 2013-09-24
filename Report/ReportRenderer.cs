﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Report.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Style = Report.Base.Style;
using TableStyle = Report.Base.TableStyle;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Report
{
    /// <summary>
    /// 
    /// </summary>
    public class ReportRenderer : ReportRendererBase
    {
        private readonly Report.Base.Report _report;

        public ReportRenderer(Report.Base.Report report)
        {
            _report = report;
        }

        public override byte[] ToPdf()
        {
            byte[] buffer;

            using (var outputMemoryStream = new MemoryStream())
            {
                using (var pdfDocument = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 25, 15, 10))
                {
                    var arialuniTff = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                    FontFactory.Register(arialuniTff);

                    //FontFactory.Register(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));

                    var pdfWriter = PdfWriter.GetInstance(pdfDocument, outputMemoryStream);

                    pdfWriter.CloseStream = false;
                    pdfDocument.Open();

                    foreach (var x in _report)
                    {
                        if (x is TextElement)
                        {
                            var element = x as TextElement;

                            if (element.Value == InternalContants.NewLine || element.Value == InternalContants.NewSection)
                            {
                                pdfDocument.Add(new Paragraph(new Phrase(" "))); //так должно быть!!!
                                continue;
                            }

                            var font = FontFactory.GetFont(arialuniTff, System.Text.Encoding.GetEncoding(855).BodyName, BaseFont.EMBEDDED);
                            font = PdfFontStyle(font, element.Style.FontType);
                            font.Size = (float)element.Style.FontSize;
                            font.Color = new BaseColor(element.Style.FontColor);


                            var paragraph = new Paragraph(element.Value, font);
                            paragraph.Alignment = PdfHorizontalAlignment(element.Style.Aligment.HorizontalAligment);
                            pdfDocument.Add(paragraph);
                        }

                        if (x is TableElement)
                        {
                            var element = x as TableElement;
                            if (element.Table.Rows.Count == 0)
                            {
                                continue;
                                //throw new Exception("The table is empty");
                            }

                            var table = new PdfPTable(element.Table.Columns.Count);

                            if (element.Headers.Count != 0)
                            {
                                var headerFont = FontFactory.GetFont(arialuniTff, System.Text.Encoding.GetEncoding(1251).BodyName, BaseFont.EMBEDDED);
                                headerFont = PdfFontStyle(headerFont, element.Style.FontType);
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


                            var tableFont = FontFactory.GetFont(arialuniTff, System.Text.Encoding.GetEncoding(1251).BodyName, BaseFont.EMBEDDED);
                            tableFont = PdfFontStyle(tableFont, element.Style.FontType);
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
                        }

                        if (x is ImageElement)
                        {
                            var element = x as ImageElement;
                            Image image = Image.GetInstance(element.Image, element.Image.RawFormat);
                            pdfDocument.Add(image);
                        }
                    }
                }

                buffer = new byte[outputMemoryStream.Position];
                outputMemoryStream.Position = 0;
                outputMemoryStream.Read(buffer, 0, buffer.Length);
            }

            return buffer;
        }

        public override byte[] ToExcel()
        {
            byte[] buffer;
            using (var outputMemoryStream = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(outputMemoryStream, SpreadsheetDocumentType.Workbook, true))
                {
                    // Add a WorkbookPart to the document.
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Spreadsheet.Workbook();

                    // Add a WorksheetPart to the WorkbookPart.
                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Spreadsheet.Worksheet(new Spreadsheet.SheetData());

                    // Add Sheets to the Workbook.
                    var sheets = document.WorkbookPart.Workbook.AppendChild<Spreadsheet.Sheets>(new Spreadsheet.Sheets());

                    // Append a new worksheet and associate it with the workbook.
                    var sheet = new Spreadsheet.Sheet()
                                    {
                                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                                        SheetId = 1,
                                        Name = "Sheet1"
                                    };

                    sheets.Append(sheet);

                    // Add a WorkbookStylesPart to the WorkbookPart
                    var stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    var shData = worksheetPart.Worksheet.GetFirstChild<Spreadsheet.SheetData>();



                    stylesPart.Stylesheet = InitializeDefaultStyles(stylesPart.Stylesheet);

                    uint styleId = 2; //обязательно 2 потому что уже додано два стиля по умолчанию с индексами 0 и 1;

                    var styles = _report.Select(o => o.Style).ToList();

                    //нельзя перенести в отдельный метод :(
                    foreach (var s in styles)
                    {
                        if (s is TableStyle)
                        {
                            var style = s as TableStyle;

                            var border = new Spreadsheet.Border();

                            var leftBorder = new Spreadsheet.LeftBorder();
                            var rightBorder = new Spreadsheet.RightBorder();
                            var topBorder = new Spreadsheet.TopBorder();
                            var bottomBorder = new Spreadsheet.BottomBorder();

                            leftBorder.Style = SetLineType(style.BorderLine);
                            rightBorder.Style = SetLineType(style.BorderLine);
                            topBorder.Style = SetLineType(style.BorderLine);
                            bottomBorder.Style = SetLineType(style.BorderLine);

                            leftBorder.Color = new Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            rightBorder.Color = new Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            topBorder.Color = new Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };
                            bottomBorder.Color = new Spreadsheet.Color { Rgb = ToHexBinaryValue(style.BorderColor) };

                            border.LeftBorder = leftBorder;
                            border.RightBorder = rightBorder;
                            border.TopBorder = topBorder;
                            border.BottomBorder = bottomBorder;

                            stylesPart.Stylesheet.Borders.AppendChild<Spreadsheet.Border>(border);



                            var patternFill = new Spreadsheet.PatternFill
                            {
                                BackgroundColor = new Spreadsheet.BackgroundColor { Rgb = ToHexBinaryValue(style.Background) },
                                ForegroundColor = new Spreadsheet.ForegroundColor { Rgb = ToHexBinaryValue(style.Foreground) },
                                PatternType = SetPatternType(style.PatternType)
                            };

                            var fill = new Spreadsheet.Fill { PatternFill = patternFill };

                            stylesPart.Stylesheet.Fills.AppendChild<Spreadsheet.Fill>(fill);
                        }
                        else
                        {
                            stylesPart.Stylesheet.Fills.AppendChild<Spreadsheet.Fill>(new Spreadsheet.Fill());
                            stylesPart.Stylesheet.Borders.AppendChild<Spreadsheet.Border>(new Spreadsheet.Border());
                        }


                        if (s != null)
                        {
                            var font = new Spreadsheet.Font
                                                {
                                                    FontName = new Spreadsheet.FontName { Val = new StringValue { Value = s.FontName } },
                                                    FontSize = new Spreadsheet.FontSize { Val = s.FontSize },
                                                    Color = new Spreadsheet.Color { Rgb = ToHexBinaryValue(s.FontColor) }
                                                };

                            switch (s.FontType)
                            {
                                case TextStyle.Bold: font.Bold = new Spreadsheet.Bold { Val = true }; break;
                                case TextStyle.Italic: font.Italic = new Spreadsheet.Italic { Val = true }; break;
                                case TextStyle.Underline: font.Underline = new Spreadsheet.Underline { Val = Spreadsheet.UnderlineValues.Single }; break;
                                case TextStyle.Normal: break;
                            }

                            stylesPart.Stylesheet.Fonts.AppendChild<Spreadsheet.Font>(font);



                            var cellFormat = new Spreadsheet.CellFormat
                                                        {
                                                            BorderId = UInt32Value.ToUInt32(styleId),
                                                            FillId = UInt32Value.ToUInt32(styleId),
                                                            FontId = UInt32Value.ToUInt32(styleId)
                                                        };

                            stylesPart.Stylesheet.CellFormats.AppendChild<Spreadsheet.CellFormat>(cellFormat);
                        }

                        styleId = styleId + 1;
                    }

                    foreach (var x in _report)
                    {
                        if (x is TextElement)
                        {
                            var element = x as TextElement;

                            var row = shData.AppendChild<Spreadsheet.Row>(new Spreadsheet.Row());

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

                            var cell = new Spreadsheet.Cell
                                           {
                                               CellValue = new Spreadsheet.CellValue(value),
                                               DataType = Spreadsheet.CellValues.String,
                                               StyleIndex = UInt32Value.FromUInt32(styleId + 1)
                                           };

                            row.AppendChild(cell);

                            workbookpart.Workbook.Save();
                        }

                        if (x is TableElement)
                        {
                            var element = x as TableElement;
                            if (element.Table.Rows.Count == 0)
                            {
                                throw new Exception("The table is empty");
                            }

                            if (element.Headers.Count != 0)
                            {
                                styleId = GetStyleId(element.HeaderStyle, styles);

                                var row = shData.AppendChild(new Spreadsheet.Row());

                                foreach (var head in element.Headers)
                                {
                                    var cell = new Spreadsheet.Cell
                                                   {
                                                       CellValue = new Spreadsheet.CellValue(head),
                                                       DataType = Spreadsheet.CellValues.String,
                                                       StyleIndex = UInt32Value.FromUInt32(styleId + 1)
                                                   };

                                    row.AppendChild<Spreadsheet.Cell>(cell);
                                }
                                workbookpart.Workbook.Save();
                            }


                            styleId = GetStyleId(element.TableStyle, styles);

                            for (int i = 0; i < element.Table.Rows.Count; i++)
                            {
                                var row = shData.AppendChild(new Spreadsheet.Row());

                                for (int j = 0; j < element.Table.Columns.Count; j++)
                                {
                                    var cell = new Spreadsheet.Cell
                                                   {
                                                       CellValue = new Spreadsheet.CellValue(element.Table.Rows[i][j].ToString()),
                                                       DataType = Spreadsheet.CellValues.String,
                                                       StyleIndex = UInt32Value.FromUInt32(styleId + 1)
                                                   };

                                    row.AppendChild(cell);
                                }
                            }

                            workbookpart.Workbook.Save();
                        }

                        if (x is ImageElement)
                        {
#warning need to implement for ImageElement
                            workbookpart.Workbook.Save();
                        }
                    }

                    var dateLine = shData.AppendChild(new Spreadsheet.Row());

                    dateLine.AppendChild(new Spreadsheet.Cell()
                                             {
                                                 CellValue = new Spreadsheet.CellValue(_report.TimeStamp.Date.ToString()),
                                                 DataType = Spreadsheet.CellValues.String,
                                                 StyleIndex = 0
                                             });

                    workbookpart.Workbook.Save();
                }

                buffer = new byte[outputMemoryStream.Position];
                outputMemoryStream.Position = 0;
                outputMemoryStream.Read(buffer, 0, buffer.Length);
            }
            return buffer;
        }

        public override byte[] ToWord()
        {
#warning Need to extend text style support

            var memoryStream = new MemoryStream();

            using (var document = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
            {
                document.AddMainDocumentPart();
                document.MainDocumentPart.Document = new Wordprocessing.Document();

                StyleDefinitionsPart stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                Wordprocessing.Body body = document.MainDocumentPart.Document.AppendChild<Wordprocessing.Body>(new Wordprocessing.Body());


                var styles = _report.Select(o => o.Style).ToList();
                stylesPart.Styles = WordStyles(styles);
                stylesPart.Styles.Save();


                foreach (var x in _report)
                {
                    if (x is TextElement)
                    {
                        var element = x as TextElement;

                        if (element.Value == InternalContants.NewLine || element.Value == InternalContants.NewSection)
                        {
                            body.AppendChild<Wordprocessing.Paragraph>(new Wordprocessing.Paragraph());
                            continue;
                        }

                        Wordprocessing.Text text = new Wordprocessing.Text(element.Value.ToString());
                        Wordprocessing.Run run = new Wordprocessing.Run(text);
                        Wordprocessing.Paragraph paragraph = new Wordprocessing.Paragraph(run);

                        paragraph.ParagraphProperties = new Wordprocessing.ParagraphProperties
                        {
                            ParagraphStyleId = new Wordprocessing.ParagraphStyleId()
                            {
                                Val = WordStyleIdFromStyleName(document, element.Style.Name)
                            },
                            WordWrap = new Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                            TextAlignment = new Wordprocessing.TextAlignment { Val = Wordprocessing.VerticalTextAlignmentValues.Auto },
                        };

                        body.Append(paragraph);
                    }

                    if (x is TableElement)
                    {
                        var element = x as TableElement;
                        if (element.Table.Rows.Count == 0)
                        {
                            throw new Exception("The table is empty");
                        }

                        Wordprocessing.Table table = new Wordprocessing.Table();

                        if (element.Headers.Count != 0)
                        {
                            Wordprocessing.TableRow row = new Wordprocessing.TableRow();
                            foreach (var head in element.Headers)
                            {
                                Wordprocessing.Text text = new Wordprocessing.Text(head.ToString());
                                Wordprocessing.Run run = new Wordprocessing.Run(text);
                                Wordprocessing.Paragraph paragraph = new Wordprocessing.Paragraph(run);

                                paragraph.ParagraphProperties = new Wordprocessing.ParagraphProperties
                                {
                                    ParagraphStyleId = new Wordprocessing.ParagraphStyleId()
                                    {
                                        Val = WordStyleIdFromStyleName(document, element.HeaderStyle.Name)
                                    },
                                    WordWrap = new Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                                    TextAlignment = new Wordprocessing.TextAlignment { Val = Wordprocessing.VerticalTextAlignmentValues.Auto },
                                };

                                Wordprocessing.TableCell cell = new Wordprocessing.TableCell(paragraph);

                                Wordprocessing.TableCellProperties cellprop = WordCellProperties(element.HeaderStyle);
                                cell.Append(cellprop);

                                row.Append(cell);
                            }
                            table.Append(row);
                        }


                        for (int i = 0; i < element.Table.Rows.Count; i++)
                        {
                            Wordprocessing.TableRow row = new Wordprocessing.TableRow();
                            for (int j = 0; j < element.Table.Columns.Count; j++)
                            {
                                Wordprocessing.Text text = new Wordprocessing.Text(element.Table.Rows[i][j].ToString());
                                Wordprocessing.Run run = new Wordprocessing.Run(text);
                                Wordprocessing.Paragraph paragraph = new Wordprocessing.Paragraph(run);

                                paragraph.ParagraphProperties = new Wordprocessing.ParagraphProperties
                                {
                                    ParagraphStyleId = new Wordprocessing.ParagraphStyleId()
                                    {
                                        Val = WordStyleIdFromStyleName(document, element.TableStyle.Name)
                                    },
                                    WordWrap = new Wordprocessing.WordWrap { Val = new OnOffValue { Value = true } },
                                    TextAlignment = new Wordprocessing.TextAlignment { Val = Wordprocessing.VerticalTextAlignmentValues.Auto },
                                };

                                Wordprocessing.TableCell cell = new Wordprocessing.TableCell(paragraph);

                                Wordprocessing.TableCellProperties cellprop = WordCellProperties(element.TableStyle);
                                cell.Append(cellprop);

                                row.Append(cell);
                            }
                            table.Append(row);
                        }
                        body.AppendChild<Wordprocessing.Table>(table);
                    }

                    if (x is ImageElement)
                    {
#warning Image element not implemented

                        var element = x as ImageElement;
                        //body.AppendChild<Wordprocessing.Paragraph>();
                        document.MainDocumentPart.Document.Save();
                    }

                    document.MainDocumentPart.Document.Save();
                }

            }

            return memoryStream.ToArray();
        }

        public override byte[] ToHtml()
        {
            var head = new StringBuilder();

            head.AppendLine("<style type=\"text/css\">");

            //стиль по умолчанию для неформатированного текста
            var style = new Style();
            head.AppendFormat("\nhtml, body : {{");
            head.AppendFormat(" font-size: {0}%;", style.FontSize * 100 / 12); //http://habrahabr.ru/post/42151/ описание единиц измерения шрифтов
            //будем отталкиваться от эталона в 12pt и масштабировать в % 
            //причем процентное соотношение будет браться от настроек браузера пользователя
            head.AppendFormat(" font-family: \"{0}\";", style.FontName);
            head.AppendFormat(" color: #{0};", ColorToRgb(style.FontColor));
            head.AppendFormat(" {0};", FontStyle(style.FontType));
            head.AppendFormat(" }}");


            var styles = _report.Select(o => o.Style).ToList();
            head.AppendFormat("\n{0}", SetCssStyles(styles));

            head.AppendLine("\n</style>");

            var document = new StringBuilder();

            document.AppendLine("<html>");
            document.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />");

            document.AppendLine("<head>");
            document.AppendLine(head.ToString());
            document.AppendLine("</head>");

            document.AppendLine("<body>");
            document.AppendLine(ToHtmlBlock());
            document.AppendLine("</body>");

            document.AppendLine("</html>");

            var memoryStream = new MemoryStream();
            var streamWriter = new StreamWriter(memoryStream);
            streamWriter.Write(document.ToString());
            streamWriter.Flush();
            memoryStream.Position = 0;

            return memoryStream.ToArray();
        }

        public virtual string ToHtmlBlock()
        {
            var body = new StringBuilder();

            body.AppendFormat("<section class=\"report\">");

            foreach (var x in _report)
            {
                if (x is TextElement)
                {
                    var element = x as TextElement;

                    if (element.Value == InternalContants.NewLine)
                    {
                        body.AppendFormat("<br />");
                        continue;
                    }

                    if (element.Value == InternalContants.NewSection)
                    {
                        body.AppendFormat("<p />");
                        continue;
                    }

                    if (element.Style.DocumentTitle != DocumentTitle.None)
                    {
                        body.AppendFormat("<h{2} class=\"{0}\">{1}</h{2}>", element.Style.Name, element.Value, ((int)element.Style.DocumentTitle));
                        continue;
                    }

                    body.AppendFormat("<span class=\"{0}\">{1}</span>", element.Style.Name, element.Value);
                }

                if (x is TableElement)
                {
                    var element = x as TableElement;
                    body.AppendLine(RenderTableToHtml(element));
                }

                if (x is ImageElement)
                {
                    var element = x as ImageElement;
                    body.AppendLine(RenderImage(element));
                }
            }

            body.AppendFormat("<span>{0}</span>", _report.TimeStamp.ToString());
            body.AppendFormat("</section>");

            return body.ToString();
        }

        protected virtual string RenderTableToHtml(TableElement tableElement)
        {
            var sb = new StringBuilder();

            if (tableElement.Headers != null && tableElement.Headers.Count > 0 && tableElement.Headers.Count != tableElement.Table.Columns.Count)
            {
                throw new Exception("Count of column titles  not equal to count of column headers.");
            }

            sb.AppendLine("<table cellpadding=\"0\" cellspacing=\"0\">");

            if (tableElement.Headers.Count != 0)
            {
                sb.AppendLine("<thead>");
                //when we have column names
                sb.AppendFormat("<tr class=\"{0}\">", tableElement.HeaderStyle.Name);

                foreach (var c in tableElement.Headers)
                {
                    sb.AppendFormat("<th class=\"{0}\">{1}</td>", tableElement.HeaderStyle.Name, c);
                }

                sb.AppendLine("</tr>");
                sb.AppendLine("</thead>");
            }

            sb.AppendLine("<tbody>");
            foreach (DataRow row in tableElement.Table.Rows)
            {
                sb.AppendFormat("<tr class=\"{0}\">", tableElement.TableStyle.Name);

                for (int i = 0; i < tableElement.Table.Columns.Count; i++)
                {
                    var cell = row[i];
                    sb.AppendFormat("<td class=\"{0}\">{1}</td>", tableElement.TableStyle.Name, cell);
                }

                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");

            return sb.ToString();
        }

        private static string RenderImage(ImageElement element)
        {
            if (element.Url != null && !String.IsNullOrEmpty(element.Url.ToString()))
            {
                return String.Format("<img src=\"{0}\" />", element.Url);
            }

            if (element.Image != null)
            {
                using (var memoryStream = new MemoryStream())
                {
                    // Convert Image to byte[]
                    element.Image.Save(memoryStream, ImageFormat.Jpeg);
                    byte[] imageBytes = memoryStream.ToArray();

                    // Convert byte[] to Base64 String
                    var base64String = Convert.ToBase64String(imageBytes);
                    return String.Format("<img src=\"{0}\" />", base64String);
                }
            }
            return String.Empty;
        }

        private static uint GetStyleId(Style style, ICollection<Style> styles)
        {
            uint styleId = 1;
            //Учитывая особенность построения списка стилей по умолчанию, 
            //а также особенность реализации и уточнения возможных ошибок,
            //удобно начинать индексацию с 1

            foreach (var s in styles)
            {
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

        private static Spreadsheet.Stylesheet InitializeDefaultStyles(Spreadsheet.Stylesheet stylesheet)
        {
            //!!!!!!!!!!!!!!!!!!!!!!!!!
            //Просьба не переделывать этот метод, это связано с непонятным багом в библиотеке Office XML SDK
            //!!!!!!!!!!!!!!!!!!!!!!!!!
            stylesheet = new Spreadsheet.Stylesheet();
            stylesheet.Borders = new Spreadsheet.Borders();
            stylesheet.Borders.AppendChild<Spreadsheet.Border>(new Spreadsheet.Border());
            stylesheet.Borders.AppendChild<Spreadsheet.Border>(new Spreadsheet.Border());

            stylesheet.Fills = new Spreadsheet.Fills();
            stylesheet.Fills.AppendChild<Spreadsheet.Fill>(new Spreadsheet.Fill());
            stylesheet.Fills.AppendChild<Spreadsheet.Fill>(new Spreadsheet.Fill());

            stylesheet.Fonts = new Spreadsheet.Fonts();
            stylesheet.Fonts.AppendChild<Spreadsheet.Font>(new Spreadsheet.Font());
            stylesheet.Fonts.AppendChild<Spreadsheet.Font>(new Spreadsheet.Font());

            stylesheet.CellFormats = new Spreadsheet.CellFormats();
            stylesheet.CellFormats.AppendChild<Spreadsheet.CellFormat>(new Spreadsheet.CellFormat());
            stylesheet.CellFormats.AppendChild<Spreadsheet.CellFormat>(new Spreadsheet.CellFormat());

            return stylesheet;
        }

        private static Spreadsheet.BorderStyleValues SetLineType(BoderLine lineType)
        {
            switch (lineType)
            {
                case BoderLine.Thin: return Spreadsheet.BorderStyleValues.Thin;
                case BoderLine.Medium: return Spreadsheet.BorderStyleValues.Medium;
                case BoderLine.Thick: return Spreadsheet.BorderStyleValues.Thick;

                default: return Spreadsheet.BorderStyleValues.None;
            }
        }

        private static Spreadsheet.PatternValues SetPatternType(PatternType patternType)
        {
            switch (patternType)
            {
                case PatternType.Solid: return Spreadsheet.PatternValues.Solid;

                default: return Spreadsheet.PatternValues.None;
            }
        }

        private static string SetCssStyles(IEnumerable<Style> stylesList)
        {
            var sb = new StringBuilder();

            foreach (var style in stylesList)
            {
                if (style is TableStyle)
                {
                    var element = style as TableStyle;
                    sb.AppendFormat(".{0} {{", style.Name);
                    sb.AppendFormat(" background-color: #{0};", ColorToRgb(element.Foreground));
                    sb.AppendLine(" }");

                    sb.AppendFormat(".{0} {{", style.Name);
                    sb.AppendFormat(" border: {0};", BorderLine(element));
                    //sb.AppendFormat(" border-left: {0};", BorderLine(element));
                    //sb.AppendFormat(" border-right: {0};", BorderLine(element));
                    //sb.AppendFormat(" border-top: {0};", BorderLine(element));
                    //sb.AppendFormat(" border-bottom: {0};", BorderLine(element));
                    sb.AppendLine(" }");
                }

                sb.AppendFormat(".{0} {{ ", style.Name);
                sb.AppendFormat(" font-size: {0}%;", style.FontSize * 100 / 12);
                sb.AppendFormat(" font-family: \"{0}\";", style.FontName);
                sb.AppendFormat(" color: #{0};", ColorToRgb(style.FontColor));
                sb.AppendFormat(" {0};", FontStyle(style.FontType));
                sb.AppendLine(" }");
            }

            return sb.ToString();
        }

        private static string BorderLine(TableStyle style)
        {
            var sb = new StringBuilder();
            switch (style.BorderLine)
            {
                case BoderLine.Thin: sb.AppendFormat("1px solid #{0}", ColorToRgb(style.BorderColor)); break;
                case BoderLine.Medium: sb.AppendFormat("2px solid #{0}", ColorToRgb(style.BorderColor)); break;
                case BoderLine.Thick: sb.AppendFormat("3px solid #{0}", ColorToRgb(style.BorderColor)); break;
                default: sb.AppendFormat("0px solid #{0}", ColorToRgb(style.BorderColor)); break;
            }
            return sb.ToString();
        }

        private static string ColorToRgb(System.Drawing.Color color)
        {
            return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }

        private static HexBinaryValue ToHexBinaryValue(System.Drawing.Color color)
        {
            return HexBinaryValue.FromString(ColorToRgb(color));
        }

        private static string FontStyle(TextStyle fontStyle)
        {
            switch (fontStyle)
            {
                case TextStyle.Bold: return "font-weight: bold";
                case TextStyle.Italic: return "font-style: italic";
                case TextStyle.Underline: return "text-decoration: underline";
                default: return "font-style: normal";
            }
        }

        private static Wordprocessing.TableCellProperties WordCellProperties(TableStyle style)
        {
            var leftBorder = new Wordprocessing.LeftBorder();
            var rightBorder = new Wordprocessing.RightBorder();
            var topBorder = new Wordprocessing.TopBorder();
            var bottomBorder = new Wordprocessing.BottomBorder();
            var insideHorizontalBorder = new Wordprocessing.InsideHorizontalBorder();
            var insideVerticalBorder = new Wordprocessing.InsideVerticalBorder();

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


            Wordprocessing.TableCellProperties cellprop = new Wordprocessing.TableCellProperties();
            Wordprocessing.Shading shading = new Wordprocessing.Shading();
            shading.Val = Wordprocessing.ShadingPatternValues.Clear;
            shading.Fill = ColorToRgb(style.Foreground);

            Wordprocessing.TableCellBorders tableCellBorders = new Wordprocessing.TableCellBorders();
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
            StyleDefinitionsPart stylePart = document.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<Wordprocessing.StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Wordprocessing.Style)s.Parent).Type == Wordprocessing.StyleValues.Paragraph))
                .Select(n => ((Wordprocessing.Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        private static Wordprocessing.Styles WordStyles(List<Style> stylesList)
        {
            Wordprocessing.Styles styles = new Wordprocessing.Styles();
            int countId = 1;
            foreach (var s in stylesList)
            {
                /*if (s is TableStyle)
                {
                    var style = s as TableStyle;
                    Wordprocessing.Style tableStyle = new Wordprocessing.Style()
                    {
                        Type = Wordprocessing.StyleValues.Table,
                        StyleId = "t"+countId.ToString(),
                        CustomStyle = true,
                        Default = false
                    };

                    var leftBorder = new Wordprocessing.LeftBorder();
                    var rightBorder = new Wordprocessing.RightBorder();
                    var topBorder = new Wordprocessing.TopBorder();
                    var bottomBorder = new Wordprocessing.BottomBorder();
                    var insideHorizontalBorder = new Wordprocessing.InsideHorizontalBorder();
                    var insideVerticalBorder = new Wordprocessing.InsideVerticalBorder();

                    leftBorder.Val  = SetWordLineType(style.BorderLine);
                    rightBorder.Val = SetWordLineType(style.BorderLine);
                    topBorder.Val = SetWordLineType(style.BorderLine);
                    bottomBorder.Val = SetWordLineType(style.BorderLine);
                    insideHorizontalBorder.Val = SetWordLineType(style.BorderLine);
                    insideVerticalBorder.Val = SetWordLineType(style.BorderLine);

                    leftBorder.Color = ColorToRgb(style.BorderColor);
                    rightBorder.Color = ColorToRgb(style.BorderColor);
                    topBorder.Color = ColorToRgb(style.BorderColor);
                    bottomBorder.Color = ColorToRgb(style.BorderColor);
                    insideHorizontalBorder.Color = ColorToRgb(style.BorderColor);
                    insideVerticalBorder.Color = ColorToRgb(style.BorderColor);

                    leftBorder.Size = SetBorderWaight(style.BorderLine);
                    rightBorder.Size = SetBorderWaight(style.BorderLine);
                    topBorder.Size = SetBorderWaight(style.BorderLine);
                    bottomBorder.Size = SetBorderWaight(style.BorderLine);
                    insideHorizontalBorder.Size = SetBorderWaight(style.BorderLine);
                    insideVerticalBorder.Size = SetBorderWaight(style.BorderLine);

                    Wordprocessing.TableProperties tabProp = new Wordprocessing.TableProperties
                    {
                        TableBorders = new Wordprocessing.TableBorders
                        {
                            LeftBorder = leftBorder,
                            RightBorder = rightBorder,
                            TopBorder = topBorder,
                            BottomBorder = bottomBorder,
                            InsideHorizontalBorder = insideHorizontalBorder,
                            InsideVerticalBorder = insideVerticalBorder,
                        },
                    };
                    Wordprocessing.TableCellProperties cellProperties = new Wordprocessing.TableCellProperties
                    {
                        Shading = new Wordprocessing.Shading 
                        { 
                            Val = Wordprocessing.ShadingPatternValues.Clear, 
                            Fill = ColorToRgb(style.Foreground),
                        },
                    };

                    tableStyle.Append(tabProp);
                    tableStyle.Append(cellProperties);
                    styles.Append(tableStyle);
                }*/

                Wordprocessing.Style paragraphStyle = new Wordprocessing.Style()
                {
                    Type = Wordprocessing.StyleValues.Paragraph,
                    StyleId = "p" + countId.ToString(),
                    CustomStyle = true,
                    Default = false,
                    StyleName = new Wordprocessing.StyleName { Val = s.Name }
                };

                Wordprocessing.StyleRunProperties styleRunProperties = new Wordprocessing.StyleRunProperties
                {
                    Color = new Wordprocessing.Color() { Val = ColorToRgb(s.FontColor) },
                    RunFonts = new Wordprocessing.RunFonts() { Ascii = s.FontName },
                    FontSize = new Wordprocessing.FontSize() { Val = (s.FontSize * 2).ToString() },
                    //Bold = new Wordprocessing.Bold(),
                    //Italic = new Wordprocessing.Italic(),
                    //Underline = new Wordprocessing.Underline() { Val = Wordprocessing.UnderlineValues.Single },
                };

                Wordprocessing.StyleRunProperties helpStyleRunProperties = WordFontStyle(s.FontType);

                paragraphStyle.Append(styleRunProperties);
                paragraphStyle.Append(helpStyleRunProperties);
                styles.Append(paragraphStyle);

                countId = countId + 1;
            }
            return styles;
        }

        private static Wordprocessing.StyleRunProperties WordFontStyle(TextStyle textStyle)
        {
            Wordprocessing.StyleRunProperties property = new Wordprocessing.StyleRunProperties();
            switch (textStyle)
            {
                case TextStyle.Bold: property.Bold = new Wordprocessing.Bold(); break;
                case TextStyle.Italic: property.Italic = new Wordprocessing.Italic(); break;
                case TextStyle.Underline: property.Underline = new Wordprocessing.Underline() { Val = Wordprocessing.UnderlineValues.Single }; break;
                default: break;
            }
            return property;
        }

        private static EnumValue<Wordprocessing.BorderValues> WordLineType(BoderLine borderLine)
        {
            switch (borderLine)
            {
                case BoderLine.Thin: return Wordprocessing.BorderValues.Single;
                case BoderLine.Medium: return Wordprocessing.BorderValues.Single;
                case BoderLine.Thick: return Wordprocessing.BorderValues.Single;
                default: return Wordprocessing.BorderValues.None;
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

        private static Font PdfFontStyle(Font font, TextStyle textStyle)
        {
            switch (textStyle)
            {
                case TextStyle.Bold: font.SetStyle(Font.BOLD); return font;
                case TextStyle.Italic: font.SetStyle(Font.ITALIC); return font;
                case TextStyle.Underline: font.SetStyle(Font.UNDERLINE); return font;
                default: font.SetStyle(Font.NORMAL); return font;
            }
        }
    }

}