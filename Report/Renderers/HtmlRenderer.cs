using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using Report.Base;

namespace Report.Renderers
{
    /// <summary>
    /// 
    /// </summary>
    public class HtmlRenderer : RendererBase
    {
        public override byte[] Render(Base.Report report)
        {
            var document = new StringBuilder();

            document.AppendLine("<html>");
            document.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />");

            document.AppendLine("<head>");
            document.AppendLine(RenderHtmlHead(report));
            document.AppendLine("</head>");

            document.AppendLine("<body>");
            document.AppendLine(ToHtmlBlock(report));
            document.AppendLine("</body>");

            document.AppendLine("</html>");

            var memoryStream = new MemoryStream();
            var streamWriter = new StreamWriter(memoryStream);
            streamWriter.Write(document.ToString());
            streamWriter.Flush();
            memoryStream.Position = 0;

            return memoryStream.ToArray();
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

        public virtual string RenderHtmlHead(IEnumerable<IElement> report)
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
            head.AppendFormat(" {0};", FontStyle(style.TextStyle));
            head.AppendFormat(" }}");


            var styles = report.Select(o => o.Style).ToList();
            head.AppendFormat("\n{0}", RenderCssStyles(styles));

            head.AppendLine("\n</style>");

            return head.ToString();
        }

        public virtual string ToHtmlBlock(Base.Report report)
        {
            var body = new StringBuilder();

            body.AppendFormat("<section class=\"report\">");

            foreach (var x in report)
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

                if (x is ComplexHeaderCell)
                {
                    var element = x as ComplexHeaderCell;
                    body.AppendLine("<div style=\"color: red; font-weight: bold;\">Rended for ComplexHeaderCell not implemented.</div>");
                
                }

                if (x is ComplexHeader)
                {
                    var element = x as ComplexHeader;
                    body.AppendLine("<div style=\"color: red; font-weight: bold;\">Rended for ComplexHeader not implemented.</div>");
                }
            }

            body.AppendFormat("<span>{0}</span>", report.TimeStamp);
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

            //sb.AppendLine("<table cellpadding=\"0\" cellspacing=\"0\">");

            sb.AppendLine("<table cellpadding=\"0\" cellspacing=\"0\" class=\"for_replace\">");

            if (tableElement.Headers.Count != 0)
            {
                //sb.AppendLine("<thead>");
                sb.AppendFormat("<thead class=\"{0}\">", tableElement.HeaderStyle.Name);

                //when we have column names
                //sb.AppendFormat("<tr class=\"{0}\">", tableElement.HeaderStyle.Name);
                sb.AppendLine("<tr>");

                foreach (var c in tableElement.Headers)
                {
                    //sb.AppendFormat("<th class=\"{0}\">{1}</td>", tableElement.HeaderStyle.Name, c);
                    sb.AppendFormat("<th>{0}</td>", c);
                }

                sb.AppendLine("</tr>");
                sb.AppendLine("</thead>");
            }

            //sb.AppendLine("<tbody>");
            sb.AppendFormat("<tbody class=\"{0}\">", tableElement.TableStyle.Name);

            foreach (DataRow row in tableElement.Table.Rows)
            {
                //sb.AppendFormat("<tr class=\"{0}\">", tableElement.TableStyle.Name);
                sb.AppendLine("<tr>");

                for (int i = 0; i < tableElement.Table.Columns.Count; i++)
                {
                    var cell = row[i];
                    //sb.AppendFormat("<td class=\"{0}\">{1}</td>", tableElement.TableStyle.Name, cell);
                    sb.AppendFormat("<td>{0}</td>", cell);
                }

                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");

            return sb.ToString();
        }

        private static string RenderCssStyles(IEnumerable<Style> styles)
        {
            var sb = new StringBuilder();

            foreach (var style in styles)
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
                //sb.AppendFormat(" font-family: \"{0}\";", style.FontName);
                sb.AppendFormat(" color: #{0};", ColorToRgb(style.FontColor));
                sb.AppendFormat(" {0};", FontStyle(style.TextStyle));
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

        private static string FontStyle(IEnumerable<TextStyleType> fontStyle)
        {
            //тут нужно доделать
            foreach (var item in fontStyle)
            {
                switch (item)
                {
                    case TextStyleType.Bold: return "font-weight: bold";
                    case TextStyleType.Italic: return "font-style: italic";
                    case TextStyleType.Underline: return "text-decoration: underline";
                    default: return "font-style: normal";
                }
            }
            return "";
        }
    }
}
