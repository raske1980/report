using System.IO;
using Report.Base;
using Report.Renderers;

namespace Report
{
    /// <summary>
    /// 
    /// </summary>
    public class ReportRenderer : IReportRenderer
    {
        private readonly Report.Base.Report _report;

        public ReportRenderer(Report.Base.Report report)
        {
            _report = report;
        }

        public byte[] ToPdf()
        {
            IRenderer renderer = new PdfRenderer();
            return renderer.Render(_report);
        }

        public byte[] ToExcel()
        {
            IRenderer renderer = new ExcelRenderer();
            return renderer.Render(_report);
        }

        public byte[] ToWord()
        {
            IRenderer renderer = new WordRenderer();
            return renderer.Render(_report);
        }

        public byte[] ToHtml()
        {
            IRenderer renderer = new HtmlRenderer();
            return renderer.Render(_report);
        }

        public virtual string ToHtmlBlock()
        {
            var renderer = new HtmlRenderer();
            return renderer.ToHtmlBlock(_report);
        }

        public void ToPdf(string path)
        {
            var bytes = ToPdf();
            File.WriteAllBytes(path, bytes);
        }

        public void ToExcel(string path)
        {
            CheckDirectory(Path.GetDirectoryName(path));
            var bytes = ToExcel();
            File.WriteAllBytes(path, bytes);
        }

        public void ToWord(string path)
        {
            CheckDirectory(Path.GetDirectoryName(path));
            var bytes = ToWord();
            File.WriteAllBytes(path, bytes);
        }

        public void ToHtml(string path)
        {
            CheckDirectory(Path.GetDirectoryName(path));
            var bytes = ToHtml();
            File.WriteAllBytes(path, bytes);
        }

        private static void CheckDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }
    }
}
