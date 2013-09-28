using System.IO;

namespace Report.Base
{
    public abstract class ReportRendererBase : IReportRenderer
    {
        public void ToPdf(string path)
        {
            var bytes = ToPdf();
            File.WriteAllBytes(path, bytes);
        }

        public void ToExcel(string path)
        {
            var bytes = ToExcel();
            File.WriteAllBytes(path, bytes);
        }

        public void ToWord(string path)
        {
            var bytes = ToWord();
            File.WriteAllBytes(path, bytes);
        }

        public void ToHtml(string path)
        {
            var bytes = ToHtml();
            File.WriteAllBytes(path, bytes);
        }

        public abstract byte[] ToPdf();
        public abstract byte[] ToExcel();
        public abstract byte[] ToWord();
        public abstract byte[] ToHtml();
    }
}
