using System;

namespace Report.Base
{
    public interface IReportRenderer
    {
        /// <summary>
        /// Converts report to PDF (.pdf)
        /// </summary>
        /// <param name="path"></param>
        void ToPdf(String path);

        /// <summary>
        /// Converts report to Excel (.xlsx)
        /// </summary>
        /// <param name="path"></param>
        void ToExcel(String path);

        /// <summary>
        /// Converts report to Word (.docx)
        /// </summary>
        /// <param name="path"></param>
        void ToWord(String path);

        /// <summary>
        /// Converts report to html (.html)
        /// </summary>
        /// <param name="path"></param>
        void ToHtml(String path);

        /// <summary>
        /// Converts report to PDF (.pdf)
        /// </summary>
        byte[] ToPdf();

        /// <summary>
        /// Converts report to Excel (.xlsx)
        /// </summary>
        byte[] ToExcel();

        /// <summary>
        /// Converts report to Word (.docx)
        /// </summary>
        byte[] ToWord();


        /// <summary>
        /// Converts report to html (.html)
        /// </summary>
        byte[] ToHtml();
    }
}
