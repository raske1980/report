using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace Report.Base
{
    public interface IReportBuilder
    {
        /// <summary>
        /// Appends a block of text.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="style"></param>
        void AppendTextBlock(String text, Style style);

        /// <summary>
        /// Append new line in report.
        /// </summary>
        void AppendNewLine();

        /// <summary>
        /// Append new line in report.
        /// </summary>
        void AppendNewSection();

        /// <summary>
        /// Appends a table.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="headers"></param>
        /// <param name="tableStyle"> </param>
        /// <param name="headerStyle"> </param>
        void AppendTable(DataTable data, IEnumerable<string> headers, Style tableStyle, Style headerStyle);

        /// <summary>
        /// Appends a table
        /// </summary>
        /// <param name="data"></param>
        /// <param name="headers"></param>
        /// <param name="takeOutNumber">Specifies a number of columns which you will take out into the header</param>
        /// <param name="tableStyle"> </param>
        /// <param name="headerStyle"> </param>
        void AppendTable(DataTable data, IEnumerable<string> headers, Int32 takeOutNumber, Style tableStyle, Style headerStyle);

        /// <summary>
        /// Appends a table
        /// </summary>
        /// <param name="data"></param>
        /// <param name="headers"></param>
        /// <param name="takeOutHeaders"></param>
        /// <param name="tableStyle"> </param>
        /// <param name="headerStyle"> </param>
        void AppendTable(DataTable data, IEnumerable<string> headers, IEnumerable<string> takeOutHeaders, Style tableStyle, Style headerStyle);

        /// <summary>
        /// Append a string list
        /// </summary>
        /// <param name="items"></param>
        /// <param name="style"></param>
        void AppendList<T>(IEnumerable<T> items, Style style);

        /// <summary>
        /// Appends the string returned by processing a composite format string, 
        /// which contains zero or more format items, to this instance. Each format 
        /// item is replaced by the string representation of a corresponding argument 
        /// in a parameter array.
        /// </summary>
        /// <param name="format">A composite format string</param>
        /// <param name="style"> </param>
        /// <param name="args">An array of objects to format</param>
        void AppendFormat(string format, Style style, params object[] args);

        /// <summary>
        /// Appends a copy of the specified string followed by the default line 
        /// terminator to the end of the current System.Text.StringBuilder object.
        /// </summary>
        /// <param name="value">The plain string or html to append.</param>
        /// <param name="style"> </param>
        void AppendLine(string value, Style style);

        void AppendImage(Image image);

        void AppendImage(string url);

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <param name="records"></param>
        /// <param name="style"></param>
        void AppendDictionary<T1, T2>(IDictionary<T1, T2> records, Style style);

        Report Build();
    }
}
