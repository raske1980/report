using System.Collections.Generic;
using System.Data;
using System;

namespace Report.Base
{
    public abstract class ReportBuilderBase : IReportBuilder
    {
        public abstract Report Build();
        //public abstract void AppendTitle(String text, Style style);
        public abstract void AppendTextBlock(String text, Style style);
        public abstract void AppendTable(DataTable data, IEnumerable<String> headers, Style tableStyle, Style headerStyle);
        public abstract void AppendTable(DataTable data, IEnumerable<String> headers, int takeOutNumber, Style tableStyle, Style headerStyle);
        public abstract void AppendTable(DataTable data, IEnumerable<String> headers, IEnumerable<String> takeOutHeaders, Style tableStyle, Style headerStyle);
        public abstract void AppendImage(System.Drawing.Image image);
        public abstract void AppendImage(String url);
        public abstract void AppendList<T>(IEnumerable<T> items, Style style);
        public abstract void AppendFormat(String format, Style style, params object[] args);
        public abstract void AppendLine(String value, Style style);
        public abstract void AppendDictionary<T1, T2>(IDictionary<T1, T2> records, Style style);
        public abstract void AppendNewLine();
        public abstract void AppendNewSection();
        
        public virtual void AppendTextBlock(String text)
        {
            AppendTextBlock(text, new Style());
        }

        public virtual void AppendTable(DataTable data)
        {
            AppendTable(data, null, new TableStyle(), new TableStyle());
        }

        public virtual void AppendTable(DataTable data, Int32 takeOutNumber, Style headerStyle)
        {
            AppendTable(data, null, takeOutNumber, new TableStyle(), headerStyle);
        }

        public virtual void AppendTable(DataTable data, IEnumerable<String> takeOutHeaders, Style headerStyle)
        {
            AppendTable(data, null, takeOutHeaders, null, headerStyle);
        }

        public virtual void AppendTable(DataView data, IEnumerable<String> headers, Style tableStyle, Style headerStyle)
        {
            AppendTable(data.Table, headers, tableStyle, headerStyle);
        }

        public virtual void AppendTable(DataView data, IEnumerable<String> headers, Int32 takeOutNumber, Style tableStyle, Style headerStyle)
        {
            AppendTable(data.Table, headers, takeOutNumber, tableStyle, headerStyle);
        }

        public virtual void AppendTable(DataView data, IEnumerable<String> headers, IEnumerable<String> takeOutHeaders, Style tableStyle, Style headerStyle)
        {
            AppendTable(data.Table, headers, takeOutHeaders, tableStyle, headerStyle);
        }

        public virtual void AppendTable(DataView data)
        {
            AppendTable(data.Table, null, new TableStyle(), new TableStyle());
        }

        public virtual void AppendTable(DataView data, Int32 takeOutNumber, Style headerStyle)
        {
            AppendTable(data.Table, null, takeOutNumber, new TableStyle(), headerStyle);
        }

        public virtual void AppendTable(DataView data, IEnumerable<String> takeOutHeaders, Style headerStyle)
        {
            AppendTable(data.Table, null, takeOutHeaders, new TableStyle(), headerStyle);
        }
    }
}
