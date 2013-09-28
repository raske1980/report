using System.Collections;
using Report.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Report
{
    /// <summary>
    /// 
    /// </summary>
    public class ReportBuilder : ReportBuilderBase
    {
        private readonly Report.Base.Report _report;

        public Style DefaultStyle { get; set; }

        public ReportBuilder()
        {
            _report = new Base.Report();
            DefaultStyle = new Style();

        }

        public virtual void Clear()
        {
            _report.Clear();
        }

        public override void AppendNewLine()
        {
            AppendLine(InternalContants.NewLine, new Style());
        }

        public override void AppendNewSection()
        {
            AppendLine(InternalContants.NewSection, new Style());
        }

        public override void AppendFormat(string format, Style style, params object[] args)
        {
            _report.Add(new TextElement
                          {
                              Value = String.Format(format, args),
                              Style = style
                          });

        }

        public override void AppendLine(string value, Style styleName)
        {
            AppendTextBlock(value, styleName);
        }

        public override void AppendTextBlock(string text, Style style)
        {
            _report.Add(new TextElement
                          {
                              Style = style,
                              Value = text
                          });
        }

        public override void AppendTable(DataTable data, IEnumerable<String> headers, IEnumerable<String> takeOutHeaders, Style tableStyle, Style headerStyle)
        {
            throw new NotImplementedException();
        }

        public override void AppendTable(DataTable data, IEnumerable<String> headers, Style tableStyle, Style headerStyle)
        {
            var tableElement = new TableElement(data)
                                   {
                                       HeaderStyle = headerStyle as TableStyle,
                                       TableStyle = tableStyle as TableStyle
                                   };

            tableElement.Headers.Clear();

            if (headers != null)
            {
                tableElement.Headers.AddRange(headers);
            }

            _report.Add(tableElement);
        }

        public override void AppendTable(DataTable dataTable, IEnumerable<String> headers, int takeOutNumber, Style tableStyle, Style headerStyle)
        {
            if (headers != null && dataTable.Columns.Count != headers.Count())
            {
                throw new Exception("Count of columns titles  not equal to count of column headers.");
            }

            if (dataTable.Rows.Count != 0)
            {
                int count = 0;
                var subHeaders = (from o in dataTable.Rows.Cast<DataRow>()
                                  select o[count].ToString()).Distinct().ToList();

                foreach (var subHeader in subHeaders)
                {
                    AppendTextBlock(subHeader, headerStyle);

                    var subRecords = (from o in dataTable.Rows.Cast<DataRow>()
                                      where o[count].ToString() == subHeader
                                      select o).ToList();

                    if (count != takeOutNumber - 1)
                    {
                        SubTable(subRecords, takeOutNumber, count, headers, tableStyle, headerStyle);
                    }
                    else
                    {
                        FillTable(subRecords, takeOutNumber, headers, tableStyle, headerStyle);
                    }
                }
            }
        }

        private void SubTable(IEnumerable<DataRow> records, int takeOutNumber, int count, IEnumerable<String> headers, Style tableStyle, Style headerStyle)
        {
            count = count + 1;

            var subHeaders = (from o in records
                              select o[count].ToString()).Distinct().ToList();

            foreach (var subHeader in subHeaders)
            {
                AppendTextBlock(subHeader, headerStyle);

                var subRecords = (from o in records
                                  where o[count].ToString() == subHeader
                                  select o).ToList();

                if (count != takeOutNumber - 1)
                {
                    SubTable(records, takeOutNumber, count, headers, tableStyle, headerStyle);
                }
                else
                {
                    FillTable(subRecords, takeOutNumber, headers, tableStyle, headerStyle);
                }
            }
        }

        private void FillTable(IEnumerable<DataRow> records, int takeOutNumber, IEnumerable<String> headers, Style tableStyle, Style headerStyle)
        {
            if (records != null && records.Any())
            {
                var dataTable = new DataTable();

                var someRecord = records.First();
                var columnCount = someRecord.ItemArray.Count();

                for (int i = takeOutNumber; i < columnCount; i++)
                {
                    var col = new DataColumn();
                    dataTable.Columns.Add(col);
                }

                foreach (var row in records)
                {
                    var r = dataTable.NewRow();

                    for (int i = takeOutNumber; i < columnCount; i++)
                    {
                        r[i - takeOutNumber] = row[i];
                    }

                    dataTable.Rows.Add(r);
                }

                var tableElement = new TableElement(dataTable);

                tableElement.HeaderStyle = headerStyle as TableStyle;
                tableElement.TableStyle = tableStyle as TableStyle;

                tableElement.Headers.Clear();

                if (headers != null && headers.Any())
                {
                    for (int i = takeOutNumber; i < headers.Count(); i++)
                    {
                        tableElement.Headers.Add(headers.ElementAt(i));
                    }
                }

                _report.Add(tableElement);
            }
        }

        public override void AppendDictionary<T1, T2>(IDictionary<T1, T2> records, Style style)
        {
            if (records == null || !records.Any())
            {
                return;
            }

            _report.Add(TableElement.Create(records, style));
        }

        public override void AppendImage(System.Drawing.Image image)
        {
            _report.Add(new ImageElement(image));
        }

        public override void AppendImage(string url)
        {
            _report.Add(new ImageElement(url));
        }

        public override Base.Report Build()
        {
            _report.TimeStamp = DateTime.Now;
            return _report;
        }

        public override void AppendList<T>(IEnumerable<T> items, Style style)
        {
            if (items == null || !items.Any())
            {
                return;
            }

            _report.Add(TableElement.Create(items, style));
        }
    }
}
