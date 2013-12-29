using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Report.Base
{
    public class TableElement : IElement
    {
        public List<string> Headers { get; private set; }
        public DataTable Table { get; set; }
        public TableStyle TableStyle { get; set; }
        public TableStyle HeaderStyle { get; set; }

        public Style Style { get { return TableStyle; } }

        public TableElement()
        {
            Table = new DataTable();
            Headers = new List<string>();
            TableStyle = new TableStyle();
            HeaderStyle = new TableStyle();
        }

        public TableElement(DataTable dataTable)
        {
            Table = dataTable;
            Headers = new List<string>();
            TableStyle = new TableStyle();
            HeaderStyle = new TableStyle();
        }

        public static TableElement Create(IEnumerable items, Style style)
        {
            var table = new DataTable();
            table.Columns.Add();

            foreach (var item in items)
            {
                var row = table.NewRow();
                row[0] = item.ToString().Replace("[", "").Replace("]", "");
                table.Rows.Add(row);
            }

            var t = new TableElement(table);
            t.TableStyle = style as TableStyle;

            return t;
        }

        public static TableElement Create<T1, T2>(IDictionary<T1, T2> records, Style tableStyle)
        {
            var table = new DataTable();
            table.Columns.Add();
            table.Columns.Add();

            foreach (var record in records)
            {
                var row = table.NewRow();
                row[0] = record.Key;
                row[1] = record.Value;
                table.Rows.Add(row);
            }

            var t = new TableElement(table);
            t.TableStyle = tableStyle as TableStyle; ;

            return t;
        }
    }
}
