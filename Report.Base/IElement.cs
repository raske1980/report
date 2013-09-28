using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace Report.Base
{
    public interface IElement
    {
        Style Style { get; }
    }

    public class TextElement : IElement
    {
        public String Value { get; set; }
        public Style Style { get; set; }

        public TextElement()
        {
            Style = new Style();
            Value = String.Empty;
        }
    }

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
                row[0] = item.ToString();
                table.Rows.Add(row);
            }

            var t = new TableElement(table);
            t.TableStyle = style as TableStyle;

            return t;
        }

        public static TableElement Create<T1, T2>(IDictionary<T1, T2> records, TableStyle tableStyle)
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
            t.TableStyle = tableStyle;

            return t;
        }
    }

    public class EmptyElement : IElement
    {
        public EmptyElement()
        {
        }

        public Style Style { get; set; }
    }

    public class ImageElement : IElement
    {
        public Image Image { get; set; }
        public Uri Url { get; set; }

        public ImageElement()
        {
        }

		  public ImageElement(System.Drawing.Image image)
		  {
				Image = image;
		  }

        public ImageElement(string url)
        {
				Image = Image.FromFile(url);
            Url = new Uri(url);
        }

        /*public ImageElement(System.Drawing.Image image)
        {
            Image = image;
        }*/

        public Style Style { get; private set; }
    }
}
