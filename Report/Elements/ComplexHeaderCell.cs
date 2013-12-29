using DocumentFormat.OpenXml.Spreadsheet;
using Report.Base;

namespace Report.Base
{
    public class ComplexHeaderCell : IElement
    {
        private string CellNameFrom { get; set; }
        private string CellNameTo { get; set; }
        private string CellText { get; set; }
        public Style Style { get; set; }

        public ComplexHeaderCell(string cellNameFrom, string cellNameTo, string cellText, Style style)
        {
            Style = style;
            CellNameFrom = cellNameFrom;
            CellNameTo = cellNameTo;
            CellText = cellText;
        }

        public ComplexHeaderCell(string cellNameFrom, string cellNameTo, string cellText)
            : this(cellNameFrom, cellNameTo, cellText, null)
        {
        }

        public ComplexHeaderCell(string cellNameFrom, string cellText)
            : this(cellNameFrom, cellText, null)
        {
        }

        public void Render(Worksheet doc)
        {
            Render(doc, 0);
        }


        public void Render(Worksheet doc, uint styleid)
        {
            if (!string.IsNullOrEmpty(CellNameTo))
                Merge.MergeTwoCells(doc, CellNameFrom, CellNameTo, CellText, styleid);
            else
                Merge.CreateSpreadsheetCellIfNotExist(doc, CellNameFrom, CellText, styleid);
        }

    }
}
