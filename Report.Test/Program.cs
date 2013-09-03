using Report.Base;
using System.Data;
using System.Drawing;

namespace Report.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var tab = new DataTable();

            for (int i = 0; i < 10; i++)
            {
                tab.Columns.Add(i.ToString());
            }

            for (int i = 0; i < 15; i++)
            {
                var row = tab.NewRow();

                for (int j = 0; j < tab.Columns.Count; j++)
                {
                    row[j] = (i * j).ToString() + " ячейка";
                }

                tab.Rows.Add(row);
            }

            var style0 = new TableStyle();

            var style1 = new TableStyle
                             {
                                 Foreground = System.Drawing.Color.Blue,
                                 PatternType = PatternType.Solid,
                                 BorderLine = BoderLine.Thick,
                                 BorderColor = Color.Orange,
                                 DocumentTitle = DocumentTitle.Title
                             };

            var style2 = new TableStyle
                             {
                                 FontColor = Color.Black,
                                 FontSize = 5,
                                 FontName = "Calibri",
                                 Foreground = System.Drawing.Color.Gray,
                                 PatternType = PatternType.Solid,
                                 BorderLine = BoderLine.Thin,
                                 BorderColor = Color.Red,
                                 DocumentTitle = DocumentTitle.Heading1
                             };

            var style3 = new TableStyle()
                             {
                                 FontColor = Color.Bisque,
                                 FontSize = 16,
                                 FontName = "C",
                                 Foreground = Color.Green,
                                 PatternType = PatternType.Solid,
                                 DocumentTitle = DocumentTitle.Heading2
                             };

            var style4 = new TableStyle()
                             {
                                 FontColor = Color.LimeGreen,
                                 FontSize = 30,
                                 FontName = "Blackadder ITC",
                                 Foreground = Color.Yellow,
                                 PatternType = PatternType.Solid,
                                 DocumentTitle = DocumentTitle.Heading3
                             };

            var tableStyle1 = new TableStyle
                                  {
                                      Foreground = System.Drawing.Color.Blue,
                                      PatternType = PatternType.Solid,
                                      BorderLine = BoderLine.Thick,
                                      BorderColor = Color.Orange,
                                      DocumentTitle = DocumentTitle.None,
                                      FontSize = 10,
                                  };

            var tableStyle2 = new TableStyle
                                  {
                                      FontColor = Color.Black,
                                      FontSize = 10,
                                      FontName = "Calibri",
                                      Foreground = System.Drawing.Color.Gray,
                                      PatternType = PatternType.Solid,
                                      BorderLine = BoderLine.Thick,
                                      BorderColor = Color.Red,
                                      DocumentTitle = DocumentTitle.None
                                  };

            var tableStyle3 = new TableStyle
                                  {
                                      FontColor = Color.Black,
                                      FontSize = 10,
                                      FontName = "Calibri",
                                      Foreground = System.Drawing.Color.Gray,
                                      PatternType = PatternType.Solid,
                                      BorderLine = BoderLine.Thick,
                                      BorderColor = Color.Red,
                                      DocumentTitle = DocumentTitle.None
                                  };

            string[] abs = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" };

            var report = new ReportBuilder();

            report.AppendTextBlock("Test title1", style0);
            report.AppendTextBlock("Test title2", style2);
            report.AppendTextBlock("Test title3", style3);
            report.AppendTextBlock("Test title4", style4);
            report.AppendTextBlock("Test title3", style3);
            report.AppendTextBlock("Test title1", style1);
            report.AppendTextBlock("Test title4", style4);
            report.AppendNewSection();


            report.AppendTable(tab, abs, tableStyle1, tableStyle2);
            report.AppendNewSection();
            report.AppendTable(tab, abs, tableStyle2, tableStyle3);
            report.AppendNewLine();
            report.AppendTable(tab);

            var reportRender = new ReportRenderer(report.Build());

            var directory = @"c:\pub\";

            reportRender.ToExcel(directory + "example.xlsx");
            reportRender.ToHtml(directory + "example.html");
            reportRender.ToPdf(directory + "example.pdf");
            reportRender.ToWord(directory + "example.docx");
        }

    }
}
