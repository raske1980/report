using System;
using System.Collections.Generic;
using System.Text;
using RNPExcelExport.Data.Collection;
using RNPExcelExport.Data;
using System.Data;
using Report;
using Report.Base;
using System.Drawing;
using Report.Merging.Item;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var headerStyle = new TableStyle
                {
                    Foreground = Color.Blue,
                    PatternType = PatternType.Solid,
                    BorderLine = BoderLine.Thick,
                    BorderColor = Color.Orange,
                    DocumentTitle = DocumentTitle.Title,
                    FontSize = 9,
                };

            var bodyStyle = new TableStyle
                {
                    FontColor = Color.Black,
                    FontSize = 11,
                    FontName = "Calibri",
                    Foreground = Color.Bisque,
                    PatternType = PatternType.Solid,
                    BorderLine = BoderLine.Thin,
                    BorderColor = Color.Red,
                    DocumentTitle = DocumentTitle.Heading1
                };
            ReportBuilder reportBuilder = new ReportBuilder();
            reportBuilder.AppendComplexHeader(Constants.DataTables.RNPTable.RNPHeader(0, 0, 2013, "6.050201 Системна інженерія", "Компютеризовані та робототехнічні системи", "бакалавр", "Технічна кібернетика", "Факультет інформатики та обчислювальної техніки", "денна", "3 роки 10 місяців", "Молодший інженер з компютерної техніки"));
            reportBuilder.AppendComplexHeader(Constants.DataTables.RNPTable.RNPTableHeader(0, 7, 3, 18, 18));
            var report = reportBuilder.Build();

            var reportRender = new ReportRenderer(report);

            var directory = @"c:\pub\";
            reportRender.ToExcel(directory + "example2.xlsx");
            reportRender.ToExcel(directory + "example3.xlsx");
        }

        public static string Column(int column)
        {
            column--;
            if (column >= 0 && column < 26)
                return ((char)('A' + column)).ToString();
            else if (column > 25)
                return Column(column / 26) + Column(column % 26 + 1);
            else
                throw new Exception("Invalid Column #" + (column + 1).ToString());
        }
    }
}
