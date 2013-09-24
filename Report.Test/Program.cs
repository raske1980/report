using System;
using Report.Base;
using System.Data;
using System.Drawing;

namespace Report.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var bigTextblock = String.Empty;
            
            #region Big textblock

            bigTextblock= @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec id nibh sed eros mollis varius vel id arcu. Phasellus facilisis, urna vitae laoreet blandit, justo neque vehicula metus, sit amet lobortis quam nulla quis eros. Vivamus in placerat nulla. Pellentesque ac elit adipiscing, ultrices magna ac, adipiscing libero. Morbi sit amet auctor est. Pellentesque ac adipiscing elit. Vivamus sollicitudin sollicitudin est consectetur cursus. Nullam ac posuere ligula. Vestibulum a auctor orci. Mauris in ipsum ornare, laoreet quam eget, porta urna. Donec quis ultricies velit. Donec suscipit non odio at gravida.

Nullam mauris arcu, posuere in rhoncus at, facilisis eu elit. Aliquam sed sapien libero. Nunc consectetur quis elit et viverra. Nunc pulvinar faucibus sem, vitae suscipit diam consequat sit amet. Sed a mi purus. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; Praesent mi felis, lobortis sit amet nibh vitae, dapibus sodales sem. Fusce ultrices eros at nisi rutrum, vel ornare nisl tincidunt.

Suspendisse tristique ante non laoreet pellentesque. Duis quis magna tempor, suscipit ligula et, fermentum neque. Nunc dignissim in sem ac cursus. Fusce ac diam facilisis, tempus erat a, imperdiet leo. Aenean varius bibendum suscipit. Aenean varius sapien velit, vel vestibulum turpis ullamcorper ut. Fusce eu auctor mi. Cras id porta lorem, vitae imperdiet arcu. Morbi auctor erat vitae tortor pulvinar blandit. Mauris augue sem, aliquet sed velit sit amet, consequat pulvinar nunc. Nam fermentum bibendum arcu, eget ultrices ipsum.

In hac habitasse platea dictumst. Integer fringilla tincidunt ligula posuere consectetur. Nunc hendrerit leo quam, nec gravida erat facilisis nec. Donec iaculis mi a nibh sodales porta. Suspendisse scelerisque magna at urna tristique tincidunt. Praesent molestie iaculis semper. Donec tincidunt enim quis risus mattis rhoncus. Sed sed purus eu lacus dapibus hendrerit. Nunc vehicula nulla et magna commodo vestibulum. Fusce ac tortor vel massa dictum bibendum. Etiam id tortor nec lacus laoreet dapibus. Morbi non scelerisque ante. Nam non porta lorem. Etiam ultrices dolor ac porta posuere.

Etiam in est ac tortor placerat tincidunt at non lorem. Praesent nibh ipsum, euismod dignissim laoreet vel, venenatis eget lacus. Aenean fringilla purus vel lacus rutrum, in dapibus dui bibendum. In sed consectetur turpis. Aenean velit sem, faucibus quis congue ac, vehicula ut leo. Sed molestie erat ut condimentum bibendum. Donec volutpat nibh vitae tempus elementum.

In ullamcorper, erat ut egestas accumsan, tellus mauris pellentesque purus, molestie euismod turpis magna ac mi. Vestibulum fringilla ullamcorper tortor, a tincidunt quam convallis eu. Aenean consequat pharetra venenatis. Maecenas pellentesque, orci quis posuere pulvinar, massa magna fermentum sem, at consequat sem nibh vel augue. Integer non libero in neque commodo ultrices. Nunc quis tristique dui. Maecenas in condimentum neque, vel porta neque. Etiam laoreet euismod purus sit amet vulputate. Mauris at ligula tellus. Proin sagittis varius arcu. Nullam nec aliquam arcu, non faucibus sem. Proin varius fringilla arcu, auctor ultricies tellus mattis ac. Sed justo felis, bibendum eget fermentum dapibus, venenatis et dui. Praesent adipiscing nulla nisl, non ultrices nulla tincidunt eu. Proin sed nunc mi. Sed ac magna ornare, suscipit augue sit amet, sollicitudin tortor.

Sed consectetur tristique risus, vel malesuada ligula vulputate molestie. Vivamus sollicitudin fringilla arcu sed facilisis. Vivamus ut venenatis dui, sit amet viverra diam. Nullam condimentum neque quis bibendum blandit. Pellentesque vestibulum libero eget dapibus aliquam. In tincidunt nibh sed nulla tempus accumsan. Proin eleifend magna pretium sapien aliquet, id pulvinar ante luctus.

Quisque vestibulum varius urna eu tempus. Phasellus laoreet, tellus sed accumsan posuere, mauris sapien aliquet ligula, sit amet aliquam ligula mauris at nulla. Nulla tincidunt nibh vel lobortis convallis. Proin tristique ligula in nibh cursus gravida. Maecenas euismod est est, ac auctor augue congue ut. Quisque ac erat rhoncus, auctor libero quis, aliquam turpis. Integer ac molestie ante. Cras non neque sed est vulputate blandit et vestibulum nibh. Fusce in rhoncus justo, sit amet luctus sem. Nam fringilla ut dolor eu pharetra. Suspendisse consectetur nibh a dictum varius. Maecenas fermentum interdum nulla, quis pulvinar lorem interdum nec.

Pellentesque vitae libero id ligula vulputate fringilla. Sed sollicitudin eu ipsum non pharetra. In pretium, massa nec consectetur sodales, mi eros suscipit lectus, ac vehicula massa odio vel dui. Cras posuere, justo pretium pellentesque sagittis, metus erat tempor felis, id convallis purus quam et diam. Pellentesque blandit leo non nulla blandit pulvinar vitae nec dolor. Sed laoreet sem vel nibh mattis, in convallis nibh consequat. Pellentesque eget varius quam, sit amet volutpat ante. Etiam convallis eros ut erat condimentum, vel vehicula diam placerat. Nunc sed enim consectetur ligula dignissim ultricies.

Fusce sed neque viverra, mattis orci ac, ornare est. Morbi pellentesque enim pellentesque, vestibulum urna vel, blandit neque. Pellentesque a pulvinar libero. Fusce quis scelerisque metus. Sed ac orci ac libero ornare luctus eget porttitor est. Donec adipiscing placerat dui, sed sagittis augue facilisis eget. Nullam felis nisi, consequat sed justo sed, fringilla convallis magna. In ac elementum diam. Donec suscipit massa ut commodo ultrices.";

            #endregion;

            var dataTable1 = new DataTable();
            var dataTable2 = new DataTable();

            for (int i = 0; i < 10; i++)
            {
                dataTable1.Columns.Add(i.ToString());
            }

            for (int i = 0; i < 15; i++)
            {
                var row = dataTable1.NewRow();

                for (int j = 0; j < dataTable1.Columns.Count; j++)
                {
                    row[j] = (i * j) + " ячейка";
                }

                dataTable1.Rows.Add(row);
            }
                  
            dataTable2.Columns.Add("Column 1");
            dataTable2.Columns.Add("Column 2");

            for (int i = 0; i < 5; i++)
            {
                var row = dataTable2.NewRow();
                row[0] = i;
                row[1] = "строка";
                dataTable2.Rows.Add(row);
            }

            var style0 = new TableStyle();

            var style1 = new TableStyle
                             {
                                 Foreground = Color.Blue,
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
                                 Foreground = Color.Gray,
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

            var style5 = new TableStyle()
            {
                FontColor = Color.Black,
                FontSize = 12,
                FontName = "Calibri",
                Foreground = Color.White,
                PatternType = PatternType.Solid,
                DocumentTitle = DocumentTitle.None
            };

            var tableStyle1 = new TableStyle
                                  {
                                      Foreground = Color.Blue,
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
                                      Foreground = Color.Gray,
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
                                      Foreground = Color.Gray,
                                      PatternType = PatternType.Solid,
                                      BorderLine = BoderLine.Thick,
                                      BorderColor = Color.Red,
                                      DocumentTitle = DocumentTitle.None
                                  };

            string[] abs = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" };

            var reportBuilder = new ReportBuilder();

            reportBuilder.AppendTextBlock("Test title1", style0);
            reportBuilder.AppendTextBlock("Test title2", style2);
            reportBuilder.AppendTextBlock("Test title3", style3);
            reportBuilder.AppendTextBlock("Test title4", style4);
            reportBuilder.AppendTextBlock("Test title3", style3);
            reportBuilder.AppendTextBlock("Test title1", style1);
            reportBuilder.AppendTextBlock("Test title4", style4);
            reportBuilder.AppendNewSection();


            reportBuilder.AppendTable(dataTable1, abs, tableStyle1, tableStyle2);
            reportBuilder.AppendNewSection();
            reportBuilder.AppendTable(dataTable1, abs, tableStyle2, tableStyle3);
            reportBuilder.AppendNewLine();
            reportBuilder.AppendTable(dataTable1);


            reportBuilder.AppendTextBlock(bigTextblock, style5);
            reportBuilder.AppendTable(dataTable2, new [] {"Колонка 1", "rjkjyrf 2"});
            reportBuilder.AppendTextBlock(bigTextblock, style5);


            var report = reportBuilder.Build();

            var reportRender = new ReportRenderer(report);

            var directory = @"c:\pub\";

            reportRender.ToExcel(directory + "example.xlsx");
            reportRender.ToHtml(directory + "example.html");
            reportRender.ToPdf(directory + "example.pdf");
            reportRender.ToWord(directory + "example.docx");
        }

    }
}
