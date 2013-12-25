using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Report.Merging.Item;
using Report.Base;

namespace Constants
{
    public static class DataTables
    {
        public static string GetColumn(int column)
        {
            column--;
            if (column >= 0 && column < 26)
                return ((char)('A' + column)).ToString();
            else if (column > 25)
                return GetColumn(column / 26) + GetColumn(column % 26 + 1);
            else
                throw new Exception("Invalid Column #" + (column + 1).ToString());
        }

        public static class RNPTable
        {
            private static ComplexHeader _RNPHeader;
            public static ComplexHeader RNPHeader(int columnOffset, int rowOffset, int startYear, string napryam, string discipline, string okr, string cafedra, string faculty, string studyForm, string studyTerm, string qualification)
            {
                if (_RNPHeader == null)
                {
                    _RNPHeader = new ComplexHeader();

                    Style titleCenter = new Style() { Aligment = new Aligment() { VerticalAligment = VerticalAligment.Center, HorizontalAligment = HorizontalAligment.Center }, FontSize = 20 };
                    titleCenter.TextStyle.Add(TextStyleType.Bold);
                    Style subTitleCenter = new Style() { Aligment = new Aligment() { VerticalAligment = VerticalAligment.Center, HorizontalAligment = HorizontalAligment.Center }, FontSize = 16 };
                    subTitleCenter.TextStyle.Add(TextStyleType.Bold);
                    Style left = new Style() { Aligment = new Aligment() { VerticalAligment = VerticalAligment.Center, HorizontalAligment = HorizontalAligment.Left }, FontSize = 11 };

                    Style leftItalicUnderLine = new Style() { Aligment = new Aligment() { VerticalAligment = VerticalAligment.Center, HorizontalAligment = HorizontalAligment.Left }, FontSize = 11 };
                    leftItalicUnderLine.TextStyle.Add(TextStyleType.Italic);
                    leftItalicUnderLine.TextStyle.Add(TextStyleType.Underline);
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(1, 1, columnOffset, rowOffset), GetCellWithOffset(21, 1, columnOffset, rowOffset), "Робочий навчальний план", titleCenter));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(1, 2, columnOffset, rowOffset), GetCellWithOffset(21, 2, columnOffset, rowOffset), string.Format("на {0}/{1} навчальний рік", startYear, (startYear + 1)), subTitleCenter));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(4, 3, columnOffset, rowOffset), GetCellWithOffset(7, 3, columnOffset, rowOffset), "Напрям підготовки (код і назва)", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(8, 3, columnOffset, rowOffset), GetCellWithOffset(13, 3, columnOffset, rowOffset), napryam, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(4, 4, columnOffset, rowOffset), GetCellWithOffset(7, 4, columnOffset, rowOffset), "Програма професійного спрямування", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(8, 4, columnOffset, rowOffset), GetCellWithOffset(13, 4, columnOffset, rowOffset), discipline, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(4, 5, columnOffset, rowOffset), GetCellWithOffset(7, 5, columnOffset, rowOffset), "Освітньо-кваліфікаційний рівень", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(8, 5, columnOffset, rowOffset), GetCellWithOffset(13, 5, columnOffset, rowOffset), okr, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(4, 6, columnOffset, rowOffset), GetCellWithOffset(7, 6, columnOffset, rowOffset), "Випускова кафедра", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(8, 6, columnOffset, rowOffset), GetCellWithOffset(13, 6, columnOffset, rowOffset), cafedra, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(14, 3, columnOffset, rowOffset), GetCellWithOffset(16, 3, columnOffset, rowOffset), "Факультет(інститут)", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(17, 3, columnOffset, rowOffset), GetCellWithOffset(21, 3, columnOffset, rowOffset), faculty, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(14, 4, columnOffset, rowOffset), GetCellWithOffset(16, 4, columnOffset, rowOffset), "Форма навчання", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(17, 4, columnOffset, rowOffset), GetCellWithOffset(21, 4, columnOffset, rowOffset), studyForm, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(14, 5, columnOffset, rowOffset), GetCellWithOffset(16, 5, columnOffset, rowOffset), "Термін навчання", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(17, 5, columnOffset, rowOffset), GetCellWithOffset(21, 5, columnOffset, rowOffset), studyTerm, leftItalicUnderLine));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(14, 6, columnOffset, rowOffset), GetCellWithOffset(16, 6, columnOffset, rowOffset), "Кваліфікація", left));
                    _RNPHeader.Add(new ComplexHeaderCell(GetCellWithOffset(17, 6, columnOffset, rowOffset), GetCellWithOffset(21, 6, columnOffset, rowOffset), qualification, leftItalicUnderLine));

                    //_RNPHeader.Add(new ComplexHeaderCell(Column(2 + columnOffset) + (1 + rowOffset).ToString(), Column(2 + columnOffset) + (7 + rowOffset).ToString(), "Назва кафедри", generalCenter));
                }

                return _RNPHeader;
            }

            private static string GetCellWithOffset(int column, int row, int columnOffset, int rowOffset)
            {
                return GetColumn(column + columnOffset) + (row + rowOffset).ToString();
            }



            private static ComplexHeader _RNPTableHeader;
            public static ComplexHeader RNPTableHeader(int columnOffset, int rowOffset, int kurs, int weekCount1, int weekCount2)
            {
                int semestr = kurs * 2 - 1; //первый семестр курса
                if (_RNPTableHeader == null)
                {
                    _RNPTableHeader = new ComplexHeader();

                    Style verticalCenter = new Style() { Aligment = new Aligment() { Rotation = 90, HorizontalAligment = HorizontalAligment.Center } };
                    Style generalCenter = new Style() { Aligment = new Aligment() { VerticalAligment = VerticalAligment.Center, HorizontalAligment = HorizontalAligment.Center } };
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(1 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(1 + columnOffset) + (7 + rowOffset).ToString(), "Назва дисципліни", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(2 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(2 + columnOffset) + (7 + rowOffset).ToString(), "Назва кафедри", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(3 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(4 + columnOffset) + (3 + rowOffset).ToString(), "Обсяг дисципліни", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(3 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(3 + columnOffset) + (7 + rowOffset).ToString(), "Кредитів ЕКТС", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(4 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(4 + columnOffset) + (7 + rowOffset).ToString(), "Годин", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(5 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(5 + columnOffset) + (7 + rowOffset).ToString(), "Самостійна робота студентів", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(6 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(13 + columnOffset) + (3 + rowOffset).ToString(), "Контрольні заходи", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(6 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(6 + columnOffset) + (7 + rowOffset).ToString(), "Екзамени", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(7 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(7 + columnOffset) + (7 + rowOffset).ToString(), "Заліки", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(8 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(8 + columnOffset) + (7 + rowOffset).ToString(), "МКР", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(9 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(9 + columnOffset) + (7 + rowOffset).ToString(), "Курсові проекти", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(10 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(10 + columnOffset) + (7 + rowOffset).ToString(), "Курсові роботи", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(11 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(11 + columnOffset) + (7 + rowOffset).ToString(), "РГР, РР", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(12 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(12 + columnOffset) + (7 + rowOffset).ToString(), "ДКР", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(13 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(13 + columnOffset) + (7 + rowOffset).ToString(), "Реферати", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (1 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (2 + rowOffset).ToString(), kurs + " курс", generalCenter));
                    //_RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (2 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (2 + rowOffset).ToString(), kurs + " курс", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (3 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (3 + rowOffset).ToString(), "Кількість годин", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(17 + columnOffset) + (4 + rowOffset).ToString(), semestr + " семестр", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (5 + rowOffset).ToString(), GetColumn(17 + columnOffset) + (5 + rowOffset).ToString(), weekCount1 + " тижнів", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(14 + columnOffset) + (6 + rowOffset).ToString(), GetColumn(14 + columnOffset) + (7 + rowOffset).ToString(), "Всього", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(15 + columnOffset) + (6 + rowOffset).ToString(), GetColumn(17 + columnOffset) + (6 + rowOffset).ToString(), "У тому числі", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(15 + columnOffset) + (7 + rowOffset).ToString(), "", "Лекції", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(16 + columnOffset) + (7 + rowOffset).ToString(), "", "Практичні", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(17 + columnOffset) + (7 + rowOffset).ToString(), "", "Лабораторні", verticalCenter));

                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(18 + columnOffset) + (4 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (4 + rowOffset).ToString(), ((semestr + 1) + " семестр"), generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(18 + columnOffset) + (5 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (5 + rowOffset).ToString(), weekCount2 + " тижнів", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(18 + columnOffset) + (6 + rowOffset).ToString(), GetColumn(18 + columnOffset) + (7 + rowOffset).ToString(), "Всього", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(19 + columnOffset) + (6 + rowOffset).ToString(), GetColumn(21 + columnOffset) + (6 + rowOffset).ToString(), "У тому числі", generalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(19 + columnOffset) + (7 + rowOffset).ToString(), "", "Лекції", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(20 + columnOffset) + (7 + rowOffset).ToString(), "", "Практичні", verticalCenter));
                    _RNPTableHeader.Add(new ComplexHeaderCell(GetColumn(21 + columnOffset) + (7 + rowOffset).ToString(), "", "Лабораторні", verticalCenter));

                    for (int i = 1; i < 22; i++)
                    {
                        _RNPTableHeader.Add(new ComplexHeaderCell(GetCellWithOffset(i, 8, columnOffset, rowOffset), "", i.ToString(), generalCenter));
                    }
                }

                return _RNPTableHeader;
            }


            private static DataTable _RNP;
            public static DataTable RNP
            {
                get
                {
                    if (_RNP == null)
                    {
                        _RNP = new DataTable();
                        for (int i = 0; i < RNPColumn.Count; i++)
                            _RNP.Columns.Add(RNPColumn[i]);
                    }

                    return _RNP;
                }
            }

            private static SortedDictionary<int, string> _RNPColumn;
            public static SortedDictionary<int, string> RNPColumn
            {
                get
                {
                    if (_RNPColumn == null)
                    {
                        _RNPColumn = new SortedDictionary<int, string>();
                        _RNPColumn.Add(Columns.RowName, "Назва дисципліни");
                        _RNPColumn.Add(Columns.DivName, "Назва кафедри");
                        _RNPColumn.Add(Columns.CreditECTS, "Кредитів ECTS");
                        _RNPColumn.Add(Columns.HourFact, "Годин");
                        //_RNPRow.Add(ExamenCount, "Аудиторних годин всього");
                        //_RNPRow.Add(5, "Лекцій");
                        //_RNPRow.Add(6, "Практичні");
                        //_RNPRow.Add(7, "Лабораторні");
                        //_RNPRow.Add(7, "Самостійна робота студентів");
                        _RNPColumn.Add(Columns.ExamenCount, "Екзамени");
                        _RNPColumn.Add(Columns.ZalikCount, "Заліки");
                        _RNPColumn.Add(Columns.MkrCount, "МКР");
                        _RNPColumn.Add(Columns.CourseProjectCount, "Курсові проекти");
                        _RNPColumn.Add(Columns.CourseWorkCount, "Курсові роботи");
                        _RNPColumn.Add(Columns.RgrCount, "РГР");
                        _RNPColumn.Add(Columns.DkrCount, "ДКР");
                        _RNPColumn.Add(Columns.ReferatCount, "Реферати");
                    }

                    return _RNPColumn;
                }
            }

            public static class Columns
            {
                public const int RowName = 0;
                public const int DivName = 1;
                public const int CreditECTS = 2;
                public const int HourFact = 3;
                public const int ExamenCount = 4;//9;
                public const int ZalikCount = 5;//10;
                public const int MkrCount = 6;//11;
                public const int CourseProjectCount = 7;//12;
                public const int CourseWorkCount = 8;//13;
                public const int RgrCount = 9;//14;
                public const int DkrCount = 10;//15;
                public const int ReferatCount = 11;//16;
            }
        }
    }
}
