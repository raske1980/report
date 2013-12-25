using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace Report
{
    public static class MergeAPI
    {
        public static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name, string text)
        {
            MergeTwoCells(worksheet, cell1Name, cell2Name, text, 0);
        }

        public static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name, string text, uint styleId)
        {
            if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
            {
                return;
            }

            // Verify if the specified cells exist, and if they do not exist, create them.
            if (styleId > 0)
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name, text, styleId);
            else
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name, text);
            //CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
            mergeCells.Append(mergeCell);

            worksheet.Save();
        }

        public static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName, string text, uint styleid)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => (r.RowIndex != null && r.RowIndex.Value == rowIndex));

            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName), CellValue = new CellValue(text), StyleIndex = styleid, DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                row.Append(cell);
                InsertNewRow(worksheet, rowIndex, row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName), CellValue = new CellValue(text), StyleIndex = styleid, DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }

        public static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName, string text)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => (r.RowIndex != null && r.RowIndex.Value == rowIndex));

            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName), CellValue = new CellValue(text), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                row.Append(cell);
                InsertNewRow(worksheet, rowIndex, row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName), CellValue = new CellValue(text), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }

        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => (r.RowIndex != null && r.RowIndex.Value == rowIndex));

            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                row.Append(cell);
                InsertNewRow(worksheet, rowIndex, row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName), DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }

        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        private static void InsertNewRow(Worksheet worksheet, uint index, Row row)
        {
            IEnumerable<Row> underRows = worksheet.Descendants<Row>().OrderByDescending(r => int.Parse(r.RowIndex)).Where(r => (r.RowIndex != null && r.RowIndex.Value < index));
            if (underRows != null && underRows.Count() > 0)
                worksheet.Descendants<SheetData>().First().InsertAfter(row, underRows.First());
            else
                worksheet.Descendants<SheetData>().First().Append(row);

        }
    }
}
