using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RNPExcelExport.Data.Collection;
using System.Data.SqlClient;
using Constants.SQLStoredPrcedures;
using SQLExecuter;
using RNPExcelExport.Data.Item;
using System.Data;
using Report;
using Constants;
using Report.Base;
using System.Drawing;

namespace RNPExcelExport.Data
{

    public static class RnpAPI
    {
        public static RNPBaseCollection GetRNPBase()
        {
            RNPBaseCollection result = null;
            using (Executer exec = new Executer())
            {
                SqlDataReader reader = exec.Execute(RNPProcedures.GetRNPBase);
                while (reader.Read())
                {
                    if (result == null)
                        result = new RNPBaseCollection();

                    RNPBase item = new RNPBase();
                    item.CourseProjectCount = Converter.ToInt16Nullable(reader["CourseProjectCount"]);
                    item.CourseWorkCount = Converter.ToInt16Nullable(reader["CourseWorkCount"]);
                    item.CreditECTS = Converter.ToDouble(reader["CreditECTS"]);
                    item.DivName = Converter.ToString(reader["DivName"]);
                    item.DkrCount = Converter.ToInt16Nullable(reader["DkrCount"]);
                    item.ExamenCount = Converter.ToInt16Nullable(reader["ExamenCount"]);
                    item.HourFact = Converter.ToInt32(reader["HourFact"]);
                    item.MkrCount = Converter.ToInt16Nullable(reader["MkrCount"]);
                    item.ReferatCount = Converter.ToInt16Nullable(reader["ReferatCount"]);
                    item.RgrCount = Converter.ToInt16(reader["RgrCount"]);
                    item.RowName = Converter.ToString(reader["RowName"]);
                    item.ZalikCount = Converter.ToInt16Nullable(reader["ZalikCount"]);

                    result.Add(item);
                }
            }

            return result;
        }

        public static DataTable GetRNPTable(RNPBaseCollection data)
        {
            if (data == null || data.Count() == 0)
                return null;

            DataTable table = Constants.DataTables.RNPTable.RNP;

            foreach (RNPBase item in data)
            {
                DataRow row = table.NewRow();

                row[Constants.DataTables.RNPTable.Columns.RowName] = item.RowName;
                row[Constants.DataTables.RNPTable.Columns.DivName] = item.DivName;
                row[Constants.DataTables.RNPTable.Columns.CourseProjectCount] = item.CourseProjectCount;
                row[Constants.DataTables.RNPTable.Columns.CourseWorkCount] = item.CourseWorkCount;
                row[Constants.DataTables.RNPTable.Columns.CreditECTS] = item.CreditECTS;
                row[Constants.DataTables.RNPTable.Columns.DkrCount] = item.DkrCount;
                row[Constants.DataTables.RNPTable.Columns.ExamenCount] = item.ExamenCount;
                row[Constants.DataTables.RNPTable.Columns.HourFact] = item.HourFact;
                row[Constants.DataTables.RNPTable.Columns.MkrCount] = item.MkrCount;
                row[Constants.DataTables.RNPTable.Columns.ReferatCount] = item.ReferatCount;
                row[Constants.DataTables.RNPTable.Columns.RgrCount] = item.RgrCount;
                row[Constants.DataTables.RNPTable.Columns.ZalikCount] = item.ZalikCount;

                table.Rows.Add(row);
            }

            return table;
        }

        public static void RenderRNPTable(ReportBuilder builder, DataTable table)
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

            builder.AppendTable(table, DataTables.RNPTable.RNPColumn.Values, bodyStyle, headerStyle);
        }
    }
}
