using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SoftArtisans.OfficeWriter.ExcelWriter; //For ExcelWriter

namespace ExportFile.Helpers
{
    public static class DataTableExtensionsHelper
    {
        public static byte[] ToCsvByteArray(this DataTable dt)
        {
            var result = new StringBuilder();
            for (var i = 0; i < dt.Columns.Count; i++)
            {
                result.Append(dt.Columns[i].ColumnName);
                result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
            }

            foreach (DataRow row in dt.Rows)
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    // TC-861 - CSV rules: http://en.wikipedia.org/wiki/Comma-separated_values#Basic_rules_and_examples
                    // 1. if the data has quote, escape the quote in the data
                    // 2. if the data contains the delimiter (in our case ','), double-quote it
                    // 3. if the data contains the new-line, double-quote it.
                    var data = row[i].ToString();
                    if (data.Contains("\""))
                    {
                        data = data.Replace("\"", "\"\"");
                    }

                    if (data.Contains(","))
                    {
                        data = String.Format("\"{0}\"", data);
                    }

                    if (data.Contains(Environment.NewLine))
                    {
                        data = String.Format("\"{0}\"", data);
                    }
                    result.Append(data);
                    result.Append(i == dt.Columns.Count - 1 ? "\n" : ",");
                }
            }
            return Encoding.Default.GetBytes(result.ToString());
        }

        // Uses Softartisans library
        public static Stream ToExcelStream(this DataTable dt)
        {
            var xla = new ExcelApplication();
            var wb = xla.Create(ExcelApplication.FileFormat.Xls);

            var firstSheet = wb.Worksheets[0];

            firstSheet.Cells[0, 0].Value = "Exported By:";
            firstSheet.Cells[1, 0].Value = "Exported Date:";

            firstSheet.Cells[0, 1].Value = "Jane Smith";
            firstSheet.Cells[1, 1].Value = DateTime.Now.ToString(CultureInfo.InvariantCulture);

            for (var c = 0; c < dt.Columns.Count; c++)
                firstSheet.Cells[3, c].Value = dt.Columns[c].ColumnName;

            var importCell = firstSheet.Cells[4, 0];
            firstSheet.ImportData(dt, importCell);

            return xla.SaveToStream(wb);
        }

        // Uses Epplus library
        public static byte[] ToXlsxByteArray(this DataTable dt)
        {
            byte[] bytes;
            using (var pck = new ExcelPackage())
            {
                //Create the worksheet
                var ws = pck.Workbook.Worksheets.Add(DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss"));

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(dt, true);

                //Format the header for column 1-3
                using (ExcelRange rng = ws.Cells["A1:C1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }

                //Write it back to the client
                bytes = pck.GetAsByteArray();
            }
            return bytes;
        }
    }
}