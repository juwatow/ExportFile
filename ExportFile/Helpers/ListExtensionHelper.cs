using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExportFile.Models;

namespace ExportFile.Helpers
{
    public static class ListExtensionHelper
    {
        public static DataTable ToDataTable(this List<UserModel> list)
        {
            // Dictionary of "user property/export title"
            // Keys must match the UserModel properties
            var headers = new Dictionary<string, string>
            {
                {"FirstName", "First Name"},
                {"LastName", "Last Name"},
                {"DateOfBirth", "Date Of Birth"}
            };

            var dtDataExport = new DataTable("DataExport");
            var properties = typeof(UserModel).GetProperties();
            // Add column objects to the table.
            foreach (var column in properties.Select(property => new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = headers[property.Name]
            }))
            {
                dtDataExport.Columns.Add(column);
            }
            foreach (var user in list)
            {
                // Once a table has been created, use the NewRow to create a DataRow.
                var row = dtDataExport.NewRow();

                // Then add the new row to the collection.
                foreach (var property in properties)
                {
                    row[headers[property.Name]] = user.GetType().GetProperty(property.Name).GetValue(user, null);
                }
                dtDataExport.Rows.Add(row);
            }
            return dtDataExport;
        }
    }
}