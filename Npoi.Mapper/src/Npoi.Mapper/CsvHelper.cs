using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Npoi.Mapper
{
    /// <summary>
    /// Provide helper methods for import and export CSV file.
    /// Example:
    ///   CsvExport myExport = new CsvHelper();
    ///
    ///   myExport.AddRow();
    ///   myExport["Region"] = "New York, USA";
    ///   myExport["Sales"] = 100000;
    ///   myExport["Date Opened"] = new DateTime(2003, 12, 31);
    ///
    ///   myExport.AddRow();
    ///   myExport["Region"] = "Sydney \"in\" Australia";
    ///   myExport["Sales"] = 50000;
    ///   myExport["Date Opened"] = new DateTime(2005, 1, 1, 9, 30, 0);
    ///
    /// Then you can do any of the following three output options:
    ///   string myCsv = myExport.Export();
    ///   myExport.ExportToFile("Somefile.csv");
    ///   byte[] myCsvData = myExport.ExportToBytes();
    /// </summary>
	public class CsvHelper // TODO: integrate CSV format
    {
        /// <summary>
        /// To keep the ordered list of column names
        /// </summary>
        private readonly List<string> _fields = new List<string>();

        /// <summary>
        /// The list of rows
        /// </summary>
        private readonly List<Dictionary<string, object>> _rows = new List<Dictionary<string, object>>();

        /// <summary>
        /// The current row
        /// </summary>
        private Dictionary<string, object> CurrentRow { get { return _rows[_rows.Count - 1]; } }

        /// <summary>
        /// Set a value on this column
        /// </summary>
        public object this[string field]
        {
            set
            {
                // Keep track of the field names, because the dictionary loses the ordering
                if (!_fields.Contains(field)) _fields.Add(field);
                CurrentRow[field] = value;
            }
        }

        /// <summary>
        /// Call this before setting any fields on a row
        /// </summary>
        public void AddRow()
        {
            _rows.Add(new Dictionary<string, object>());
        }

        /// <summary>
        /// Add a list of typed objects, maps object properties to CsvFields
        /// </summary>
        public void AddRows<T>(IEnumerable<T> list)
        {
            foreach (var obj in list)
            {
                AddRow();
                var values = obj.GetType().GetProperties();
                foreach (var value in values)
                {
                    this[value.Name] = value.GetValue(obj, null);
                }
            }
        }

        /// <summary>
        /// Converts a value to how it should output in a csv file
        /// If it has a comma, it needs surrounding with double quotes
        /// Eg Sydney, Australia -> "Sydney, Australia"
        /// Also if it contains any double quotes ("), then they need to be replaced with quad quotes[sic] ("")
        /// Eg "Dangerous Dan" McGrew -> """Dangerous Dan"" McGrew"
        /// </summary>
        public static string MakeValueCsvFriendly(object value)
        {
            if (value == null) return "";

            if (value is DateTime)
            {
                if (((DateTime)value).TimeOfDay.TotalSeconds < 0.1)
                    return ((DateTime)value).ToString("yyyy-MM-dd");
                return ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
            }
            string output = value.ToString().Trim();
            if (output.Contains(",") || output.Contains("\"") || output.Contains("\n") || output.Contains("\r"))
                output = '"' + output.Replace("\"", "\"\"") + '"';

            if (output.Length > 30000) //cropping value for stupid Excel
            {
                if (output.EndsWith("\""))
                {
                    output = output.Substring(0, 30000);
                    if (output.EndsWith("\"") && !output.EndsWith("\"\"")) //rare situation when cropped line ends with a '"'
                        output += "\""; //add another '"' to escape it
                    output += "\"";
                }
                else
                    output = output.Substring(0, 30000);
            }
            return output;
        }

        /// <summary>
        /// Outputs all rows as a CSV, returning one string at a time
        /// </summary>
        private IEnumerable<string> ExportToLines()
        {
            yield return "sep=,";

            // The header
            yield return string.Join(",", _fields);

            // The rows
            foreach (Dictionary<string, object> row in _rows)
            {
                foreach (string k in _fields.Where(f => !row.ContainsKey(f)))
                {
                    row[k] = null;
                }
                yield return string.Join(",", _fields.Select(field => MakeValueCsvFriendly(row[field])));
            }
        }

        /// <summary>
        /// Output all rows as a CSV returning a string
        /// </summary>
        public string Export()
        {
            StringBuilder sb = new StringBuilder();

            foreach (string line in ExportToLines())
            {
                sb.AppendLine(line);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Exports to a file
        /// </summary>
        public void ExportToFile(string path)
        {
            File.WriteAllLines(path, ExportToLines(), Encoding.UTF8);
        }

        /// <summary>
        /// Exports as raw UTF8 bytes
        /// </summary>
        public byte[] ExportToBytes()
        {
            var data = Encoding.UTF8.GetBytes(Export());
            return Encoding.UTF8.GetPreamble().Concat(data).ToArray();
        }
    }
}
