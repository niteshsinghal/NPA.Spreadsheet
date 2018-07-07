using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Ionic.Zlib;

namespace NPA.Spreadsheet
{
    /// <summary>
    /// Defines extensions for IList&lt;IList&lt;string>> types (table).
    /// </summary>
    public static class TableExtensions
    {
        /// <summary>
        /// Trim empty lines at the beginning and at the end, cut left and
        /// right empty columns, and fold cells when the row is shorter than
        /// its siblings. This function does not remove empty rows.
        /// </summary>
        public static void Normalize(this IList<IList<string>> @this)
        {
            // Trim empty rows before the header
            while (@this.Count > 0 && @this[0].Count == 0)
                @this.RemoveAt(0);

            // Trim empty rows at the end
            while (@this.Count > 0 && @this[@this.Count - 1].Count == 0)
                @this.RemoveAt(@this.Count - 1);

            // Normalize row lengths
            var max = @this.MaxColumns();
            foreach (var row in @this)
            {
                while (row.Count < max)
                    row.Add("");
            }

            if (@this.Count > 0)
            {
                // Remove left empty columns
                for (var i = 0; i < @this[0].Count; i++)
                {
                    if (@this.IsColumnEmpty(i))
                        @this.RemoveColumn(i);
                    else
                        break;
                }

                // Remove right empty columns
                for (var i = @this[0].Count - 1; i >= 0; i--)
                {
                    if (@this.IsColumnEmpty(i))
                        @this.RemoveColumn(i);
                    else
                        break;
                }
            }
        }

        /// <summary>
        /// Remove empty rows from the table.
        /// </summary>
        public static void RemoveEmptyRows(this IList<IList<string>> @this)
        {
            for (var index = 0; index < @this.Count; index++)
            {
                var row = @this[index];
                if (row.IsRowEmpty())
                    @this.Remove(row);
            }
        }

        /// <summary>
        /// Check if an entire column is empty.
        /// </summary>
        public static bool IsColumnEmpty(this IList<IList<string>> @this, int n)
        {
            var empties = @this.Count(row => n < row.Count && string.IsNullOrEmpty(row[n]));
            return empties == @this.Count;
        }

        /// <summary>
        /// Check if an entire row is empty.
        /// </summary>
        public static bool IsRowEmpty(this IList<string> @this)
        {
            var empties = @this.Count(string.IsNullOrEmpty);
            return empties == @this.Count;
        }

        /// <summary>
        /// Add a table column (cells will have empty strings).
        /// </summary>
        public static void AddColumn(this IList<IList<string>> @this, int index)
        {
            foreach (var row in @this)
            {
                row.Insert(index, "");
            }
        }

        /// <summary>
        /// Remove an entire column.
        /// </summary>
        public static void RemoveColumn(this IList<IList<string>> @this, int n)
        {
            foreach (var row in @this)
            {
                if (n < row.Count)
                    row.RemoveAt(n);
            }
        }

        /// <summary>
        /// Return the max column length.
        /// </summary>
        public static int MaxColumns(this IList<IList<string>> @this)
        {
            return @this.Select(row => row.Count).Concat(new[] {0}).Max();
        }

        /// <summary>
        /// Convert the current table to a CSV string.
        /// </summary>
        public static string ToString(this IList<IList<string>> @this)
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new StreamWriter(stream))
                {
                    @this.Serialize(writer);
                    writer.Flush();
                    stream.Seek(0, SeekOrigin.Begin);
                }
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Serialize the current table to a CSV file.
        /// </summary>
        public static void Serialize(this IList<IList<string>> @this, StreamWriter writer, char separator = ';')
        {
            @this.Aggregate(writer, (w, row) =>
            {
                var list = row.Aggregate(new List<string>(), (l, v) =>
                {
                    // - Fields that contain double quote characters must be surounded by double-quotes,
                    // and the embedded double-quotes must each be represented by a pair of consecutive
                    // double quotes.
                    if (v.Contains('\"'))
                    {
                        l.Add("\"" + v.Replace("\"", "\"\"") + "\"");
                    }
                    // - Fields with embedded commas must be delimited with double-quote characters.
                    // - A field that contains embedded line-breaks must be surounded by double-quotes.
                    // - Fields with leading or trailing spaces must be delimited with double-quote characters.
                    else if (v.Contains(separator) 
                            || v.Contains(Environment.NewLine)
                            || (v.Length > 0 && char.IsWhiteSpace(v.First()) || char.IsWhiteSpace(v.Last()))
                            )
                    {
                        l.Add("\"" + v + "\"");
                    }
                    // - Any other field will be written without double-quotes.
                    else
                    {
                        l.Add(v);
                    }
                    return l;
                });

                w.WriteLine(string.Join(separator.ToString(), list));
                return w;
            });
        }
    }
}
