using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPA.Spreadsheet
{
    internal class CsvReader : IReader
    {
        private char _separator;
        private IList<IList<string>> _table;
        private IList<string> _row = new List<string>();
        private string _value = string.Empty;

        private interface IState
        {
            IState ReadLine(CsvReader @this, ref string line);
        };

        private class FieldsState : IState
        {
            public IState ReadLine(CsvReader @this, ref string line)
            {
                if (@this._value.Length > 0)
                {
                    // consumes trailing whitespace
                    while (line.Length > 0 && line[0] != @this._separator)
                    {
                        if (!char.IsWhiteSpace(line[0]))
                            throw new ApplicationException("Invalid CSV file format!");

                        line = line.Substring(1);
                    }
                }

                var lastChar = default(char);
                while (line.Length > 0)
                {
                    lastChar = line[0];

                    if (lastChar == @this._separator)
                    {
                        @this.AddValue();
                        line = line.Substring(1);
                    }
                    else if (lastChar == '"')
                    {
                        @this._value = @this._value.TrimStart();
                        if (@this._value.Length > 0)
                        {
                            throw new ApplicationException("Invalid state!");
                        }

                        // chomp first dquote
                        @this._value += lastChar;
                        line = line.Substring(1);
                        return new DquoteState();
                    }
                    else
                    {
                        @this._value += line[0];
                        line = line.Substring(1);
                    }
                }

                if (@this._value.Length > 0
                    || lastChar == @this._separator)
                {
                    @this.AddValue();
                }

                @this.AddRow();
                return this;
            }
        }

        private class DquoteState : IState
        {
            public IState ReadLine(CsvReader @this, ref string line)
            {
                if (@this._value.Length > 1)
                    @this._value += Environment.NewLine;

                while (line.Length > 0)
                {
                    if (line.StartsWith("\"\""))
                    {
                        @this._value += "\"";
                        line = line.Substring(2);
                    }
                    else if (line[0] == '"')
                    {
                        @this._value += line[0];
                        line = line.Substring(1);

                        if (line.Length == 0)
                        {
                            @this.AddValue();
                            @this.AddRow();
                        }

                        return new FieldsState();
                    }
                    else
                    {
                        @this._value += line[0];
                        line = line.Substring(1);
                    }
                }

                return this;
            }
        }

        private void AddValue()
        {
            // remove spaces from the beginning
            while (_value.Length > 0 && char.IsWhiteSpace(_value.First()))
                _value = _value.Substring(1);

            // remove trailing spaces
            while (_value.Length > 0 && char.IsWhiteSpace(_value.Last()))
                _value = _value.Substring(0, _value.Length - 1);

            // remove dquotes
            if (_value.StartsWith("\""))
            {
                _value = _value.Substring(1);
                _value = _value.Substring(0, _value.Length - 1);                
            }

            _row.Add(_value);
            _value = string.Empty;
        }

        private void AddRow()
        {
            _table.Add(_row);
            _row = new List<string>();
        }

        private IState _state = new FieldsState();

        public CsvReader()
            : this(default(char))
        {

        }

        public CsvReader(char separator)
        {
            if (separator != default(char)
                && separator != ';'
                && separator != ',')
            {
                throw new ArgumentException("Invalid separator: should be default, comma or semicolon only");
            }

            _separator = separator;
        }

        public IList<IList<string>> Read(FileInfo inputFile)
        {
            _table = new List<IList<string>>();

            using (var reader = new StreamReader(inputFile.FullName, Encoding.Default, true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    ProcessLine(line);
                }

                if (_row.Count > 0 || _value.Length > 0)
                    throw new ApplicationException("Invalid CSV file");
            }

            return _table;
        }
        public IList<IList<string>> ReadFirstRow(FileInfo inputFile)
        {
            _table = new List<IList<string>>();

            using (var reader = new StreamReader(inputFile.FullName, Encoding.Default, true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    ProcessLine(line);
                    break;
                }

                if (_row.Count > 0 || _value.Length > 0)
                    throw new ApplicationException("Invalid CSV file");
            }

            return _table;
        }
        private void ProcessLine(string line)
        {
            if (_table.Count == 0 && _separator == default(char))
                InspectFirstRow(line);

            while (line.Length > 0)
            {
                _state = _state.ReadLine(this, ref line);
            }
        }

        private void InspectFirstRow(string row)
        {
            while (row.Length > 0)
            {
                if (row.StartsWith("\"\""))
                    row = row.Substring(2);
                else if (row.First() == '"')
                {
                    row = row.Substring(1);
                    while (row.Length > 0 && row.First() != '"')
                        row = row.Substring(1);
                    if (row.Length > 0 && row.First() == '"')
                        row = row.Substring(1);
                }
                else if (row.First() == ';' || row.First() == ',')
                {
                    if (_separator != row.First())
                    {
                        if (_separator != default(char))
                            throw new ApplicationException("Ambiguous CSV file format");

                        _separator = row.First();
                    }
                    row = row.Substring(1);
                }
                else
                {
                    row = row.Substring(1);
                }
            }
        }
    }
}