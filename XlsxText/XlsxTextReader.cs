using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;

namespace XlsxText
{
    public class XlsxTextCellReference
    {
        private string _value;
        public string Value
        {
            get => _value;
            set
            {
                value = value?.Trim().ToUpper();
                if (value == null || !Regex.IsMatch(value, @"^[A-Z]+\d+$"))
                    throw new Exception("Invalid value of cell reference");

                int row = int.Parse(new Regex(@"\d+$").Match(value).Value);
                if (row < 1)
                    throw new Exception("Invalid value of cell reference");

                string colValue = new Regex(@"^[A-Z]+").Match(value).Value;
                int col = 0;
                for (int i = colValue.Length - 1, multiple = 1; i >= 0; --i, multiple *= 26)
                {
                    int n = colValue[i] - 'A' + 1;
                    col += (n * multiple);
                }

                _value = value;
                Row = row;
                Col = col;
            }
        }
        /// <summary>
        /// Number of rows, starting from 1
        /// </summary>
        public int Row { get; private set; }
        /// <summary>
        /// Number of columns, starting from 1
        /// </summary>
        public int Col { get; private set; }

        public XlsxTextCellReference(string value)
        {
            Value = value;
        }

        public XlsxTextCellReference(int row, int col) : this(RowColToValue(row, col))
        {
        }
        public override string ToString() { return Value; }

        /// <summary>
        /// Row and Col convert to Value
        /// </summary>
        /// <param name="row">Number of rows, starting from 1</param>
        /// <param name="col">Number of columns, starting from 1</param>
        /// <returns></returns>
        public static string RowColToValue(int row, int col)
        {
            if (row < 1 || col < 1) return null;
            string colValue = "";
            while (col > 0)
            {
                var n = col % 26;
                if (n == 0) n = 26;
                colValue = (char)('A' + n - 1) + colValue;
                col = (col - n) / 26;
            }
            return colValue + row;
        }
    }
    public class XlsxTextCell
    {
        /// <summary>
        /// Represents a single cell reference in a SpreadsheetML document
        /// </summary>
        public XlsxTextCellReference Reference { get; private set; }
        /// <summary>
        /// Number of rows, starting from 1
        /// </summary>
        public int Row => Reference.Row;
        /// <summary>
        /// Number of columns, starting from 1
        /// </summary>
        public int Col => Reference.Col;

        private string _value = "";
        /// <summary>
        /// Value of cell
        /// </summary>
        public string Value
        {
            get => _value;
            private set => _value = value ?? "";
        }

        internal XlsxTextCell(string reference, string value)
        {
            Reference = new XlsxTextCellReference(reference);
            Value = value;
        }
    }

    public class XlsxTextSheetReader
    {
        public XlsxTextReader Archive { get; private set; }
        public string Name { get; private set; }

        private Dictionary<string, KeyValuePair<string, string>> _mergeCells = new Dictionary<string, KeyValuePair<string, string>>();
        public List<XlsxTextCell> Row { get; private set; } = new List<XlsxTextCell>();

        private XlsxTextSheetReader(XlsxTextReader archive, string name, ZipArchiveEntry archiveEntry)
        {
            if (archive == null)
                throw new ArgumentNullException(nameof(archive));
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            if (archiveEntry == null)
                throw new ArgumentNullException(nameof(archiveEntry));

            Archive = archive;
            Name = name;

            Load(archiveEntry);
        }

        public static XlsxTextSheetReader Create(XlsxTextReader archive, string name, ZipArchiveEntry archiveEntry)
        {
            return new XlsxTextSheetReader(archive, name, archiveEntry);
        }

        private void Load(ZipArchiveEntry archiveEntry)
        {
            _mergeCells.Clear();
            using (XmlReader mergeCellsReader = XmlReader.Create(archiveEntry.Open(), new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true }))
            {
                mergeCellsReader.MoveToContent();
                if (mergeCellsReader.NodeType == XmlNodeType.Element && mergeCellsReader.Name == "worksheet" && mergeCellsReader.ReadToDescendant("mergeCells") && mergeCellsReader.ReadToDescendant("mergeCell"))
                {
                    do
                    {
                        string[] references = mergeCellsReader["ref"].Split(':');
                        _mergeCells.Add(references[0], new KeyValuePair<string, string>(references[1], null));
                    } while (mergeCellsReader.ReadToNextSibling("mergeCell"));
                }
            }

            _reader = XmlReader.Create(archiveEntry.Open(), new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true });
            if (_reader.ReadToDescendant("worksheet") && _reader.ReadToDescendant("sheetData"))
            {
                _reader = _reader.ReadSubtree();
            }
        }

        public string GetNumFmtValue(int cellStyle, string rawValue)
        {
            string formatCode = Archive.GetNumFmtCode(cellStyle);
            if (formatCode == null) return null;

            Trace.TraceWarning("Can not parse the value of numFmt: " + formatCode);
            return null;
            string value = "";
            return value;
        }

        private bool _isReading = false;
        private XmlReader _reader;
        public bool Read()
        {
            Row.Clear();
            if (_isReading ? _reader.ReadToNextSibling("row") : _reader.ReadToDescendant("row"))
            {
                _isReading = true;
                // read a row
                XmlReader rowReader = _reader.ReadSubtree();
                if (rowReader.ReadToDescendant("c"))
                {
                    do
                    {
                        string reference = rowReader["r"], style = rowReader["s"], type = rowReader["t"], value = null;
                        XmlReader cellReader = rowReader.ReadSubtree();

                        if (type == "inlineStr" && cellReader.ReadToDescendant("is") && cellReader.ReadToDescendant("t"))
                            value = cellReader.ReadElementContentAsString();
                        else if (cellReader.ReadToDescendant("v"))
                            value = cellReader.ReadElementContentAsString();
                        while (cellReader.Read()) ;

                        if (type == "d")
                        {
                            Trace.TraceWarning("Can not parse the cell " + reference + "'s value of date type");
                        }
                        else if (type == "e")
                        {
                            Trace.TraceWarning("Can not parse the cell " + reference + "'s value of error type");
                        }
                        else if (type == "s")
                        {
                            value = Archive.GetSharedString(int.Parse(value));
                        }
                        else if (type == "inlineStr")
                        {

                        }
                        else if (type == "b" && type == "n" && type == "str")
                        {

                        }
                        else
                        {
                            if (value != null)  // this cell's value is NumberFormat
                            {
                                if (style != null)
                                {
                                    value = GetNumFmtValue(int.Parse(style), value);
                                    if(value == null)
                                        Trace.TraceWarning("Can not parse the cell " + reference + "'s value of NumberFormat type. Please replace with string type.");
                                }
                            }
                            else
                            {
                                XlsxTextCellReference curr = new XlsxTextCellReference(reference);
                                foreach (var mergeCell in _mergeCells)
                                {
                                    XlsxTextCellReference begin = new XlsxTextCellReference(mergeCell.Key);
                                    if (curr.Row >= begin.Row && curr.Col >= begin.Col && !(curr.Row == begin.Row && curr.Col == begin.Col))
                                    {
                                        XlsxTextCellReference end = new XlsxTextCellReference(mergeCell.Value.Key);
                                        if (curr.Row <= end.Row && curr.Col <= end.Col)
                                            value = mergeCell.Value.Value;
                                    }
                                }
                            }
                        }

                        if (value == null) continue;

                        if (_mergeCells.TryGetValue(reference, out _))
                            _mergeCells[reference] = new KeyValuePair<string, string>(_mergeCells[reference].Key, value);

                        Row.Add(new XlsxTextCell(reference, value));

                    } while (rowReader.ReadToNextSibling("c"));
                }
                return true;
            }

            return false;
        }
    }
    public class XlsxTextReader
    {
        public const string RelationshipPart = "xl/_rels/workbook.xml.rels";
        public const string WorkbookPart = "xl/workbook.xml";
        public const string SharedStringsPart = "xl/sharedStrings.xml";
        public const string StylesPart = "xl/styles.xml";
        private readonly Dictionary<int, string> StandardNumFmts = new Dictionary<int, string>()
        {
            { 0, "General" },
            { 1, "0" },
            { 2, "0.00" },
            { 3, "#,##0" },
            { 4, "#,##0.00" },
            { 9, "0%" },
            { 10, "0.00%" },
            { 11, "0.00E+00" },
            { 12, "# ?/?" },
            { 13, "# ??/??" },
            { 14, "mm-dd-yy" },
            { 15, "d-mmm-yy" },
            { 16, "d-mmm" },
            { 17, "mmm-yy" },
            { 18, "h:mm AM/PM" },
            { 19, "h:mm:ss AM/PM" },
            { 20, "h:mm" },
            { 21, "h:mm:ss" },
            { 22, "m/d/yy h:mm" },
            { 37, "#,##0 ;(#,##0)" },
            { 38, "#,##0 ;[Red](#,##0)" },
            { 39, "#,##0.00;(#,##0.00)" },
            { 40, "#,##0.00;[Red](#,##0.00)" },
            { 45, "mm:ss" },
            { 46, "[h]:mm:ss" },
            { 47, "mmss.0" },
            { 48, "##0.0E+0" },
            { 49, "@" }
        };

        private ZipArchive _archive;
        private Dictionary<string, string> _rels = new Dictionary<string, string>();
        private List<KeyValuePair<string, string>> _sheets = new List<KeyValuePair<string, string>>();
        private List<string> _sharedStrings = new List<string>();
        private Dictionary<int, string> _numFmts;
        private List<int> _cellXfs = new List<int>();

        public int SheetsCount => _sheets.Count;
        public string GetSharedString(int index) => 0 <= index && index < _sharedStrings.Count ? _sharedStrings[index] : null;    
        public string GetNumFmtCode(int cellStyle)
        {
            if (0 <= cellStyle && cellStyle < _cellXfs.Count && _numFmts.TryGetValue(_cellXfs[cellStyle], out var formatCode))
                return formatCode;
            return null;
        }
        private XlsxTextReader(Stream stream)
        {
            _archive = new ZipArchive(stream, ZipArchiveMode.Read);
            Load();
        }
        private XlsxTextReader(string path) : this(new FileStream(path, FileMode.Open))
        {
        }

        public static XlsxTextReader Create(Stream stream)
        {
            return new XlsxTextReader(stream);
        }
        public static XlsxTextReader Create(string path)
        {
            return new XlsxTextReader(path);
        }

        private void Load()
        {
            _rels.Clear();
            using (Stream stream = _archive.GetEntry(RelationshipPart).Open())
            {
                using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true }))
                {
                    reader.MoveToContent();
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "Relationships" && reader.ReadToDescendant("Relationship"))
                    {
                        do
                        {
                            _rels.Add(reader["Id"], "xl/" + reader["Target"]);
                        } while (reader.ReadToNextSibling("Relationship"));
                    }
                }
            }

            _sheets.Clear();
            using (Stream stream = _archive.GetEntry(WorkbookPart).Open())
            {
                using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true }))
                {
                    reader.MoveToContent();
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "workbook" && reader.ReadToDescendant("sheets") && reader.ReadToDescendant("sheet"))
                    {
                        do
                        {
                            _sheets.Add(new KeyValuePair<string, string>(reader["name"], _rels[reader["r:id"]]));
                        } while (reader.ReadToNextSibling("sheet"));
                    }
                }
            }

            _sharedStrings.Clear();
            using (Stream stream = _archive.GetEntry(SharedStringsPart)?.Open())
            {
                if (stream != null)
                {
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true }))
                    {
                        reader.MoveToContent();
                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "sst" && reader.ReadToDescendant("si"))
                        {
                            do
                            {
                                XmlReader inner = reader.ReadSubtree();
                                if (inner.Read() && inner.Read())
                                {
                                    if (inner.NodeType == XmlNodeType.Element && inner.Name == "t")
                                    {
                                        _sharedStrings.Add(inner.ReadElementContentAsString());
                                        while (inner.Read()) ;
                                    }
                                    else if (inner.NodeType == XmlNodeType.Element && inner.Name == "r")
                                    {
                                        string value = "";
                                        do
                                        {
                                            XmlReader inner2 = inner.ReadSubtree();
                                            if (inner2.ReadToDescendant("t"))
                                            {
                                                do
                                                {
                                                    value += inner2.ReadElementContentAsString();
                                                } while (inner2.ReadToNextSibling("t"));
                                            }
                                        } while (inner.ReadToNextSibling("r"));
                                        _sharedStrings.Add(value);
                                    }
                                }
                            } while (reader.ReadToNextSibling("si"));
                        }
                    }
                }
            }

            _numFmts = new Dictionary<int, string>(StandardNumFmts);
            _cellXfs.Clear();
            using (Stream stream = _archive.GetEntry(StylesPart)?.Open())
            {
                if (stream != null)
                {
                    using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true }))
                    {
                        if (reader.ReadToDescendant("styleSheet") && reader.ReadToDescendant("numFmts") && reader.ReadToDescendant("numFmt"))
                        {
                            do
                            {
                                _numFmts[int.Parse(reader["numFmtId"])] = reader["formatCode"];
                            } while (reader.ReadToNextSibling("numFmt"));
                        }
                        if(reader.ReadToNextSibling("cellXfs") && reader.ReadToDescendant("xf"))
                        {
                            do
                            {
                                _cellXfs.Add(int.Parse(reader["numFmtId"]));
                            } while (reader.ReadToNextSibling("xf"));
                        }
                    }
                }
            }
        }

        public XlsxTextSheetReader SheetReader { get; private set; }
        private int _readIndex = 0;
        public bool Read()
        {
            SheetReader = null;
            if (_readIndex < SheetsCount)
            {
                // create a sheet reader
                SheetReader = XlsxTextSheetReader.Create(this, _sheets[_readIndex].Key, _archive.GetEntry(_sheets[_readIndex].Value));
                ++_readIndex;
                return true;
            }
            return false;
        }
    }
}
