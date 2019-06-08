using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
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
                string expr = value?.Trim().ToUpper() ?? "", rowValue = "", colValue = "";
                for (int i = 0; i < expr.Length; ++i)
                {
                    if (!('A' <= expr[i] && expr[i] <= 'Z'))
                    {
                        rowValue = expr.Substring(i);
                        colValue = expr.Substring(0, i);
                        break;
                    }
                }
                if (rowValue == "" || colValue == "")
                    throw new Exception("Invalid value of cell reference");

                long row;
                if (!long.TryParse(rowValue, out row) || row < 1)
                    throw new Exception("Invalid value of cell reference");

                long col = 0;
                for (int i = colValue.Length - 1, multiple = 1; i >= 0; --i, multiple *= 26)
                {
                    int n = colValue[i] - 'A' + 1;
                    col += (n * multiple);
                }

                _value = expr;
                Row = row;
                Col = col;
            }
        }
        /// <summary>
        /// Number of rows, starting from 1
        /// </summary>
        public long Row { get; private set; }
        /// <summary>
        /// Number of columns, starting from 1
        /// </summary>
        public long Col { get; private set; }

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
        public long Row => Reference.Row;
        /// <summary>
        /// Number of columns, starting from 1
        /// </summary>
        public long Col => Reference.Col;

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
        public long CellCount { get; private set; } = 0;
        public long RowCount { get; private set; } = 0;

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
            /**
             * <worksheet>
             *     <sheetData>
             *         <row r="1">
             *              <c r="A1" s="11"><v>2</v></c>
             *              <c r="B1" s="11"><v>3</v></c>
             *              <c r="C1" s="11"><v>4</v></c>
             *              <c r="D1" t="s"><v>0</v></c>
             *              <c r="E1" t="inlineStr"><is><t>This is inline string example</t></is></c>
             *              <c r="D1" t="d"><v>1976-11-22T08:30</v></c>
             *              <c r="G1"><f>SUM(A1:A3)</f><v>9</v></c>
             *              <c r="H1" s="11"/>
             *          </row>
             *     </sheetData>
             * <worksheet>
             */
            using (XmlReader reader = XmlReader.Create(archiveEntry.Open()))
            {
                string[] names = new string[4];
                long count = 0;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (count < 4) names[count] = reader.Name;
                        ++count;

                        if (count == 3 && names[0] == "worksheet" && names[1] == "sheetData" && names[2] == "row")
                        {
                            ++RowCount;
                        }
                        else if (count == 4 && names[0] == "worksheet" && names[1] == "sheetData" && names[2] == "row" && names[3] == "c")
                        {
                            ++CellCount;
                        }

                        if (reader.IsEmptyElement) --count;
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        --count;
                    }
                }
            }

            /**
             * <worksheet>
             *     <mergeCells count="5">
             *         <mergeCell ref="A1:B2"/>
             *         <mergeCell ref="C1:E5"/>
             *         <mergeCell ref="A3:B6"/>
             *         <mergeCell ref="A7:C7"/>
             *         <mergeCell ref="A8:XFD9"/>
             *     </mergeCells>
             * <worksheet>
             */
            using (XmlReader reader = XmlReader.Create(archiveEntry.Open()))
            {
                string[] names = new string[3];
                long count = 0;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (count < 3) names[count] = reader.Name;
                        ++count;

                        if (count == 3 && names[0] == "worksheet" && names[1] == "mergeCells" && names[2] == "mergeCell")
                        {
                            string[] references = reader["ref"]?.Split(':');
                            if (references?.Length == 2) _mergeCells.Add(references[0], new KeyValuePair<string, string>(references[1], null));
                        }

                        if (reader.IsEmptyElement) --count;
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        --count;
                    }
                }
            }

            /**
             * <worksheet>
             *     <sheetData>
             *         <row r="1">
             *              <c r="A1" s="11"><v>2</v></c>
             *              <c r="B1" s="11"><v>3</v></c>
             *              <c r="C1" s="11"><v>4</v></c>
             *              <c r="D1" t="s"><v>0</v></c>
             *              <c r="E1" t="inlineStr"><is><t>This is inline string example</t></is></c>
             *              <c r="D1" t="d"><v>1976-11-22T08:30</v></c>
             *              <c r="G1"><f>SUM(A1:A3)</f><v>9</v></c>
             *              <c r="H1" s="11"/>
             *          </row>
             *     </sheetData>
             * <worksheet>
             */
            _reader = XmlReader.Create(archiveEntry.Open());
            if (!(_reader.ReadToDescendant("worksheet") && _reader.ReadToDescendant("sheetData")))
            {
                _reader.Close();
            }
        }
        public string GetNumFmtValue(int cellStyle, string rawValue)
        {
            string formatCode = Archive.GetNumFmtCode(cellStyle);
            if (formatCode == XlsxTextReader.StandardNumFmts[0])
                return rawValue;
            else
                return null;
        }

        private XmlReader _reader;
        public bool Read()
        {
            /**
             * <row r="1">
             *     <c r="A1" s="11"><v>2</v></c>
             *     <c r="B1" s="11"><v>3</v></c>
             *     <c r="C1" s="11"><v>4</v></c>
             *     <c r="D1" t="s"><v>0</v></c>
             *     <c r="E1" t="inlineStr"><is><t>This is inline string example</t></is></c>
             *     <c r="D1" t="d"><v>1976-11-22T08:30</v></c>
             *     <c r="G1"><f>SUM(A1:A3)</f><v>9</v></c>
             *     <c r="H1" s="11"/>
             * </row>
             */
            Row.Clear();
            string[] names = new string[4];
            long count = 0;
            long rowIndex;
            string reference = "", value = "", type = "", style = "";
            while (_reader.Read())
            {
                if (_reader.NodeType == XmlNodeType.Element)
                {
                    if (count < 4) names[count] = _reader.Name;
                    ++count;

                    if (count == 1 && names[0] == "row")
                        long.TryParse(_reader["r"], out rowIndex);
                    if (count == 2 && names[0] == "row" && names[1] == "c")
                    {
                        reference = _reader["r"] ?? "";
                        type = _reader["t"] ?? "";
                        style = _reader["s"] ?? "";

                        if (_reader.IsEmptyElement)
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
                            Row.Add(new XlsxTextCell(reference, value));
                            reference = value = type = style = "";
                        }
                    }
                    if (_reader.IsEmptyElement) --count;
                }
                else if (_reader.NodeType == XmlNodeType.Text)
                {
                    if (names[0] == "row" && names[1] == "c" && ((count == 3 && names[2] == "v") || (count == 4 && names[3] == "is" && names[4] == "t")))
                    {
                        value = _reader.Value;
                        if (type == "d")
                        {
                            Debug.Print("Can not parse the cell " + reference + "'s value of date type");
                        }
                        else if (type == "e")
                        {
                            Debug.Print("Can not parse the cell " + reference + "'s value of error type");
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
                            string rawValue = value;
                            if (int.TryParse(style, out int styleId))
                            {
                                value = GetNumFmtValue(styleId, rawValue);
                                if (value == null)
                                    Debug.Print("Can not parse the cell " + reference + "'s value of NumberFormat type. Please replace with string type.");
                            }
                        }

                        if (value != null)
                        {
                            if (_mergeCells.TryGetValue(reference, out _))
                                _mergeCells[reference] = new KeyValuePair<string, string>(_mergeCells[reference].Key, value);

                            Row.Add(new XlsxTextCell(reference, value));
                            reference = value = type = style = "";
                        }
                    }
                }
                else if (_reader.NodeType == XmlNodeType.EndElement)
                {
                    --count;
                    if (count == 0 && _reader.Name == "row")
                        return true;
                    else if (count < 0)
                        return false;
                }
            }

            return false;
        }
    }
    public class XlsxTextReader : IDisposable
    {
        public const string RelationshipPart = "xl/_rels/workbook.xml.rels";
        public const string WorkbookPart = "xl/workbook.xml";
        public const string SharedStringsPart = "xl/sharedStrings.xml";
        public const string StylesPart = "xl/styles.xml";
        public static readonly Dictionary<int, string> StandardNumFmts = new Dictionary<int, string>()
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
                if (stream != null)
                {
                    /**
                     * xl/styles.xml
                     * <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                     *     <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                     *     <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                     *     <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
                     *     <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                     *     <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                     * </Relationships>
                     */
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string[] names = new string[2];
                        long count = 0;
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (count < 2) names[count] = reader.Name;
                                ++count;

                                if (count == 2 && names[0] == "Relationships" && names[1] == "Relationship")
                                {
                                    string id = reader["Id"] ?? "";
                                    string code = reader["Target"] ?? "";
                                    if (id != "" && code != "")
                                        _rels.Add(id, "xl/" + code);
                                }

                                if (reader.IsEmptyElement) --count;
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                --count;
                            }
                        }
                    }
                }
            }

            _sheets.Clear();
            using (Stream stream = _archive.GetEntry(WorkbookPart).Open())
            {
                if (stream != null)
                {
                    /**
                     * <workbook>
                     *     <sheets>
                     *         <sheet name="Example1" sheetId="1" r:id="rId1"/>
                     *         <sheet name="Example2" sheetId="6" r:id="rId2"/>
                     *         <sheet name="Example3" sheetId="7" r:id="rId3"/>
                     *         <sheet name="Example4" sheetId="8" r:id="rId4"/>
                     *     </sheets>
                     * <workbook>
                     */
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string[] names = new string[3];
                        long count = 0;
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (count < 3) names[count] = reader.Name;
                                ++count;

                                if (count == 3 && names[0] == "workbook" && names[1] == "sheets" && names[2] == "sheet")
                                {
                                    string name = reader["name"] ?? "";
                                    if (name != "" && _rels.TryGetValue(reader["r:id"] ?? "", out string url))
                                        _sheets.Add(new KeyValuePair<string, string>(name, url));
                                }

                                if (reader.IsEmptyElement) --count;
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                --count;
                            }
                        }
                    }
                }
            }

            _sharedStrings.Clear();
            using (Stream stream = _archive.GetEntry(SharedStringsPart)?.Open())
            {
                if (stream != null)
                {
                    /**
                     * xl/sharedStrings.xml
                     * <sst>
                     *     <si><t>共享字符串1</t></si>
                     *     <si><r><t>共享富文本字符串1</t></r><r><t>共享富文本字符串2</t></r></si>
                     * </sst>
                     */
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string[] names = new string[4];
                        long count = 0;
                        string expr = "";
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && !reader.IsEmptyElement)
                            {
                                if (count < 4) names[count] = reader.Name;
                                ++count;
                            }
                            else if (reader.NodeType == XmlNodeType.Text)
                            {
                                if ((count == 3 || count == 4) && names[0] == "sst" && names[1] == "si")
                                {
                                    if (count == 3 && names[2] == "t")
                                    {
                                        expr = reader.Value;
                                    }
                                    else if (count == 4 && names[2] == "r" && names[3] == "t")
                                    {
                                        expr += reader.Value;
                                    }
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                --count;
                                if (count == 1 && reader.Name == "si")
                                {
                                    _sharedStrings.Add(expr);
                                    expr = "";
                                }
                            }
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
                    /**
                     * xl/styles.xml
                     * <styleSheet>
                     *     <numFmts count="2">
                     *         <numFmt numFmtId="8" formatCode="&quot;¥&quot;#,##0.00;[Red]&quot;¥&quot;\-#,##0.00"/>
                     *         <numFmt numFmtId="176" formatCode="&quot;$&quot;#,##0.00_);\(&quot;$&quot;#,##0.00\)"/>
                     *     </numFmts>
                     *     <cellXfs count="3">
                     *         <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                     *         <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
                     *         <xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" applyNumberFormat="1"/>
                     *     </cellXfs>
                     * </styleSheet>
                     */
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string[] names = new string[3];
                        long count = 0;
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (count < 3) names[count] = reader.Name;
                                ++count;

                                if (count == 3 && names[0] == "styleSheet")
                                {
                                    if (names[1] == "numFmts" && names[2] == "numFmt" && int.TryParse(reader["numFmtId"], out int nid))
                                    {
                                        string code = reader["formatCode"] ?? "";
                                        if (code != "") _numFmts[nid] = code;
                                    }
                                    else if (names[1] == "cellXfs" && names[2] == "xf" && int.TryParse(reader["numFmtId"], out int xid))
                                    {
                                        _cellXfs.Add(xid);
                                    }
                                }

                                if (reader.IsEmptyElement) --count;
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                --count;
                            }
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
        public void Dispose()
        {
            _archive.Dispose();
        }
    }

}
