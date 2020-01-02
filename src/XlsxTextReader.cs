using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace XlsxTextReader
{
    /// <summary>
    /// 单元格引用
    /// </summary>
    public class Reference
    {
        /// <summary>
        /// 行号, 从1开始
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 列号, 从1开始
        /// </summary>
        public short Column { get; }

        public string Value
        {
            get
            {
                GetValue(Row, Column, out string value);
                return value;
            }
        }

        public Reference(string value)
        {
            if (GetRowCol(value, out int row, out short column))
            {
                Row = row;
                Column = column;
            }
            else
                throw new Exception("无效引用值: " + value);
        }

        public Reference(int row, short column)
        {
            if (row < 0 && column < 0)
                throw new Exception("无效引用范围：" + row + ',' + column);
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 引用号获取行列值
        /// </summary>
        /// <param name="value">引用值</param>
        /// <param name="row">行号, 从1开始</param>
        /// <param name="col">列号, 从1开始</param>
        /// <returns></returns>
        public static bool GetRowCol(string value, out int row, out short column)
        {
            row = 0;
            column = 0;
            for (int i = 0; i < value.Length; ++i)
            {
                if ('A' <= value[i] && value[i] <= 'Z')
                    column = (short)(column * 26 + (value[i] - 'A') + 1);
                else
                    return int.TryParse(value.Substring(i), out row);
            }

            return false;
        }

        /// <summary>
        /// 行列号获取引用值
        /// </summary>
        /// <param name="row">行号, 从1开始</param>
        /// <param name="column">列号, 从1开始</param>
        /// <param name="value">引用值</param>
        /// <returns></returns>
        public static bool GetValue(int row, int column, out string value)
        {
            if (row < 1 || column < 1)
            {
                value = null;
                return false;
            }

            value = "";
            while (column > 0)
            {
                int c = (column - 1) % 26 + 'A';
                value = (char)c + value;
                column = (column - (c - 'A' + 1)) / 26;
            }
            value += row;

            return true;
        }
    }

    /// <summary>
    /// 单元格
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// 单元格引用
        /// </summary>
        public Reference Reference { get; }

        /// <summary>
        /// 合并单元格末端引用
        /// </summary>
        public Reference EndReference { get; }

        /// <summary>
        /// 是否是合并单元格
        /// </summary>
        public bool isMergeCell { get => EndReference.Row >= Reference.Row && EndReference.Column >= Reference.Column && (EndReference.Row > Reference.Row || EndReference.Column > Reference.Column); }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }

        public Cell(Reference reference, string value, Reference endReference = null)
        {
            Reference = reference;
            EndReference = endReference ?? reference;
            Value = value;
        }
    }

    /// <summary>
    /// 工作表
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// 工作簿
        /// </summary>
        public Workbook Workbook { get; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name { get; }

        protected Worksheet(Workbook workbook, string name)
        {
            Workbook = workbook;
            Name = name;
        }
        public virtual IEnumerable<List<Cell>> Read() { yield break; }
    }

    public class Workbook : IDisposable
    {
        private class WorkbookImpl : Workbook
        {
            /// <summary>
            /// 工作表
            /// </summary>
            private class WorksheetImpl : Worksheet
            {
                /// <summary>
                /// 工作表part
                /// </summary>
                private ZipArchiveEntry _part;
                /// <summary>
                /// 合并单元格
                /// </summary>
                private List<Cell> _mergeCells;

                public WorksheetImpl(WorkbookImpl workbook, string name, ZipArchiveEntry part) : base(workbook, name) => _part = part;

                private void Load()
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
                     *     <mergeCells count="5">
                     *         <mergeCell ref="A1:B2"/>
                     *         <mergeCell ref="C1:E5"/>
                     *         <mergeCell ref="A3:B6"/>
                     *         <mergeCell ref="A7:C7"/>
                     *         <mergeCell ref="A8:XFD9"/>
                     *     </mergeCells>
                     * <worksheet>
                     */
                    _mergeCells = new List<Cell>();
                    using (XmlReader reader = XmlReader.Create(_part.Open()))
                    {
                        if (reader.ReadToNextSibling("worksheet") && reader.ReadToDescendant("mergeCells") && reader.ReadToDescendant("mergeCell"))
                        {
                            do
                            {
                                string[] refs = reader["ref"].Split(':');
                                _mergeCells.Add(new Cell(new Reference(refs[0]), null, new Reference(refs[1])));
                            } while (reader.ReadToNextSibling("mergeCell"));
                        }
                    }
                }

                public override IEnumerable<List<Cell>> Read()
                {
                    if (_mergeCells == null)
                        Load();

                    /**
                     * <worksheet>
                     *     <sheetData>
                     *         <row r="1">
                     *             <c r="A1" s="11">
                     *                 <v>2</v>
                     *             </c>
                     *             <c r="E1" t="inlineStr">
                     *                 <is>
                     *                     <t>This is inline string example</t>
                     *                 </is>
                     *             </c>
                     *             <c r="G1">
                     *                 <f>SUM(A1:A3)</f>
                     *                 <v>9</v>
                     *             </c>
                     *             <c r="H1" s="11"/>
                     *         </row>
                     *     </sheetData>
                     * <worksheet>
                     */
                    using (XmlReader reader = XmlReader.Create(_part.Open()))
                    {
                        if (reader.ReadToNextSibling("worksheet") && reader.ReadToDescendant("sheetData") && reader.ReadToDescendant("row"))
                        {
                            do
                            {
                                if (reader.ReadToDescendant("c"))
                                {
                                    List<Cell> row = new List<Cell>();
                                    do
                                    {
                                        string r = reader["r"], t = reader["t"], s = reader["s"], v = null;

                                        if (!reader.IsEmptyElement)
                                        {
                                            int[] tree = { 0, 0 };
                                            int depth = 0;
                                            while (reader.Read())
                                            {
                                                switch (reader.NodeType)
                                                {
                                                    case XmlNodeType.Element:
                                                        switch (depth)
                                                        {
                                                            case 0:
                                                                tree[0] = reader.Name == "v" ? 1 : reader.Name == "is" ? 2 : 0;
                                                                break;
                                                            case 1:
                                                                tree[1] = reader.Name == "t" ? 1 : 0;
                                                                break;
                                                        }
                                                        if (!reader.IsEmptyElement) ++depth;
                                                        break;
                                                    case XmlNodeType.EndElement:
                                                        --depth;
                                                        break;
                                                    case XmlNodeType.SignificantWhitespace:
                                                    case XmlNodeType.Text:
                                                        switch (depth)
                                                        {
                                                            case 1:
                                                                if (tree[0] == 1)
                                                                    v = reader.Value;
                                                                break;
                                                            case 2:
                                                                if (tree[0] == 2 && tree[1] == 1)
                                                                    v = v == null ? reader.Value : v + reader.Value;
                                                                break;
                                                        }
                                                        break;
                                                }
                                                if (depth < 0)
                                                    break;
                                            }
                                        }

                                        string value;
                                        switch (t)
                                        {
                                            case "n":
                                            case "str":
                                            case "inlineStr":
                                                value = v;
                                                break;
                                            case "b":
                                                value = v == "0" ? "FALSE" : "TRUE";
                                                break;
                                            case "s":
                                                value = Workbook._sharedStrings[int.Parse(v)];
                                                break;
                                            case "e":
                                                throw new Exception(r + ": 单元格有错误");
                                            case "d":
                                                throw new Exception(r + ": 不支持解析时间类型的值");
                                            case null:
                                                if (s != null && v != null)
                                                {
                                                    string formatCode = Workbook._numFmts[Workbook._cellXfs[int.Parse(s)]];
                                                    if (formatCode == BuiltinNumFmts[0] || formatCode == BuiltinNumFmts[49])
                                                        value = v;
                                                    else
                                                        throw new Exception(r + ": 不支持解析: " + formatCode);
                                                }
                                                else
                                                    value = v;
                                                break;
                                            default:
                                                throw new Exception(r + ": 不支持类型: " + t);
                                        }

                                        Cell cell = new Cell(new Reference(r), value);
                                        foreach (Cell mergeCell in _mergeCells)
                                        {
                                            if (mergeCell.Reference.Row == cell.Reference.Row && mergeCell.Reference.Column == cell.Reference.Column)
                                                mergeCell.Value = cell.Value;
                                            else if (mergeCell.Reference.Row <= cell.Reference.Row && cell.Reference.Row <= mergeCell.EndReference.Row
                                                && mergeCell.Reference.Column <= cell.Reference.Column && cell.Reference.Column <= mergeCell.EndReference.Column)
                                                cell = new Cell(cell.Reference, mergeCell.Value, new Reference(mergeCell.EndReference.Row, mergeCell.EndReference.Column));
                                        }
                                        row.Add(cell);

                                    } while (reader.ReadToNextSibling("c"));
                                    if (row.Count > 0)
                                        yield return row;
                                }
                            } while (reader.ReadToNextSibling("row"));
                        }
                    }
                }
            }

            public WorkbookImpl(Stream stream) : base(stream) { }
            public WorkbookImpl(string path) : base(path) { }

            private void Load()
            {
                _rels = new Dictionary<string, string>();
                using (Stream stream = _archive.GetEntry(RelationshipPart).Open())
                {
                    /**
                     * xl/_rels/workbook.xml.rels
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
                        if (reader.ReadToNextSibling("Relationships") && reader.ReadToDescendant("Relationship"))
                        {
                            do
                                _rels.Add(reader["Id"], "xl/" + reader["Target"]);
                            while (reader.ReadToNextSibling("Relationship"));
                        }
                    }
                }

                _worksheets = new Dictionary<string, string>();
                using (Stream stream = _archive.GetEntry(WorkbookPart).Open())
                {
                    /**
                     * xl/workbook.xml
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
                        if (reader.ReadToNextSibling("workbook") && reader.ReadToDescendant("sheets") && reader.ReadToDescendant("sheet"))
                        {
                            do
                                _worksheets.Add(reader["name"], _rels[reader["r:id"]]);
                            while (reader.ReadToNextSibling("sheet"));
                        }
                    }
                }

                _sharedStrings = new List<string>();
                using (Stream stream = _archive.GetEntry(SharedStringsPart)?.Open())
                {
                    if (stream != null)
                    {
                        /**
                         * xl/sharedStrings.xml
                         * <sst>
                         *     <si>
                         *         <t>共享字符串1</t>
                         *     </si>
                         *     <si>
                         *         <r>
                         *             <t>共享富文本字符串1</t>
                         *         </r>
                         *         <r>
                         *             <t>共享富文本字符串2</t>
                         *         </r>
                         *     </si>
                         * </sst>
                         */
                        using (XmlReader reader = XmlReader.Create(stream))
                        {
                            string value = "";
                            int[] tree = { 0, 0, 0, 0 };
                            int depth = 0;
                            while (reader.Read())
                            {
                                switch (reader.NodeType)
                                {
                                    case XmlNodeType.Element:
                                        switch (depth)
                                        {
                                            case 0:
                                                tree[0] = reader.Name == "sst" ? 1 : 0;
                                                break;
                                            case 1:
                                                tree[1] = reader.Name == "si" ? 1 : 0;
                                                break;
                                            case 2:
                                                tree[2] = reader.Name == "t" ? 1 : reader.Name == "r" ? 2 : 0;
                                                break;
                                            case 3:
                                                tree[3] = reader.Name == "t" ? 1 : 0;
                                                break;
                                        }
                                        if (!reader.IsEmptyElement) ++depth;
                                        break;
                                    case XmlNodeType.EndElement:
                                        if (depth == 2 && tree[0] == 1 && tree[1] == 1)
                                        {
                                            _sharedStrings.Add(value);
                                            value = "";
                                        }
                                        --depth;
                                        break;
                                    case XmlNodeType.SignificantWhitespace:
                                    case XmlNodeType.Text:
                                        switch (depth)
                                        {
                                            case 3:
                                                if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1)
                                                    value = reader.Value;
                                                break;
                                            case 4:
                                                if (tree[0] == 1 && tree[1] == 1 && tree[2] == 2 && tree[3] == 1)
                                                    value += reader.Value;
                                                break;
                                        }
                                        break;
                                }
                            }
                        }
                    }
                }

                _numFmts = new Dictionary<int, string>(BuiltinNumFmts);
                _cellXfs = new List<int>();
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
                            if (reader.ReadToNextSibling("styleSheet"))
                            {
                                bool nfmtRead = false, xfsRead = false;
                                while (reader.Read() && (!nfmtRead || !xfsRead))
                                {
                                    if (reader.NodeType == XmlNodeType.Element)
                                    {
                                        if (reader.Name == "numFmts")
                                        {
                                            if (reader.ReadToDescendant("numFmt"))
                                            {
                                                do
                                                    _numFmts[int.Parse(reader["numFmtId"])] = reader["formatCode"];
                                                while (reader.ReadToNextSibling("numFmt"));
                                            }
                                            nfmtRead = true;
                                        }
                                        else if (reader.Name == "cellXfs")
                                        {
                                            if (reader.ReadToDescendant("xf"))
                                            {
                                                do
                                                    _cellXfs.Add(int.Parse(reader["numFmtId"]));
                                                while (reader.ReadToNextSibling("xf"));
                                            }
                                            xfsRead = true;
                                        }
                                        else
                                            reader.Skip();
                                    }
                                }
                            }
                        }
                    }
                }
            }

            public override List<Worksheet> Read()
            {
                if (_worksheets == null)
                    Load();
                List<Worksheet> worksheets = new List<Worksheet>();
                foreach (KeyValuePair<string, string> keyValue in _worksheets)
                    worksheets.Add(new WorksheetImpl(this, keyValue.Key, _archive.GetEntry(keyValue.Value)));
                return worksheets;
            }

            public new void Dispose() => _archive.Dispose();
        }

        public const string RelationshipPart = "xl/_rels/workbook.xml.rels";
        public const string WorkbookPart = "xl/workbook.xml";
        public const string SharedStringsPart = "xl/sharedStrings.xml";
        public const string StylesPart = "xl/styles.xml";
        public static readonly ReadOnlyDictionary<int, string> BuiltinNumFmts = new ReadOnlyDictionary<int, string>(new Dictionary<int, string>()
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
        });

        protected readonly ZipArchive _archive;
        protected Dictionary<string, string> _rels;
        protected Dictionary<string, string> _worksheets;
        protected List<string> _sharedStrings;
        protected List<int> _cellXfs;
        protected Dictionary<int, string> _numFmts;

        protected Workbook(Stream stream) => _archive = new ZipArchive(stream, ZipArchiveMode.Read);
        protected Workbook(string path) : this(new FileStream(path, FileMode.Open)) { }

        public virtual List<Worksheet> Read() => new List<Worksheet>();
        public void Dispose() { }

        public static Workbook Open(Stream stream) => new WorkbookImpl(stream);
        public static Workbook Open(string path) => new WorkbookImpl(path);
    }
}
