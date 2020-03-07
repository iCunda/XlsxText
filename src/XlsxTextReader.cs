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
        /// 值
        /// </summary>
        public string Value { get; set; }

        public Cell(Reference reference, string value)
        {
            Reference = reference;
            Value = value;
        }
    }

    /// <summary>
    /// 合并单元格
    /// </summary>
    public class MergeCell : Cell
    {
        /// <summary>
        /// 始点单元格引用
        /// </summary>
        public Reference Begin { get; }
        /// <summary>
        /// 末点单元格引用
        /// </summary>
        public Reference End { get; }
        public MergeCell(Reference reference, string value, Reference end) : base(reference, value)
        {
            Begin = reference;
            End = end;
        }
        public MergeCell(Reference reference, string value, Reference begin, Reference end) : base(reference, value)
        {
            Begin = begin;
            End = end;
        }
    }

    /// <summary>
    /// 工作表
    /// </summary>
    public abstract class Worksheet
    {
        /// <summary>
        /// 工作簿
        /// </summary>
        public Workbook Workbook { get; }
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// 最大行号，负数为无限制
        /// </summary>
        public int MaxRow = -1;
        /// <summary>
        /// 最大列号，负数为无限制
        /// </summary>
        public short MaxColumn = -1;
        protected Worksheet(Workbook workbook, string name)
        {
            Workbook = workbook;
            Name = name;
        }

        public abstract IEnumerable<List<Cell>> Read();
    }

    /// <summary>
    /// 工作簿
    /// </summary>
    public abstract class Workbook : IDisposable
    {
        /// <summary>
        /// 工作表实现
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
            private MergeCell[] _mergeCells;

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
                _mergeCells = new MergeCell[0];
                using (XmlReader reader = XmlReader.Create(_part.Open()))
                {
                    int[] tree = { 0, 0 };
                    bool read = false;
                    int count = 0;
                    while (!read && reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                switch (reader.Depth)
                                {
                                    case 0:
                                        tree[0] = reader.Name == "worksheet" ? 1 : 0;
                                        break;
                                    case 1:
                                        tree[1] = reader.Name == "mergeCells" ? 1 : 0;
                                        if (tree[0] == 1 && tree[1] == 1)
                                            _mergeCells = new MergeCell[int.Parse(reader["count"])];
                                        break;
                                    case 2:
                                        if (tree[0] == 1 && tree[1] == 1 && reader.Name == "mergeCell")
                                        {
                                            string[] refs = reader["ref"].Split(':');
                                            _mergeCells[count++] = new MergeCell(new Reference(refs[0]), null, new Reference(refs[1]));
                                        }
                                        break;
                                }
                                break;
                            case XmlNodeType.EndElement:
                                if (tree[0] == 1 && tree[1] == 1 && reader.Depth == 1)
                                    read = true;
                                break;
                        }
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
                    int[] tree = { 0, 0, 0, 0, 0, 0 };
                    bool eof = false;
                    int row = 0;
                    List<Cell> rowCells = null;
                    string r = null, t = null, s = null, v = null;
                    while (!eof && reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                switch (reader.Depth)
                                {
                                    case 0:
                                        tree[0] = reader.Name == "worksheet" ? 1 : 0;
                                        break;
                                    case 1:
                                        tree[1] = reader.Name == "sheetData" ? 1 : 0;
                                        break;
                                    case 2:
                                        tree[2] = reader.Name == "row" ? 1 : 0;
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1)
                                        {
                                            row = int.Parse(reader["r"]);
                                            rowCells = new List<Cell>();
                                        }
                                        break;
                                    case 3:
                                        tree[3] = reader.Name == "c" ? 1 : 0;
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1 && tree[3] == 1)
                                        {
                                            r = reader["r"];
                                            t = reader["t"];
                                            s = reader["s"];
                                            v = null;
                                        }
                                        break;
                                    case 4:
                                        tree[4] = reader.Name == "v" ? 1 : reader.Name == "is" ? 2 : 0;
                                        break;
                                    case 5:
                                        tree[5] = reader.Name == "t" ? 1 : 0;
                                        break;
                                }
                                break;
                            case XmlNodeType.EndElement:
                                switch (reader.Depth)
                                {
                                    case 1:
                                        if (tree[0] == 1 && tree[1] == 1)
                                            eof = true;
                                        break;
                                    case 2:
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1)
                                        {
                                            if (MaxRow < 0 || row <= MaxRow)
                                            {
                                                List<MergeCell> mergeCells = new List<MergeCell>();
                                                foreach (MergeCell mergeCell in _mergeCells)
                                                {
                                                    if ((MaxColumn < 0 || mergeCell.Begin.Column <= MaxColumn) && mergeCell.Begin.Row <= row && row <= mergeCell.End.Row)
                                                        mergeCells.Add(mergeCell);
                                                }
                                                mergeCells.Sort((x, y) => x.Begin.Column - y.Begin.Column);
                                                rowCells.Sort((x, y) => x.Reference.Column - y.Reference.Column);

                                                if (mergeCells.Count > 0 || rowCells.Count > 0)
                                                {
                                                    List<Cell> newRowCells = new List<Cell>();
                                                    for (short column = 1, i1 = 0, i2 = 0; (MaxColumn < 0 || column <= MaxColumn) && (i1 < rowCells.Count || i2 < mergeCells.Count);)
                                                    {
                                                        if (i2 < mergeCells.Count && mergeCells[i2].Begin.Column == column)
                                                        {
                                                            if (i1 < rowCells.Count && rowCells[i1].Reference.Column == column) mergeCells[i2].Value = rowCells[i1].Value;
                                                            for (; (MaxColumn < 0 || column <= MaxColumn) && column <= mergeCells[i2].End.Column; ++column)
                                                            {
                                                                if (i1 < rowCells.Count && rowCells[i1].Reference.Column == column) ++i1;
                                                                newRowCells.Add(new Cell(new Reference(row, column), mergeCells[i2].Value));
                                                            }
                                                            ++i2;
                                                        }
                                                        else
                                                        {
                                                            if (i1 < rowCells.Count && rowCells[i1].Reference.Column == column)
                                                            {
                                                                newRowCells.Add(rowCells[i1]);
                                                                ++i1;
                                                            }
                                                            else
                                                                newRowCells.Add(new Cell(new Reference(row, column), ""));
                                                            ++column;
                                                        }
                                                    }
                                                    yield return newRowCells;
                                                }
                                            }

                                            if (MaxRow >= 0 && row >= MaxRow) eof = true;
                                            row = 0;
                                        }
                                        break;
                                    case 3:
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1 && tree[3] == 1)
                                        {
                                            Reference reference = new Reference(r);
                                            if (MaxColumn < 0 || reference.Column <= MaxColumn)
                                            {
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
                                                            int numFmtId = Workbook._cellXfs[int.Parse(s)];
                                                            if (Workbook._numFmts.TryGetValue(numFmtId, out string formatCode))
                                                            {
                                                                if (formatCode == BuiltinNumFmts[0] || formatCode == BuiltinNumFmts[49])
                                                                    value = v;
                                                                else
                                                                    throw new Exception(r + ": 不支持解析: formatCode=" + formatCode);
                                                            }
                                                            else
                                                                throw new Exception(r + ": 不支持解析: numFmtId=" + numFmtId);
                                                        }
                                                        else
                                                            value = v;
                                                        break;
                                                    default:
                                                        throw new Exception(r + ": 不支持类型: " + t);
                                                }
                                                rowCells.Add(new Cell(reference, value ?? ""));
                                            }
                                        }
                                        break;
                                }
                                break;
                            case XmlNodeType.SignificantWhitespace:
                            case XmlNodeType.Text:
                                switch (reader.Depth)
                                {
                                    case 5:
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1 && tree[3] == 1 && tree[4] == 1)
                                            v = reader.Value;
                                        break;
                                    case 6:
                                        if (tree[0] == 1 && tree[1] == 1 && tree[2] == 1 && tree[3] == 1 && tree[4] == 2 && tree[5] == 1)
                                            v = v == null ? reader.Value : v + reader.Value;
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 工作簿实现
        /// </summary>
        private class WorkbookImpl : Workbook
        {
            private readonly ZipArchive _archive;

            public WorkbookImpl(Stream stream) => _archive = new ZipArchive(stream, ZipArchiveMode.Read);
            public WorkbookImpl(string path) : this(new FileStream(path, FileMode.Open)) { }

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
                        int[] tree = { 0 };
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    switch (reader.Depth)
                                    {
                                        case 0:
                                            tree[0] = reader.Name == "Relationships" ? 1 : 0;
                                            break;
                                        case 1:
                                            if (tree[0] == 1 && reader.Name == "Relationship")
                                                _rels.Add(reader["Id"], "xl/" + reader["Target"]);
                                            break;
                                    }
                                    break;
                            }
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
                        int[] tree = { 0, 0 };
                        bool read = false;
                        while (!read && reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    switch (reader.Depth)
                                    {
                                        case 0:
                                            tree[0] = reader.Name == "workbook" ? 1 : 0;
                                            break;
                                        case 1:
                                            tree[1] = reader.Name == "sheets" ? 1 : 0;
                                            break;
                                        case 2:
                                            if (tree[0] == 1 && tree[1] == 1 && reader.Name == "sheet")
                                                _worksheets.Add(reader["name"], _rels[reader["r:id"]]);
                                            break;
                                    }
                                    break;
                                case XmlNodeType.EndElement:
                                    if (tree[0] == 1 && tree[1] == 1 && reader.Depth == 1)
                                        read = true;
                                    break;
                            }
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
                            while (reader.Read())
                            {
                                switch (reader.NodeType)
                                {
                                    case XmlNodeType.Element:
                                        switch (reader.Depth)
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
                                        break;
                                    case XmlNodeType.EndElement:
                                        if (tree[0] == 1 && tree[1] == 1 && reader.Depth == 1)
                                        {
                                            _sharedStrings.Add(value);
                                            value = "";
                                        }
                                        break;
                                    case XmlNodeType.SignificantWhitespace:
                                    case XmlNodeType.Text:
                                        switch (reader.Depth)
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
                            int[] tree = { 0, 0 };
                            bool read1 = false, read2 = false;
                            while ((!read1 || !read2) && reader.Read())
                            {
                                switch (reader.NodeType)
                                {
                                    case XmlNodeType.Element:
                                        switch (reader.Depth)
                                        {
                                            case 0:
                                                tree[0] = reader.Name == "styleSheet" ? 1 : 0;
                                                break;
                                            case 1:
                                                tree[1] = reader.Name == "numFmts" ? 1 : reader.Name == "cellXfs" ? 2 : 0;
                                                break;
                                            case 2:
                                                if (tree[0] == 1 && tree[1] == 1 && reader.Name == "numFmt")
                                                    _numFmts[int.Parse(reader["numFmtId"])] = reader["formatCode"];
                                                else if (tree[0] == 1 && tree[1] == 2 && reader.Name == "xf")
                                                    _cellXfs.Add(int.Parse(reader["numFmtId"]));
                                                break;
                                        }
                                        break;
                                    case XmlNodeType.EndElement:
                                        if (tree[0] == 1 && tree[1] == 1 && reader.Depth == 1)
                                            read1 = true;
                                        else if (tree[0] == 1 && tree[1] == 2 && reader.Depth == 1)
                                            read2 = true;
                                        break;
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

            public override void Dispose() => _archive.Dispose();
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

        protected Dictionary<string, string> _rels;
        protected Dictionary<string, string> _worksheets;
        protected List<string> _sharedStrings;
        protected List<int> _cellXfs;
        protected Dictionary<int, string> _numFmts;

        public abstract List<Worksheet> Read();
        public abstract void Dispose();

        public static Workbook Open(Stream stream) => new WorkbookImpl(stream);
        public static Workbook Open(string path) => new WorkbookImpl(path);
    }
}
