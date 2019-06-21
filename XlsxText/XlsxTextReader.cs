using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace XlsxText
{
    /// <summary>
    /// 单元格引用
    /// </summary>
    public class Refer
    {
        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; }
        /// <summary>
        /// 行, 从1开始
        /// </summary>
        public int Row { get; }
        /// <summary>
        /// 列, 从1开始
        /// </summary>
        public int Col { get; }

        public Refer(string value)
        {
            if (GetRowCol(value, out int row, out int col))
            {
                Value = value;
                Row = row;
                Col = col;
            }
            else
                throw new Exception("无效引用: " + value);
        }

        public Refer(int row, int col)
        {
            if (GetValue(row, col, out string value))
            {
                Value = value;
                Row = row;
                Col = col;
            }
            else
                throw new Exception("无效引用: " + value);
        }

        /// <summary>
        /// 行列获取引用
        /// </summary>
        /// <param name="row">行, 从1开始</param>
        /// <param name="col">列, 从1开始</param>
        /// <param name="value">引用</param>
        /// <returns></returns>
        public static bool GetValue(int row, int col, out string value)
        {
            if (row < 1 || col < 1)
            {
                value = null;
                return false;
            }

            value = "";
            while (col > 0)
            {
                int c = (col - 1) % ('Z' - 'A' + 1) + 'A';
                value = (char)c + value;
                col = (col - (c - 'A' + 1)) / 26;
            }
            value += row;

            return true;
        }

        /// <summary>
        /// 引用获取行列
        /// </summary>
        /// <param name="value">引用</param>
        /// <param name="row">行, 从1开始</param>
        /// <param name="col">列, 从1开始</param>
        /// <returns></returns>
        public static bool GetRowCol(string value, out int row, out int col)
        {
            row = 0;
            col = 0;
            for (int i = 0; i < value.Length; ++i)
            {
                if ('A' <= value[i] && value[i] <= 'Z')
                    col = col * ('Z' - 'A' + 1) + value[i] - 'A' + 1;
                else
                    return int.TryParse(value.Substring(i), out row);
            }

            return false;
        }
    }
    /// <summary>
    /// 单元格
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// 引用
        /// </summary>
        public Refer Refer { get; set; }
        /// <summary>
        /// 结束引用，用于合并单元格
        /// </summary>
        public Refer EndRefer { get; set; }
        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }

        public Cell(string refer, string value)
        {
            EndRefer = Refer = new Refer(refer);
            Value = value;
        }
        public Cell(Refer refer, string value)
        {
            EndRefer = Refer = refer;
            Value = value;
        }
    }

    /// <summary>
    /// 工作表
    /// </summary>
    public class Worksheet : IDisposable
    {
        /// <summary>
        /// 工作簿
        /// </summary>
        public Workbook Workbook { get; private set; }
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; private set; }

        public Worksheet(Workbook workbook, string name, ZipArchiveEntry archiveEntry)
        {
            Workbook = workbook;
            Name = name;

            Load(archiveEntry);
        }

        private void Load(ZipArchiveEntry archiveEntry)
        {
            RowCount = 0;
            _mergeCells = new Dictionary<string, Cell>();
            using (Stream stream = archiveEntry?.Open())
            {
                if (stream != null)
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
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        int depth = 0;
                        int[] flags = new int[3];
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                {
                                    if (depth == 0)
                                        flags[0] = reader.Name == "worksheet" ? 1 : 0;
                                    else if (depth == 1)
                                        flags[1] = reader.Name == "sheetData" ? 2 : reader.Name == "mergeCells" ? 21 : 0;
                                    else if (depth == 2)
                                        flags[2] = reader.Name == "row" ? 3 : reader.Name == "mergeCell" ? 31 : 0;

                                    if (depth == 2 && flags[0] == 1 && flags[1] == 2 && flags[2] == 3)
                                    {
                                        ++RowCount;
                                    }
                                    else if (depth == 2 && flags[0] == 1 && flags[1] == 21 && flags[2] == 31)
                                    {
                                        string[] refs = reader["ref"].Split(':');
                                        Cell mergeCell = new Cell(refs[0], null);
                                        mergeCell.EndRefer = new Refer(refs[1]);
                                        _mergeCells.Add(mergeCell.Refer.Value, mergeCell);
                                    }

                                    if (!reader.IsEmptyElement) ++depth;
                                    break;
                                }
                                case XmlNodeType.EndElement:
                                {
                                    --depth;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (RowCount == 0) return;
            Stream stream2 = archiveEntry?.Open();
            if (stream2 != null)
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
                _reader = XmlReader.Create(stream2);
                int depth = 0;
                int[] flags = new int[2];
                while (_reader.Read())
                {
                    if (_reader.NodeType == XmlNodeType.Element)
                    {
                        if (depth == 0)
                            flags[0] = _reader.Name == "worksheet" ? 1 : 0;
                        else if (depth == 1)
                            flags[1] = _reader.Name == "sheetData" ? 2 : 0;

                        if (depth == 1 && flags[0] == 1 && flags[1] == 2)
                            break;

                        if (!_reader.IsEmptyElement) ++depth;
                    }
                    if (_reader.NodeType == XmlNodeType.EndElement)
                    {
                        --depth;
                    }
                }
            }
        }

        /// <summary>
        /// 解析单元格的值
        /// </summary>
        /// <param name="reference">引用</param>
        /// <param name="rawValue">原值</param>
        /// <param name="type">类型</param>
        /// <param name="style">样式</param>
        /// <returns></returns>
        private string ParseCellValue(string reference, string rawValue, string type, string style)
        {
            string value = null;
            if (rawValue == null)
            {
                value = null;
            }
            else if (type == "s")
            {
                if (int.TryParse(rawValue, out int shareStringIndex))
                    value = Workbook.GetSharedString(shareStringIndex);
            }
            else if (type == "d")
            {
                Debug.Print("不支持解析时间类型的值：" + reference);
                value = rawValue;
            }
            else if (type == "e")
            {
                value = "";
            }
            else if (type == "b" || type == "n" || type == "str" || type == "inlineStr")
            {
                value = rawValue;
            }
            else
            {
                if (type != null)
                {
                    Debug.Print("{0}: 不支持解析类型为\"{1}\"的值", reference, type);
                    value = rawValue;
                }
                else if (style != null)
                {
                    if (int.TryParse(style, out int styleIndex))
                        value = Workbook.GetNumFmtValue(styleIndex, rawValue);
                }
                else
                {
                    value = rawValue;
                }
            }
            return value;
        }

        private XmlReader _reader;
        private Dictionary<string, Cell> _mergeCells;
        public bool Read(out List<Cell> row)
        {
            row = null;
            if (_reader != null)
            {
                row = new List<Cell>();
                (string t, string s) curCell = (null, null);

                /**
                * <sheetData>
                *     <row r="1">
                *          <c r="A1" s="11"><v>2</v></c>
                *          <c r="B1" s="11"><v>3</v></c>
                *          <c r="C1" s="11"><v>4</v></c>
                *          <c r="D1" t="s"><v>0</v></c>
                *          <c r="E1" t="inlineStr"><is><t>This is inline string example</t></is></c>
                *          <c r="D1" t="d"><v>1976-11-22T08:30</v></c>
                *          <c r="G1"><f>SUM(A1:A3)</f><v>9</v></c>
                *          <c r="H1" s="11"/>
                *      </row>
                * </sheetData>
                */
                int depth = 0;
                int[] flags = new int[4];
                while (_reader.Read())
                {
                    if (_reader.NodeType == XmlNodeType.Element)
                    {
                        if (depth == 0)
                            flags[0] = _reader.Name == "row" ? 1 : 0;
                        else if (depth == 1)
                            flags[1] = _reader.Name == "c" ? 2 : 0;
                        else if (depth == 2)
                            flags[2] = _reader.Name == "v" ? 3 : _reader.Name == "is" ? 31 : 0;
                        else if (depth == 3)
                            flags[3] = _reader.Name == "t" ? 41 : 0;

                        if (depth == 1 && flags[0] == 1 && flags[1] == 2)
                        {
                            row.Add(new Cell(_reader["r"], null));
                            curCell.t = _reader["t"];
                            curCell.s = _reader["s"];

                            foreach (var kv in _mergeCells)
                            {
                                if (kv.Value.Refer.Row <= row[row.Count - 1].Refer.Row && row[row.Count - 1].EndRefer.Row <= kv.Value.EndRefer.Row
                                    && kv.Value.Refer.Col <= row[row.Count - 1].Refer.Col && row[row.Count - 1].EndRefer.Col <= kv.Value.EndRefer.Col)
                                {
                                    row[row.Count - 1] = kv.Value;
                                    break;
                                }
                            }
                        }

                        if (!_reader.IsEmptyElement) ++depth;
                    }
                    else if (_reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (depth == 1 && flags[0] == 1)
                            return true;
                        else if (depth == 0)
                        {
                            _reader.Close();
                            _reader = null;
                            return false;
                        }
                        --depth;
                    }
                    else if (_reader.NodeType == XmlNodeType.Text)
                    {
                        if ((depth == 3 && flags[0] == 1 && flags[1] == 2 && flags[2] == 3)
                            || (depth == 4 && flags[0] == 1 && flags[1] == 2 && flags[2] == 31 && flags[3] == 41))
                        {
                            string value = ParseCellValue(row[row.Count - 1].Refer.Value, _reader.Value, curCell.t, curCell.s);
                            if (_mergeCells.TryGetValue(row[row.Count - 1].Refer.Value, out Cell mergeCell))
                                row[row.Count - 1] = mergeCell;
                            row[row.Count - 1].Value = value;
                        }
                    }
                }
            }
            return false;
        }

        public void Close()
        {
            _mergeCells.Clear();
            _reader?.Close();
            Workbook = null;
        }

        public void Dispose() => Close();
    }

    /// <summary>
    /// 工作簿
    /// </summary>
    public class Workbook : IDisposable
    {
        public const string RelationshipPart = "xl/_rels/workbook.xml.rels";
        public const string WorkbookPart = "xl/workbook.xml";
        public const string SharedStringsPart = "xl/sharedStrings.xml";
        public const string StylesPart = "xl/styles.xml";
        public static readonly Dictionary<int, string> StdNumFmts = new Dictionary<int, string>()
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

        private readonly ZipArchive _archive;
        private Dictionary<string, string> _rels;
        private List<KeyValuePair<string, string>> _worksheets;
        private List<string> _sharedStrings;
        private Dictionary<int, string> _numFmts;
        private List<int> _cellXfs;

        public Workbook(Stream stream)
        {
            _archive = new ZipArchive(stream, ZipArchiveMode.Read);
            Load();
        }

        public Workbook(string path) : this(new FileStream(path, FileMode.Open)) { }

        private void Load()
        {
            _rels = new Dictionary<string, string>();
            using (Stream stream = _archive?.GetEntry(RelationshipPart)?.Open())
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
                        int depth = 0;
                        int[] flags = new int[2];
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                {
                                    if (depth == 0)
                                        flags[0] = reader.Name == "Relationships" ? 1 : 0;
                                    else if (depth == 1)
                                        flags[1] = reader.Name == "Relationship" ? 2 : 0;

                                    if (depth == 1 && flags[0] == 1 && flags[1] == 2)
                                        _rels.Add(reader["Id"], "xl/" + reader["Target"]);

                                    if (!reader.IsEmptyElement) ++depth;
                                    break;
                                }
                                case XmlNodeType.EndElement:
                                {
                                    --depth;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            _worksheets = new List<KeyValuePair<string, string>>();
            using (Stream stream = _archive?.GetEntry(WorkbookPart)?.Open())
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
                        int depth = 0;
                        int[] flags = new int[3];
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                {
                                    if (depth == 0)
                                        flags[0] = reader.Name == "workbook" ? 1 : 0;
                                    else if (depth == 1)
                                        flags[1] = reader.Name == "sheets" ? 2 : 0;
                                    else if (depth == 2)
                                        flags[2] = reader.Name == "sheet" ? 3 : 0;

                                    if (depth == 2 && flags[0] == 1 && flags[1] == 2 && flags[2] == 3)
                                        _worksheets.Add(new KeyValuePair<string, string>(reader["name"], _rels[reader["r:id"]]));

                                    if (!reader.IsEmptyElement) ++depth;
                                    break;
                                }
                                case XmlNodeType.EndElement:
                                {
                                    --depth;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            _sharedStrings = new List<string>();
            using (Stream stream = _archive?.GetEntry(SharedStringsPart)?.Open())
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
                        int depth = 0;
                        int[] flags = new int[4];
                        string value = "";
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                {
                                    if (depth == 0)
                                        flags[0] = reader.Name == "sst" ? 1 : 0;
                                    else if (depth == 1)
                                        flags[1] = reader.Name == "si" ? 2 : 0;
                                    else if (depth == 2)
                                        flags[2] = reader.Name == "t" ? 3 : reader.Name == "r" ? 31 : 0;
                                    else if (depth == 3)
                                        flags[3] = reader.Name == "t" ? 41 : 0;

                                    if (!reader.IsEmptyElement) ++depth;
                                    break;
                                }
                                case XmlNodeType.EndElement:
                                {
                                    if (depth == 2 && flags[0] == 1 && flags[1] == 2)
                                    {
                                        _sharedStrings.Add(value);
                                        value = "";
                                    }
                                    --depth;
                                    break;
                                }
                                case XmlNodeType.Text:
                                {
                                    if (depth == 3 && flags[0] == 1 && flags[1] == 2 && flags[2] == 3)
                                    {
                                        value = reader.Value;
                                    }
                                    else if (depth == 4 && flags[0] == 1 && flags[1] == 2 && flags[2] == 31 && flags[3] == 41)
                                    {
                                        value += reader.Value;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            _numFmts = new Dictionary<int, string>(StdNumFmts);
            _cellXfs = new List<int>();
            using (Stream stream = _archive?.GetEntry(StylesPart)?.Open())
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
                        int depth = 0;
                        int[] flags = new int[3];
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                {
                                    if (depth == 0)
                                        flags[0] = reader.Name == "styleSheet" ? 1 : 0;
                                    else if (depth == 1)
                                        flags[1] = reader.Name == "numFmts" ? 2 : reader.Name == "cellXfs" ? 21 : 0;
                                    else if (depth == 2)
                                        flags[2] = reader.Name == "numFmt" ? 3 : reader.Name == "xf" ? 31 : 0;

                                    if (depth == 2 && flags[0] == 1)
                                    {
                                        if (flags[1] == 2 && flags[2] == 3)
                                            _numFmts[int.Parse(reader["numFmtId"])] = reader["formatCode"];
                                        else if (flags[1] == 21 && flags[2] == 31)
                                            _cellXfs.Add(int.Parse(reader["numFmtId"]));
                                    }

                                    if (!reader.IsEmptyElement) ++depth;
                                    break;
                                }
                                case XmlNodeType.EndElement:
                                {
                                    --depth;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        public string GetSharedString(int index) => 0 <= index && index < _sharedStrings.Count ? _sharedStrings[index] : null;
        public string GetNumFmtValue(int cellStyle, string rawValue)
        {
            if (0 <= cellStyle && cellStyle < _cellXfs.Count && _numFmts.TryGetValue(_cellXfs[cellStyle], out var formatCode))
            {
                if (formatCode == StdNumFmts[0])
                    return rawValue;
                else
                    return rawValue;
            }
            return null;
        }

        /// <summary>
        /// 工作表数量
        /// </summary>
        public int WorksheetCount => _worksheets.Count;

        private int _readIndex = 0;
        /// <summary>
        /// 读取一张表
        /// </summary>
        /// <returns></returns>
        public bool Read(out Worksheet sheet)
        {
            sheet = null;
            if (_readIndex < _worksheets.Count)
            {
                sheet = new Worksheet(this, _worksheets[_readIndex].Key, _archive.GetEntry(_worksheets[_readIndex].Value));
                ++_readIndex;
                return true;
            }
            return false;
        }

        public void Close()
        {
            _cellXfs.Clear();
            _numFmts.Clear();
            _sharedStrings.Clear();
            _worksheets.Clear();
            _rels.Clear();
            _archive.Dispose();
        }

        public void Dispose() => Close();
    }
}
