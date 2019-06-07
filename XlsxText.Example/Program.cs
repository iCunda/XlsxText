using System;
using System.Diagnostics;

namespace XlsxText.Example
{
    class Program
    {
        public const string ResourcePath = "../../../Resource";
        static void Main(string[] args)
        {
            XlsxTextReader xlsx = XlsxTextReader.Create(ResourcePath + "/example.xlsx");
            while (xlsx.Read())
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                XlsxTextSheetReader sheetReader = xlsx.SheetReader;
                Console.WriteLine("Sheet Name: " + sheetReader.Name + ", 行数: " + sheetReader.RowCount + ", 单元格数: " + sheetReader.CellCount);

                while (sheetReader.Read())
                {
                    if (sheetReader.Row.Count == 0)
                        continue;
                    foreach (var cell in sheetReader.Row)
                    {
                        Console.Write(cell.Value + "\t");
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
                sw.Stop();
                Console.WriteLine("sw总共花费{0}ms.", sw.Elapsed.TotalMilliseconds);
            }

            Console.ReadKey();
        }
    }
}
