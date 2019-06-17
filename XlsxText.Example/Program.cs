using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace XlsxText.Example
{
    class Program
    {
        public const string ResourcePath = "../../../Resource";
        static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            using (Workbook workbook = new Workbook(ResourcePath + "/example.xlsx"))
            {
                while (workbook.Read(out Worksheet worksheet))
                {
                    Console.WriteLine("Sheet Name: " + worksheet.Name + ", 行数: " + worksheet.RowCount);

                    while (worksheet.Read(out List<Cell> row))
                    {
                        foreach (var cell in row)
                        {
                            Console.Write(cell.Value + "\t");
                        }
                        Console.WriteLine();
                    }
                    Console.WriteLine();
                }
            }
            sw.Stop();
            Console.WriteLine("总共消耗{0}ms.", sw.Elapsed.TotalMilliseconds);
            Console.ReadKey();
        }
    }
}
