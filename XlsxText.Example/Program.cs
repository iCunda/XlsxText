using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                XlsxTextSheetReader sheetReader = xlsx.SheetReader;
                Console.WriteLine("Sheet Name: " + sheetReader.Name);

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
            }
            Console.ReadKey();
        }
    }
}
