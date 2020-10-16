using System;
using System.Collections.Generic;

#if DEBUG
namespace XlsxTextReader
{
    class Program
    {
        static void Main(string[] args)
        {
            using (Workbook workbook = Workbook.Open(@"doc\example.xlsx"))
            {
                foreach (Worksheet worksheet in workbook.Read())
                {
                    Console.WriteLine("----------- 表：" + worksheet.Name + " -----------");
                    foreach (List<Cell> row in worksheet.Read())
                    {
                        foreach (Cell cell in row)
                            Console.Write(cell.Reference.Value + ": " + cell.Value + "\t");
                        Console.WriteLine();
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}
#endif