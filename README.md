# XlsxText
Quickly read the text in *.xlsx.

----------
example：

    using (Workbook workbook = new Workbook("D:/example.xlsx"))
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
    Console.ReadKey();
