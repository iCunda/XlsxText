# XlsxText
Quickly read the text in *.xlsx.

----------
example：

    XlsxTextReader xlsx = XlsxTextReader.Create(@"D:\example.xlsx");
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
