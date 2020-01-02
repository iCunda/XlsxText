# XlsxTextReader
Quickly read the text in *.xlsx.

----------
example：

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
