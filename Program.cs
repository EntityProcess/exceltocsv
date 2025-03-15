using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // For XLSX
using NPOI.HSSF.UserModel; // For XLS
using System;
using System.IO;

class Program
{
	static void Main(string[] args)
	{
		if (args.Length != 1 || string.IsNullOrEmpty(args[0]))
		{
			Console.Error.WriteLine("Usage: ExcelToCsv <excel-file-path>");
			Environment.Exit(1);
		}

		string filePath = args[0];
		if (!File.Exists(filePath))
		{
			Console.Error.WriteLine($"File not found: {filePath}");
			Environment.Exit(1);
		}

		try
		{
			IWorkbook? workbook = null; // Initialize workbook to null
			using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
					workbook = new HSSFWorkbook(fileStream); // XLS
				else if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
					workbook = new XSSFWorkbook(fileStream); // XLSX
				else
				{
					Console.Error.WriteLine("Unsupported file format. Use .xls or .xlsx.");
					Environment.Exit(1);
				}

				// Process all sheets
				for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
				{
					ISheet sheet = workbook.GetSheetAt(sheetIndex);
					string sheetName = workbook.GetSheetName(sheetIndex);
					
					// Output sheet header
					Console.WriteLine($"Sheet: {sheetName}");
					
					for (int rowIdx = 0; rowIdx <= sheet.LastRowNum; rowIdx++)
					{
						IRow row = sheet.GetRow(rowIdx);
						if (row == null) continue;

						var cells = new string[row.LastCellNum];
						for (int colIdx = 0; colIdx < row.LastCellNum; colIdx++)
						{
							ICell cell = row.GetCell(colIdx);
							cells[colIdx] = cell?.ToString()?.Replace(",", " ") ?? ""; // Replace commas
						}
						Console.WriteLine(string.Join(",", cells)); // Output CSV row
					}
					
					// Add a blank line between sheets for better readability
					Console.WriteLine();
				}
			} // workbook is disposed here with the fileStream
		}
		catch (Exception ex)
		{
			Console.Error.WriteLine($"Error processing file: {ex.Message}");
			Environment.Exit(1);
		}
	}
}
