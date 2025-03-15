using NPOI.HSSF.UserModel; // For XLS
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // For XLSX
using System.Text;

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
			using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				IWorkbook? workbook = null;
				try
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
							if (row == null) 
							{
								Console.WriteLine();
								continue;
							}

							var sb = new StringBuilder();
							for (int colIdx = 0; colIdx < row.LastCellNum; colIdx++)
							{
								if (colIdx > 0) sb.Append(',');
								
								ICell cell = row.GetCell(colIdx);
								if (cell != null)
								{
									try
									{
										string cellValue = GetCellValueSafely(cell);
										sb.Append(cellValue);
									}
									catch (Exception ex)
									{
										// If there's an error processing a cell, output empty value
										Console.Error.WriteLine($"Warning: Error processing cell at row {rowIdx + 1}, column {colIdx + 1}: {ex.Message}");
									}
								}
							}
							Console.WriteLine(sb.ToString());
						}
						
						// Add a blank line between sheets for better readability
						Console.WriteLine();
					}
				}
				finally
				{
					workbook?.Close();
				}
			}
		}
		catch (Exception ex)
		{
			Console.Error.WriteLine($"Error processing file: {ex.Message}");
			Environment.Exit(1);
		}
	}

	private static string GetCellValueSafely(ICell cell)
	{
		try
		{
			string value = "";
			switch (cell.CellType)
			{
				case CellType.Numeric:
					if (DateUtil.IsCellDateFormatted(cell))
						value = cell.DateCellValue?.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) ?? "";
					else
						value = cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
					break;
				case CellType.String:
					value = cell.StringCellValue;
					break;
				case CellType.Boolean:
					value = cell.BooleanCellValue.ToString();
					break;
				case CellType.Formula:
					try
					{
						value = cell.StringCellValue;
					}
					catch
					{
						try
						{
							value = cell.NumericCellValue.ToString();
						}
						catch
						{
							value = cell.CellFormula;
						}
					}
					break;
				default:
					value = cell?.ToString() ?? "";
					break;
			}
			return value?.Replace(",", " ")?.Replace("\n", " ")?.Replace("\r", "") ?? "";
		}
		catch
		{
			return "";
		}
	}
}
