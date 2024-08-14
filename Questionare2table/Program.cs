using OfficeOpenXml;
using System.Data;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("请同时拖放Excel文件和namelist.txt到此程序上.");
            return;
        }

        // 获取文件路径
        var filePaths = GetFilePaths(args);
        if (filePaths.ExcelFilePath == null || filePaths.NameListFilePath == null)
        {
            Console.WriteLine("请确保拖放了正确的Excel文件和namelist.txt文件.");
            return;
        }

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filePaths.ExcelFilePath));
            var worksheet = package.Workbook.Worksheets[0];
            var dataTable = ExtractDataTableFromWorksheet(worksheet);
            ProcessDataAndSaveResults(dataTable, filePaths.NameListFilePath, filePaths.ExcelFilePath);
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"文件未找到: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理过程中发生错误: {ex.Message}");
        }
    }

    private static (string ExcelFilePath, string NameListFilePath) GetFilePaths(string[] args)
    {
        string excelFilePath = null;
        string nameListFilePath = null;

        foreach (var filePath in args)
        {
            if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                excelFilePath = filePath;
            }
            else if (Path.GetFileName(filePath).Equals("namelist.txt", StringComparison.OrdinalIgnoreCase))
            {
                nameListFilePath = filePath;
            }
        }

        return (excelFilePath, nameListFilePath);
    }

    private static DataTable ExtractDataTableFromWorksheet(ExcelWorksheet worksheet)
    {
        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        var dt = new DataTable();
        for (int col = 1; col <= colCount; col++)
        {
            if (!columnsToDelete.Contains(col))
            {
                dt.Columns.Add($"Column{col}");
            }
        }

        for (int row = 1; row <= rowCount; row++)
        {
            var dataRow = dt.NewRow();
            int dataColumnIndex = 0;
            for (int col = 1; col <= colCount; col++)
            {
                if (!columnsToDelete.Contains(col))
                {
                    string cellValue = worksheet.Cells[row, col].Text;
                    cellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, @".*—", "");
                    dataRow[dataColumnIndex++] = cellValue;
                }
            }
            dt.Rows.Add(dataRow);
        }

        return dt;
    }

    private static void ProcessDataAndSaveResults(DataTable dataTable, string nameListFilePath, string excelFilePath)
    {
        var nameList = File.ReadAllLines(nameListFilePath);
        var names = dataTable.Rows[0].ItemArray.Select(x => x.ToString()).ToList();
        var scores = dataTable.AsEnumerable().Skip(1).Select(r => r.ItemArray.Select(x => Convert.ToInt32(x)).ToArray()).ToList();
        var uniqueNames = names.Distinct().ToList();
        int numQuestions = scores[0].Length / uniqueNames.Count;

        var result = uniqueNames.ToDictionary(name => name, name => new int[numQuestions]);

        foreach (var scoreRow in scores)
        {
            var nameCounts = uniqueNames.ToDictionary(name => name, name => 0);

            for (int i = 0; i < names.Count; i++)
            {
                var name = names[i];
                if (result.ContainsKey(name))
                {
                    int index = nameCounts[name]++;
                    result[name][index] += scoreRow[i];
                }
            }
        }

        SaveResultsToExcel(result, numQuestions, excelFilePath);
    }

    private static void SaveResultsToExcel(Dictionary<string, int[]> result, int numQuestions, string excelFilePath)
    {
        var outputDt = new DataTable();
        outputDt.Columns.Add("姓名");
        for (int i = 1; i <= numQuestions; i++)
        {
            outputDt.Columns.Add($"T{i}");
        }

        foreach (var kvp in result)
        {
            var row = outputDt.NewRow();
            row["姓名"] = kvp.Key;
            for (int i = 0; i < kvp.Value.Length; i++)
            {
                row[$"T{i + 1}"] = kvp.Value[i];
            }
            outputDt.Rows.Add(row);
        }

        string outputFilePath = Path.Combine(Path.GetDirectoryName(excelFilePath), "final.xlsx");
        using var outputPackage = new ExcelPackage(new FileInfo(outputFilePath));
        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Sheet1");

        for (int i = 0; i < outputDt.Columns.Count; i++)
        {
            outputWorksheet.Cells[1, i + 1].Value = outputDt.Columns[i].ColumnName;
        }

        for (int i = 0; i < outputDt.Rows.Count; i++)
        {
            for (int j = 0; j < outputDt.Columns.Count; j++)
            {
                outputWorksheet.Cells[i + 2, j + 1].Value = outputDt.Rows[i][j];
            }
        }

        outputPackage.Save();
        Console.WriteLine($"结果已保存到 {outputFilePath}");
    }
}