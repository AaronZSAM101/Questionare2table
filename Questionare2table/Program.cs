using OfficeOpenXml;
using System.Data;
using System.Text.Json;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("请同时拖放Excel文件和namelist.txt到此程序上.");
            return;
        }

        var config = LoadConfig();
        var filePaths = GetFilePaths(args);

        if (filePaths.ExcelFilePath == null || filePaths.NameListFilePath == null)
        {
            Console.WriteLine("请确保拖放了正确的Excel文件和namelist.txt文件.");
            return;
        }

        try
        {
            List<int> columnsToDelete;
            if (string.IsNullOrEmpty(config.ColumnsToDelete))
            {
                Console.WriteLine("请输入要删除的列，以逗号分隔 (例如: A,B,C):");
                var columnsToDeleteInput = Console.ReadLine();
                columnsToDelete = ParseColumns(columnsToDeleteInput);
                config.ColumnsToDelete = columnsToDeleteInput;
                SaveConfig(config);
            }
            else
            {
                Console.WriteLine($"使用上次的配置，删除以下列: {config.ColumnsToDelete}");
                columnsToDelete = ParseColumns(config.ColumnsToDelete);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filePaths.ExcelFilePath));
            var worksheet = package.Workbook.Worksheets[0];
            var dataTable = ExtractDataTableFromWorksheet(worksheet, columnsToDelete);
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

    private static Config LoadConfig()
    {
        var configPath = "config.json";
        if (!File.Exists(configPath))
        {
            return new Config();
        }

        var configJson = File.ReadAllText(configPath);
        return JsonSerializer.Deserialize<Config>(configJson);
    }

    private static void SaveConfig(Config config)
    {
        var configJson = JsonSerializer.Serialize(config);
        File.WriteAllText("config.json", configJson);
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

    private static List<int> ParseColumns(string columnsToDeleteInput)
    {
        var columnsToDelete = new List<int>();

        foreach (var column in columnsToDeleteInput.Split(','))
        {
            int columnNumber = GetColumnNumber(column.Trim());
            columnsToDelete.Add(columnNumber);
        }

        return columnsToDelete;
    }

    private static int GetColumnNumber(string columnName)
    {
        int columnNumber = 0;
        int multiplier = 1;

        for (int i = columnName.Length - 1; i >= 0; i--)
        {
            char letter = columnName[i];
            columnNumber += (letter - 'A' + 1) * multiplier;
            multiplier *= 26;
        }

        return columnNumber;
    }

    private static DataTable ExtractDataTableFromWorksheet(ExcelWorksheet worksheet, List<int> columnsToDelete)
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
    private class Config
    {
        public string ColumnsToDelete { get; set; }
    }
}