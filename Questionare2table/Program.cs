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
                Console.WriteLine("请输入要过滤的列，以逗号分隔 (例如: A,B,C):");
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

        // 建立DataTable列，仅添加需要的列
        for (int col = 1; col < colCount; col++)  // 从第一列开始，并排除最后一列
        {
            if (!columnsToDelete.Contains(col) && col != colCount)  // 排除指定列和最后一列
            {
                dt.Columns.Add($"Column{col}");
            }
        }

        // 填充DataTable行
        for (int row = 1; row <= rowCount; row++)
        {
            var dataRow = dt.NewRow();
            int dataColIndex = 0;
            for (int col = 1; col < colCount; col++)  // 从第一列开始，并排除最后一列
            {
                if (!columnsToDelete.Contains(col) && col != colCount)  // 排除指定列和最后一列
                {
                    string cellValue = worksheet.Cells[row, col].Text;
                    cellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, @".*—", "");  // 正则表达式替换
                    dataRow[dataColIndex] = cellValue;
                    dataColIndex++;
                }
            }
            dt.Rows.Add(dataRow);
        }

        return dt;
    }
    private static void ProcessDataAndSaveResults(DataTable dataTable, string nameListFilePath, string excelFilePath)
    {
        var names = dataTable.Rows[0].ItemArray.Skip(1).Select(x => x.ToString()).ToList(); // 第一行的姓名列表，跳过第1列（此列是已删除列之后的“姓名”列）

        var scores = dataTable.AsEnumerable().Skip(1).Select(r => r.ItemArray.Skip(1).Select(x =>
        {
            if (int.TryParse(x.ToString(), out int result))
            {
                return result;
            }
            return 0; // 非数字内容返回0
        }).ToArray()).ToList();

        var uniqueNames = names.Distinct().ToList();
        int numQuestions = scores[0].Length / uniqueNames.Count;
        var result = uniqueNames.ToDictionary(name => name, name => new int[numQuestions]);

        for (int rowIndex = 0; rowIndex < scores.Count; rowIndex++)
        {
            var scoreRow = scores[rowIndex];
            var selfName = dataTable.Rows[rowIndex + 1][0].ToString(); // 获取当前行的姓名 (第1列)

            var nameCountTracker = uniqueNames.ToDictionary(name => name, name => 0);

            for (int colIndex = 0; colIndex < names.Count; colIndex++)
            {
                var name = names[colIndex];

                if (result.ContainsKey(name))
                {
                    // 检查是否为自评（即当前行的姓名与列名相同），如果是则跳过
                    if (selfName != name)
                    {
                        int index = nameCountTracker[name]++;
                        result[name][index] += scoreRow[colIndex];
                    }
                }
            }
        }

        SaveResultsToExcel(result, numQuestions, excelFilePath);
    }
    private static void SaveResultsToExcel(Dictionary<string, int[]> result, int numQuestions, string excelFilePath)
    {
        var outputDt = new DataTable();
        outputDt.Columns.Add("姓名", typeof(string));

        // 为每个问题添加列
        for (int i = 1; i <= numQuestions; i++)
        {
            outputDt.Columns.Add($"T{i}", typeof(int));
        }

        // 添加平均分列
        outputDt.Columns.Add("平均分", typeof(double));

        // 计算平均分并填充数据表
        int totalParticipants = result.Keys.Count;
        foreach (var kvp in result)
        {
            var row = outputDt.NewRow();
            row["姓名"] = kvp.Key;
            int sum = 0;

            // 填充题目分数并计算总和
            for (int i = 0; i < kvp.Value.Length; i++)
            {
                row[$"T{i + 1}"] = kvp.Value[i];
                sum += kvp.Value[i];
            }

            // 计算平均分并填入最后一列
            double average = (double)sum / (totalParticipants - 1);
            row["平均分"] = average;
            outputDt.Rows.Add(row);
        }

        // 保存结果到Excel文件
        string outputFilePath = Path.Combine(Path.GetDirectoryName(excelFilePath), "final.xlsx");
        using var outputPackage = new ExcelPackage(new FileInfo(outputFilePath));
        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Sheet1");

        // 填充Excel表头
        for (int i = 0; i < outputDt.Columns.Count; i++)
        {
            outputWorksheet.Cells[1, i + 1].Value = outputDt.Columns[i].ColumnName;
        }

        // 填充Excel内容并确保数值以数值类型存储
        for (int i = 0; i < outputDt.Rows.Count; i++)
        {
            for (int j = 0; j < outputDt.Columns.Count; j++)
            {
                var value = outputDt.Rows[i][j];
                outputWorksheet.Cells[i + 2, j + 1].Value = value;
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