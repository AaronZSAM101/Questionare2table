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

        string excelFilePath = null;
        string nameListFilePath = null;

        // 判断哪个文件是Excel文件，哪个是namelist.txt
        foreach (var filePath in args)
        {
            if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                excelFilePath = filePath;
            }
            else if (Path.GetFileName(filePath).Equals(".txt", StringComparison.OrdinalIgnoreCase))
            {
                nameListFilePath = filePath;
            }
        }

        if (excelFilePath == null || nameListFilePath == null)
        {
            Console.WriteLine("请确保拖放了正确的Excel文件和namelist.txt文件.");
            return;
        }

        try
        {

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // 获取表格中的所有数据
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // 创建DataTable来存储数据
                DataTable dt = new DataTable();
                for (int col = 1; col <= colCount; col++)
                {
                    dt.Columns.Add($"Column{col}");
                }

                // 读取数据到DataTable
                for (int row = 1; row <= rowCount; row++)
                {
                    var dataRow = dt.NewRow();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        // 使用正则表达式替换不需要的部分
                        cellValue = System.Text.RegularExpressions.Regex.Replace(cellValue, @".*—", "");
                        dataRow[col - 1] = cellValue;
                    }
                    dt.Rows.Add(dataRow);
                }

                // 删除前8列
                for (int i = 0; i < 8; i++)
                {
                    dt.Columns.RemoveAt(0);
                }

                // 读取姓名列表
                string[] nameList = File.ReadAllLines(nameListFilePath);

                // 获取姓名和对应的分数
                var names = dt.Rows[0].ItemArray.Select(x => x.ToString()).ToList();
                var scores = dt.AsEnumerable().Skip(1).Select(r => r.ItemArray.Select(x => Convert.ToInt32(x)).ToArray()).ToList();

                // 获取唯一的姓名列表
                var uniqueNames = names.Distinct().ToList();

                // 计算每个姓名对应的题目数量
                int numQuestions = scores[0].Length / uniqueNames.Count;

                // 初始化字典存储结果
                var result = new Dictionary<string, int[]>();
                foreach (var name in uniqueNames)
                {
                    result[name] = new int[numQuestions];
                }

                // 遍历每行分数进行累加
                foreach (var scoreRow in scores)
                {
                    var nameCounts = new Dictionary<string, int>();
                    foreach (var name in uniqueNames)
                    {
                        nameCounts[name] = 0;
                    }

                    for (int i = 0; i < names.Count; i++)
                    {
                        var name = names[i];
                        if (result.ContainsKey(name))
                        {
                            int index = nameCounts[name];
                            result[name][index] += scoreRow[i];
                            nameCounts[name]++;
                        }
                    }
                }

                // 生成题号列表，题号以"T"开头
                var questionTitles = Enumerable.Range(1, numQuestions).Select(i => $"T{i}").ToList();

                // 创建新的DataTable保存结果
                DataTable outputDt = new DataTable();
                outputDt.Columns.Add("姓名");

                foreach (var title in questionTitles)
                {
                    outputDt.Columns.Add(title);
                }

                // 填充结果到DataTable
                foreach (var kvp in result)
                {
                    var row = outputDt.NewRow();
                    row["姓名"] = kvp.Key;
                    for (int i = 0; i < kvp.Value.Length; i++)
                    {
                        row[questionTitles[i]] = kvp.Value[i];
                    }
                    outputDt.Rows.Add(row);
                }

                // 保存结果到Excel文件
                string outputFilePath = Path.Combine(Path.GetDirectoryName(excelFilePath), "final.xlsx");
                using (var outputPackage = new ExcelPackage(new FileInfo(outputFilePath)))
                {
                    var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Sheet1");

                    // 将DataTable数据写入新的Excel文件
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

                    // 保存文件
                    outputPackage.Save();
                }

                Console.WriteLine($"结果已保存到 {outputFilePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理过程中发生错误: {ex.Message}");
        }
    }
}