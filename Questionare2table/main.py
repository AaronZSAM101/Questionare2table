import pandas as pd

# 读取原始Excel文件并删除无用列
file_path = 'data.xlsx'
df = pd.read_excel(file_path, header=None)

df = df.replace(to_replace=r".*—", value='', regex=True)
df.drop(columns=[0, 1, 2, 3, 4, 5, 6, 7], inplace=True)

# 读取姓名列表
with open('namelist.txt', 'r', encoding='utf-8') as f:
    name_list = [line.strip() for line in f]

# 获取姓名和对应的分数
names = df.iloc[0, :].tolist()  # 获取第一行作为姓名
scores = df.iloc[1:, :].values  # 获取后面的所有行作为分数

# 获取唯一的姓名列表
unique_names = list(set(names))

# 计算每个姓名对应的题目数量
num_questions = len(scores[0]) // len(unique_names)

# 初始化字典存储结果
result = {name: [0] * num_questions for name in unique_names}

# 遍历每行分数进行累加
for score_row in scores:
    name_counts = {name: 0 for name in unique_names}  # 记录每个姓名的题目顺序
    for name, score in zip(names, score_row):
        if name in result:
            index = name_counts[name]  # 累加到正确的位置
            result[name][index] += score
            name_counts[name] += 1

# 生成题号列表，题号以"T"开头
question_titles = [f"T{i+1}" for i in range(num_questions)]

# 将结果转换为DataFrame，并添加标题
output_df = pd.DataFrame.from_dict(result, orient='index', columns=question_titles)

# 添加"姓名"列标题
output_df.index.name = '姓名'

# 重置索引，使"姓名"成为DataFrame的一列
output_df.reset_index(inplace=True)

# 保存结果到Excel文件
output_file = 'final.xlsx'
output_df.to_excel(output_file, index=False)

print(f"结果已保存到 {output_file}")