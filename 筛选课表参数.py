import csv

# 输入输出文件名
input_file = '所有的课表参数.csv'
output_file = '筛选的课表参数.csv'

# 需要的年级
target_years = {'2025', '2024', '2023', '2022'}

with open(input_file, encoding='utf-8') as fin, open(output_file, 'w', encoding='utf-8-sig', newline='') as fout:
    reader = csv.reader(fin)
    writer = csv.writer(fout)
    header = next(reader)
    # 只保留前8列
    writer.writerow(header[:8])
    for row in reader:
        if len(row) < 8:
            continue
        if row[3] in target_years:
            writer.writerow(row[:8])

print(f'已筛选并保存到 {output_file}')
