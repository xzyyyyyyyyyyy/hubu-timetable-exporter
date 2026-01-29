
import csv
import openpyxl
from app_paths import resolve_input_path, resolve_output_path


def main():
    # 输入输出文件名
    input_file = resolve_input_path('所有的课表参数.csv')
    output_file = resolve_output_path('筛选的课表参数.csv')

    # 需要的年级
    target_years = {'2025', '2024', '2023', '2022'}

    rows = []
    with open(input_file, encoding='utf-8') as fin, open(output_file, 'w', encoding='utf-8-sig', newline='') as fout:
        reader = csv.reader(fin)
        writer = csv.writer(fout)
        header = next(reader)
        writer.writerow(header[:8])
        rows.append(header[:8])
        for row in reader:
            if len(row) < 8:
                continue
            if row[3] in target_years:
                writer.writerow(row[:8])
                rows.append(row[:8])

    # 导出为Excel并自动调整列宽行高
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "筛选课表参数"
    for row in rows:
        ws.append(row)
    # 自动调整列宽
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 50))
    # 自动调整行高
    for row in ws.iter_rows():
        max_lines = 1
        for cell in row:
            if cell.value:
                lines = str(cell.value).count("\n") + 1
                max_lines = max(max_lines, lines)
        ws.row_dimensions[cell[0].row].height = max(15, min(max_lines * 15, 120))
    excel_output = output_file.replace('.csv', '.xlsx')
    wb.save(excel_output)
    print(f'已筛选并保存到 {output_file} 和 {excel_output}')
    print(f'已筛选并保存到 {output_file}')


if __name__ == "__main__":
    main()
