xnxq01id_list = [
    "2025-2026-1", "2024-2025-2", "2024-2025-1", "2023-2024-2", "2023-2024-1", "2022-2023-2", "2022-2023-1",
]
kbjcmsid_list = [
    ("04185F9CDDC04BC2AF96C38D2B31EB68", "长江新区校区作息时间"),
    ("952E3FC5FF563C09E053459072CA5D79", "武昌校区作息时间")
]
import requests
import openpyxl
from bs4 import BeautifulSoup
import csv

def choose_from_list(options, label, key='name'):
    print(f"请选择{label}：")
    for idx, item in enumerate(options):
        print(f"{idx+1}. {item[key] if isinstance(item, dict) else item}")
    while True:
        try:
            sel = int(input(f"输入序号(1-{len(options)}): "))
            if 1 <= sel <= len(options):
                return options[sel-1]
        except Exception:
            pass
        print("输入有误，请重新输入。")


def get_jsessionid():
    print("请在浏览器登录教务系统并进入课表查询界面，按F12打开开发者工具，切换到Application/存储/Storage->Cookies，找到JSESSIONID（路径为根目录），复制其值并粘贴到下方：")
    return input("请输入JSESSIONID: ").strip()

BASE_URL = "https://jwxt.hubu.edu.cn"
KB_URL = BASE_URL + "/kkglAction.do?method=toKbforward"

# 自动批量采集课表，参数来自csv
def get_kb_html(session, xnxq01id, kbjcmsid, yxbh, rxnf, zy, bjbh, xx04mc):
    data = {
        'method': 'toKbforward',
        'type': 'xx04',
        'isview': '1',
        'zc': '',
        'xnxq01id': xnxq01id,
        'kbjcmsid': kbjcmsid,
        'yxbh': yxbh,
        'rxnf': rxnf,
        'zy': zy,
        'bjbh': bjbh,
        'xx04id': bjbh,
        'xx04mc': xx04mc,
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0',
    }
    resp = session.post(KB_URL, headers=headers, data=data)
    return resp.text


def extract_full_cell(td):
    """提取一个td中所有kbcontent/kbcontent1的全部内容（包括隐藏的），结构化输出课程名/老师/教室/周次"""
    cell_texts = []
    divs = td.find_all(lambda tag: tag.name == "div" and tag.get("class", [""])[0].startswith("kbcontent"))
    for div in divs:
        # 结构化提取
        course_name = teacher = room = weeks = ""
        fonts = div.find_all("font")
        for f in fonts:
            t = f.get_text(separator="", strip=True)
            if t and "------------------------------" not in t and "课程编号" not in t and "方式" not in t:
                course_name = t
                break
        teacher_font = div.find('font', {'title': '老师'})
        if teacher_font:
            teacher = teacher_font.get_text(strip=True)
        room_font = div.find('font', {'title': '教室'})
        if room_font:
            room = room_font.get_text(strip=True)
        weeks_font = div.find('font', {'title': '周次(节次)'})
        if weeks_font:
            weeks = weeks_font.get_text(strip=True)
        if not weeks:
            for f in fonts:
                if f.has_attr("title") and "节" in f["title"]:
                    weeks = f["title"]
                    break
        # 合并输出
        info = " / ".join([x for x in [course_name, teacher, room, weeks] if x])
        if info:
            cell_texts.append(info)
    return "\n\n".join(cell_texts) if cell_texts else ""

def parse_kb_to_matrix(html):
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", {"id": "kbtable"})
    if table is None:
        raise ValueError("未找到课表table")
    ths = table.find_all("tr")[0].find_all("th")
    days = [th.get_text(strip=True) for th in ths[1:]]
    matrix = []
    trs = table.find_all("tr")[1:-1]
    for tr in trs:
        tds = tr.find_all("td")
        if not tds: continue
        section = tds[0].get_text(separator="", strip=True)
        row = [section]
        for td in tds[1:]:
            cell = extract_full_cell(td)
            row.append(cell)
        matrix.append(row)
    return days, matrix

def export_matrix_to_excel(days, matrix, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "课表"
    ws.append(["节次"] + days)
    for row in matrix:
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
        ws.row_dimensions[cell.row].height = max(15, min(max_lines * 15, 120))
    wb.save(filename)


def main():
    jsessionid = get_jsessionid()
    session = requests.Session()
    session.cookies.update({'JSESSIONID': jsessionid})
    # 读取参数csv
    with open("筛选的课表参数.csv", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = []
        for r in reader:
            # 统一去除key的BOM和空白
            row = {k.strip().replace('\ufeff',''): v for k, v in r.items()}
            rows.append(row)
    # 院系
    yx_list = []
    yx_seen = set()
    for row in rows:
        if row['yxbh'] and row['院系'] and (row['yxbh'], row['院系']) not in yx_seen:
            yx_list.append({'yxbh': row['yxbh'], 'name': row['院系']})
            yx_seen.add((row['yxbh'], row['院系']))
    yx = choose_from_list(yx_list, '院系')
    # 年级
    nj_list = []
    nj_seen = set()
    for row in rows:
        if row['yxbh'] == yx['yxbh'] and row['rxnf'] and (row['rxnf'], row['年级']) not in nj_seen:
            nj_list.append({'rxnf': row['rxnf'], 'name': row['年级']})
            nj_seen.add((row['rxnf'], row['年级']))
    nj = choose_from_list(nj_list, '年级')
    # 专业
    zy_list = []
    zy_seen = set()
    for row in rows:
        if row['yxbh'] == yx['yxbh'] and row['rxnf'] == nj['rxnf'] and row['zy'] and (row['zy'], row['专业']) not in zy_seen:
            zy_list.append({'zy': row['zy'], 'name': row['专业']})
            zy_seen.add((row['zy'], row['专业']))
    zy = choose_from_list(zy_list, '专业')
    # 班级
    bj_list = []
    for row in rows:
        if row['yxbh'] == yx['yxbh'] and row['rxnf'] == nj['rxnf'] and row['zy'] == zy['zy']:
            bj_list.append({'bjbh': row['bjbh'], 'name': row['班级'], 'xnxq01id': row.get('xnxq01id',''), 'kbjcmsid': row.get('kbjcmsid','')})
    bj = choose_from_list(bj_list, '班级')
    # 学年学期（固定列表选择）
    xnxq = choose_from_list([{'name': x} for x in xnxq01id_list], '学年学期')
    # 校区（固定列表选择）
    kbjc = choose_from_list([{'name': x[0], 'desc': x[1]} for x in kbjcmsid_list], '校区（04开头为长江新区校区，95开头为武昌校区）')
    # 找到最终参数行（只用bjbh匹配）
    param_row = None
    for row in rows:
        if row['bjbh'] == bj['bjbh']:
            param_row = row
            break
    if not param_row:
        print('未找到对应参数，退出')
        return
    # 采集课表
    try:
        html = get_kb_html(session, xnxq['name'], kbjc['name'], param_row['yxbh'], param_row['rxnf'], param_row['zy'], param_row['bjbh'], param_row['班级'])
        days, matrix = parse_kb_to_matrix(html)
        fname = f"{param_row['班级']}_{param_row['rxnf']}_{xnxq['name']}.xlsx".replace("/", "_").replace("[", "").replace("]", "").replace(" ", "")
        export_matrix_to_excel(days, matrix, fname)
        print(f"已保存: {fname}")
    except Exception as e:
        print(f"采集失败: {e}")

if __name__ == "__main__":
    main()
