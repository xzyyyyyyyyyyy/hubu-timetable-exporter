
import requests
import openpyxl
from bs4 import BeautifulSoup
import csv

# 可选学年学期和校区列表
xnxq01id_list = [
    "2025-2026-1", "2024-2025-2", "2024-2025-1", "2023-2024-2", "2023-2024-1",
]
kbjcmsid_list = [
    ("04185F9CDDC04BC2AF96C38D2B31EB68", "长江新区校区作息时间"),
    ("952E3FC5FF563C09E053459072CA5D79", "武昌校区作息时间")
]

def get_jsessionid():
    print("请在浏览器登录教务系统并进入课表查询界面，按F12打开开发者工具，切换到Application/存储/Storage->Cookies，找到JSESSIONID（路径为根目录），复制其值并粘贴到下方：")
    return input("请输入JSESSIONID: ").strip()

BASE_URL = "https://jwxt.hubu.edu.cn"
KB_URL = BASE_URL + "/kkglAction.do?method=toKbforward"

# 采集单个课表HTML
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
    cell_texts = []
    divs = td.find_all(lambda tag: tag.name == "div" and tag.get("class", [""])[0].startswith("kbcontent"))
    for div in divs:
        raw_txt = div.get_text(separator="\n", strip=True)
        if raw_txt:
            cell_texts.append(raw_txt)
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
            row = {k.strip().replace('\ufeff',''): v for k, v in r.items()}
            rows.append(row)
    # 批量采集
    fail_list = []
    import os
    root_dir = "全部课表导出"
    for row in rows:
        zy = row.get('专业', row.get('zy', '未知专业'))
        nj = row.get('年级', row.get('rxnf', '未知年级'))
        for xnxq01id in xnxq01id_list:
            for kbjcmsid, kbjc_desc in kbjcmsid_list:
                try:
                    html = get_kb_html(
                        session,
                        xnxq01id,
                        kbjcmsid,
                        row['yxbh'],
                        row['rxnf'],
                        row['zy'],
                        row['bjbh'],
                        row['班级']
                    )
                    days, matrix = parse_kb_to_matrix(html)
                    dir_path = os.path.join(root_dir, zy, nj)
                    os.makedirs(dir_path, exist_ok=True)
                    fname = f"{row['班级']}_{nj}_{xnxq01id}_{kbjc_desc}.xlsx".replace("/", "_").replace("[", "").replace("]", "").replace(" ", "")
                    file_path = os.path.join(dir_path, fname)
                    export_matrix_to_excel(days, matrix, file_path)
                    print(f"已保存: {file_path}")
                except Exception as e:
                    print(f"采集失败: {row['班级']} {nj} {xnxq01id} {kbjc_desc}，原因: {e}")
                    fail_row = row.copy()
                    fail_row['xnxq01id'] = xnxq01id
                    fail_row['kbjcmsid'] = kbjcmsid
                    fail_row['kbjc_desc'] = kbjc_desc
                    fail_list.append(fail_row)
    print(f"全部完成，失败数量: {len(fail_list)}")
    if fail_list:
        with open("采集失败列表.csv", "w", encoding="utf-8-sig", newline="") as fout:
            writer = csv.DictWriter(fout, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(fail_list)
        print("失败详情已保存到 采集失败列表.csv")

if __name__ == "__main__":
    main()
