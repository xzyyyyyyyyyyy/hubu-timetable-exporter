# 脚本功能：
# 使用Selenium自动化采集湖北大学教务系统“理论课表-各类课表查询”页面下所有院系、年级、专业、班级的参数组合，并导出为CSV。
# 操作流程：手动登录后自动切换iframe，依次遍历所有下拉选项，记录所有有效参数。

from selenium import webdriver  # Selenium主库

from selenium.webdriver.common.by import By  # 元素定位方式
from selenium.webdriver.support.ui import Select, WebDriverWait  # 下拉选择与等待
from selenium.webdriver.support import expected_conditions as EC  # 等待条件
import time, csv  # 时间控制与CSV导出
from selenium.webdriver.chrome.service import Service  # Chrome驱动服务
from selenium.common.exceptions import StaleElementReferenceException  # 处理动态页面异常
from app_paths import resolve_output_path


def main():
    # chromedriver路径（需根据实际环境修改）
    CHROMEDRIVER_PATH = r"G:\chromedriver-win64\chromedriver-win64\chromedriver.exe"
    # 目标教务系统地址
    url = "https://jwxt.hubu.edu.cn/"

    # 启动Chrome浏览器并打开目标网址
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service)
    driver.get(url)

    # 需人工登录，确保页面已切换到“理论课表-各类课表查询”
    input("请手动登录系统，切换到学期理论课表-各类课表查询后，按回车继续...")

    # 依次切换三层iframe，等待学年学期和课表层次下拉框加载
    wait = WebDriverWait(driver, 60)
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "Frame1")))
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "llsykb_find")))
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "llsykb_find")))
    elem_xnxq = wait.until(EC.presence_of_element_located((By.ID, "xnxq01id")))
    elem_kbjc = wait.until(EC.presence_of_element_located((By.ID, "kbjcmsid")))

    # 记录当前学年学期和课表层次参数
    xnxq01id = Select(elem_xnxq).first_selected_option.get_attribute("value")
    kbjcmsid = Select(elem_kbjc).first_selected_option.get_attribute("value")

    output_csv = resolve_output_path("所有的课表参数.csv")
    with open(output_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        # 写入表头
        writer.writerow(["yxbh", "院系", "rxnf", "年级", "zy", "专业", "bjbh", "班级", "xnxq01id", "kbjcmsid"])

        # 采集所有院系选项
        yx_select = Select(driver.find_element(By.ID, "yxbh"))
        yx_value_text_list = []
        for opt in yx_select.options:
            value = opt.get_attribute("value")
            text = opt.text
            if value and "请选择" not in text:
                yx_value_text_list.append((value, text))
        wait = WebDriverWait(driver, 30)
        for yxbh, yxmc in yx_value_text_list:
            try:
                yx_select = Select(driver.find_element(By.ID, "yxbh"))
                yx_options = [opt.get_attribute("value") for opt in yx_select.options]
            except StaleElementReferenceException:
                continue
            if yxbh not in yx_options:
                continue
            # 选择院系
            wait.until(EC.presence_of_element_located((By.XPATH, f'//select[@id="yxbh"]/option[@value="{yxbh}"]')))
            try:
                yx_select.select_by_value(yxbh)
            except StaleElementReferenceException:
                continue
            time.sleep(1.5)

            # 采集年级
            try:
                rxnf_select = Select(driver.find_element(By.ID, "rxnf"))
                rxnf_value_text_list = []
                for opt in rxnf_select.options:
                    value = opt.get_attribute("value")
                    text = opt.text
                    if value and "请选择" not in text:
                        rxnf_value_text_list.append((value, text))
            except StaleElementReferenceException:
                continue
            for rxnf, njmc in rxnf_value_text_list:
                try:
                    rxnf_select = Select(driver.find_element(By.ID, "rxnf"))
                    rxnf_options = [opt.get_attribute("value") for opt in rxnf_select.options]
                except StaleElementReferenceException:
                    continue
                if rxnf not in rxnf_options:
                    continue
                # 选择年级
                wait.until(EC.presence_of_element_located((By.XPATH, f'//select[@id="rxnf"]/option[@value="{rxnf}"]')))
                try:
                    rxnf_select.select_by_value(rxnf)
                except StaleElementReferenceException:
                    continue
                time.sleep(1.5)

                # 采集专业
                try:
                    zy_select = Select(driver.find_element(By.ID, "zy"))
                    zy_value_text_list = []
                    for opt in zy_select.options:
                        value = opt.get_attribute("value")
                        text = opt.text
                        if value and "请选择" not in text:
                            zy_value_text_list.append((value, text))
                except StaleElementReferenceException:
                    continue
                for zy, zymc in zy_value_text_list:
                    try:
                        zy_select = Select(driver.find_element(By.ID, "zy"))
                        zy_options = [opt.get_attribute("value") for opt in zy_select.options]
                    except StaleElementReferenceException:
                        continue
                    if zy not in zy_options:
                        continue
                    # 选择专业
                    wait.until(EC.presence_of_element_located((By.XPATH, f'//select[@id="zy"]/option[@value="{zy}"]')))
                    try:
                        zy_select.select_by_value(zy)
                    except StaleElementReferenceException:
                        continue
                    time.sleep(1.5)

                    # 采集班级
                    try:
                        bj_select = Select(driver.find_element(By.ID, "bjbh"))
                        bj_value_text_list = []
                        for opt in bj_select.options:
                            value = opt.get_attribute("value")
                            text = opt.text
                            if value and "请选择" not in text:
                                bj_value_text_list.append((value, text))
                    except StaleElementReferenceException:
                        continue
                    for bjbh, bjmc in bj_value_text_list:
                        try:
                            bj_select = Select(driver.find_element(By.ID, "bjbh"))
                            bj_options = [opt.get_attribute("value") for opt in bj_select.options]
                        except StaleElementReferenceException:
                            continue
                        if bjbh not in bj_options:
                            continue
                        # 选择班级
                        wait.until(EC.presence_of_element_located((By.XPATH, f'//select[@id="bjbh"]/option[@value="{bjbh}"]')))
                        try:
                            bj_select.select_by_value(bjbh)
                        except StaleElementReferenceException:
                            continue
                        time.sleep(1.5)
                        # 打印并写入当前参数组合
                        print(yxbh, yxmc, rxnf, njmc, zy, zymc, bjbh, bjmc, xnxq01id, kbjcmsid)
                        writer.writerow([yxbh, yxmc, rxnf, njmc, zy, zymc, bjbh, bjmc, xnxq01id, kbjcmsid])

    # 可选：打印当前页面源码（调试用）
    print(driver.page_source)
    # 关闭浏览器
    driver.quit()
    print(f"采集完成，结果已保存到 {output_csv}")


if __name__ == "__main__":
    main()