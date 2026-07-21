from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import cloudscraper
from bs4 import BeautifulSoup
import pandas as pd
import time

# ==============================
# 配置
# ==============================
# 比赛ID
contest_id = "133876"

school_name = "南京邮电大学"

# ID对应的名单
name_map = {
    "nju": "张三",
    }

# ==============================
# Selenium获取排名
# ==============================

def get_rank_by_selenium():

    print("启动浏览器获取排名...")

    driver = webdriver.Chrome(
        service=Service(
            r"D:\chromedriver\chromedriver.exe"
        )
    )

    wait = WebDriverWait(
        driver,
        20
    )

    driver.get(
        f"https://ac.nowcoder.com/acm/contest/{contest_id}#rank"
    )

    time.sleep(5)

    user_list = []

    # 搜索学校

    try:

        search_box = wait.until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "input[placeholder*='输入用户名称']"
                )
            )
        )

        search_box.clear()

        search_box.send_keys(
            school_name
        )

        search_box.send_keys(
            Keys.ENTER
        )

        print(
            "已搜索学校"
        )

        time.sleep(5)

    except Exception as e:

        print(
            "搜索失败:",
            e
        )

    # 获取排名表

    rows = driver.find_elements(
        By.CSS_SELECTOR,
        "tr"
    )

    for row in rows:

        if school_name not in row.text:
            continue

        try:

            cols = row.find_elements(
                By.TAG_NAME,
                "td"
            )

            if len(cols)<5:
                continue

            name_element = cols[1].find_element(
                By.TAG_NAME,
                "a"
            )

            user_list.append(
                {
                    "排名":
                    cols[0].text.strip(),

                    "昵称":
                    name_element.text.strip(),

                    "主页":
                    name_element.get_attribute(
                        "href"
                    ),

                    "AC数量":
                    cols[3].text.strip(),

                    "总分":
                    cols[4].text.strip()

                }
            )
        except:
            continue

    print(
        f"找到 {len(user_list)} 名同学"
    )

    driver.quit()

    return user_list

# ==============================
# cloudscraper获取Rating
# ==============================

def get_rating(user_list):
    print(
        "\n开始获取Rating..."
    )

    scraper = cloudscraper.create_scraper(
        browser={
            "browser":"chrome",
            "platform":"windows",
            "mobile":False
        }
    )

    headers = {

        "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        " AppleWebKit/537.36 Chrome/120 Safari/537.36"

    }

    final_data=[]

    for user in user_list:

        print(
            "查询:",
            user["昵称"],
            end=" "
        )

        try:

            resp = scraper.get(
                user["主页"],
                headers=headers,
                timeout=20
            )

            soup = BeautifulSoup(
                resp.text,
                "html.parser"
            )

            rating = soup.select_one(
                "a[class*='state-num']"
            )

            if rating:

                user["Rating"] = (
                    rating.text.strip()
                )

            else:

                user["Rating"] = (
                    "暂无"
                )

        except Exception as e:

            print(
                "失败:",
                e
            )

            user["Rating"] = (
                "获取失败"
            )

        print(
            user["Rating"]
        )

        final_data.append(
            user
        )

        time.sleep(0.5)

    return final_data

# ==============================
# 添加真实姓名
# ==============================

def add_real_name(data):

    for user in data:

        user["姓名"] = (
            name_map.get(
                user["昵称"],
                "未知"
            )
        )
    return data

# ==============================
# 保存Excel
# ==============================

def save_excel(data):

    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    if data:

        df = pd.DataFrame(data)

        # 调整列顺序
        df = df[
            [
                "排名",
                "姓名",
                "昵称",
                "AC数量",
                "总分",
                "Rating"
            ]
        ]

        output_file = (
            "Rating详情.xlsx"
        )

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                index=False,
                sheet_name="Sheet1"
            )

            worksheet = writer.sheets["Sheet1"]

            # ==========================
            # 单元格格式
            # ==========================

            align_format = Alignment(
                wrap_text=True,
                vertical="center",
                horizontal="center"
            )

            # ==========================
            # 中文宽度计算
            # ==========================

            def get_string_width(text):

                width = 0

                for char in str(text):

                    if ord(char) > 127:

                        width += 2.0

                    else:

                        width += 1.1

                return width

            # ==========================
            # 设置单元格格式和行高
            # ==========================

            for row in worksheet.iter_rows(

                min_row=1,

                max_row=worksheet.max_row,

                min_col=1,

                max_col=worksheet.max_column

            ):

                max_lines = 1

                for cell in row:

                    cell.alignment = align_format

                    if cell.value is not None:

                        val_str = str(cell.value)

                        line_count = len(
                            val_str.split("\n")
                        )

                        if line_count > max_lines:

                            max_lines = line_count

                row_idx = row[0].row

                if row_idx == 1:

                    worksheet.row_dimensions[
                        row_idx
                    ].height = 25

                else:

                    worksheet.row_dimensions[
                        row_idx
                    ].height = max(
                        22,
                        max_lines * 19
                    )

            # ==========================
            # 设置列宽
            # ==========================

            for col in worksheet.columns:

                max_len = 0

                col_letter = get_column_letter(
                    col[0].column
                )

                for cell in col:

                    if cell.value is not None:

                        lines = str(
                            cell.value
                        ).split("\n")

                        for line in lines:

                            line_width = (
                                get_string_width(line)
                            )

                            if line_width > max_len:

                                max_len = line_width

                worksheet.column_dimensions[
                    col_letter
                ].width = max(
                    max_len + 4,
                    12
                )

        print(
            f"\n🎉 任务完成！文件已保存为: {output_file}"
        )
    else:
        print(
            "\n❌ 未能提取到任何数据，请检查登录状态或搜索词。"
        )
# ==============================
# 主函数
# ==============================

def main():

    users = get_rank_by_selenium()

    if len(users)==0:
        print(
            "没有找到用户"
        )
        return

    data = get_rating(
        users
    )

    data = add_real_name(
        data
    )

    save_excel(
        data
    )

if __name__ == "__main__":
    main()
