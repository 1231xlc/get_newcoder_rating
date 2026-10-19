# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import cloudscraper
from bs4 import BeautifulSoup

import pandas as pd
import time
import re
import json
import html as html_lib
import os
import atexit
import platform
import shutil
import subprocess
import tempfile
import urllib.request
import zipfile


# ==============================
# 配置
# ==============================
# EdgeDriver 路径。留空时脚本会尝试临时下载匹配的驱动
path = r""

# 指定比赛 ID。若不设置该变量，脚本会自动选择最近结束的目标比赛
# contest_id = "130107"

# 用于在排名页面搜索的学校名称或关键词
school_name = "你的学校名称"

# 牛客昵称到真实姓名的映射
name_map = {
    "牛客昵称1": "姓名1",
    "牛客昵称2": "姓名2",
}

# ==============================
# Edge / edgedriver 版本检测与下载
# ==============================
EDGEDRIVER_DOWNLOAD_HOSTS = (
    "https://msedgedriver.microsoft.com",
    "https://msedgedriver.azureedge.net",
)


def _version_tuple(version):
    """把版本号字符串转成可比较的元组"""
    try:
        return tuple(
            int(part)
            for part in re.findall(r"\d+", str(version))[:4]
        )
    except (TypeError, ValueError):
        return ()


def get_local_edge_version():
    """获取本地 Edge 浏览器版本，优先读取注册表，失败时读取安装目录"""
    try:
        import winreg

        registry_keys = (
            (
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Edge\BLBeacon"
            ),
            (
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Edge\BLBeacon"
            ),
            (
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon"
            ),
        )

        for root, sub_key in registry_keys:
            try:
                with winreg.OpenKey(root, sub_key) as key:
                    version, _ = winreg.QueryValueEx(
                        key,
                        "version"
                    )
                version = str(version).strip()
                if re.fullmatch(r"\d+(\.\d+){3}", version):
                    return version
            except OSError:
                continue
    except ImportError:
        pass

    # 注册表中找不到时，从 Edge 安装目录的版本文件夹推断
    base_dirs = []

    for env_name in (
        "PROGRAMFILES(X86)",
        "PROGRAMFILES",
        "LOCALAPPDATA"
    ):
        base = os.environ.get(env_name)
        if base:
            base_dirs.append(
                os.path.join(
                    base,
                    "Microsoft",
                    "Edge",
                    "Application"
                )
            )

    versions = []

    for base in base_dirs:
        if not os.path.isdir(base):
            continue

        for name in os.listdir(base):
            if re.fullmatch(r"\d+(\.\d+){3}", name):
                versions.append(name)

    if versions:
        return max(
            versions,
            key=_version_tuple
        )

    return None


def get_driver_version(driver_path):
    """执行 msedgedriver --version 并解析出版本号"""
    if not driver_path or not os.path.isfile(driver_path):
        return None

    try:
        result = subprocess.run(
            [driver_path, "--version"],
            capture_output=True,
            text=True,
            timeout=20,
            creationflags=getattr(
                subprocess,
                "CREATE_NO_WINDOW",
                0
            )
        )

        output = (result.stdout or "") + (result.stderr or "")

        match = re.search(
            r"\d+(?:\.\d+){3}",
            output
        )

        return match.group(0) if match else None

    except Exception:
        return None


def _download_bytes(url, timeout=60):
    """下载 URL 内容并返回 bytes"""
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        }
    )

    with urllib.request.urlopen(
        request,
        timeout=timeout
    ) as response:
        return response.read()


def _get_latest_driver_version(major_version):
    """查询与本地 Edge 主版本号对应的最新 edgedriver 版本"""
    for host in EDGEDRIVER_DOWNLOAD_HOSTS:

        url = (
            f"{host}/LATEST_RELEASE_"
            f"{major_version}"
        )

        try:
            data = _download_bytes(
                url,
                timeout=30
            )
        except Exception:
            continue

        for encoding in ("utf-16", "utf-8-sig", "utf-8"):

            try:
                text = data.decode(encoding).strip()
            except (UnicodeDecodeError, LookupError):
                continue

            match = re.search(
                r"\d+(?:\.\d+){3}",
                text
            )

            if match:
                return match.group(0)

    return None


def _get_driver_zip_urls(version):
    """生成该版本 edgedriver 压缩包的候选下载地址"""
    machine = platform.machine().lower()

    if machine in ("arm64", "aarch64"):
        arch = "arm64"
    else:
        arch = "win64"

    return [
        f"{host}/{version}/edgedriver_{arch}.zip"
        for host in EDGEDRIVER_DOWNLOAD_HOSTS
    ]


def _replace_file(src, dst):
    """用新文件替换目标文件，目标被占用时先重命名备份再写入"""
    dst = os.path.abspath(dst)

    target_dir = os.path.dirname(dst)
    if target_dir:
        os.makedirs(
            target_dir,
            exist_ok=True
        )

    try:
        os.replace(src, dst)
        return
    except OSError:
        pass

    backup = (
        f"{dst}.bak."
        f"{int(time.time())}"
    )

    try:
        os.replace(dst, backup)
    except OSError:
        backup = None

    shutil.move(src, dst)

    if backup:
        try:
            os.remove(backup)
        except OSError:
            pass


def download_edgedriver(target_path=None):
    """
    下载与本地 Edge 版本对应的 edgedriver。

    target_path 为 None 时下载到临时目录并返回临时文件路径，
    脚本退出时自动清理；否则安装到 target_path。
    """
    edge_version = get_local_edge_version()

    if not edge_version:
        raise RuntimeError(
            "无法检测本地 Edge 浏览器版本"
        )

    major_version = edge_version.split(".")[0]

    versions = [edge_version]

    latest_version = _get_latest_driver_version(
        major_version
    )

    if (
        latest_version
        and latest_version not in versions
    ):
        versions.append(latest_version)

    print(
        f"本地 Edge 版本: {edge_version}，"
        f"准备下载对应的 edgedriver..."
    )

    last_error = None

    for version in versions:

        for url in _get_driver_zip_urls(version):

            try:
                zip_data = _download_bytes(
                    url,
                    timeout=120
                )

                download_dir = tempfile.mkdtemp(
                    prefix="edgedriver_dl_"
                )

                try:
                    zip_path = os.path.join(
                        download_dir,
                        "edgedriver.zip"
                    )

                    with open(zip_path, "wb") as f:
                        f.write(zip_data)

                    with zipfile.ZipFile(zip_path) as zf:
                        exe_name = next(
                            (
                                name
                                for name in zf.namelist()
                                if name.lower().endswith(
                                    "msedgedriver.exe"
                                )
                            ),
                            None
                        )

                        if not exe_name:
                            raise RuntimeError(
                                "压缩包中未找到 msedgedriver.exe"
                            )

                        zf.extract(
                            exe_name,
                            download_dir
                        )

                    extracted = os.path.join(
                        download_dir,
                        exe_name
                    )

                    if target_path:
                        _replace_file(
                            extracted,
                            target_path
                        )

                        print(
                            f"[OK] edgedriver {version} 已安装到: "
                            f"{target_path}"
                        )

                        return os.path.abspath(
                            target_path
                        )

                    # 没有配置 path：保留临时目录，退出时清理
                    atexit.register(
                        shutil.rmtree,
                        download_dir,
                        ignore_errors=True
                    )

                    print(
                        f"[OK] 已临时下载 edgedriver {version}: "
                        f"{extracted}"
                    )

                    return extracted

                finally:
                    # 安装到目标路径后即可删除下载目录
                    if target_path:
                        shutil.rmtree(
                            download_dir,
                            ignore_errors=True
                        )

            except Exception as e:
                last_error = e
                continue

    raise RuntimeError(
        f"下载 edgedriver 失败: {last_error}"
    )


# ==============================
# 确认是否安装edgedriver
# ==============================
def ensure_edgedriver():
    """
    检查 edgedriver 是否存在且与本地 Edge 版本匹配。

    - path 已配置：不匹配或不存在时下载并替换为该路径的 edgedriver
    - path 未配置：临时下载对应版本的 edgedriver 供本次使用
    """
    configured_path = globals().get("path")

    # 没有 path 变量（或为空）：临时下载
    if not configured_path:
        print(
            "未配置 path 变量，"
            "临时下载与本地 Edge 匹配的 edgedriver..."
        )

        try:
            return download_edgedriver(None)
        except Exception as e:
            print(
                "临时下载 edgedriver 失败:",
                e
            )
            return None

    configured_path = str(configured_path)

    edge_version = get_local_edge_version()
    driver_exists = os.path.isfile(configured_path)

    if driver_exists:
        driver_version = get_driver_version(
            configured_path
        )

        if not edge_version:
            print(
                "无法检测本地 Edge 版本，"
                "直接使用现有 edgedriver"
            )
            return configured_path

        if (
            driver_version
            # edgedriver 与浏览器主版本号一致即可兼容
            and driver_version.split(".")[0]
            == edge_version.split(".")[0]
        ):
            print(
                f"edgedriver 版本 {driver_version} "
                f"与本地 Edge {edge_version} 匹配"
            )
            return configured_path

        print(
            f"edgedriver 版本 {driver_version or '未知'} "
            f"与本地 Edge {edge_version} 不匹配，"
            f"正在下载替换..."
        )
    else:
        print(
            f"未找到 edgedriver: {configured_path}，"
            f"正在下载..."
        )

    try:
        return download_edgedriver(
            configured_path
        )
    except Exception as e:
        print(
            "下载 edgedriver 失败:",
            e
        )

        if os.path.isfile(configured_path):
            print(
                "继续使用现有 edgedriver"
            )
            return configured_path

        return None



def get_latest_nowcoder_contest_id():

    print("正在获取最近完成的牛客比赛...")


    scraper = cloudscraper.create_scraper(
        browser={
            "browser": "chrome",
            "platform": "windows",
            "mobile": False
        }
    )


    headers = {
        "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }


    url = "https://ac.nowcoder.com/acm/contest/vip-index?topCategoryFilter=13"


    resp = scraper.get(
        url,
        headers=headers,
        timeout=40
    )


    html = resp.text


    now = int(time.time()*1000)


    contests=[]


    blocks = re.split(
        r'data-id="',
        html
    )


    for block in blocks[1:]:


        match = re.match(
            r'(\d+)',
            block
        )

        if not match:
            continue


        cid = match.group(1)


        pos = block.find(
            'data-json="'
        )


        if pos < 0:
            continue


        pos += len(
            'data-json="'
        )


        end = block.find(
            '"',
            pos
        )


        if end < 0:
            continue


        raw = block[pos:end]


        data_text = html_lib.unescape(
            html_lib.unescape(raw)
        )


        try:

            data=json.loads(
                data_text
            )

        except:

            continue



        name=data.get(
            "contestName",
            ""
        )


        start=data.get(
            "contestStartTime",
            0
        )


        finish=data.get(
            "contestEndTime",
            0
        )


        if not name or not finish:
            continue



        # 过滤比赛类型
        if not (
            "周赛" in name
            or "小白月赛" in name
            or "挑战赛" in name
            or "练习赛" in name
        ):
            continue



        # 只选择已经结束的比赛
        if finish < now:

            contests.append(
                (
                    finish,
                    cid,
                    name
                )
            )



    if not contests:

        raise Exception(
            "没有找到已结束比赛"
        )


    # 按结束时间倒序
    contests.sort(
        reverse=True
    )


    finish,cid,name = contests[0]


    print(
        f"最近完成比赛: {name}"
    )

    print(
        f"contest_id={cid}"
    )


    return cid

try:
    contest_id
except NameError:
    contest_id = get_latest_nowcoder_contest_id()

# ==============================
# Selenium获取排名
# ==============================

def get_rank_by_selenium():

    print("启动浏览器获取排名...")

    driver_path = ensure_edgedriver()

    if driver_path:
        driver = webdriver.Edge(service=Service(driver_path))
    else:
        # Selenium 4.6+ 内置管理器自动处理
        driver = webdriver.Edge()


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

def get_uid_from_profile_url(url):
    """从主页链接中提取牛客用户 ID"""
    match = re.search(
        r"/profile/(\d+)",
        url or ""
    )

    return match.group(1) if match else None


def format_rating_change(value):
    """把数值格式化成带符号的变化值，例如 +4 / -12"""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None

    if number.is_integer():
        integer = int(number)
        if integer == 0:
            return "0"
        return f"{integer:+d}"

    return f"{number:+g}"


def get_latest_rating_change(uid, scraper, headers):
    """
    通过牛客 rating-history 接口获取最近一场比赛的 Rating 变化。
    """
    if not uid:
        return None

    url = (
        "https://ac.nowcoder.com/acm/contest/"
        "rating-history"
        f"?token=&uid={uid}"
        f"&_={int(time.time() * 1000)}"
    )

    resp = scraper.get(
        url,
        headers=headers,
        timeout=20
    )

    payload = resp.json()
    history = payload.get("data") or []

    if not history:
        return None

    # 接口按时间升序返回，这里取时间最大的一条
    latest = max(
        history,
        key=lambda item: item.get("time", 0)
    )

    return format_rating_change(
        latest.get("changeValue")
    )


def parse_rating_change_from_element(element):
    """备用方案：从页面元素文本中解析 (+4) / (-12)"""
    if element is None:
        return None

    match = re.search(
        r"\(\s*([+-]?\d+)\s*\)",
        element.get_text(
            " ",
            strip=True
        )
    )

    return match.group(1) if match else None


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

        name = name_map.get(user["昵称"], "未录入该用户姓名")
        print("查询:", user["昵称"], f"({name})", end=" ")

        rating_element = None

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

            # 评级信息，兼容 a / div 两种写法
            rating_element = soup.select_one(
                "div.state-num, a.state-num"
            )

            if rating_element is None:
                rating_element = soup.select_one(
                    "[class*='state-num']"
                )

            if rating_element:

                score_element = rating_element.select_one(
                    ".rate-score2"
                )

                if score_element:
                    user["Rating"] = (
                        score_element.get_text(
                            strip=True
                        )
                    )
                else:
                    user["Rating"] = re.sub(
                        r"\([^)]*\)",
                        "",
                        rating_element.get_text(
                            " ",
                            strip=True
                        )
                    ).strip()

                if not user["Rating"]:
                    user["Rating"] = "暂无"

            else:
                user["Rating"] = "暂无"

        except Exception as e:

            print(
                "失败:",
                e
            )

            user["Rating"] = "获取失败"

        # 最近一场比赛的 Rating 变化（优先读取 rating-history 接口）
        change = None

        uid = get_uid_from_profile_url(
            user.get("主页")
        )

        try:
            change = get_latest_rating_change(
                uid,
                scraper,
                headers
            )
        except Exception as e:
            print(
                "获取 Rating 变化失败:",
                e
            )

        if change is None:
            # 备用方案：页面元素中若直接渲染了 (+4) 则解析
            change = parse_rating_change_from_element(
                rating_element
            )

        user["Rating变化"] = (
            change if change else "获取失败"
        )

        print(
            f"{user['Rating']} (变化: {user['Rating变化']})"
        )

        final_data.append(
            user
        )

        time.sleep(0.5)

    return final_data

# ==============================
# 添加成员姓名
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
                "当前Rating",
                "最近一场Rating变化"
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
