from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time


def get_detailed_rating():
    print(1)
    # 1. 启动浏览器
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    print(2)
    wait = WebDriverWait(driver, 20)

    # 2. 登录与搜索
    contest_id = "130107"  # 你的比赛ID
    driver.get(f"https://ac.nowcoder.com/acm/contest/{contest_id}#rank")

    print("⚠️ 请在弹出的浏览器中完成登录，并确保排名列表已加载...")
    # 如果 15 秒不够登录，请自行调大这个数字
    time.sleep(15)

    # 自动搜索学校
    try:
        search_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='输入用户名称']")))
        search_box.clear()
        search_box.send_keys("江苏第二师范学院")
        time.sleep(3)
        search_box.send_keys(Keys.ENTER)  # 💡 关键：模拟按下回车键
        print("🔍 已提交搜索，等待列表刷新...")
        time.sleep(5)  # 等待搜索结果加载完成
    except Exception as e:
        print("搜索框加载失败，请手动在浏览器搜索‘江苏第二师范学院’。")

    # 3. 第一步：抓取列表中的基本信息和主页 URL
    user_list = []
    # 重新获取搜索后的行
    rows = driver.find_elements(By.CSS_SELECTOR, "tr")
    for row in rows:
        # 只处理包含学校名称的有效行
        if "江苏第二师范学院" in row.text:
            try:
                cols = row.find_elements(By.TAG_NAME, "td")
                # 精确抓取第二列中的 <a> 标签（昵称和链接）
                name_element = cols[1].find_element(By.TAG_NAME, "a")
                user_list.append({
                    "排名": cols[0].text.strip(),
                    "昵称": name_element.text.strip(),
                    "主页": name_element.get_attribute("href"),
                    "AC数量": cols[3].text.strip(),
                    "总分": cols[4].text.strip()
                })
            except:
                continue

    print(f"📊 找到 {len(user_list)} 名同学，准备进入详情页抓取 Rating...")

    # 4. 第二步：穿梭进入每个人的主页抓取 Rating
    final_data = []
    for user in user_list:
        print(f"正在查询: {user['昵称']} ...", end="")
        driver.get(user['主页'])

        try:
            # 💡 核心修改：使用 CSS 选择器匹配包含 'rate-score' 的 a 标签
            # 这种方式可以兼容 rate-score1, rate-score2 等所有颜色的分数
            time.sleep(2)  # 给页面一点点加载 JS 的时间
            rating_selector = "a[class*='state-num']"
            rating_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, rating_selector)))

            user['Rating'] = rating_element.text.strip()
        except:
            # 备选方案：如果上面的没找到，尝试通过父级 .state-num 定位
            try:
                user['Rating'] = driver.find_element(By.CSS_SELECTOR, ".state-num a").text.strip()
            except:
                user['Rating'] = "暂无(未参加周赛)"

        print(f" 结果: {user['Rating']}")
        final_data.append(user)
        time.sleep(1)  # 访问频率控制，保护账号

    # 5. 保存结果
    if final_data:
        df = pd.DataFrame(final_data)
        # 调整列顺序
        df = df[["排名", "昵称", "AC数量", "总分", "Rating"]]
        output_file = "江苏第二师范学院_Rating详情表.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\n🎉 任务完成！文件已保存为: {output_file}")
    else:
        print("\n❌ 未能提取到任何数据，请检查登录状态或搜索词。")

    driver.quit()


if __name__ == "__main__":
    get_detailed_rating()
