# 牛客竞赛成绩与 Rating 自动统计工具

## 项目简介

本项目用于自动获取牛客（NowCoder）ACM 比赛中指定学校成员的比赛成绩，并查询选手个人 Rating，最终生成格式化 Excel 报表。

项目采用：

- Selenium：获取动态加载的比赛排名数据
- cloudscraper：访问用户主页获取 Rating
- pandas + openpyxl：生成格式化 Excel 文件

---

# 功能特点

## 1. 自动获取比赛排名

支持获取：

- 排名
- 牛客昵称
- 用户主页
- AC数量
- 比赛总分

通过 Selenium 模拟浏览器操作，解决牛客排名页面动态加载问题。

---

## 2. 自动获取 Rating

对于排名中的每位选手：

- 自动访问个人主页
- 解析当前竞赛 Rating

相比逐个使用 Selenium 访问，速度更快。

---

## 3. 昵称映射真实姓名

支持配置：

```python
name_map = {
    "牛客昵称": "真实姓名"
}
```

生成 Excel 时自动替换显示。

---

## 4. Excel 自动格式化

输出：

```
Rating详情表.xlsx
```

包含：

|排名|姓名|昵称|AC数量|总分|Rating|
|-|-|-|-|-|-|

自动设置：

- 居中显示
- 自动换行
- 中文宽度优化
- 自适应列宽

---

# 技术栈

- Python
- Selenium
- ChromeDriver
- cloudscraper
- BeautifulSoup4
- pandas
- openpyxl

---

# 环境安装

Python：

```
Python >= 3.10
```

安装依赖：

```bash
pip install selenium cloudscraper beautifulsoup4 pandas openpyxl webdriver-manager
```

chromedriver.exe安装路径

```
https://storage.googleapis.com/chrome-for-testing-public/150.0.7871.129/win64/chrome-win64.zip
```

Ps：Google Chrome自行下载，chromedriver.exe需求的版本信息通过在Chrome的网址搜索栏中输入下列代码后自行查看

```
chrome://version/
```

---

# 使用方法

## 1. 设置比赛 ID

修改：

```python
contest_id = "133876"
```

例如：

```
https://ac.nowcoder.com/acm/contest/137561
```

其中：

```
133876
```

为比赛 ID。

---

## 2. 设置学校名称

修改：

```python
school_name = "南京邮电大学"
```

---

## 3. 配置姓名映射

示例：

```python
name_map = {
    "nju": "张三",
}
```

---

## 4. 配置 ChromeDriver

修改：

```python
Service(
    r"D:\chromedriver\chromedriver.exe"
)
```

填写本机 ChromeDriver 路径。

---

## 5. 运行

```python
main.py
```

程序流程：

```
启动浏览器
    ↓
打开牛客比赛排名
    ↓
搜索学校
    ↓
获取成员列表
    ↓
cloudscraper获取Rating
    ↓
生成Excel
```

---

# 输出示例

控制台：

```
找到17名同学

查询: nju 1600

任务完成！
文件已保存:
Rating详情表.xlsx
```

---

# 项目结构

```
NowCoder-Rating-Crawler

│
├── main.py
├── README.md
├── requirements.txt
└── chromedriver.exe
```

---

# 注意事项

## 登录

牛客排名页面可能需要登录。

第一次运行时：

1. 等待浏览器打开
2. 手动完成登录
3. 程序继续执行

---

## ChromeDriver

确保：

Chrome版本

与

ChromeDriver版本

匹配。

---

## 使用规范

本项目用于：

- ACM竞赛成绩统计
- 学校竞赛管理
- 学习研究

请遵守牛客网站相关规则。

---



# Author

徐落尘
