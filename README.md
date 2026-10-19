# 牛客比赛排名与 Rating 查询工具


一个基于 Python 的牛客（NowCoder）比赛数据查询脚本，可自动获取最近结束的指定类型比赛，抓取排名列表，并查询参赛用户的当前 Rating 和最近一场比赛的 Rating 变化，最后将整理后的数据导出为 Excel。

## 功能特性

- **自动选择比赛**：未手动指定 `contest_id` 时，自动从牛客比赛列表中查找最近结束的周赛、小白月赛、挑战赛或练习赛。
- **指定比赛查询**：可在脚本中设置比赛 ID，查询对应比赛的排名。
- **按关键词搜索**：通过牛客排名页的用户搜索框输入学校名称或其他关键词，并筛选匹配的排名行。
- **获取 Rating**：抓取用户主页中的当前 Rating。
- **获取 Rating 变化**：优先通过 Rating 历史接口查询最近一场比赛的变化值，并在必要时尝试从页面文本中解析。
- **昵称映射真实姓名**：可维护昵称与姓名的映射表，便于输出队伍或学校内部统计表。
- **导出 Excel**：生成 `Rating详情.xlsx`，包含排名、姓名、昵称、AC 数量、总分、当前 Rating 和最近一场 Rating 变化，并自动调整列宽、行高和单元格对齐方式。
- **EdgeDriver 版本处理**：检测本地 Microsoft Edge 版本，尝试下载匹配主版本的 EdgeDriver；也支持使用已有的驱动路径。

## 运行流程

1. 检测或下载与本机 Edge 浏览器版本匹配的 EdgeDriver。
2. 获取指定比赛 ID；如果没有指定，则从牛客比赛列表中自动选择最近结束的目标比赛。
3. 使用 Selenium 打开比赛排名页面，并按配置的关键词搜索。
4. 提取匹配用户的排名、昵称、主页、AC 数量和总分。
5. 使用 `cloudscraper` 和 `BeautifulSoup` 查询用户当前 Rating，并查询最近一场比赛的 Rating 变化。
6. 根据昵称映射表补充姓名。
7. 将结果保存到当前工作目录下的 `Rating详情.xlsx`。

## 环境要求

- Windows（脚本中的 Edge 版本检测使用 Windows 注册表；EdgeDriver 下载逻辑按 Windows 平台构建）
- Python 3.9 或更高版本（建议使用较新的 Python 3）
- Microsoft Edge 浏览器
- 可访问牛客网站及 Microsoft EdgeDriver 下载服务的网络环境

## 安装依赖

在项目目录打开 PowerShell 或终端，执行：

```bash
pip install selenium cloudscraper beautifulsoup4 pandas openpyxl
```

脚本还使用 Python 标准库中的模块，无需单独安装，包括 `time`、`re`、`json`、`html`、`os`、`atexit`、`platform`、`shutil`、`subprocess`、`tempfile`、`urllib` 和 `zipfile`。

## 配置说明

打开脚本，在配置区域修改以下变量：

```python
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
```

### 1. EdgeDriver 路径 `path`

- **留空**：脚本尝试自动检测本机 Edge 版本，并下载对应的 EdgeDriver 到临时目录供本次运行使用；程序退出时会清理临时文件。
- **填写路径**：例如 `r"C:\WebDriver\msedgedriver.exe"`。如果驱动不存在或主版本号不匹配，脚本会尝试下载并替换到该路径。
- 如果自动下载失败，脚本会尝试使用已存在的驱动；若未配置可用驱动，则会回退到 Selenium 自带的驱动管理机制。

### 2. 比赛 ID `contest_id`

如果需要查询指定比赛，可以在脚本中取消相应注释并填写比赛 ID，例如：

```python
contest_id = "130107"
```

如果没有定义 `contest_id`，脚本会自动从牛客比赛列表中选择最近结束、且名称包含以下关键词之一的比赛：

- 周赛
- 小白月赛
- 挑战赛
- 练习赛

### 3. 学校名称 `school_name`

该字段用于在牛客比赛排名页面的搜索框中输入关键词，并筛选表格中包含该关键词的行。请根据牛客页面实际显示的内容填写。若关键词不匹配，脚本可能无法提取到用户。

### 4. 昵称与姓名映射 `name_map`

将牛客昵称映射到真实姓名，格式如下：

```python
name_map = {
    "昵称A": "张三",
    "昵称B": "李四",
}
```

未录入映射的昵称会使用默认提示文本。请根据自己的实际名单维护此字典。

## 运行方法

假设脚本文件名为 `main.py`，在脚本所在目录执行：

```bash
python main.py
```

运行结束后，如果成功获取到数据，会在当前工作目录生成：

```text
Rating详情.xlsx
```

Excel 表格包含以下字段：

| 字段 | 含义 |
| --- | --- |
| 排名 | 用户在所查询比赛中的排名 |
| 姓名 | 根据 `name_map` 映射得到的姓名 |
| 昵称 | 牛客昵称 |
| AC数量 | 排名页面显示的 AC 数量 |
| 总分 | 排名页面显示的比赛总分 |
| 当前Rating | 用户主页显示的当前 Rating |
| 最近一场Rating变化 | 最近一场比赛的 Rating 变化值，例如 `+4`、`-12` 或 `0` |

## 常见问题与排查

### 1. 没有找到用户

- 检查 `school_name` 是否与牛客排名页面的搜索内容匹配。
- 确认该比赛排名页存在对应数据。
- 检查牛客页面结构是否发生变化。

### 2. EdgeDriver 下载失败或浏览器无法启动

- 确认 Microsoft Edge 已安装且可以正常启动。
- 检查网络是否可以访问 EdgeDriver 下载地址。
- 可以手动下载与 Edge 主版本号匹配的 EdgeDriver，并将 `path` 设置为 `msedgedriver.exe` 的完整路径。
- 如果脚本仍无法启动浏览器，请检查 Selenium 版本和浏览器版本。

### 3. Rating 显示“暂无”或“获取失败”

- 用户主页可能没有公开显示 Rating。
- 牛客页面结构或接口可能发生变化。
- 网络请求可能超时，或网站限制了请求。
- 脚本会尝试通过 Rating 历史接口读取最近一场比赛的变化；如果接口不可用，也会尝试从主页文本中解析。

### 4. Excel 没有生成

脚本只有在成功提取到用户数据时才会保存 Excel。若未提取到数据，会输出提示信息。请优先检查搜索关键词、比赛 ID 和页面数据是否正常。

## 注意事项

- 本项目依赖牛客网站当前的页面结构和接口；网站改版后，CSS 选择器或数据解析逻辑可能需要更新。
- 脚本会访问牛客比赛列表、比赛排名页和用户主页，请确保网络连接正常，并合理控制请求频率。
- 请遵守牛客网站的使用规则及相关服务条款，不要将脚本用于未经授权的数据收集。
- 用户昵称与真实姓名映射由使用者自行维护；提交到公开仓库前，请检查并移除不应公开的个人信息。
- 自动下载 EdgeDriver 需要访问 Microsoft 提供的下载服务；在受限网络环境下，建议手动配置驱动路径。

## 技术栈

- [Python](https://www.python.org/)
- [Selenium](https://www.selenium.dev/)
- [Cloudscraper](https://github.com/VeNoMouS/cloudscraper)
- [Beautiful Soup 4](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/)

