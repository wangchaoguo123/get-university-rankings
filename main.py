from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import shutil
import os
import pandas as pd
import bs4
import openpyxl
from datetime import datetime
import time
from dataclasses import dataclass
import logging
import os
from logging.handlers import RotatingFileHandler
from datetime import datetime as _dt

# 日志配置：写入 logs/ 目录，文件名为 Logging_YYYYMMDD.log，超过 5MB 时轮转
log_level_name = os.getenv('LOG_LEVEL', 'INFO')
log_level = getattr(logging, log_level_name.upper(), logging.INFO)

base_dir = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(base_dir, 'logs')
os.makedirs(log_dir, exist_ok=True)
date_str = _dt.now().strftime('%Y%m%d')
LOG_FILENAME = f'Logging_{date_str}.log'
log_path = os.path.join(log_dir, LOG_FILENAME)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)  # 根 logger 接收所有级别，由 handler 过滤
# 清理可能存在的默认处理器
for h in list(root_logger.handlers):
    root_logger.removeHandler(h)

# 当日志文件超过 5MB 时轮转，新文件会以 .1 .2 ... 形式保存，保留最近 5 个备份
max_bytes = 5 * 1024 * 1024
backup_count = 5


# 更简洁的文件日志格式（去掉模块名，便于快速查看）
formatter = logging.Formatter('%(asctime)s %(levelname)s: %(message)s')

# 单一日志文件，包含所有级别
all_path = os.path.join(log_dir, f'Logging_{date_str}.log')
file_handler = RotatingFileHandler(all_path, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8')
file_handler.setFormatter(formatter)
# 只将关键性日志写入文件（INFO 及以上），避免过多 DEBUG 细节
file_handler.setLevel(logging.INFO)
root_logger.addHandler(file_handler)

# 控制台输出级别由环境变量 LOG_LEVEL 控制（默认 WARNING）
console_level = getattr(logging, os.getenv('LOG_LEVEL', 'WARNING').upper(), logging.WARNING)
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
console_handler.setLevel(console_level)
root_logger.addHandler(console_handler)

# 尝试使用 Edge 浏览器（Chromium），失败时回退到 Chrome
def _start_browser():
    # 优先 Edge
    try:
        edge_options = webdriver.EdgeOptions()
        # 尝试设置 detach（Chromium-based browsers 可能支持）
        try:
            edge_options.add_experimental_option("detach", True)
        except Exception:
            pass
        logger = logging.getLogger(__name__)
        logger.info("尝试启动 Edge 浏览器...")
        return webdriver.Edge(options=edge_options)
    except Exception as e:
        logger = logging.getLogger(__name__)
        logger.warning(f"启动 Edge 失败，回退到 Chrome: {e}")
        # 回退 Chrome
        chrome_options = webdriver.ChromeOptions()
        try:
            chrome_options.add_experimental_option("detach", True)
        except Exception:
            pass
        return webdriver.Chrome(options=chrome_options)

url_template = "https://www.shanghairanking.cn/rankings/{rankType}/{year}"
# 启动浏览器 bcur
browser = _start_browser()
# 浏览器最大化（可选）
# browser.maximize_window()

# 先创建一个数组，保存结果
contents = []

@dataclass
class RankSchool:
    Name: str
    shortName: str

# 创建结构体数组
rankSchools = [
    RankSchool("中国大学排名", "bcur"),
    RankSchool("中国高职院校排名", "bcvcr"),
    RankSchool("世界大学学术排名", "arwu"),
    RankSchool("中国最好学科排名", "bcsr"),
    RankSchool("中国大学专业排名", "bcmr"),
    RankSchool("世界一流学科排名", "gras"),
    RankSchool("全球体育类院系学术排名", "grsssd"),
]

# 检索数据取出放进数组里面
def get_data(page, rank_type=None):
    # 获取全部网页信息
    html = browser.page_source
    soup = BeautifulSoup(html, "html.parser")
    # 找到tbody的子节点
    tbody = soup.find('tbody')
    if not tbody:
        logging.getLogger(__name__).warning("页面中未找到 tbody，可能页面结构已变化")
        return
    # 预定义列模式（用于构建行时决定列数与列名）
    schema_map = {
        'bcur': ["排名", "学校名称", "省市", "类型", "总分", "办学层次"],
        'bcvcr': ["排名", "学校名称", "省市", "总分"],
        'arwu': ["排名", "学校名称", "国家/地区", "国家/地区排名", "总分", "校友获奖"],
    }

    for tr in tbody.find_all('tr', recursive=False):
        # 判断tr是否属于子节点中
        if not isinstance(tr, bs4.element.Tag):
            continue
        tds = tr.find_all('td')
        try:
            # 获取文本，清除换行符，清空空格，容错多个可能的类名
            name_tag = tr.find(class_="name-cn") or tr.find(class_="name") or tr.find('a')
            name = name_tag.get_text(strip=True) if name_tag else ""
            texts = [td.get_text(strip=True) for td in tds]

            # 决定使用的列名
            if rank_type and rank_type in schema_map:
                cols = schema_map[rank_type]
            else:
                # 根据 tds 长度做简单推断
                if len(texts) == 6:
                    cols = schema_map['bcur']
                elif len(texts) == 4:
                    cols = schema_map['bcvcr']
                else:
                    cols = [f"col_{i}" for i in range(len(texts))]

            # 构建行：优先使用 name_tag 填充学校名称列
            row = []
            for i, colname in enumerate(cols):
                if any(k in colname for k in ("学校", "名称", "name")):
                    # 尽量使用明确的 name_tag，否则回退到文本序列中的常见位置
                    if name:
                        v = name
                    else:
                        v = texts[i] if i < len(texts) else ""
                else:
                    v = texts[i] if i < len(texts) else ""
                row.append(v)

            contents.append(row)
        except Exception as e:
            logging.getLogger(__name__).warning(f"跳过一行（第{page}页）解析失败: {e}")
    logging.getLogger(__name__).info("开始爬取第 %s 页，总计获取到 %s 个学校排名", page, len(contents))

def get_all():
    # 保留向后兼容的无参调用（默认当前打开页面）
    page = 1
    while page <= 20:  # 循环页数
        # 默认不指定 rank_type（会在 get_data 内基于列数推断），或者外部改为传入
        try:
            # 若全局有 rank_type 变量则传入，否则不传
            get_data(page, globals().get('current_rank_type'))
        except TypeError:
            get_data(page)
        try:
            next_page = browser.find_element(By.CSS_SELECTOR, 'li.ant-pagination-next>a')
            # 若下一页不存在或被禁用则停止
            cls = next_page.get_attribute('class') or ''
            if 'disabled' in cls:
                break
            next_page.click()
            time.sleep(1)
            page += 1
        except Exception:
            break

def get_all_for_year(rankType: str, year: int):
    global contents
    # 暴露当前 rank_type 以便 get_all/get_data 使用
    globals()['current_rank_type'] = rankType
    contents = []
    target = url_template.format(rankType=rankType, year=year)
    logging.getLogger(__name__).info(f"访问 {target} ...")
    browser.get(target)
    browser.implicitly_wait(5)
    get_all()

def build_dataframe(data, rank_type=None, columns=None):
    """
    将抓取的数据转换为 DataFrame，并根据排名类型自动选择列名与类型转换。

    参数:
      - data: list[list] | list[dict] | pd.DataFrame
      - rank_type: 可选，传入短名（如 'bcur','bcvcr','arwu'）以选择预定义列名
      - columns: 可选，显式列名列表，优先于 rank_type
    """
    # 预定义的列映射（短名 -> 列名）
    schema_map = {
        'bcur': ["排名", "学校名称", "省市", "类型", "总分", "办学层次"],
        'bcvcr': ["排名", "学校名称", "省市", "总分"],
        'arwu': ["排名", "学校名称", "国家/地区", "国家/地区排名", "总分", "校友获奖"],
    }

    # 已经是 DataFrame，直接复制并尽量处理列
    if isinstance(data, pd.DataFrame):
        df = data.copy()
    else:
        # 空数据
        if not data:
            df = pd.DataFrame(columns=columns if columns else [])
        else:
            # 列表中的元素为 dict
            if isinstance(data[0], dict):
                df = pd.DataFrame(data)
            else:
                # 行列表，确定最大列数
                maxcols = max((len(row) for row in data), default=0)
                if columns:
                    cols = columns
                elif rank_type and rank_type in schema_map:
                    cols = schema_map[rank_type]
                elif maxcols == 6:
                    cols = schema_map['bcur']
                elif maxcols == 4:
                    cols = schema_map['bcvcr']
                else:
                    cols = [f"col_{i}" for i in range(maxcols)]
                df = pd.DataFrame(data, columns=cols)

    # 清理字符串两端空白（兼容缺少 DataFrame.applymap 的环境）
    for col in df.columns:
        try:
            if df[col].dtype == object:
                df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
        except Exception:
            df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    # 智能数值转换：根据列名关键词进行转换
    def _try_numeric(colname, series):
        name = str(colname)
        if "排名" in name or "名次" in name or name.lower().startswith("rank"):
            return pd.to_numeric(series, errors='coerce').fillna(0).astype(int)
        if any(k in name for k in ("总分", "分", "score", "得分")):
            return pd.to_numeric(series, errors='coerce')
        if any(k in name for k in ("层次", "level")):
            return pd.to_numeric(series, errors='coerce')
        return series

    for c in df.columns:
        try:
            df[c] = _try_numeric(c, df[c])
        except Exception:
            # 转换失败则保持原值
            pass

    return df

def save_dataframe_to_file(df: pd.DataFrame, filename: str):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base_dir, "output")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, filename)
    df.to_excel(path, index=False)
    logging.getLogger(__name__).info(f"已保存: {path}")

if __name__ == '__main__':
    # 获取用户输入的排名类型
    inputData = input(
        """请输入排名类型：
        1.中国大学排名
        2.中国高职院校排名
        3.世界大学学术排名
        4.中国最好学科排名(未使用)
        5.中国大学专业排名(未使用)
        6.世界一流学科排名(未使用)
        7.全球体育类院系学术排名(未使用)
        请输入对应数字（默认1）: """
    ).strip()

    if inputData not in {"1", "2", "3", "4", "5", "6", "7"}:
        logging.getLogger(__name__).warning("输入无效，默认使用中国大学排名")
        inputData = "1"
    
    rankType = rankSchools[int(inputData) - 1].shortName
    logging.getLogger(__name__).info(f"已选择排名类型: {rankSchools[int(inputData) - 1].Name}")
    # 最近三年（包含当前年）
    current_year = datetime.now().year
    years = [current_year - i for i in range(0, 3)]
    year_dfs = {}
    for y in years:
        logging.getLogger(__name__).info(f"开始爬取 {y} 年排名...")
        get_all_for_year(rankType, y)
        df = build_dataframe(contents, rank_type=rankType)
        year_dfs[y] = df
        # save_dataframe_to_file(df, f"{y}_{rankSchools[int(inputData) - 1].Name}.xlsx")

    # 写入同一 Excel，不同 sheet
    combined_name = f"{rankSchools[int(inputData) - 1].Name}_{years[-1]}-{years[0]}.xlsx"
    base_dir = os.path.dirname(os.path.abspath(__file__))
    combined_path = os.path.join(base_dir, "output", combined_name)
    with pd.ExcelWriter(combined_path) as writer:
        for y in years:
            sheet = str(y)
            year_dfs[y].to_excel(writer, sheet_name=sheet, index=False)
    logging.getLogger(__name__).info(f"已将近3年排名写入 {combined_path}")

    # 关闭浏览器
    browser.quit()