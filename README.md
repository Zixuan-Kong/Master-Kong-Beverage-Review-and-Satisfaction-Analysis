项目说明（中文）

这个项目用于分析“康师傅双周报数据20250718.csv”中的情感（正面/负面/中性）并输出以下产物：

主要脚本
- `calculate_satisfaction.py`  
  计算各 Brand + FirstName 维度的样本数与满意度（满意度按 (正面-负面)/(正面+负面) *100% 计算），并输出 `satisfaction_results.xlsx`。

- `generate_product_reports.py`  
  为每个 Brand 生成明细表（每个 Brand 一个 Excel sheet）并绘制每个 Brand 的评价分布图，输出：
  - `product_reports.xlsx`（每个 Brand 一个 sheet）
  - `plots/*.png`（各 Brand 的堆积条形图）

- `generate_negative_wordclouds.py`（现已扩展为同时生成正/负词云）
  - 为每个 Brand 生成“去停用词”后的负面/正面词云图片，保存到 `plots/wordclouds/negative/` 和 `plots/wordclouds/positive/`。
  - 输出关键词表 `negative_words.xlsx` 和 `positive_words.xlsx`（每行: Brand, word, count）。
  - 停用词文件：`stopwords.txt`（脚本会在缺失时生成一个基础停用词列表，建议按需编辑并补充）。

依赖（requirements）
- pandas
- numpy
- matplotlib
- openpyxl
- jieba
- wordcloud

你可以用下面的 `requirements.txt`（项目已包含）安装依赖：

PowerShell 示例（若系统 PATH 中已配置 python，请使用 `python`）：

```powershell
# 使用系统 Python
python -m pip install -r requirements.txt
# 或使用完整 Python 路径（示例为本机配置）
C:/Users/86136/AppData/Local/Programs/Python/Python312/python.exe -m pip install -r requirements.txt
```

运行脚本

1) 计算满意度并生成基本 Excel（satisfaction）

```powershell
# 在工作目录（含 CSV 的文件夹）运行
python calculate_satisfaction.py
# 或使用完整路径
C:/Users/86136/AppData/Local/Programs/Python/Python312/python.exe calculate_satisfaction.py
```

2) 生成每个 Brand 的报告与图

```powershell
python generate_product_reports.py
```

3) 生成正/负词云（并使用 `stopwords.txt`）

```powershell
python generate_negative_wordclouds.py
```

输出路径（在工作目录）
- `satisfaction_results.xlsx` — 满意度明细与品牌汇总（由 `calculate_satisfaction.py` 生成）
- `product_reports.xlsx` — 每个 Brand 的 sheet（由 `generate_product_reports.py` 生成）
- `plots/` — 各 Brand 的分布图（png）
- `plots/wordclouds/negative/` — 负面词云 png
- `plots/wordclouds/positive/` — 正面词云 png
- `negative_words.xlsx`, `positive_words.xlsx` — 词频表
- `stopwords.txt` — 可编辑的停用词列表

字体与中文显示说明
- matplotlib/wordcloud 在默认配置下可能会因为缺少中文字体而在图片中出现方块或缺字。脚本中已默认使用系统常见字体路径 `C:\Windows\Fonts\msyh.ttc`（微软雅黑）。
- 如果您系统中没有该字体，请将脚本中的 `font_path` 修改为可用的中文字体文件（例如 `simsun.ttc`, `msyh.ttc`, `simhei.ttf` 等），或将字体文件复制到该路径。

常见问题与排查
- 如果运行时报错提示 `ModuleNotFoundError`，请先安装依赖（见上方 `pip install -r requirements.txt`）。
- 如果 matplotlib 保存图片时提示字体缺失（console 会有 "Glyph missing" 警告），可以：
  1. 在 Windows 的 `C:\Windows\Fonts` 找到合适字体，修改脚本中 `font_path` 到该文件；
  2. 或在脚本中加入：
     ```python
     import matplotlib
     matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei']
     matplotlib.rcParams['axes.unicode_minus'] = False
     ```

可选的改进/下一步建议
- 扩充 `stopwords.txt`（我可以替您附上一个更完整的中文停用词列表）。
- 把每个 Brand 的 top N 关键词自动按主题（口感/价格/物流/包装/服务）归类并生成汇总表。
- 生成 PPT 报告或单页 PDF 汇总（包含若干 Brand 的图和结论）。

如果您希望我：
- 帮您扩展 `stopwords.txt` 并重新生成词云；
- 或把 top 10 负面关键词归类并加入 `product_reports.xlsx` 的 sheet 中；
告知我您选择的项，我会继续执行并把结果放在工作区。