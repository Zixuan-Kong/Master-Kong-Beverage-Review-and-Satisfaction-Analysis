import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

CSV_FILE = '康师傅双周报数据20250718.csv'
OUTPUT_EXCEL = 'product_reports.xlsx'
PLOTS_DIR = 'plots'

os.makedirs(PLOTS_DIR, exist_ok=True)

# 读取数据
print('读取CSV...')
df = pd.read_csv(CSV_FILE, encoding='utf-8')

# 规范化情感标签，处理可能的错别字/空值
print('规范化Sentiment列...')
df['Sentiment'] = df['Sentiment'].astype(str).str.strip()
df['Sentiment'] = df['Sentiment'].replace({'positive':'正面','negative':'负面','neutral':'中性'})
# 简单拼写修正（中文）
df['Sentiment'] = df['Sentiment'].replace({'正 面': '正面', '负 面': '负面', '中 性': '中性'})

# 确保FirstName和Brand存在
print('检查Brand和FirstName...')
df['Brand'] = df['Brand'].fillna('未知')
df['FirstName'] = df['FirstName'].fillna('未知')

# 分组计算：每个Brand下，每个FirstName的正/中/负数量、样本数、满意度
print('计算统计量...')
summary_rows = []
brands = sorted(df['Brand'].unique())
for brand in brands:
    sub = df[df['Brand'] == brand]
    grp = sub.groupby('FirstName')['Sentiment'].value_counts().unstack(fill_value=0)
    # 确保存在三列
    for col in ['正面','中性','负面']:
        if col not in grp.columns:
            grp[col] = 0
    grp = grp[['正面','中性','负面']]
    grp['sample_count'] = grp.sum(axis=1)
    # 计算满意度
    grp['avg_sentiment'] = ((grp['正面'] - grp['负面']) / (grp['正面'] + grp['负面']).replace(0, np.nan)) * 100
    grp = grp.reset_index()
    for _, r in grp.iterrows():
        summary_rows.append({
            'Brand_mapped': brand,
            'FirstName': r['FirstName'],
            'positive': int(r['正面']),
            'neutral': int(r['中性']),
            'negative': int(r['负面']),
            'sample_count': int(r['sample_count']),
            'avg_sentiment': None if pd.isna(r['avg_sentiment']) else float(r['avg_sentiment'])
        })

summary_df = pd.DataFrame(summary_rows)

# 将每个Brand写入Excel的单独sheet
print('写入Excel...')
with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    for brand in brands:
        sheet_df = summary_df[summary_df['Brand_mapped'] == brand].copy()
        if sheet_df.empty:
            continue
        sheet_df.to_excel(writer, sheet_name=brand[:31], index=False)

# 生成每个Brand的图表并保存
print('绘图并保存PNG...')
for brand in brands:
    sheet_df = summary_df[summary_df['Brand_mapped'] == brand].copy()
    if sheet_df.empty:
        continue
    # 仅取样本数>0的FirstName做展示，按样本数排序
    plot_df = sheet_df[sheet_df['sample_count']>0].sort_values('sample_count', ascending=True)
    if plot_df.empty:
        continue
    y = plot_df['FirstName']
    pos = plot_df['positive']
    neu = plot_df['neutral']
    neg = plot_df['negative']

    fig, ax = plt.subplots(figsize=(8, max(4, 0.35*len(y))))
    ax.barh(y, pos, color='#5DA5FF', label='正面')
    ax.barh(y, neu, left=pos, color='#C0C0C0', label='中性')
    ax.barh(y, neg, left=pos+neu, color='#FF8A8A', label='负面')

    # 在每段中间添加数值标签（正面和负面）
    for i, (p, n) in enumerate(zip(pos, neg)):
        ax.text(p/2, i, str(p), va='center', ha='center', color='white', fontsize=8)
        ax.text(p+neu.iloc[i]+n/2, i, str(n), va='center', ha='center', color='white', fontsize=8)
    # 在正面-负面旁显示样本总数
    for i, sc in enumerate(plot_df['sample_count']):
        ax.text(plot_df['positive'].iloc[i]+plot_df['neutral'].iloc[i]+plot_df['negative'].iloc[i]+1, i, str(sc), va='center', ha='left', fontsize=8)

    ax.set_title(f'{brand} - 评价分布')
    ax.set_xlabel('评价数量')
    ax.legend(loc='lower right')
    plt.tight_layout()
    png_path = os.path.join(PLOTS_DIR, f"{brand[:50]}.png")
    fig.savefig(png_path, dpi=150)
    plt.close(fig)

print('全部完成。生成文件：', OUTPUT_EXCEL, '和目录：', PLOTS_DIR)