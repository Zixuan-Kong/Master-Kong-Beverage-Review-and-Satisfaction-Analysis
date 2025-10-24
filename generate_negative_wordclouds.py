import os
import pandas as pd
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt

CSV_FILE = '康师傅双周报数据20250718.csv'
OUT_DIR = 'plots\\wordclouds'
OUT_XLSX_NEG = 'negative_words.xlsx'
OUT_XLSX_POS = 'positive_words.xlsx'
STOPWORDS_FILE = 'stopwords.txt'
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, 'negative'), exist_ok=True)
os.makedirs(os.path.join(OUT_DIR, 'positive'), exist_ok=True)

print('读取CSV...')
df = pd.read_csv(CSV_FILE, encoding='utf-8')

# 载入或创建停用词文件
if not os.path.exists(STOPWORDS_FILE):
    default_sw = ['的','了','和','是','我','也','很','都','就','也很','还','还要','没有','感觉','有点','可以']
    with open(STOPWORDS_FILE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(default_sw))
    print('已创建默认停用词文件:', STOPWORDS_FILE)
stopwords = set()
with open(STOPWORDS_FILE, 'r', encoding='utf-8') as f:
    for line in f:
        w = line.strip()
        if w:
            stopwords.add(w)

def gen_wordcloud_and_freqs(texts, out_png_dir, out_xlsx_path, label):
    rows = []
    if not texts:
        return rows
    text = '\n'.join(texts)
    words = jieba.cut(text)
    words = [w for w in words if len(w.strip())>1 and w not in stopwords]
    if not words:
        return rows
    word_text = ' '.join(words)
    wc = WordCloud(font_path='C:\\\\Windows\\\\Fonts\\\\msyh.ttc', width=800, height=600, background_color='white', max_words=200)
    wc.generate(word_text)
    png_path = os.path.join(out_png_dir, f"{label[:50]}.png")
    wc.to_file(png_path)

    freqs = {}
    for w in words:
        freqs[w] = freqs.get(w, 0) + 1
    sorted_items = sorted(freqs.items(), key=lambda x: x[1], reverse=True)
    for word, count in sorted_items[:200]:
        rows.append({'Brand': label, 'word': word, 'count': count})
    return rows

print('开始为每个品牌生成去停用词的正/负向词云...')
neg_df = df[df['Sentiment'].astype(str).str.contains('负')].copy()
pos_df = df[df['Sentiment'].astype(str).str.contains('正')].copy()
neg_df['Content'] = neg_df['Content'].astype(str)
pos_df['Content'] = pos_df['Content'].astype(str)

brands = sorted(set(neg_df['Brand'].dropna().unique().tolist() + pos_df['Brand'].dropna().unique().tolist()))
neg_rows_all = []
pos_rows_all = []
for brand in brands:
    neg_texts = neg_df[neg_df['Brand'] == brand]['Content'].tolist()
    pos_texts = pos_df[pos_df['Brand'] == brand]['Content'].tolist()
    neg_rows = gen_wordcloud_and_freqs(neg_texts, os.path.join(OUT_DIR, 'negative'), OUT_XLSX_NEG, brand)
    pos_rows = gen_wordcloud_and_freqs(pos_texts, os.path.join(OUT_DIR, 'positive'), OUT_XLSX_POS, brand)
    neg_rows_all.extend(neg_rows)
    pos_rows_all.extend(pos_rows)

print('保存关键词到Excel...')
if neg_rows_all:
    pd.DataFrame(neg_rows_all).to_excel(OUT_XLSX_NEG, index=False)
if pos_rows_all:
    pd.DataFrame(pos_rows_all).to_excel(OUT_XLSX_POS, index=False)

print('词云生成完成，保存在', OUT_DIR)