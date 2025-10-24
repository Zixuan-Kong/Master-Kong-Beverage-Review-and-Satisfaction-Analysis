import pandas as pd
import numpy as np

# 读取CSV文件
df = pd.read_csv('康师傅双周报数据20250718.csv', encoding='utf-8')

# 打印列名和数据前几行，用于调试
print("列名:", df.columns.tolist())
print("\n前几行数据:")
print(df.head())

def calculate_satisfaction(data):
    if isinstance(data, pd.Series):
        positive = (data == '正面').sum()
        negative = (data == '负面').sum()
        if positive + negative == 0:
            return np.nan
        return ((positive - negative) / (positive + negative)) * 100
    return np.nan

# 按Brand和FirstName分组计算满意度
print("\n开始计算分组满意度...")
grouped = df.groupby(['Brand', 'FirstName'])
counts = grouped.size().reset_index(name='sample_count')
sentiments = grouped['Sentiment'].apply(calculate_satisfaction).reset_index(name='avg_sentiment')
result = pd.merge(counts, sentiments, on=['Brand', 'FirstName'])
result.columns = ['Brand_mapped', 'FirstName', 'sample_count', 'avg_sentiment']

# 计算整体品牌满意度
print("\n开始计算品牌整体满意度...")
brand_grouped = df.groupby('Brand')
brand_counts = brand_grouped.size().reset_index(name='Total_Count')
brand_sentiments = brand_grouped['Sentiment'].apply(calculate_satisfaction).reset_index(name='Overall_Satisfaction')
brand_summary = pd.merge(brand_counts, brand_sentiments, on='Brand')

# 保存结果到Excel文件
with pd.ExcelWriter('satisfaction_results.xlsx') as writer:
    result.to_excel(writer, sheet_name='Detailed Results', index=False)
    brand_summary.to_excel(writer, sheet_name='Brand Summary', index=False)

print('分析完成！结果已保存到satisfaction_results.xlsx')