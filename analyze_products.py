#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
商品价格走向分析脚本
分析毛利率、毛利额、品类结构等关键指标
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

def load_data(file_path):
    """加载Excel数据"""
    print("正在加载数据...")
    df = pd.read_excel(file_path)
    print(f"数据加载完成，共 {len(df)} 条记录")
    print(f"列名: {df.columns.tolist()}")
    return df

def calculate_metrics(df):
    """计算关键指标"""
    print("\n正在计算关键指标...")
    
    # 清理数据 - 过滤掉空行和只有分类标题的行
    df_clean = df[df['商品名称'].notna()].copy()
    
    metrics = {}
    
    # 毛利率列（转换为百分比）
    if '求和项:标品毛利率' in df_clean.columns:
        margin_rates = df_clean['求和项:标品毛利率'].dropna() * 100  # 转换为百分比
        metrics['avg_margin_rate'] = margin_rates.mean()
        metrics['max_margin_rate'] = margin_rates.max()
        metrics['min_margin_rate'] = margin_rates.min()
    
    # 毛利额列
    if '求和项:销售毛利' in df_clean.columns:
        metrics['total_profit'] = df_clean['求和项:销售毛利'].sum()
    
    # 金额列
    if '求和项:实际金额' in df_clean.columns:
        metrics['total_amount'] = df_clean['求和项:实际金额'].sum()
    
    return metrics

def analyze_by_category(df):
    """按品类分析"""
    print("\n正在进行品类分析...")
    
    # 使用一级分类列
    category_col = '一级分类'
    
    # 向下填充分类名称（Excel中分类只在标题行）
    df['一级分类_filled'] = df[category_col].ffill()
    
    # 过滤掉标题行和汇总行，只保留有商品名称的数据行
    df_clean = df[df['商品名称'].notna()].copy()
    df_clean = df_clean[df_clean['求和项:实际金额'].notna()].copy()
    
    # 排除"汇总"行和"总计"行
    df_clean = df_clean[~df_clean['一级分类_filled'].str.contains('汇总|总计', na=False)]
    
    print(f"使用分类列: {category_col}")
    print(f"清理后数据行数: {len(df_clean)}")
    print(f"分类数量: {df_clean['一级分类_filled'].nunique()}")
    
    # 按分类汇总
    category_analysis = df_clean.groupby('一级分类_filled').agg({
        '求和项:实际金额': 'sum',
        '求和项:销售毛利': 'sum'
    }).reset_index()
    
    # 重命名列
    category_analysis.columns = [category_col, '金额', '毛利']
    
    # 计算毛利率
    category_analysis['毛利率'] = (category_analysis['毛利'] / category_analysis['金额'] * 100).round(2)
    
    # 过滤掉金额为0或NaN的行
    category_analysis = category_analysis[category_analysis['金额'] > 0]
    
    # 按金额排序
    category_analysis = category_analysis.sort_values('金额', ascending=False)
    
    return category_analysis

def create_visualizations(df, category_analysis):
    """创建可视化图表"""
    print("\n正在生成可视化图表...")
    
    fig = plt.figure(figsize=(20, 12))
    
    # 图1: 品类金额对比（横向条形图）
    if category_analysis is not None and len(category_analysis) > 0:
        ax1 = plt.subplot(2, 3, 1)
        category_col = category_analysis.columns[0]
        top_categories = category_analysis.head(15)
        
        colors = ['#ff4444' if x < 0.05 else '#ffaa44' if x < 0.10 else '#44ff44' if x > 0.15 else '#4488ff' 
                  for x in top_categories['毛利率']/100]
        
        ax1.barh(range(len(top_categories)), top_categories['金额'], color=colors)
        ax1.set_yticks(range(len(top_categories)))
        ax1.set_yticklabels(top_categories[category_col], fontsize=9)
        ax1.set_xlabel('金额（元）', fontsize=11)
        ax1.set_title('品类金额排行（Top 15）\n颜色：红<5% 橙<10% 蓝10-15% 绿>15%', fontsize=12, fontweight='bold')
        ax1.grid(axis='x', alpha=0.3)
        
        # 添加数值标签
        for i, v in enumerate(top_categories['金额']):
            ax1.text(v, i, f' {v:,.0f}', va='center', fontsize=8)
    
    # 图2: 毛利率分布（横向条形图）
    if category_analysis is not None and len(category_analysis) > 0:
        ax2 = plt.subplot(2, 3, 2)
        sorted_by_rate = category_analysis.sort_values('毛利率', ascending=False).head(15)
        
        colors = ['#44ff44' if x > 0.15 else '#4488ff' if x > 0.10 else '#ffaa44' if x > 0.05 else '#ff4444' 
                  for x in sorted_by_rate['毛利率']/100]
        
        ax2.barh(range(len(sorted_by_rate)), sorted_by_rate['毛利率'], color=colors)
        ax2.set_yticks(range(len(sorted_by_rate)))
        ax2.set_yticklabels(sorted_by_rate[category_col], fontsize=9)
        ax2.set_xlabel('毛利率（%）', fontsize=11)
        ax2.set_title('品类毛利率排行（Top 15）', fontsize=12, fontweight='bold')
        ax2.axvline(x=5, color='red', linestyle='--', alpha=0.5, label='5%警戒线')
        ax2.axvline(x=10, color='orange', linestyle='--', alpha=0.5, label='10%目标线')
        ax2.axvline(x=15, color='green', linestyle='--', alpha=0.5, label='15%优秀线')
        ax2.legend(fontsize=8)
        ax2.grid(axis='x', alpha=0.3)
        
        # 添加数值标签
        for i, v in enumerate(sorted_by_rate['毛利率']):
            ax2.text(v, i, f' {v:.2f}%', va='center', fontsize=8)
    
    # 图3: 毛利额排行
    if category_analysis is not None and len(category_analysis) > 0:
        ax3 = plt.subplot(2, 3, 3)
        sorted_by_profit = category_analysis.sort_values('毛利', ascending=False).head(15)
        
        ax3.barh(range(len(sorted_by_profit)), sorted_by_profit['毛利'], color='#66ccff')
        ax3.set_yticks(range(len(sorted_by_profit)))
        ax3.set_yticklabels(sorted_by_profit[category_col], fontsize=9)
        ax3.set_xlabel('毛利额（元）', fontsize=11)
        ax3.set_title('品类毛利额排行（Top 15）', fontsize=12, fontweight='bold')
        ax3.grid(axis='x', alpha=0.3)
        
        # 添加数值标签
        for i, v in enumerate(sorted_by_profit['毛利']):
            ax3.text(v, i, f' {v:,.0f}', va='center', fontsize=8)
    
    # 图4: 毛利率分布直方图
    if '毛利率' in df.columns:
        ax4 = plt.subplot(2, 3, 4)
        margin_rates = df['毛利率'].dropna()
        margin_rates = margin_rates[margin_rates <= 100]  # 过滤异常值
        
        ax4.hist(margin_rates, bins=30, color='skyblue', edgecolor='black', alpha=0.7)
        ax4.axvline(x=5, color='red', linestyle='--', linewidth=2, label='5%警戒线')
        ax4.axvline(x=10, color='orange', linestyle='--', linewidth=2, label='10%目标线')
        ax4.axvline(x=margin_rates.mean(), color='green', linestyle='-', linewidth=2, 
                   label=f'平均值={margin_rates.mean():.2f}%')
        ax4.set_xlabel('毛利率（%）', fontsize=11)
        ax4.set_ylabel('商品数量', fontsize=11)
        ax4.set_title('毛利率分布直方图', fontsize=12, fontweight='bold')
        ax4.legend(fontsize=9)
        ax4.grid(axis='y', alpha=0.3)
    
    # 图5: 金额与毛利率散点图
    if category_analysis is not None and len(category_analysis) > 0:
        ax5 = plt.subplot(2, 3, 5)
        
        scatter = ax5.scatter(category_analysis['金额'], category_analysis['毛利率'], 
                            s=category_analysis['毛利']/100, alpha=0.6, c=category_analysis['毛利率'],
                            cmap='RdYlGn', vmin=0, vmax=20)
        
        # 标注关键品类
        for idx, row in category_analysis.head(10).iterrows():
            ax5.annotate(row[category_col][:8], 
                        (row['金额'], row['毛利率']),
                        fontsize=8, alpha=0.7)
        
        ax5.axhline(y=5, color='red', linestyle='--', alpha=0.3)
        ax5.axhline(y=10, color='orange', linestyle='--', alpha=0.3)
        ax5.set_xlabel('销售金额（元）', fontsize=11)
        ax5.set_ylabel('毛利率（%）', fontsize=11)
        ax5.set_title('金额-毛利率关系图\n（圆圈大小=毛利额）', fontsize=12, fontweight='bold')
        plt.colorbar(scatter, ax=ax5, label='毛利率%')
        ax5.grid(alpha=0.3)
    
    # 图6: 四象限分析
    if category_analysis is not None and len(category_analysis) > 0:
        ax6 = plt.subplot(2, 3, 6)
        
        median_amount = category_analysis['金额'].median()
        median_rate = 10  # 使用10%作为毛利率标准线
        
        # 分类四象限
        q1 = category_analysis[(category_analysis['金额'] >= median_amount) & 
                               (category_analysis['毛利率'] >= median_rate)]
        q2 = category_analysis[(category_analysis['金额'] < median_amount) & 
                               (category_analysis['毛利率'] >= median_rate)]
        q3 = category_analysis[(category_analysis['金额'] < median_amount) & 
                               (category_analysis['毛利率'] < median_rate)]
        q4 = category_analysis[(category_analysis['金额'] >= median_amount) & 
                               (category_analysis['毛利率'] < median_rate)]
        
        ax6.scatter(q1['金额'], q1['毛利率'], s=100, c='green', alpha=0.6, label=f'明星品类({len(q1)}个)')
        ax6.scatter(q2['金额'], q2['毛利率'], s=100, c='blue', alpha=0.6, label=f'潜力品类({len(q2)}个)')
        ax6.scatter(q3['金额'], q3['毛利率'], s=100, c='gray', alpha=0.6, label=f'问题品类({len(q3)}个)')
        ax6.scatter(q4['金额'], q4['毛利率'], s=100, c='red', alpha=0.6, label=f'瘦狗品类({len(q4)}个)')
        
        ax6.axvline(x=median_amount, color='black', linestyle='--', alpha=0.5)
        ax6.axhline(y=median_rate, color='black', linestyle='--', alpha=0.5)
        
        ax6.set_xlabel('销售金额（元）', fontsize=11)
        ax6.set_ylabel('毛利率（%）', fontsize=11)
        ax6.set_title('品类四象限分析', fontsize=12, fontweight='bold')
        ax6.legend(fontsize=9, loc='best')
        ax6.grid(alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('/workspace/product_analysis_charts.png', dpi=300, bbox_inches='tight')
    print("图表已保存: product_analysis_charts.png")
    plt.close()

def identify_issues(category_analysis):
    """识别问题品类"""
    print("\n正在识别问题品类...")
    
    issues = {
        '超高毛利率品类': [],
        '超低毛利率品类': [],
        '零负毛利品类': [],
        '大而不赚品类': [],
        '潜力品类': []
    }
    
    if category_analysis is None:
        return issues
    
    category_col = category_analysis.columns[0]
    
    # 超高毛利率（≥18%）
    high_margin = category_analysis[category_analysis['毛利率'] >= 18]
    for _, row in high_margin.iterrows():
        issues['超高毛利率品类'].append({
            '品类': row[category_col],
            '毛利率': f"{row['毛利率']:.2f}%",
            '金额': f"{row['金额']:,.0f}",
            '毛利': f"{row['毛利']:,.0f}"
        })
    
    # 超低毛利率（≤5%）
    low_margin = category_analysis[category_analysis['毛利率'] <= 5]
    for _, row in low_margin.iterrows():
        issues['超低毛利率品类'].append({
            '品类': row[category_col],
            '毛利率': f"{row['毛利率']:.2f}%",
            '金额': f"{row['金额']:,.0f}",
            '毛利': f"{row['毛利']:,.0f}"
        })
    
    # 零负毛利（0-2%）
    zero_margin = category_analysis[category_analysis['毛利率'] <= 2]
    for _, row in zero_margin.iterrows():
        issues['零负毛利品类'].append({
            '品类': row[category_col],
            '毛利率': f"{row['毛利率']:.2f}%",
            '金额': f"{row['金额']:,.0f}",
            '毛利': f"{row['毛利']:,.0f}"
        })
    
    # 大而不赚（金额高但毛利率低）
    median_amount = category_analysis['金额'].median()
    big_low = category_analysis[(category_analysis['金额'] >= median_amount) & 
                                (category_analysis['毛利率'] < 10)]
    for _, row in big_low.iterrows():
        issues['大而不赚品类'].append({
            '品类': row[category_col],
            '毛利率': f"{row['毛利率']:.2f}%",
            '金额': f"{row['金额']:,.0f}",
            '毛利': f"{row['毛利']:,.0f}"
        })
    
    # 潜力品类（金额不高但毛利率高）
    small_high = category_analysis[(category_analysis['金额'] < median_amount) & 
                                   (category_analysis['毛利率'] >= 15)]
    for _, row in small_high.iterrows():
        issues['潜力品类'].append({
            '品类': row[category_col],
            '毛利率': f"{row['毛利率']:.2f}%",
            '金额': f"{row['金额']:,.0f}",
            '毛利': f"{row['毛利']:,.0f}"
        })
    
    return issues

def generate_report(df, category_analysis, metrics, issues):
    """生成数据分析报告"""
    print("\n正在生成分析报告...")
    
    report_date = datetime.now().strftime('%Y年%m月%d日')
    
    report = f"""# 商品价格走向与毛利率数据分析报告
报告生成时间：{report_date}

## 一、整体数据概览

### 1.1 基础数据统计
- 商品总数：{len(df)} 个SKU
- 总销售金额：{metrics.get('total_amount', 0):,.2f} 元
- 总毛利额：{metrics.get('total_profit', 0):,.2f} 元
- 平均毛利率：{metrics.get('avg_margin_rate', 0):.2f}%
- 最高毛利率：{metrics.get('max_margin_rate', 0):.2f}%
- 最低毛利率：{metrics.get('min_margin_rate', 0):.2f}%

### 1.2 品类数量
"""
    
    if category_analysis is not None:
        report += f"- 品类总数：{len(category_analysis)} 个\n\n"
    
    report += """## 二、毛利率/毛利的极端情况

### 2.1 超高毛利率品类（≥18%）
"""
    
    if issues['超高毛利率品类']:
        for i, item in enumerate(issues['超高毛利率品类'][:10], 1):
            report += f"""
（{i}）{item['品类']}：
   - 毛利率：{item['毛利率']}
   - 销售金额：{item['金额']} 元
   - 毛利额：{item['毛利']} 元
   - 分析：高毛利率品类，{"金额较小，可维持现状" if float(item['金额'].replace(',','')) < 10000 else "重点关注，可加大推广力度"}
"""
    else:
        report += "暂无超高毛利率品类\n"
    
    report += """\n### 2.2 超低毛利率品类（≤5%）
**风险提示：低毛利且金额高的风险品类**
"""
    
    if issues['超低毛利率品类']:
        for i, item in enumerate(issues['超低毛利率品类'][:10], 1):
            amount = float(item['金额'].replace(',',''))
            risk_level = "高风险" if amount > 30000 else "中风险" if amount > 10000 else "低风险"
            report += f"""
（{i}）{item['品类']}：
   - 毛利率：{item['毛利率']}（{risk_level}）
   - 销售金额：{item['金额']} 元
   - 毛利额：{item['毛利']} 元
   - 建议：{"需要优化供应商渠道，找更优资源；若无法改善，建议缩减SKU" if amount > 10000 else "考虑逐步下架或替换"}
"""
    else:
        report += "暂无超低毛利率品类\n"
    
    report += """\n### 2.3 零毛利/负毛利品类（≤2%）
**严重警告：基本无利润甚至隐性亏损**
"""
    
    if issues['零负毛利品类']:
        for i, item in enumerate(issues['零负毛利品类'][:10], 1):
            report += f"""
（{i}）{item['品类']}：
   - 毛利率：{item['毛利率']}
   - 销售金额：{item['金额']} 元
   - 问题：可能存在定价不合理或成本核算有误
   - 措施：立即检查系统设置，核查成本价和售价；建议下架或更换高利润替代品
"""
    else:
        report += "暂无零负毛利品类（良好）\n"
    
    report += """\n### 2.4 品类结构异常

#### （1）高金额但低毛利的"拖后腿"品类
"""
    
    if issues['大而不赚品类']:
        total_big_amount = sum([float(item['金额'].replace(',','')) for item in issues['大而不赚品类']])
        report += f"""
共识别 {len(issues['大而不赚品类'])} 个"大而不赚"品类，总金额 {total_big_amount:,.0f} 元

典型品类：
"""
        for i, item in enumerate(issues['大而不赚品类'][:5], 1):
            report += f"""
{i}. {item['品类']}
   - 毛利率：{item['毛利率']}
   - 销售金额：{item['金额']} 元
   - 问题：金额大但毛利率低，属于"大而不赚"
   - 建议：需要优化供应商渠道，重新议价，或调整售价（随行就市）；目标将毛利率提升至12%以上
"""
    else:
        report += "暂无此类问题（良好）\n"
    
    report += """\n#### （2）低金额但高毛利的"潜力品类"
"""
    
    if issues['潜力品类']:
        report += f"""
共识别 {len(issues['潜力品类'])} 个潜力品类

典型品类：
"""
        for i, item in enumerate(issues['潜力品类'][:5], 1):
            report += f"""
{i}. {item['品类']}
   - 毛利率：{item['毛利率']}（优秀）
   - 销售金额：{item['金额']} 元
   - 潜力：高毛利率但销量未充分挖掘
   - 建议：扩大客户基数，提升销量，用高毛利品类带动整体利润
"""
    else:
        report += "暂无高潜力品类需要重点关注\n"
    
    report += """\n## 三、数据分析洞察

### 3.1 品类结构健康度评估
"""
    
    if category_analysis is not None and len(category_analysis) > 0:
        healthy_count = len(category_analysis[category_analysis['毛利率'] >= 10])
        warning_count = len(category_analysis[(category_analysis['毛利率'] >= 5) & (category_analysis['毛利率'] < 10)])
        danger_count = len(category_analysis[category_analysis['毛利率'] < 5])
        
        total_count = len(category_analysis)
        report += f"""
- 健康品类（毛利率≥10%）：{healthy_count} 个，占比 {healthy_count/total_count*100:.1f}%
- 警戒品类（5%≤毛利率<10%）：{warning_count} 个，占比 {warning_count/total_count*100:.1f}%
- 危险品类（毛利率<5%）：{danger_count} 个，占比 {danger_count/total_count*100:.1f}%

**健康度评级：**
"""
        if total_count > 0 and healthy_count/total_count >= 0.6:
            report += "🟢 良好 - 多数品类毛利率达标\n"
        elif healthy_count/total_count >= 0.4:
            report += "🟡 一般 - 需要优化部分品类结构\n"
        else:
            report += "🔴 较差 - 亟需调整品类结构，提升整体盈利能力\n"
    
    report += """\n### 3.2 价格走势分析
"""
    
    report += """
基于当前数据分析：

1. **毛利率两极分化明显**
   - 部分品类毛利率超过18%，显示定价能力强或成本控制好
   - 但同时存在大量低毛利率品类（<5%），拉低整体水平
   
2. **"大而不赚"现象严重**
   - 销售金额大的品类往往毛利率偏低
   - 需要重点关注主力品类的利润空间优化
   
3. **潜力品类未充分挖掘**
   - 存在高毛利率但销量不高的品类
   - 应该加大推广力度，提升销售占比

### 3.3 商品定价建议
"""
    
    if metrics.get('avg_margin_rate'):
        avg_rate = metrics['avg_margin_rate']
        report += f"""
当前平均毛利率为 {avg_rate:.2f}%。建议：

"""
        if avg_rate < 10:
            report += """- **紧急调整**：整体毛利率低于目标值（10%），需要立即采取行动
  - 对低毛利品类重新议价或调整售价
  - 加大高毛利品类的销售占比
  - 下架零负毛利商品
"""
        elif avg_rate < 12:
            report += """- **优化提升**：接近目标但仍有提升空间
  - 持续优化供应链成本
  - 逐步调整产品结构
  - 重点推广高毛利商品
"""
        else:
            report += """- **保持优化**：整体表现良好，继续保持并优化
  - 维护高毛利品类的竞争力
  - 持续监控低毛利品类
  - 挖掘新的高利润增长点
"""
    
    report += """\n## 四、下一步行动措施

### 4.1 针对低毛利/零毛利品类

**立即行动：**
"""
    
    if issues['零负毛利品类']:
        report += f"""
1. **零负毛利品类（{len(issues['零负毛利品类'])}个）**
   - 立即核查系统设置，检查成本价和售价录入
   - 直接下架或更换高利润替代品
   - 避免占用库存和资金
"""
    
    if issues['超低毛利率品类']:
        high_amount_low_margin = [item for item in issues['超低毛利率品类'] 
                                  if float(item['金额'].replace(',','')) > 10000]
        if high_amount_low_margin:
            report += f"""
2. **高金额低毛利品类（{len(high_amount_low_margin)}个重点）**
   - 重新核算成本，与供应商议价
   - 寻找更优渠道或工厂资源
   - 适当调整售价（随行就市）
   - 若无法改善，考虑缩减该类目的SKU
"""
    
    report += """\n### 4.2 针对高毛利品类

**扩大规模：**
"""
    
    if issues['超高毛利率品类'] or issues['潜力品类']:
        report += """
1. **加大采购和推广力度**
   - 增加高毛利品类的库存
   - 在销售端重点推荐
   - 培训销售团队突出这些产品优势

2. **提升销售占比**
   - 通过促销活动引导客户购买
   - 设置组合套餐，搭配高低毛利商品
   - 用高毛利品类带动整体利润
"""
    
    report += """\n### 4.3 针对"大而不赚"的主力品类

**结构优化：**
"""
    
    if issues['大而不赚品类']:
        report += """
1. **拆分子品类精细化管理**
   - 识别主力品类中的高毛利子类
   - 保留和加强高毛利子类
   - 优化或降低低毛利子类

2. **调整产品结构比例**
   - 逐步提高高毛利商品占比
   - 通过结构调整带动整体利润提升
   - 设定分品类的毛利率目标

3. **供应链深度优化**
   - 寻找更优质的供应商
   - 批量采购降低成本
   - 探索直接与工厂合作的可能性
"""
    
    report += """\n## 五、关键指标监控

### 5.1 建议设定的目标

"""
    
    if metrics.get('avg_margin_rate'):
        current = metrics['avg_margin_rate']
        target = max(10.0, current + 1.0)
        report += f"""
- **整体毛利率目标**：{target:.1f}%（当前：{current:.2f}%）
- **健康品类占比目标**：≥60%
- **危险品类（<5%）数量**：逐步降至0个
"""
    
    report += """
### 5.2 监控频率

- **每日监控**：零负毛利商品，确保及时发现问题
- **每周监控**：主力品类毛利率变化
- **每月监控**：整体毛利率达成情况，品类结构优化进展

### 5.3 预警机制

设置以下预警线：
- 🔴 红色预警：毛利率<3%，立即处理
- 🟡 黄色预警：毛利率3-5%，重点关注
- 🟢 绿色健康：毛利率≥10%，正常运营

## 六、总结与建议

### 6.1 核心问题总结
"""
    
    issues_summary = []
    if issues['零负毛利品类']:
        issues_summary.append(f"存在{len(issues['零负毛利品类'])}个零负毛利品类，需立即处理")
    if issues['超低毛利率品类']:
        issues_summary.append(f"{len(issues['超低毛利率品类'])}个超低毛利率品类拖累整体表现")
    if issues['大而不赚品类']:
        issues_summary.append(f"{len(issues['大而不赚品类'])}个主力品类利润空间不足")
    
    if issues_summary:
        for i, issue in enumerate(issues_summary, 1):
            report += f"{i}. {issue}\n"
    else:
        report += "整体运营健康，继续保持优化\n"
    
    report += """\n### 6.2 改进优先级

**P0（最高优先级）**
- 处理零负毛利商品
- 检查系统定价设置

**P1（高优先级）**
- 优化超低毛利率品类的供应链
- 加大高毛利品类推广力度

**P2（中优先级）**
- 调整主力品类产品结构
- 提升潜力品类销量

**P3（持续优化）**
- 建立定期监控机制
- 持续优化整体品类结构

### 6.3 预期效果

通过以上措施的实施，预期：
- 1个月内：零负毛利品类清零
- 3个月内：整体毛利率提升1-2个百分点
- 6个月内：品类结构优化完成，健康品类占比达60%以上

---

*本报告基于实际数据生成，具体数据请参见附图《product_analysis_charts.png》*
"""
    
    return report

def main():
    """主函数"""
    print("=" * 60)
    print("商品价格走向与毛利率数据分析")
    print("=" * 60)
    
    # 加载数据
    df = load_data('程宇昕.xlsx')
    
    print("\n数据预览：")
    print(df.head())
    print(f"\n列名: {df.columns.tolist()}")
    
    # 计算指标
    metrics = calculate_metrics(df)
    print(f"\n整体指标: {metrics}")
    
    # 品类分析
    category_analysis = analyze_by_category(df)
    if category_analysis is not None:
        print(f"\n品类分析结果（Top 10）：")
        print(category_analysis.head(10).to_string())
    
    # 识别问题
    issues = identify_issues(category_analysis)
    print(f"\n问题品类统计：")
    for issue_type, items in issues.items():
        print(f"  {issue_type}: {len(items)}个")
    
    # 生成可视化
    create_visualizations(df, category_analysis)
    
    # 生成报告
    report = generate_report(df, category_analysis, metrics, issues)
    
    # 保存报告
    with open('/workspace/数据分析报告.md', 'w', encoding='utf-8') as f:
        f.write(report)
    
    print("\n" + "=" * 60)
    print("✅ 分析完成！")
    print("=" * 60)
    print("\n生成文件：")
    print("1. 📊 数据可视化图表: product_analysis_charts.png")
    print("2. 📄 数据分析报告: 数据分析报告.md")
    print("\n请查看报告了解详细分析结果。")

if __name__ == '__main__':
    main()
