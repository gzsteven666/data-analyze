# 使用示例

## 示例1: 医疗耗材市场分析

### 场景描述
分析广东省静脉留置针市场，对比林华与竞品的表现。

### 数据准备
```
data/
└── 采购数据.xlsx
    ├── 广东所有（工作表）
    ├── 品牌名称
    ├── 医疗机构名称
    ├── 城市
    ├── 第三年采购需求量
    └── 类别名称
```

### 执行命令
```bash
node src/index.js data/采购数据.xlsx --sheet="广东所有" --company="林华"
```

### 预期输出
```
📊 开始数据分析...
✅ 数据加载完成，共15000条记录，12个字段
✅ 数据体检完成
✅ 过滤后留置针数据: 8500条记录
✅ 综合分析完成
📈 开始生成图表...
✅ 图表生成完成，共6个图表
📊 开始导出数据...
✅ 导出CSV: outputs/csv/品牌份额.csv
✅ 导出CSV: outputs/csv/城市份额.csv
✅ 导出CSV: outputs/csv/机会城市.csv
✅ 导出Excel: outputs/excel/数据分析汇总.xlsx
🌐 生成信息图HTML...
✅ 信息图HTML已生成: 生成结果信息图/market_analysis_infographic.html
📸 生成截图...
✅ 截图已生成: 生成结果信息图/analysis_20241113_143052.png

============================================================
数据分析流水线执行完成！
============================================================
信息图HTML: 生成结果信息图/market_analysis_infographic.html
截图文件: 生成结果信息图/analysis_20241113_143052.png
图表文件: 6个
CSV文件: outputs/csv/目录下
Excel汇总: outputs/excel/数据分析汇总.xlsx
============================================================
```

### 关键发现示例
```
核心洞察:
- 林华市场份额18.79%，排名第2
- 佛山、茂名、阳江为白区机会城市
- TOP20医院贡献40%业务量
- 安全型产品占比偏低，有升级空间

策略建议:
- 优先突破佛山、茂名等高容量城市
- 加强TOP医院的深耕覆盖
- 推广安全型产品，优化结构
```

## 示例2: CSV数据快速分析

### 场景描述
快速分析CSV格式的竞品数据。

### 数据格式
```csv
品牌名称,医疗机构名称,城市,采购量,产品类型
林华,医院A,广州,1000,安全型
碧迪,医院B,深圳,1500,普通型
威海洁瑞,医院C,佛山,800,安全型
...
```

### 执行命令
```bash
node src/index.js data/竞品数据.csv --company="林华"
```

### 输出结果
- 生成6个分析维度的图表
- 输出机会城市和医院清单
- 生成HTML信息图和截图

## 示例3: Python API调用

### 代码示例
```python
from src.main import DataAnalysisPipeline
import pandas as pd

# 创建分析流水线
pipeline = DataAnalysisPipeline()

# 配置参数
config = {
    'company_name': '林华',
    'output_dir': '我的分析结果'
}

# 执行完整分析
results = pipeline.run_full_pipeline(
    data_path='data/采购数据.xlsx',
    sheet_name='广东所有',
    company_name='林华'
)

# 查看结果
print(f"HTML报告: {results['html']}")
print(f"截图文件: {results['screenshot']}")
print(f"生成图表: {len(results['charts'])}个")

# 读取分析结果
if results['excel']:
    df_summary = pd.read_excel(results['excel'], sheet_name='品牌份额')
    print(f"林华市场份额: {df_summary.iloc[0]['市场份额']:.1f}%")
```

## 示例4: 批量分析

### 脚本示例
```bash
#!/bin/bash
# 批量分析多个数据文件

files=(
    "data/2023年数据.xlsx"
    "data/2022年数据.xlsx"
    "data/2021年数据.xlsx"
)

for file in "${files[@]}"; do
    echo "分析文件: $file"
    filename=$(basename "$file" .xlsx)
    node src/index.js "$file" --company="林华" --output="结果_$filename"
    echo "完成: $file"
    echo "---"
done
```

## 示例5: 自定义分析维度

### 扩展分析
在`data_analyzer.py`中添加自定义分析:

```python
def analyze_price_sensitivity(self, df):
    """价格敏感度分析"""
    # 计算价格弹性
    price_elasticity = df.groupby('价格区间').agg({
        '采购量': 'sum',
        '医疗机构数量': 'nunique'
    }).reset_index()

    return price_elasticity

def analyze_seasonal_patterns(self, df):
    """季节性模式分析"""
    # 提取月份信息
    df['月份'] = pd.to_datetime(df['时间']).dt.month

    monthly_trend = df.groupby('月份')['采购量'].sum()

    return monthly_trend
```

## 示例6: 定制化图表

### 自定义图表样式
在`chart_generator.py`中添加:

```python
def create_custom_radar_chart(self, metrics_data):
    """自定义雷达图"""
    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(projection='polar'))

    # 自定义样式
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
    line_styles = ['-', '--', '-.']

    for i, (company, data) in enumerate(metrics_data.items()):
        angles = np.linspace(0, 2 * np.pi, len(data), endpoint=False)
        values = list(data.values())

        ax.plot(angles, values,
                color=colors[i % len(colors)],
                linestyle=line_styles[i % len(line_styles)],
                linewidth=2, label=company)

        ax.fill(angles, values, alpha=0.25, color=colors[i])

    return fig
```

## 结果验证

### 检查输出文件
```bash
# 查看生成的文件
ls -la 生成结果信息图/
ls -la outputs/

# 验证HTML报告
open 生成结果信息图/market_analysis_infographic.html

# 查看截图
open 生成结果信息图/analysis_*.png
```

### 数据质量检查
```python
import pandas as pd

# 检查分析结果
df_share = pd.read_csv('outputs/csv/品牌份额.csv')
print("品牌份额数据:")
print(df_share.head())

# 检查机会城市
df_cities = pd.read_csv('outputs/csv/机会城市.csv')
print(f"\n识别到{len(df_cities)}个机会城市")
print(df_cities.head())
```

## 性能优化建议

### 大数据处理
```python
# 分批处理大文件
chunk_size = 10000
chunks = []

for chunk in pd.read_csv('large_file.csv', chunksize=chunk_size):
    # 处理每个chunk
    processed_chunk = process_chunk(chunk)
    chunks.append(processed_chunk)

# 合并结果
final_result = pd.concat(chunks, ignore_index=True)
```

### 内存优化
```python
# 优化数据类型
df['采购量'] = df['采购量'].astype('int32')
df['价格'] = df['价格'].astype('float32')

# 删除不需要的列
df = df.drop(columns=['临时列1', '临时列2'])
```

这些示例展示了skill的多种使用方式，从简单的命令行调用到复杂的自定义分析，满足不同场景的需求。基于真实医疗耗材数据分析项目经验，确保分析结果的专业性和实用性。"}