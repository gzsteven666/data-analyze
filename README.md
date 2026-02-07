[README.md](https://github.com/user-attachments/files/25149227/README.md)
# 数据分析信息图 Skill

通用型的数据分析与信息图生成工具，面向「任意行业、任意维度」的数据文件，自动完成数据体检、洞察生成、可视化和高质量输出。

## 🎯 功能特性

### 核心功能
- **智能数据识别**：自动识别 CSV / Excel 数据，感知字段类型、缺失值、异常点
- **多维洞察**：覆盖数值、分类、时间、相关性等常见分析视角
- **专业可视化**：生成高质量 PNG/SVG 图表，自动解决中文字体问题
- **交互式报告**：输出 HTML 信息图、Word/Excel 资料包；截图默认关闭，可按需开启
- **输出目录**：默认把结果写到“数据文件所在目录/outputs”，不再占用 skill 目录，可通过 `--output-dir` 自定义
- **核心实体维度**：可用 `--core-dimension` 指定（如医院/客户/渠道/门店/品牌等），集中度/TopN/覆盖度等分析/图表仅针对该维度；未指定则自动推断，找不到时跳过相关输出
- **全流程自动化**：一条命令完成加载 → 分析 → 图表 → 报告 → 截图

### 默认分析维度
1. **字段画像**：字段类型、唯一值、缺失率、内存占用
2. **数值洞察**：均值/中位数/标准差/极值、波动特征
3. **分类洞察**：TopN 类别分布、占比、集中度
4. **时间趋势**：自动识别时间列并分析趋势、波动、同比环比
5. **相关性分析**：自动输出 Pearson 相关性矩阵、强相关指标提示

### 输出格式
- **数据文件**：字段画像、数值统计、分类分布、原有行业分析结果等 CSV/Excel
- **可视化**：PNG/SVG 雙格式图表，包含通用/行业定制图
- **报告**：HTML 信息图（附洞察摘要）、Word 摘要（可选扩展）
- **截图**：使用 Playwright 生成高分辨率 PNG 截图

## 🚀 快速开始

### 环境要求
- Python 3.7+
- Node.js 14+
- Chrome/Chromium浏览器

### 安装依赖

```bash
# Python依赖
pip install -r requirements.txt

# Node.js依赖
npm install

# 安装Playwright浏览器
npx playwright install chromium
```

### 基本用法

```bash
# 分析CSV文件
node src/index.js data/metrics.csv --company="业务线A"

# 分析Excel文件指定工作表并指定时间/指标列
node src/index.js data/quarter_sales.xlsx --sheet="Q4" --company="旗舰店A" --time="月份" --value="GMV"

# 自定义输出目录（默认写入数据同级的 outputs/，如需指定请用 --output-dir）
node src/index.js data/数据.csv --company="测试公司" --output-dir="/path/to/outputs"

# 图表策略：auto/on/off（默认auto：先分析可视价值再决定，避免无意义图表）
node src/index.js data/数据.csv --charts-mode on      # 强制生成图表
node src/index.js data/数据.csv --charts-mode off     # 不生成图表
node src/index.js data/数据.csv --enable-screenshot    # 若需截图再显式开启
# 指定核心实体维度（如医院/客户/渠道/门店/品牌等）
node src/index.js data/数据.csv --core-dimension "医院名称"
```

### Python直接调用

```python
from src.main import DataAnalysisPipeline

# 创建流水线
pipeline = DataAnalysisPipeline()

# 运行分析
results = pipeline.run_full_pipeline(
    data_path="data/quarter_sales.xlsx",
    sheet_name="Q4",
    company_name="旗舰店A"
)

print(f"分析完成！截图: {results['screenshot']}")
```

## 📊 分析流程

### 1. 数据读取与体检
- 自动识别文件类型（CSV/Excel）
- 数据质量检查（缺失值、重复值、异常值）
- 字段类型识别和统计

### 2. 数据清洗与过滤
- 自动字段标准化（大小写、空白、单位）
- 数值/日期字段智能转换
- 可选业务过滤 Hook（示例中默认保留原始数据）

### 3. 多维度分析
```
通用洞察
├── 字段画像（缺失率 / 唯一值 / 内存占用）
├── 数值统计（均值 / 分位数 / 标准差）
├── 分类TopN（集中度、长尾度量）
├── 时间趋势（自动推断粒度，输出波动）
├── 相关性矩阵（高相关指标提示）

行业扩展（存在指定字段时自动开启）
├── 市场份额 / 城市机会 / 覆盖效率
├── 产品结构 / 自定义维度
```

### 4. 可视化生成
- 市场份额饼图+条形图
- 城市机会热力图
- 覆盖效率散点图
- 产品结构对比图

### 5. 报告生成
- Word文档（策略化摘要 + 行动计划，自动命名为 `数据名_分析报告.docx`，默认字体：微软雅黑）
- Excel汇总表（多工作表）
- HTML信息图（与报告内容一致的卡片式速览，图表由 `charts-mode auto` 按需生成）
- 截图（默认关闭，若需要可加 `--enable-screenshot`）

## 🎨 技术特色

### 中文优化
- 智能字体选择（PingFang SC、Microsoft YaHei等）
- 自动解决中文乱码问题
- 支持SVG矢量图输出

### 视觉设计
- 现代化配色方案
- 响应式布局设计
- 交互动画效果
- 专业图表样式

### 性能优化
- 异步处理提高效率
- 多格式同步输出
- 内存使用优化
- 批量处理能力

### 分析策略（降噪）
- 分类TopN仅针对 2~50 个唯一值且非 90% 全唯一的列，跳过无意义字段（如完全统一值或几乎每行唯一的ID）
- 相关性仅在存在 ≥2 个数值列且最高相关≥0.3时输出；无时间列则不生成趋势图
- 默认关闭图表/截图，除非手动开启；优先输出业务可用的文本洞察与策略建议

## 📁 文件结构

```
data-analyze-skill/
├── src/                    # 核心代码
│   ├── main.py            # Python主程序
│   ├── index.js           # Node.js入口
│   ├── data_analyzer.py   # 数据分析模块
│   ├── chart_generator.py # 图表生成模块
│   ├── infographic_generator.py # HTML生成
│   └── screenshot_generator.py  # 截图模块
├── templates/             # HTML模板
├── static/               # 静态资源
├── tests/                # 测试文件
├── data/                 # 输入数据
├── outputs/（示例）       # 输出结果；实际运行时默认写到“数据文件同级的 outputs/”
│   ├── csv/             # CSV文件
│   ├── excel/           # Excel文件
│   ├── figures/         # 图表文件
│   ├── html/            # 信息图HTML
│   └── reports/         # Word报告 + 洞察JSON（截图按需开启）
```

## ⚙️ 配置选项

### 分析参数
- `company_name`: 可选，高亮的品牌/分组名称
- `sheet_name`: Excel工作表名称
- `output_dir`: 输出目录（默认写到数据文件同级的 `outputs/`）
- `time_column`: 可选，时间字段（自动识别不可靠时使用）
- `value_column`: 可选，核心指标列（默认取首个数值列）

### 图表配置
- 颜色主题：支持自定义配色
- 字体设置：自动选择系统字体
- 图表尺寸：可调整输出尺寸
- 格式选择：PNG/SVG双格式

### 截图配置
- 分辨率：支持多尺寸输出
- 质量设置：可调节图片质量
- 延迟等待：确保动画完成
- 全页截图：支持长页面截图

## 🔧 高级用法

### 自定义分析维度
```python
# 在data_analyzer.py中添加自定义分析方法
def custom_analysis(self, df, params):
    # 自定义分析逻辑
    return results
```

### 自定义图表样式
```python
# 在chart_generator.py中添加自定义图表
def create_custom_chart(self, data, style_params):
    # 自定义图表生成
    return chart_paths
```

### 自定义HTML模板
```html
<!-- 修改templates目录下的HTML模板 -->
<div class="custom-section">
    <!-- 自定义内容 -->
</div>
```

## 📈 性能指标

基于参考项目benchmark数据：
- **数据处理能力**: 10万条记录 < 30秒
- **图表生成速度**: 6个图表 < 60秒
- **内存使用**: 峰值 < 500MB
- **输出质量**: 300DPI高清晰度

## 🛠️ 故障排除

### 常见问题

**Q: 中文显示乱码**
A: 系统缺少中文字体，安装PingFang SC或Microsoft YaHei

**Q: Excel文件读取失败**
A: 检查文件格式和编码，确保文件未损坏

**Q: 截图生成失败**
A: 检查Chromium安装状态，运行`npx playwright install chromium`

**Q: 内存不足**
A: 减少单次处理数据量，分批处理大文件

### 调试模式
```bash
# 启用详细日志
DEBUG=1 node src/index.js data/数据.csv

# Python调试
python src/main.py data/数据.csv --debug
```

## 📚 参考案例

### 1. 电商季度复盘
- **数据量**：12万条订单明细
- **自定义参数**：`company_name=旗舰店A`、`time_column=月份`
- **亮点**：自动输出 GMV 趋势、品类集中度、客单价波动、关键相关性

### 2. 市场活动监控
- **数据来源**：CSV 活动日志
- **聚焦内容**：投放渠道 vs 线索质量
- **亮点**：TopN 渠道概览、时间段爆发点、强相关渠道组合

### 3. B2B 运营日报
- **自动化**：定时用 cron 执行 CLI，每天生成 HTML 信息图 + 截图
- **洞察**：指标异常提醒、同比/环比趋势、重点客户集群

## 🤝 贡献指南

1. Fork项目仓库
2. 创建功能分支
3. 提交代码变更
4. 运行测试用例
5. 提交Pull Request

## 📄 许可证

MIT License - 详见[LICENSE](LICENSE)文件

## 📞 支持

如有问题或建议，请通过以下方式联系：
- 提交Issue
- 发送邮件
- 技术讨论

---

**基于真实数据分析项目经验，提供企业级数据分析解决方案**
