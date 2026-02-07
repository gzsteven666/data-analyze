#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表生成模块
生成各种类型的数据可视化图表
"""

import matplotlib
matplotlib.use("Agg")  # 使用无界面后端，避免本地GUI依赖
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
import os
from matplotlib.patches import Rectangle
import matplotlib.patches as mpatches
from matplotlib import font_manager

class ChartGenerator:
    """图表生成器类"""

    def __init__(self, output_dir='outputs/figures'):
        self.output_dir = output_dir
        self.setup_style()
        self.create_output_dir()

    def setup_style(self):
        """设置图表样式"""
        # 中文字体设置
        candidates = [
            "PingFang SC", "Heiti SC", "Songti SC", "Hiragino Sans GB",
            "Noto Sans CJK SC", "STHeiti", "Microsoft YaHei", "SimHei", "SimSun"
        ]

        installed = {f.name for f in font_manager.fontManager.ttflist}

        for font_name in candidates:
            if any(font_name in name for name in installed):
                plt.rcParams['font.family'] = font_name
                break

        plt.rcParams['axes.unicode_minus'] = False
        plt.rcParams['figure.dpi'] = 200
        plt.rcParams['savefig.dpi'] = 200

        # 颜色方案（简洁、易读）
        self.colors = ['#4C78A8', '#54A24B', '#E45756', '#F58518', '#72B7B2', '#E39C36', '#B279A2', '#9D755D']
        self.company_colors = {
            '林华': '#FF6B6B',
            '苏州碧迪': '#4ECDC4',
            '明州斯睿': '#45B7D1',
            '威海洁瑞': '#96CEB4',
            '广东百合': '#FFEAA7'
        }

    def create_output_dir(self):
        """创建输出目录"""
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(f"{self.output_dir}/png", exist_ok=True)
        os.makedirs(f"{self.output_dir}/svg", exist_ok=True)

    def save_figure(self, fig, filename):
        """保存图表为PNG和SVG格式"""
        # PNG格式
        png_path = f"{self.output_dir}/png/{filename}.png"
        fig.savefig(png_path, format='png', bbox_inches='tight', dpi=300)

        # SVG格式
        svg_path = f"{self.output_dir}/svg/{filename}.svg"
        fig.savefig(svg_path, format='svg', bbox_inches='tight')

        return {'png': png_path, 'svg': svg_path}

    def create_market_share_chart(self, data, title='品牌市场份额', highlight_company='林华'):
        """
        创建市场份额图表

        Args:
            data: 份额数据
            title: 图表标题
            highlight_company: 高亮公司

        Returns:
            dict: 图表路径
        """
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))

        # 饼图
        colors = [self.company_colors.get(brand, self.colors[i % len(self.colors)])
                 for i, brand in enumerate(data['品牌名称'])]

        wedges, texts, autotexts = ax1.pie(data['总量'], labels=data['品牌名称'],
                                          autopct='%1.1f%%', colors=colors, startangle=90)
        ax1.set_title(f'{title} - 占比分布', fontsize=14, fontweight='bold')

        # 条形图
        bars = ax2.barh(data['品牌名称'], data['市场份额'], color=colors)
        ax2.set_xlabel('市场份额 (%)', fontsize=12)
        ax2.set_title(f'{title} - 排名', fontsize=14, fontweight='bold')

        # 高亮目标公司
        if highlight_company in data['品牌名称'].values:
            idx = data[data['品牌名称'] == highlight_company].index[0]
            bars[idx].set_color('#FF4444')
            bars[idx].set_edgecolor('black')
            bars[idx].set_linewidth(2)

        # 添加数值标签
        for i, (bar, share) in enumerate(zip(bars, data['市场份额'])):
            ax2.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
                    f'{share:.1f}%', va='center', fontsize=10)

        plt.tight_layout()
        paths = self.save_figure(fig, 'market_share')
        plt.close()

        return paths

    def create_city_heatmap(self, city_data, company_name='林华'):
        """
        创建城市机会热力图

        Args:
            city_data: 城市数据
            company_name: 公司名称

        Returns:
            dict: 图表路径
        """
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 12))

        # 城市容量排名
        top_cities = city_data.nlargest(15, '总容量')
        bars1 = ax1.bar(range(len(top_cities)), top_cities['总容量'],
                       color='#4ECDC4', alpha=0.8)
        ax1.set_xticks(range(len(top_cities)))
        ax1.set_xticklabels(top_cities['城市'], rotation=45, ha='right')
        ax1.set_ylabel('市场容量', fontsize=12)
        ax1.set_title('TOP15城市市场容量', fontsize=14, fontweight='bold')

        # 添加数值标签
        for bar, value in zip(bars1, top_cities['总容量']):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + value*0.01,
                    f'{value:,.0f}', ha='center', va='bottom', fontsize=9)

        # 林华份额
        bars2 = ax2.bar(range(len(top_cities)), top_cities['林华份额'],
                       color='#FF6B6B', alpha=0.8)
        ax2.set_xticks(range(len(top_cities)))
        ax2.set_xticklabels(top_cities['城市'], rotation=45, ha='right')
        ax2.set_ylabel('林华市场份额 (%)', fontsize=12)
        ax2.set_title(f'{company_name}在TOP15城市的份额', fontsize=14, fontweight='bold')

        # 添加份额标签
        for bar, value in zip(bars2, top_cities['林华份额']):
            ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                    f'{value:.1f}%', ha='center', va='bottom', fontsize=9)

        plt.tight_layout()
        paths = self.save_figure(fig, 'city_opportunities')
        plt.close()

        return paths

    def create_coverage_scatter(self, coverage_data, company_name='林华'):
        """
        创建覆盖-份额散点图

        Args:
            coverage_data: 覆盖数据
            company_name: 公司名称

        Returns:
            dict: 图表路径
        """
        fig, ax = plt.subplots(figsize=(12, 8))

        # 散点图
        scatter = ax.scatter(coverage_data['覆盖医院数'], coverage_data['单院均量'],
                           s=coverage_data['总量']/1000, alpha=0.6, c=self.colors[:len(coverage_data)])

        # 高亮目标公司
        company_data = coverage_data[coverage_data['品牌名称'] == company_name]
        if len(company_data) > 0:
            ax.scatter(company_data['覆盖医院数'], company_data['单院均量'],
                      s=company_data['总量']/500, c='red', alpha=0.8,
                      edgecolors='black', linewidth=2, label=company_name)

        ax.set_xlabel('覆盖医院数量', fontsize=12)
        ax.set_ylabel('单院均量', fontsize=12)
        ax.set_title('品牌覆盖-效率分析（气泡大小=总业务量）', fontsize=14, fontweight='bold')

        # 添加品牌标签
        for _, row in coverage_data.iterrows():
            ax.annotate(row['品牌名称'],
                       (row['覆盖医院数'], row['单院均量']),
                       xytext=(5, 5), textcoords='offset points',
                       fontsize=9, alpha=0.8)

        ax.grid(True, alpha=0.3)
        ax.legend()

        plt.tight_layout()
        paths = self.save_figure(fig, 'coverage_analysis')
        plt.close()

        return paths

    def create_product_structure_chart(self, structure_data):
        """
        创建产品结构图表

        Args:
            structure_data: 结构数据

        Returns:
            dict: 图表路径
        """
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))

        # 准备数据
        if '产品类型' in structure_data.columns:
            df_plot = structure_data.set_index('产品类型')
        else:
            df_plot = structure_data

        # 堆叠条形图
        df_plot.plot(kind='bar', stacked=True, ax=ax1,
                    color=[self.company_colors.get(col, self.colors[i])
                          for i, col in enumerate(df_plot.columns)])
        ax1.set_title('产品结构对比（总量）', fontsize=14, fontweight='bold')
        ax1.set_xlabel('产品类型', fontsize=12)
        ax1.set_ylabel('采购量', fontsize=12)
        ax1.legend(title='品牌', bbox_to_anchor=(1.05, 1), loc='upper left')
        ax1.tick_params(axis='x', rotation=0)

        # 百分比堆叠图
        df_pct = df_plot.div(df_plot.sum(axis=1), axis=0) * 100
        df_pct.plot(kind='bar', stacked=True, ax=ax2,
                   color=[self.company_colors.get(col, self.colors[i])
                         for i, col in enumerate(df_plot.columns)])
        ax2.set_title('产品结构对比（占比）', fontsize=14, fontweight='bold')
        ax2.set_xlabel('产品类型', fontsize=12)
        ax2.set_ylabel('占比 (%)', fontsize=12)
        ax2.legend(title='品牌', bbox_to_anchor=(1.05, 1), loc='upper left')
        ax2.tick_params(axis='x', rotation=0)

        plt.tight_layout()
        paths = self.save_figure(fig, 'product_structure')
        plt.close()

        return paths

    def create_competition_radar(self, metrics_data, company_name='林华'):
        """
        创建竞争雷达图

        Args:
            metrics_data: 指标数据
            company_name: 公司名称

        Returns:
            dict: 图表路径
        """
        fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(projection='polar'))

        # 准备数据
        categories = list(metrics_data.keys())
        values = list(metrics_data.values())

        # 角度计算
        angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
        values += values[:1]  # 闭合
        angles += angles[:1]

        # 绘制雷达图
        ax.plot(angles, values, 'o-', linewidth=2, label=company_name, color='#FF6B6B')
        ax.fill(angles, values, alpha=0.25, color='#FF6B6B')

        # 设置标签
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories)
        ax.set_ylim(0, max(values) * 1.1)
        ax.set_title('竞争力雷达图', fontsize=16, fontweight='bold', pad=20)

        # 添加网格
        ax.grid(True)

        plt.tight_layout()
        paths = self.save_figure(fig, 'competition_radar')
        plt.close()

        return paths

    def create_numeric_overview_chart(self, numeric_data, top_n=8):
        """创建数值字段概览图（仅在显著时调用）"""
        data = numeric_data.sort_values('均值', ascending=False).head(top_n)
        if data.empty:
            return {}

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

        ax1.barh(data['字段'], data['均值'], color=self.colors[:len(data)])
        ax1.set_title('Top数值指标均值', fontsize=14, fontweight='bold')
        ax1.set_xlabel('均值')
        ax1.invert_yaxis()
        for i, value in enumerate(data['均值']):
            ax1.text(value, i, f"{value:.2f}", va='center', ha='left', fontsize=9)

        ax2.barh(data['字段'], data['标准差'], color=self.colors[-len(data):])
        ax2.set_title('波动度（标准差）', fontsize=14, fontweight='bold')
        ax2.set_xlabel('标准差')
        ax2.invert_yaxis()
        for i, value in enumerate(data['标准差']):
            ax2.text(value, i, f"{value:.2f}", va='center', ha='left', fontsize=9)

        plt.tight_layout()
        paths = self.save_figure(fig, 'numeric_overview')
        plt.close()
        return paths

    # 市场视角核心图表（按指标汇总）
    def create_core_share_chart(self, data, dim_name):
        dfx = data.head(15).copy()
        if dfx.empty:
            return {}
        metric_col = [c for c in dfx.columns if c not in [dim_name, '份额(%)']][0]
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.barh(dfx[dim_name], dfx[metric_col], color=self.colors[0])
        ax.invert_yaxis()
        ax.set_title(f"{dim_name} 份额（按{metric_col}，前15）", fontsize=14, fontweight='bold')
        ax.set_xlabel(metric_col)
        for i, (v, s) in enumerate(zip(dfx[metric_col], dfx.get('份额(%)', [None]*len(dfx)))):
            label = f" {v:,.0f}"
            if s is not None:
                label += f"（{s:.2f}%）"
            ax.text(v, i, label, va="center", fontsize=9)
        plt.tight_layout()
        return self.save_figure(fig, 'core_share_top15')

    def create_city_share_chart(self, data):
        dfx = data.head(15).copy()
        if dfx.empty:
            return {}
        metric_col = [c for c in dfx.columns if c not in ['城市', '份额(%)']][0]
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.barh(dfx['城市'], dfx[metric_col], color=self.colors[1])
        ax.invert_yaxis()
        ax.set_title("城市分布（按量，前15）", fontsize=14, fontweight='bold')
        ax.set_xlabel(metric_col)
        for i, v in enumerate(dfx[metric_col]):
            ax.text(v, i, f" {v:,.0f}", va="center", fontsize=9)
        plt.tight_layout()
        return self.save_figure(fig, 'city_share_top15')

    def create_category_share_chart(self, data):
        dfx = data.head(12).copy()
        if dfx.empty:
            return {}
        metric_col = [c for c in dfx.columns if c not in ['目录名称', '份额(%)']][0]
        fig, ax = plt.subplots(figsize=(9, 6))
        ax.barh(dfx['目录名称'], dfx[metric_col], color=self.colors[2])
        ax.invert_yaxis()
        ax.set_title("目录份额（前12）", fontsize=14, fontweight='bold')
        ax.set_xlabel(metric_col)
        for i, (v, s) in enumerate(zip(dfx[metric_col], dfx.get('份额(%)', [None]*len(dfx)))):
            label = f" {v:,.0f}"
            if s is not None:
                label += f"（{s:.2f}%）"
            ax.text(v, i, label, va="center", fontsize=8)
        plt.tight_layout()
        return self.save_figure(fig, 'category_share_top12')

    def create_major_share_chart(self, data):
        dfx = data.head(10).copy()
        if dfx.empty:
            return {}
        metric_col = [c for c in dfx.columns if c not in ['产品大类', '份额(%)']][0]
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(dfx['产品大类'], dfx[metric_col], color=self.colors[3])
        ax.set_title("产品大类分布（前10）", fontsize=14, fontweight='bold')
        ax.set_ylabel(metric_col)
        ax.tick_params(axis="x", labelrotation=35)
        for lbl in ax.get_xticklabels():
            lbl.set_ha("right")
        for x, v, s in zip(range(len(dfx)), dfx[metric_col], dfx.get('份额(%)', [None]*len(dfx))):
            label = f"{v:,.0f}"
            if s is not None:
                label += f"\\n{s:.2f}%"
            ax.text(x, v, label, ha="center", va="bottom", fontsize=9)
        plt.tight_layout()
        return self.save_figure(fig, 'major_share_top10')

    def create_coverage_chart(self, data, dim_name):
        dfx = data.head(15).copy()
        if dfx.empty or '覆盖医院数' not in dfx.columns or '单院均量' not in dfx.columns:
            return {}
        fig, ax1 = plt.subplots(figsize=(8, 5))
        ax1.bar(dfx[dim_name], dfx['覆盖医院数'], color=self.colors[4])
        ax1.set_ylabel("覆盖医院数")
        ax1.set_title(f"{dim_name} 覆盖与单院均量（前15）", fontsize=14, fontweight='bold')
        ax1.tick_params(axis='x', labelrotation=40)
        for lbl in ax1.get_xticklabels():
            lbl.set_ha("right")
        ax2 = ax1.twinx()
        ax2.plot(dfx[dim_name], dfx['单院均量'], color=self.colors[5], marker="o")
        ax2.set_ylabel("单院均量")
        plt.tight_layout()
        return self.save_figure(fig, 'coverage_top15')

    def create_categorical_topn_chart(self, categorical_data, top_n=10):
        """创建分类TopN分布图（仅在显著时调用）"""
        data = categorical_data.sort_values('数量', ascending=False).head(top_n).copy()
        if data.empty:
            return {}

        data['字段类别'] = data['字段'] + ' - ' + data['类别']
        fig, ax = plt.subplots(figsize=(14, 8))
        sns.barplot(data=data, x='数量', y='字段类别', palette=self.colors[:len(data)], ax=ax)
        ax.set_title('分类TopN分布', fontsize=14, fontweight='bold')
        ax.set_xlabel('数量')
        ax.set_ylabel('')
        for p, (_, row) in zip(ax.patches, data.iterrows()):
            ax.text(p.get_width() + max(data['数量']) * 0.01, p.get_y() + p.get_height()/2,
                    f"{row['占比(%)']:.1f}%", va='center', fontsize=9)

        plt.tight_layout()
        paths = self.save_figure(fig, 'categorical_topn')
        plt.close()
        return paths

    def create_correlation_heatmap(self, corr_matrix):
        """创建相关性热力图"""
        fig, ax = plt.subplots(figsize=(10, 8))
        sns.heatmap(corr_matrix, cmap='RdYlBu', annot=True, fmt='.2f',
                    linewidths=0.5, ax=ax, cbar_kws={'shrink': .5})
        ax.set_title('指标相关性矩阵', fontsize=14, fontweight='bold')
        plt.tight_layout()
        paths = self.save_figure(fig, 'correlation_heatmap')
        plt.close()
        return paths

    def create_time_series_chart(self, time_trend):
        """创建时间趋势图"""
        if time_trend.empty:
            return {}

        fig, ax = plt.subplots(figsize=(14, 6))
        ax.plot(time_trend['时间'], time_trend['数值'], marker='o',
                color='#4ECDC4', linewidth=2)
        ax.fill_between(time_trend['时间'], time_trend['数值'],
                        color='#4ECDC4', alpha=0.1)
        ax.set_title('时间趋势', fontsize=14, fontweight='bold')
        ax.set_xlabel('时间')
        ax.set_ylabel('数值')
        ax.grid(True, alpha=0.3)
        plt.tight_layout()
        paths = self.save_figure(fig, 'time_trend')
        plt.close()
        return paths

    def generate_all_charts(self, analysis_results, company_name='林华'):
        """
        生成所有图表

        Args:
            analysis_results: 分析结果
            company_name: 公司名称

        Returns:
            dict: 所有图表路径
        """
        chart_paths = {}

        try:
            core_dim = analysis_results.get('核心维度')

            # 市场部优先图：核心维度/城市/目录/大类/覆盖
            core_share = analysis_results.get('核心维度分布')
            if core_share is not None and not core_share.empty and core_dim:
                chart_paths['core_share'] = self.create_core_share_chart(core_share, core_dim)

            city_share = analysis_results.get('城市分布')
            if city_share is not None and not city_share.empty:
                chart_paths['city_share'] = self.create_city_share_chart(city_share)

            category_share = analysis_results.get('目录分布')
            if category_share is not None and not category_share.empty:
                chart_paths['category_share'] = self.create_category_share_chart(category_share)

            major_share = analysis_results.get('大类分布')
            if major_share is not None and not major_share.empty:
                chart_paths['major_share'] = self.create_major_share_chart(major_share)

            coverage = analysis_results.get('覆盖分析')
            if coverage is not None and not coverage.empty and core_dim:
                chart_paths['coverage'] = self.create_coverage_chart(coverage, core_dim)

            # 机会/结构/时间/相关性：有意义时再补
            corr_matrix = analysis_results.get('相关性矩阵')
            if corr_matrix is not None:
                chart_paths['correlation_heatmap'] = self.create_correlation_heatmap(corr_matrix)

            time_trend = analysis_results.get('时间趋势')
            if time_trend is not None and not time_trend.empty and len(time_trend) > 2:
                chart_paths['time_trend'] = self.create_time_series_chart(time_trend)

            if '产品结构' in analysis_results and not analysis_results['产品结构'].empty:
                chart_paths['product_structure'] = self.create_product_structure_chart(
                    analysis_results['产品结构'])

            # 机会城市/医院等（若后续需要，可在有字段时扩展）
            if '机会城市' in analysis_results and analysis_results.get('机会城市') is not None and not analysis_results['机会城市'].empty:
                chart_paths['city_opportunities'] = self.create_city_heatmap(
                    analysis_results['机会城市'], company_name)

            # 兜底：若未生成任何业务图，再回退到数值/分类概览
            if not chart_paths:
                numeric_data = analysis_results.get('数值列统计')
                if numeric_data is not None and not numeric_data.empty and len(numeric_data) > 1:
                    has_variation = (numeric_data['标准差'].fillna(0).abs() > 0).any()
                    if has_variation:
                        chart_paths['numeric_overview'] = self.create_numeric_overview_chart(numeric_data)
                categorical_data = analysis_results.get('分类分布')
                if core_dim and categorical_data is not None and not categorical_data.empty:
                    data_core = categorical_data[categorical_data['字段'] == core_dim]
                    if not data_core.empty:
                        total = data_core['数量'].sum()
                        if total > 0 and data_core['数量'].max() > 1:
                            chart_paths['categorical_topn'] = self.create_categorical_topn_chart(data_core)

        except Exception as e:
            print(f"图表生成过程中出现错误: {e}")

        return chart_paths
