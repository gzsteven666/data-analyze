#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
信息图HTML生成器（精简版）
将分析结果以“数即结论”的方式呈现，结构与Word报告保持一致，不强行出无意义板块。
"""

import os
import base64
from datetime import datetime
from jinja2 import Template


class InfographicGenerator:
    """信息图生成器"""

    def __init__(self, template_dir='templates', output_dir='生成结果信息图'):
        self.template_dir = template_dir
        self.output_dir = output_dir
        self.create_output_dir()

    def create_output_dir(self):
        """创建输出目录"""
        os.makedirs(self.output_dir, exist_ok=True)

    def encode_image_to_base64(self, image_path):
        """将图片转换为base64编码"""
        try:
            with open(image_path, "rb") as image_file:
                encoded = base64.b64encode(image_file.read()).decode()
                return f"data:image/png;base64,{encoded}"
        except Exception:
            return ""

    def build_context(self, analysis_results, chart_paths=None, company_name=''):
        """从分析结果构建用于渲染HTML的上下文"""
        core_dim = analysis_results.get('核心维度') or '核心维度'
        core_metric = analysis_results.get('核心指标列') or '核心指标'
        concentration = analysis_results.get('集中度分析')
        head_names = analysis_results.get('头部名单') or []
        tail_names = analysis_results.get('尾部名单') or []
        price_info = analysis_results.get('价格分析') or {}
        structure_summary = analysis_results.get('结构概览')
        trend_summary = analysis_results.get('时间趋势摘要')

        def fmt_list(names, max_n=5):
            names = [str(n) for n in names if n]
            if not names:
                return ""
            if len(names) > max_n:
                return "、".join(names[:max_n]) + " 等"
            return "、".join(names)

        # 价格段落
        price_summary = None
        if price_info:
            corr = price_info.get('相关系数')
            relation = price_info.get('相关描述')
            corr_txt = None
            if corr is not None and relation:
                corr_txt = f"{price_info.get('价格列')} 与 {core_metric} 的相关系数约 {corr:.2f}，{relation}"
            lphv = fmt_list(price_info.get('低价高量') or [])
            hplv = fmt_list(price_info.get('高价低量') or [])
            segments = []
            if lphv:
                segments.append(f"低价高量代表：{lphv}")
            if hplv:
                segments.append(f"高价低量代表：{hplv}")
            price_summary = {
                'corr': corr_txt,
                'quadrant': "；".join(segments) if segments else ""
            }

        # 结构段落
        structure_text = None
        if structure_summary:
            main_type = structure_summary.get('主类型')
            main_share = structure_summary.get('主类型占比')
            others = structure_summary.get('其他占比')
            if main_type is not None and main_share is not None:
                if others:
                    structure_text = f"主要来自 {main_type}（约 {main_share:.1f}%），其余：{others}"
                else:
                    structure_text = f"几乎全部来自 {main_type}（约 {main_share:.1f}%），结构相对单一"

        # 时间趋势
        trend_text = None
        if trend_summary:
            start_v = trend_summary.get('起始')
            end_v = trend_summary.get('当前')
            dir_text = trend_summary.get('方向')
            if start_v is not None and end_v is not None and dir_text:
                trend_text = f"趋势：{dir_text}（{start_v:,.1f} → {end_v:,.1f}）"

        # 核心结论：总量/CR3/CR5/头部/主力/覆盖
        core_share = analysis_results.get('核心维度分布')
        city_share = analysis_results.get('城市分布')
        cat_share = analysis_results.get('目录分布')
        coverage_df = analysis_results.get('覆盖分析')
        total_val = None
        core_top = None
        city_top = None
        cat_top = None
        coverage_leader = None
        if core_share is not None and not core_share.empty:
            total_val = float(core_share.iloc[:, 1].sum())
            core_top = core_share.iloc[0]
        if city_share is not None and not city_share.empty:
            city_top = city_share.iloc[0]
        if cat_share is not None and not cat_share.empty:
            cat_top = cat_share.iloc[0]
        if coverage_df is not None and not coverage_df.empty:
            coverage_leader = coverage_df.iloc[0]

        core_summary = []
        if total_val is not None:
            txt = f"总量约 {total_val:,.0f}"
            if concentration:
                txt += f"，CR3≈ {concentration['Top3占比']:.1f}%，CR5≈ {concentration['Top5占比']:.1f}%"
            core_summary.append(txt + "。")
        if core_top is not None:
            core_summary.append(f"龙头：{core_top[core_dim]}，约 {core_top['份额(%)']:.2f}% / {core_top.iloc[1]:,.0f}。")
        if cat_top is not None:
            core_summary.append(f"主力目录：{cat_top['目录名称']}（{cat_top.iloc[1]:,.0f}）。")
        if city_top is not None:
            core_summary.append(f"主力城市：{city_top['城市']}（{city_top.iloc[1]:,.0f}）。")
        if coverage_leader is not None and '覆盖医院数' in coverage_leader and '城市覆盖数' in coverage_leader:
            core_summary.append(
                f"覆盖：{coverage_leader[core_dim]} 覆盖 {int(coverage_leader['覆盖医院数'])} 家医院、"
                f"{int(coverage_leader['城市覆盖数'])} 个城市，单院均量约 {coverage_leader['单院均量']:,.0f}。"
            )

        # 数据速览 bullet（通用）
        overview = []
        overview.append(f"按 {core_dim} 汇总 {core_metric}")
        if concentration:
            max_ratio = concentration.get('最大值_中位数倍数')
            if max_ratio:
                overview.append(f"最大值约为中位数的 {max_ratio:.1f} 倍")
            overview.append(
                f"集中度：Top1 {concentration['Top1占比']:.1f}% / Top3 {concentration['Top3占比']:.1f}% / Top5 {concentration['Top5占比']:.1f}%"
            )
            overview.append(
                f"覆盖80%需约 {concentration['覆盖80所需实体数']} 个{core_dim}，覆盖90%需约 {concentration['覆盖90所需实体数']} 个"
            )
            overview.append(
                f"离群：高值 {concentration['高值离群数']} 个，低值 {concentration['低值离群数']} 个"
            )
        if head_names:
            overview.append(f"头部{core_dim}：{fmt_list(head_names)}")
        if tail_names:
            overview.append(f"尾部试点：{fmt_list(tail_names)}")
        if price_summary and price_summary.get('corr'):
            txt = price_summary['corr']
            if price_summary.get('quadrant'):
                txt += "；" + price_summary['quadrant']
            overview.append(txt)
        if structure_text:
            overview.append(f"结构：{structure_text}")
        if trend_text:
            overview.append(trend_text)

        # 核心图表（base64）
        chart_images = []
        if chart_paths:
            title_map = {
                'core_share': '核心维度份额',
                'city_share': '城市分布',
                'category_share': '目录份额',
                'major_share': '产品大类分布',
                'coverage': '覆盖与单院均量',
            }
            for key in ['core_share', 'city_share', 'category_share', 'major_share', 'coverage']:
                val = chart_paths.get(key)
                png = None
                if isinstance(val, dict):
                    png = val.get('png')
                elif isinstance(val, (list, tuple)) and len(val) > 0:
                    png = val[0]
                if png and os.path.exists(png):
                    chart_images.append({
                        'title': title_map.get(key, key),
                        'dataurl': self.encode_image_to_base64(png)
                    })

        return {
            'title': f"{company_name}数据分析报告" if company_name else "数据分析报告",
            'subtitle': f"核心维度：{core_dim} | 核心指标：{core_metric}",
            'generation_time': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'overview': overview[:8],
            'core_summary': core_summary,
            'core_dim': core_dim,
            'head_list': fmt_list(head_names),
            'tail_list': fmt_list(tail_names),
            'concentration': concentration,
            'price_summary': price_summary,
            'structure_text': structure_text,
            'trend_text': trend_text,
            'chart_images': chart_images
        }

    def generate_infographic_html(self, analysis_results, chart_paths, insights, company_name=''):
        """生成信息图HTML文本"""
        context = self.build_context(analysis_results, chart_paths, company_name)
        template = Template(self.create_html_template())
        return template.render(**context)

    def save_infographic(self, html_content, filename=None):
        """保存HTML到文件"""
        if not filename:
            filename = "market_analysis_infographic"
        filepath = os.path.join(self.output_dir, f"{filename}.html")
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)
        return filepath

    def create_html_template(self):
        """新的简洁模板：卡片+列表，不强行出图"""
        return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        :root {
            --bg: #f5f7fb;
            --panel: #ffffff;
            --card: #ffffff;
            --text: #111827;
            --muted: #6b7280;
            --accent: #2563eb;
            --border: #e5e7eb;
        }
        body {
            font-family: 'Microsoft YaHei', 'PingFang SC', 'Helvetica', sans-serif;
            background: var(--bg);
            margin: 0;
            padding: 0 0 48px 0;
            color: var(--text);
        }
        .container {
            max-width: 1100px;
            margin: 0 auto;
            padding: 28px 20px;
        }
        .panel {
            background: var(--panel);
            border-radius: 18px;
            padding: 24px;
            box-shadow: 0 14px 28px rgba(0,0,0,0.12);
        }
        .header h1 { margin: 0 0 6px 0; font-size: 28px; }
        .header .subtitle { margin: 0 0 12px 0; color: var(--muted); }
        .chips { margin-top: 6px; font-size: 12px; color: var(--muted); }
        .chip { display: inline-block; background: rgba(37,99,235,0.08); color: var(--text); padding: 4px 10px; border-radius: 999px; margin-right: 8px; }
        .section {
            margin-top: 24px;
            background: var(--card);
            border-radius: 14px;
            padding: 16px 18px;
            border: 1px solid var(--border);
            box-shadow: 0 8px 16px rgba(0,0,0,0.06);
        }
        .section h2 { margin: 0 0 10px 0; font-size: 18px; color: var(--text); }
        ul { margin: 0; padding-left: 18px; color: var(--text); }
        li { margin-bottom: 8px; line-height: 1.6; color: var(--text); }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; }
        .card {
            background: #f8fafc;
            border-radius: 12px;
            padding: 12px 14px;
            border: 1px solid var(--border);
        }
        .card .label { font-size: 12px; color: var(--muted); margin-bottom: 4px; }
        .card .value { font-size: 18px; font-weight: 600; color: var(--text); }
        p { margin: 6px 0; color: var(--text); line-height: 1.6; }
    </style>
</head>
<body>
        <div class="container">
        <div class="panel header">
            <h1>{{ title }}</h1>
            <div class="subtitle">{{ subtitle }}</div>
            <div class="chips">
                <span class="chip">生成时间: {{ generation_time }}</span>
                <span class="chip">核心维度: {{ core_dim }}</span>
            </div>
        </div>

        <div class="section">
            <h2>数据速览</h2>
            <ul>
                {% for item in overview %}
                <li>{{ item }}</li>
                {% endfor %}
            </ul>
        </div>

        {% if core_summary %}
        <div class="section">
            <h2>核心结论</h2>
            <ul>
                {% for item in core_summary %}
                <li>{{ item }}</li>
                {% endfor %}
            </ul>
        </div>
        {% endif %}

        {% if concentration %}
        <div class="section">
            <h2>集中度与结构</h2>
            <div class="grid">
                <div class="card">
                    <div class="label">Top1/Top3/Top5</div>
                    <div class="value">{{ concentration.Top1占比|round(1) }}% / {{ concentration.Top3占比|round(1) }}% / {{ concentration.Top5占比|round(1) }}%</div>
                </div>
                <div class="card">
                    <div class="label">覆盖80%/90%需要</div>
                    <div class="value">{{ concentration.覆盖80所需实体数 }} / {{ concentration.覆盖90所需实体数 }} 个</div>
                </div>
                {% if concentration.最大值_中位数倍数 %}
                <div class="card">
                    <div class="label">长尾程度</div>
                    <div class="value">最大值约为中位数 {{ concentration.最大值_中位数倍数|round(1) }} 倍</div>
                </div>
                {% endif %}
                <div class="card">
                    <div class="label">离群</div>
                    <div class="value">高值 {{ concentration.高值离群数 }} / 低值 {{ concentration.低值离群数 }}</div>
                </div>
            </div>
            {% if head_list %}
            <p>头部{{ core_dim }}：{{ head_list }}</p>
            {% endif %}
            {% if tail_list %}
            <p>尾部试点：{{ tail_list }}</p>
            {% endif %}
        </div>
        {% endif %}

        {% if price_summary and price_summary.corr %}
        <div class="section">
            <h2>价格-量</h2>
            <ul>
                <li>{{ price_summary.corr }}</li>
                {% if price_summary.quadrant %}<li>{{ price_summary.quadrant }}</li>{% endif %}
            </ul>
        </div>
        {% endif %}

        {% if structure_text %}
        <div class="section">
            <h2>结构机会</h2>
            <p>{{ structure_text }}</p>
        </div>
        {% endif %}

        {% if trend_text %}
        <div class="section">
            <h2>时间趋势</h2>
            <p>{{ trend_text }}</p>
        </div>
        {% endif %}

        {% if chart_images %}
        <div class="section">
            <h2>核心图表</h2>
            {% for item in chart_images %}
            <div class="card" style="margin-bottom:12px;">
                <div class="label">{{ item.title }}</div>
                <img src="{{ item.dataurl }}" alt="{{ item.title }}" style="max-width:100%; border-radius:6px; border:1px solid var(--border);" />
            </div>
            {% endfor %}
        </div>
        {% endif %}
    </div>
</body>
</html>
        """
