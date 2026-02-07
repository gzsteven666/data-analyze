#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据分析核心模块
基于参考项目的数据分析流程实现
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib import font_manager
import warnings
warnings.filterwarnings('ignore')

class DataAnalyzer:
    """数据分析器类"""

    def __init__(self):
        self.setup_chinese_font()
        self.setup_plot_style()

    def preprocess_data(self, df):
        """
        预处理与列名归一：
        - 去除列名首尾空白
        - 常见别名标准化：品牌/企业、城市、产品大类
        - 关键数值列尝试转为数值型，便于自动识别核心指标
        """
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        # 统一品牌/企业 -> 品牌名称
        if '品牌名称' not in df.columns:
            for cand in ['申报企业名称', '申报企业', '企业', '厂家', '生产企业', '企业名称', '品牌']:
                if cand in df.columns:
                    df['品牌名称'] = df[cand].astype(str).str.strip()
                    break

        # 注册证产品名称别名
        if '注册证产品名称' not in df.columns:
            for cand in ['产品名称', '注册证名称', '产品']:
                if cand in df.columns:
                    df['注册证产品名称'] = df[cand].astype(str).str.strip()
                    break

        # 统一城市
        if '城市' not in df.columns:
            for cand in ['地市', '地区', '地区名称', '城市']:
                if cand in df.columns:
                    if cand == '地区名称':
                        df['城市'] = df[cand].astype(str).str.split('>').str[1].fillna(df[cand].astype(str)).str.strip()
                    else:
                        df['城市'] = df[cand].astype(str).str.strip()
                    break

        # 产品大类派生自目录名称
        if '产品大类' not in df.columns and '目录名称' in df.columns:
            df['产品大类'] = df['目录名称'].astype(str).str.split('-', n=1).str[0].str.strip()

        # 关键数值列尝试转数值
        numeric_hints = ['协议采购量', '采购需求量', '第三年采购需求量', '数量', '总量', '合同量', 'GMV', '金额', '采购量']
        for col in numeric_hints:
            if col in df.columns:
                converted = pd.to_numeric(df[col], errors='coerce')
                # 仅在存在有效数值时覆盖
                if converted.notna().any():
                    df[col] = converted

        # 去除对象列首尾空格
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()

        return df

    def setup_chinese_font(self):
        """配置中文字体，解决乱码问题"""
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

    def setup_plot_style(self):
        """设置图表样式"""
        plt.style.use('seaborn-v0_8')
        sns.set_palette("husl")

    def get_numeric_columns(self, df):
        """获取数值字段列表"""
        return df.select_dtypes(include=[np.number]).columns.tolist()

    def get_categorical_columns(self, df):
        """获取分类字段列表"""
        return df.select_dtypes(include=['object', 'category']).columns.tolist()

    def detect_core_dimension(self, df, preferred=None):
        """
        自动/手动确定核心实体维度：
        - 优先使用传入的 preferred
        - 否则按照常见业务字段候选列表匹配
        """
        if preferred and preferred in df.columns:
            return preferred
        candidates = [
            '申报企业名称', '申报企业', '品牌名称', '品牌', '厂家', '生产企业', '企业名称',
            '医疗机构名称', '医院名称', '客户', '客户名称', '渠道', '渠道名称', '门店', '门店名称',
            '品类', '产品', '商品', 'SKU', '区域', '城市', '省份', '地区', '院点'
        ]
        for col in candidates:
            if col in df.columns:
                return col
        # fallback: 最大非唯一的分类列
        cat_cols = self.get_categorical_columns(df)
        if cat_cols:
            # 过滤掉几乎全唯一的列
            best = None
            best_unique = 0
            total_rows = len(df)
            for c in cat_cols:
                u = df[c].nunique(dropna=True)
                if total_rows > 0 and u / total_rows >= 0.9:
                    continue
                if u > best_unique:
                    best_unique = u
                    best = c
            if best:
                return best
        return None

    def analyze_concentration(self, df, entity_col, metric_col):
        """
        按核心实体维度分析集中度与头部/长尾结构

        Args:
            df: 数据
            entity_col: 核心实体列（如医院/客户/渠道等）
            metric_col: 核心指标列（如总约定量/GMV等）

        Returns:
            dict 或 None
        """
        if not entity_col or not metric_col:
            return None
        if entity_col not in df.columns or metric_col not in df.columns:
            return None

        series = df.groupby(entity_col)[metric_col].sum().sort_values(ascending=False)
        series = series[series > 0]
        if series.empty:
            return None

        total = series.sum()
        n_entities = len(series)
        median = float(series.median())
        max_v = float(series.iloc[0])
        max_to_median = max_v / median if median > 0 else None

        def share(k):
            if n_entities == 0:
                return 0
            return float(series.iloc[:min(k, n_entities)].sum() / total * 100)

        top1_share = share(1)
        top3_share = share(3)
        top5_share = share(5)

        # 覆盖80%/90%所需实体数
        cumsum = series.cumsum()
        # 覆盖80%/90%所需实体数（按位置）
        need_80_count = n_entities
        need_90_count = n_entities
        frac = cumsum / total
        if (frac >= 0.8).any():
            # idxmax 返回的是实体标签，这里用位置索引
            pos_80 = (frac >= 0.8).to_numpy().argmax()
            need_80_count = int(pos_80 + 1)
        if (frac >= 0.9).any():
            pos_90 = (frac >= 0.9).to_numpy().argmax()
            need_90_count = int(pos_90 + 1)

        # 使用均值±1.5*标准差识别离群
        mean_v = float(series.mean())
        std_v = float(series.std())
        low_thresh = mean_v - 1.5 * std_v
        high_thresh = mean_v + 1.5 * std_v
        low_outliers = int((series < low_thresh).sum())
        high_outliers = int((series > high_thresh).sum())

        return {
            '实体数': n_entities,
            '总量': float(total),
            '中位数': median,
            '最大值': max_v,
            '最大值_中位数倍数': max_to_median,
            'Top1占比': top1_share,
            'Top3占比': top3_share,
            'Top5占比': top5_share,
            '覆盖80所需实体数': need_80_count,
            '覆盖90所需实体数': need_90_count,
            '均值': mean_v,
            '标准差': std_v,
            '低值离群数': low_outliers,
            '高值离群数': high_outliers,
            '低值阈值': low_thresh,
            '高值阈值': high_thresh
        }

    def _normalize_score(self, series):
        """
        将数值序列归一化到 0-100 分。
        """
        s = pd.to_numeric(series, errors='coerce').fillna(0.0)
        min_v = float(s.min())
        max_v = float(s.max())
        if max_v - min_v < 1e-9:
            return pd.Series(np.full(len(s), 50.0), index=s.index)
        return ((s - min_v) / (max_v - min_v) * 100.0).clip(0.0, 100.0)

    def build_opportunity_priority(
        self,
        df,
        entity_col,
        total_col,
        share_col,
        target_volume_col=None,
        top_n=20,
    ):
        """
        机会优先级打分（影响 × 可行性 × 投入效率）
        - 影响分：潜在增量规模（总量 × 份额缺口）
        - 可行性分：已有基础（目标品牌量）或份额缺口替代
        - 投入效率分：按总量近似投入强度，投入越低分越高
        """
        if df is None or df.empty:
            return pd.DataFrame()
        required = {entity_col, total_col, share_col}
        if not required.issubset(df.columns):
            return pd.DataFrame()

        work = df[[c for c in [entity_col, total_col, share_col, target_volume_col] if c and c in df.columns]].copy()
        work = work.dropna(subset=[entity_col, total_col])
        if work.empty:
            return pd.DataFrame()

        work[entity_col] = work[entity_col].astype(str).str.strip()
        work = work[work[entity_col] != '']
        work = work[~work[entity_col].str.lower().isin(['nan', 'none', 'null', 'na', 'n/a'])]
        if work.empty:
            return pd.DataFrame()

        work[total_col] = pd.to_numeric(work[total_col], errors='coerce').fillna(0.0)
        work[share_col] = pd.to_numeric(work[share_col], errors='coerce').fillna(0.0).clip(0.0, 100.0)
        work = work[work[total_col] > 0].copy()
        if work.empty:
            return pd.DataFrame()

        # 影响：潜在增量（总量 * 未覆盖份额）
        gap_ratio = (100.0 - work[share_col]) / 100.0
        work['_impact_raw'] = work[total_col] * gap_ratio

        # 可行性：优先参考已有基础（目标品牌量）；无则用份额缺口替代
        if target_volume_col and target_volume_col in work.columns:
            work[target_volume_col] = pd.to_numeric(work[target_volume_col], errors='coerce').fillna(0.0)
            work['_feas_raw'] = np.log1p(work[target_volume_col])
        else:
            work['_feas_raw'] = gap_ratio

        # 投入：以总量近似投入强度（规模越大通常投入越高）
        work['_invest_raw'] = np.log1p(work[total_col])

        work['影响分'] = self._normalize_score(work['_impact_raw'])
        work['可行性分'] = self._normalize_score(work['_feas_raw'])
        invest_score = self._normalize_score(work['_invest_raw'])
        work['投入效率分'] = (100.0 - invest_score).clip(0.0, 100.0)

        # 立方均值近似“影响 × 可行性 × 投入效率”后的可读分值
        product = (work['影响分'] * work['可行性分'] * work['投入效率分']).clip(lower=0.0)
        work['综合优先级分'] = np.power(product, 1.0 / 3.0)

        def _priority_level(score):
            if score >= 70:
                return '高优先'
            if score >= 55:
                return '中高优先'
            if score >= 40:
                return '中优先'
            return '观察'

        def _reason(row):
            i = row['影响分']
            f = row['可行性分']
            e = row['投入效率分']
            if i >= 70 and f >= 60:
                base = '潜在增量大且已有基础，适合优先推进'
            elif i >= 70 and f < 60:
                base = '潜在增量大但基础较弱，建议先小范围试点'
            elif i < 70 and f >= 60:
                base = '增量中等但落地快，可作为稳健补充'
            else:
                base = '增量与可行性一般，建议低成本跟踪'
            if e < 40:
                return base + '，预计投入强度较高'
            if e >= 70:
                return base + '，投入效率较好'
            return base

        work['优先级'] = work['综合优先级分'].map(_priority_level)
        work['优先级理由'] = work.apply(_reason, axis=1)

        sort_cols = ['综合优先级分', '影响分', '可行性分']
        work = work.sort_values(sort_cols, ascending=False).reset_index(drop=True)
        if top_n and top_n > 0:
            work = work.head(top_n)

        keep_cols = [entity_col, total_col, share_col]
        if target_volume_col and target_volume_col in work.columns:
            keep_cols.append(target_volume_col)
        keep_cols += ['影响分', '可行性分', '投入效率分', '综合优先级分', '优先级', '优先级理由']
        return work[keep_cols]

    def generate_field_overview(self, df):
        """生成字段画像"""
        overview = []
        memory_usage = df.memory_usage(deep=True)
        for col in df.columns:
            series = df[col]
            sample = series.dropna()
            overview.append({
                '字段': col,
                '类型': str(series.dtype),
                '缺失率(%)': round(series.isna().mean() * 100, 2),
                '唯一值数': int(series.nunique(dropna=True)),
                '非空数量': int(series.count()),
                '内存占用(MB)': round(memory_usage[col] / 1024**2, 3),
                '示例值': str(sample.iloc[0]) if not sample.empty else ''
            })
        return pd.DataFrame(overview)

    def summarize_numeric_columns(self, df):
        """生成数值字段统计"""
        numeric_cols = self.get_numeric_columns(df)
        rows = []
        for col in numeric_cols:
            series = df[col].dropna()
            if series.empty:
                continue
            rows.append({
                '字段': col,
                '均值': series.mean(),
                '中位数': series.median(),
                '最小值': series.min(),
                '最大值': series.max(),
                '标准差': series.std(),
                '非空数量': int(series.count())
            })
        return pd.DataFrame(rows)

    def summarize_categorical_columns(self, df, top_n=5, core_dim=None):
        """生成分类字段TopN分布，核心实体维度优先"""
        categorical_cols = self.get_categorical_columns(df)
        rows = []
        total_rows = len(df)
        # 如果指定了核心维度，则仅针对该维度做TopN，其它维度跳过
        target_cols = [core_dim] if core_dim and core_dim in categorical_cols else categorical_cols
        for col in categorical_cols:
            if col not in target_cols:
                continue
            unique_cnt = df[col].nunique(dropna=True)
            # 跳过无信息或几乎全唯一的列，避免输出无意义的TopN
            if unique_cnt <= 1:
                continue
            if total_rows > 0 and unique_cnt / total_rows >= 0.9:
                continue

            counts = df[col].astype(str).value_counts(dropna=True).head(top_n)
            for category, count in counts.items():
                rows.append({
                    '字段': col,
                    '类别': category,
                    '数量': int(count),
                    '占比(%)': round(count / total_rows * 100, 2)
                })
        return pd.DataFrame(rows)

    def compute_correlation_matrix(self, df):
        """计算相关性矩阵"""
        numeric_cols = self.get_numeric_columns(df)
        if len(numeric_cols) < 2:
            return None
        corr = df[numeric_cols].corr()
        # 如果全部为NaN或仅对角线为1，视为无有效相关性
        if corr.replace(1.0, 0).abs().sum().sum() == 0:
            return None
        if corr.isnull().all().all():
            return None
        # 若最高相关系数低于0.3，视为无实质相关性
        corr_abs = corr.abs().copy()
        corr_abs.values[[range(len(corr_abs))]*2] = 0
        if corr_abs.max().max() < 0.3:
            return None
        return corr

    def detect_datetime_columns(self, df):
        """检测日期字段"""
        datetime_cols = []
        for col in df.columns:
            series = df[col]
            if np.issubdtype(series.dtype, np.datetime64):
                datetime_cols.append(col)
                continue
            if series.dtype == object:
                parsed = pd.to_datetime(series, errors='coerce', infer_datetime_format=True)
                if parsed.notna().mean() > 0.7:
                    df[col] = parsed
                    datetime_cols.append(col)
        return datetime_cols

    def analyze_time_series(self, df, time_column=None, value_column=None):
        """生成时间趋势分析"""
        candidate_cols = self.detect_datetime_columns(df)
        if time_column and time_column in df.columns:
            candidate_cols = [time_column] + [col for col in candidate_cols if col != time_column]

        if not candidate_cols:
            return None

        time_col = candidate_cols[0]
        data = df.copy()
        data[time_col] = pd.to_datetime(data[time_col], errors='coerce')
        data = data.dropna(subset=[time_col])
        if data.empty:
            return None

        numeric_cols = self.get_numeric_columns(data)
        metric_col = value_column if value_column in numeric_cols else (numeric_cols[0] if numeric_cols else None)

        data = data.sort_values(time_col)
        span_days = (data[time_col].max() - data[time_col].min()).days if len(data) > 1 else 0
        if span_days > 730:
            freq = 'Q'
        elif span_days > 180:
            freq = 'M'
        elif span_days > 30:
            freq = 'W'
        else:
            freq = 'D'

        if metric_col:
            grouped = data.groupby(pd.Grouper(key=time_col, freq=freq))[metric_col].sum().reset_index()
            grouped.columns = ['时间', '数值']
        else:
            grouped = data.groupby(pd.Grouper(key=time_col, freq=freq)).size().reset_index(name='数值')
            grouped.rename(columns={time_col: '时间'}, inplace=True)

        return grouped

    def load_data(self, file_path, sheet_name=None):
        """
        加载数据文件

        Args:
            file_path: 文件路径
            sheet_name: Excel工作表名称（可选）

        Returns:
            DataFrame: 加载的数据
        """
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path, encoding='utf-8-sig')
        elif file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            raise ValueError(f"不支持的文件格式: {file_path}")

    def data_health_check(self, df):
        """
        数据体检

        Args:
            df: 输入数据

        Returns:
            dict: 体检报告
        """
        report = {
            '基本信息': {
                '记录数': len(df),
                '字段数': len(df.columns),
                '内存使用': f"{df.memory_usage(deep=True).sum() / 1024**2:.2f} MB"
            },
            '缺失值': df.isnull().sum().to_dict(),
            '重复记录': df.duplicated().sum(),
            '数据类型': df.dtypes.to_dict(),
            '唯一值数量': df.nunique().to_dict()
        }

        # 数值字段统计
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            report['数值统计'] = df[numeric_cols].describe().to_dict()

        return report

    def filter_iv_catheter(self, df, product_col='品种名称'):
        """
        过滤留置针数据

        Args:
            df: 原始数据
            product_col: 产品名称字段

        Returns:
            DataFrame: 过滤后的留置针数据
        """
        mask = df[product_col].str.contains('留置针|静脉留置针', na=False, regex=True)
        return df[mask].copy()

    def calculate_market_share(self, df, group_cols, quantity_col='第三年采购需求量'):
        """
        计算市场份额

        Args:
            df: 数据
            group_cols: 分组字段
            quantity_col: 数量字段

        Returns:
            DataFrame: 包含市场份额的结果
        """
        result = df.groupby(group_cols)[quantity_col].agg(['sum', 'count']).reset_index()
        result.columns = list(group_cols) + ['总量', '覆盖数']

        total = result['总量'].sum()
        result['市场份额'] = result['总量'] / total * 100
        result = result.sort_values('总量', ascending=False)

        return result

    def identify_opportunity_cities(self, df, company_name='林华',
                                   quantity_col='第三年采购需求量',
                                   city_col='城市', company_col='品牌名称'):
        """
        识别机会城市（白区）

        Args:
            df: 数据
            company_name: 目标公司名称
            quantity_col: 数量字段
            city_col: 城市字段
            company_col: 公司字段

        Returns:
            DataFrame: 机会城市列表
        """
        city_stats = df.groupby(city_col).agg({
            quantity_col: 'sum',
            company_col: lambda x: (x == company_name).sum()
        }).reset_index()
        city_stats.columns = ['城市', '总容量', '林华覆盖数']

        # 计算林华份额
        company_share = df.groupby([city_col, company_col])[quantity_col].sum().unstack(fill_value=0)
        if company_name in company_share.columns:
            city_stats['林华份额'] = company_share[company_name] / company_share.sum(axis=1) * 100
        else:
            city_stats['林华份额'] = 0

        # 识别白区：容量大但份额低的城市
        opportunity = city_stats[
            (city_stats['总容量'] >= city_stats['总容量'].quantile(0.7)) &
            (city_stats['林华份额'] <= city_stats['林华份额'].quantile(0.3))
        ].sort_values('总容量', ascending=False)

        return opportunity

    def analyze_hospital_opportunities(self, df, company_name='林华',
                                     quantity_col='第三年采购需求量',
                                     hospital_col='医疗机构名称',
                                     company_col='品牌名称'):
        """
        分析医院机会

        Args:
            df: 数据
            company_name: 目标公司名称
            quantity_col: 数量字段
            hospital_col: 医院字段
            company_col: 公司字段

        Returns:
            DataFrame: 医院机会列表
        """
        hospital_stats = df.groupby(hospital_col).agg({
            quantity_col: 'sum',
            company_col: lambda x: list(x.unique())
        }).reset_index()
        hospital_stats.columns = ['医院名称', '总容量', '品牌列表']

        # 计算每个医院的林华份额
        hospital_share = df.groupby([hospital_col, company_col])[quantity_col].sum().unstack(fill_value=0)
        if company_name in hospital_share.columns:
            hospital_stats['林华份额'] = hospital_share[company_name] / hospital_share.sum(axis=1) * 100
        else:
            hospital_stats['林华份额'] = 0

        hospital_stats['覆盖状态'] = hospital_stats['林华份额'].apply(
            lambda x: '未覆盖' if x == 0 else '低份额' if x < 10 else '已覆盖'
        )

        # 机会医院：高容量但低份额或未覆盖
        opportunities = hospital_stats[
            (hospital_stats['总容量'] >= hospital_stats['总容量'].quantile(0.8)) &
            (hospital_stats['林华份额'] < 20)
        ].sort_values('总容量', ascending=False)

        return opportunities

    def analyze_product_structure(self, df, structure_col='类别名称',
                               quantity_col='第三年采购需求量',
                               company_col='品牌名称'):
        """
        分析产品结构（安全型vs普通型）

        Args:
            df: 数据
            structure_col: 产品类别字段
            quantity_col: 数量字段
            company_col: 公司字段

        Returns:
            DataFrame: 产品结构分析
        """
        # 判断安全型
        df['产品类型'] = df[structure_col].apply(
            lambda x: '安全型' if pd.notna(x) and '安全' in str(x) else '普通型'
        )

        structure_analysis = df.groupby(['产品类型', company_col])[quantity_col].sum().unstack(fill_value=0)
        structure_pct = structure_analysis.div(structure_analysis.sum(axis=1), axis=0) * 100

        return structure_pct.reset_index()

    def create_comprehensive_analysis(self, df, company_name=None, config=None):
        """
        创建综合分析

        Args:
            df: 数据
            company_name: 目标公司名称
            config: 额外配置

        Returns:
            dict: 综合分析结果
        """
        results = {}
        config = config or {}
        df = self.preprocess_data(df)
        core_dim = self.detect_core_dimension(df, config.get('core_dimension'))
        results['核心维度'] = core_dim
        target_brand = config.get('target_brand') or company_name

        # 通用分析
        results['字段概览'] = self.generate_field_overview(df)
        numeric_summary = self.summarize_numeric_columns(df)
        if not numeric_summary.empty:
            results['数值列统计'] = numeric_summary

        # 核心指标列：优先使用配置中的 value_column，否则选首个数值列
        numeric_cols = self.get_numeric_columns(df)
        core_metric = None
        preferred_metrics = []
        if config.get('value_column'):
            preferred_metrics.append(config.get('value_column'))
        preferred_metrics += ['协议采购量', '采购需求量', '第三年采购需求量', '数量', '总量', '合同量', 'GMV', '金额', '采购量']
        for col in preferred_metrics:
            if col in numeric_cols:
                core_metric = col
                break
        if core_metric is None and numeric_cols:
            core_metric = numeric_cols[0]
        results['核心指标列'] = core_metric

        # 数据层级与可用维度
        has_product = '注册证产品名称' in df.columns and df['注册证产品名称'].nunique(dropna=True) > 1
        has_catalog = '目录名称' in df.columns and df['目录名称'].nunique(dropna=True) > 1
        has_major = '产品大类' in df.columns and df['产品大类'].nunique(dropna=True) > 1
        has_city = '城市' in df.columns and df['城市'].nunique(dropna=True) > 1
        hosp_col = None
        for cand in ['医院名称', '医疗机构名称']:
            if cand in df.columns:
                hosp_col = cand
                break

        # 通用口径聚合（按市场视角）
        if core_dim and core_metric and core_dim in df.columns and core_metric in df.columns:
            core_share = (
                df.groupby(core_dim, as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            total_core = core_share[core_metric].sum()
            if total_core > 0:
                core_share['份额(%)'] = core_share[core_metric] / total_core * 100
            results['核心维度分布'] = core_share

        if '城市' in df.columns and core_metric and core_metric in df.columns:
            city_share = (
                df.groupby('城市', as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            total_city = city_share[core_metric].sum()
            if total_city > 0:
                city_share['份额(%)'] = city_share[core_metric] / total_city * 100
            results['城市分布'] = city_share

        if '目录名称' in df.columns and core_metric and core_metric in df.columns:
            category_share = (
                df.groupby('目录名称', as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            total_cat = category_share[core_metric].sum()
            if total_cat > 0:
                category_share['份额(%)'] = category_share[core_metric] / total_cat * 100
            results['目录分布'] = category_share

        if '产品大类' in df.columns and core_metric and core_metric in df.columns:
            major_share = (
                df.groupby('产品大类', as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            total_major = major_share[core_metric].sum()
            if total_major > 0:
                major_share['份额(%)'] = major_share[core_metric] / total_major * 100
            results['大类分布'] = major_share

        # 覆盖分析（需核心维度、医院或门店）
        if core_dim and core_metric and hosp_col and core_dim in df.columns and core_metric in df.columns:
            coverage_df = (
                df.groupby(core_dim)
                .agg(
                    总量=(core_metric, 'sum'),
                    覆盖医院数=(hosp_col, 'nunique'),
                    城市覆盖数=('城市', 'nunique') if '城市' in df.columns else (core_metric, 'count'),
                )
                .reset_index()
            )
            coverage_df['单院均量'] = coverage_df['总量'] / coverage_df['覆盖医院数'].replace(0, np.nan)
            total_cov = coverage_df['总量'].sum()
            if total_cov > 0:
                coverage_df['份额(%)'] = coverage_df['总量'] / total_cov * 100
            coverage_df = coverage_df.sort_values('总量', ascending=False).reset_index(drop=True)
            results['覆盖分析'] = coverage_df

        # 城市-品牌长表与城市Top3（需城市+核心维度+核心指标）
        if '城市' in df.columns and core_dim and core_metric and core_dim in df.columns and core_metric in df.columns:
            city_brand = (
                df.groupby(['城市', core_dim], as_index=False)[core_metric].sum()
                .sort_values(['城市', core_metric], ascending=[True, False])
            )
            city_total = (
                city_brand.groupby('城市', as_index=False)[core_metric].sum()
                .rename(columns={core_metric: '城市总量'})
            )
            city_brand = city_brand.merge(city_total, on='城市', how='left')
            city_brand['份额(%)'] = city_brand[core_metric] / city_brand['城市总量'] * 100
            results['城市品牌分布'] = city_brand

            records = []
            for city, g in city_brand.groupby('城市'):
                g2 = g.sort_values(core_metric, ascending=False).reset_index(drop=True)
                row = {'城市': city, '城市总量': g2['城市总量'].iloc[0]}
                cr3 = 0.0
                for i in range(3):
                    if i < len(g2):
                        row[f"Top{i+1}_品牌"] = g2.loc[i, core_dim]
                        row[f"Top{i+1}_数量"] = g2.loc[i, core_metric]
                        row[f"Top{i+1}_份额(%)"] = g2.loc[i, '份额(%)']
                        cr3 += g2.loc[i, '份额(%)']
                    else:
                        row[f"Top{i+1}_品牌"] = None
                        row[f"Top{i+1}_数量"] = None
                        row[f"Top{i+1}_份额(%)"] = None
                row["CR3(%)"] = cr3
                records.append(row)
            results['城市Top3'] = pd.DataFrame(records)

        # 医院TOP（如有医院列）
        hosp_col2 = None
        for cand in ['医院名称', '医疗机构名称']:
            if cand in df.columns:
                hosp_col2 = cand
                break
        if hosp_col2 and core_metric and core_metric in df.columns:
            hosp_top = (
                df.groupby([hosp_col2, '城市'] if '城市' in df.columns else [hosp_col2], as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            results['医院TOP'] = hosp_top

        # 产品/注册证名称分布（如有）
        if '注册证产品名称' in df.columns and core_metric and core_metric in df.columns:
            prod = (
                df.groupby('注册证产品名称', as_index=False)[core_metric].sum()
                .sort_values(core_metric, ascending=False)
                .reset_index(drop=True)
            )
            results['产品分布'] = prod
            # 品牌 x 产品结构
            prod_brand = (
                df.groupby(['品牌名称', '注册证产品名称'], as_index=False)[core_metric].sum()
                .sort_values(['品牌名称', core_metric], ascending=[True, False])
            )
            total_pb = prod_brand.groupby('品牌名称')[core_metric].transform('sum')
            prod_brand['份额(%)'] = prod_brand[core_metric] / total_pb * 100
            results['品牌产品分布'] = prod_brand
            # 城市 x 产品（如有城市）
            if '城市' in df.columns and has_city:
                prod_city = (
                    df.groupby(['城市', '注册证产品名称'], as_index=False)[core_metric].sum()
                    .sort_values(['城市', core_metric], ascending=[True, False])
                )
                city_total = prod_city.groupby('城市')[core_metric].transform('sum')
                prod_city['份额(%)'] = prod_city[core_metric] / city_total * 100
                results['城市产品分布'] = prod_city
            # 品牌内部产品Top
            prod_brand_top = prod_brand.sort_values(['品牌名称', core_metric], ascending=[True, False])
            results['品牌产品Top'] = prod_brand_top
            # 品牌 x 产品结构
            prod_brand = (
                df.groupby(['品牌名称', '注册证产品名称'], as_index=False)[core_metric].sum()
                .sort_values(['品牌名称', core_metric], ascending=[True, False])
            )
            total_pb = prod_brand.groupby('品牌名称')[core_metric].transform('sum')
            prod_brand['份额(%)'] = prod_brand[core_metric] / total_pb * 100
            results['品牌产品分布'] = prod_brand
            # 城市 x 产品（如有城市）
            if '城市' in df.columns and has_city:
                prod_city = (
                    df.groupby(['城市', '注册证产品名称'], as_index=False)[core_metric].sum()
                    .sort_values(['城市', core_metric], ascending=[True, False])
                )
                city_total = prod_city.groupby('城市')[core_metric].transform('sum')
                prod_city['份额(%)'] = prod_city[core_metric] / city_total * 100
                results['城市产品分布'] = prod_city
            # 品牌内部产品Top（便于看主打）
            prod_brand_top = prod_brand.sort_values(['品牌名称', core_metric], ascending=[True, False])
            results['品牌产品Top'] = prod_brand_top

        categorical_summary = self.summarize_categorical_columns(df, core_dim=core_dim)
        if not categorical_summary.empty:
            results['分类分布'] = categorical_summary

        corr_matrix = self.compute_correlation_matrix(df)
        if corr_matrix is not None:
            results['相关性矩阵'] = corr_matrix

        time_trend = self.analyze_time_series(
            df,
            time_column=config.get('time_column'),
            value_column=config.get('value_column')
        )
        if time_trend is not None and not time_trend.empty:
            results['时间趋势'] = time_trend

        # 1. 市场份额分析
        if {'品牌名称', '第三年采购需求量'}.issubset(df.columns):
            results['品牌份额'] = self.calculate_market_share(df, ['品牌名称'])

        if {'城市', '第三年采购需求量'}.issubset(df.columns):
            results['城市份额'] = self.calculate_market_share(df, ['城市'])

        # 2. 机会识别
        if {'城市', '品牌名称', '第三年采购需求量'}.issubset(df.columns):
            results['机会城市'] = self.identify_opportunity_cities(df, company_name or '')

        if {'医疗机构名称', '品牌名称', '第三年采购需求量'}.issubset(df.columns):
            results['机会医院'] = self.analyze_hospital_opportunities(df, company_name or '')

        # 3. 产品结构分析
        if {'类别名称', '品牌名称', '第三年采购需求量'}.issubset(df.columns):
            results['产品结构'] = self.analyze_product_structure(df)

        # 4. 覆盖分析
        if {'品牌名称', '医疗机构名称', '第三年采购需求量'}.issubset(df.columns):
            coverage_analysis = df.groupby('品牌名称').agg({
                '医疗机构名称': 'nunique',
                '第三年采购需求量': ['sum', 'mean']
            }).reset_index()
            coverage_analysis.columns = ['品牌名称', '覆盖医院数', '总量', '单院均量']
            results['覆盖分析'] = coverage_analysis.sort_values('总量', ascending=False)

        # 5. 集中度分析（基于核心实体维度和核心指标列）
        concentration = self.analyze_concentration(df, core_dim, core_metric)
        if concentration is not None:
            results['集中度分析'] = concentration

        # 头部/尾部名单（基于核心维度 + 核心指标）
        head_names = []
        tail_names = []
        def _clean_entity_names(values):
            cleaned = []
            for v in values:
                if pd.isna(v):
                    continue
                s = str(v).strip()
                if not s:
                    continue
                if s.lower() in {"nan", "none", "null", "na", "n/a"}:
                    continue
                cleaned.append(s)
            return cleaned
        if core_dim and core_metric and core_dim in df.columns and core_metric in df.columns:
            group = df.groupby(core_dim)[core_metric].sum().sort_values(ascending=False)
            total = group.sum()
            if total > 0:
                frac = group.cumsum() / total
                head_80 = list(frac[frac <= 0.8].index)
                if not head_80 and len(group) > 0:
                    head_80 = [group.index[0]]
                head_names = _clean_entity_names(head_80[:5])
                tail_names = _clean_entity_names(list(group.sort_values(ascending=True).head(5).index))
        results['头部名单'] = head_names
        results['尾部名单'] = tail_names

        # 白区/机会：目标品牌欠份额的城市/医院
        if target_brand and core_dim and core_metric and '城市' in df.columns and core_dim in df.columns and core_metric in df.columns:
            cb = results.get('城市品牌分布')
            if cb is not None and not cb.empty:
                city_total = cb[['城市', '城市总量']].drop_duplicates()
                tgt = cb[cb[core_dim] == target_brand][['城市', core_metric]].rename(columns={core_metric: '目标品牌量'})
                city_white = city_total.merge(tgt, on='城市', how='left').fillna({'目标品牌量': 0})
                city_white['目标品牌份额(%)'] = city_white['目标品牌量'] / city_white['城市总量'] * 100
                city_white = city_white.sort_values(['目标品牌份额(%)', '城市总量'], ascending=[True, False])
                results['城市白区'] = city_white
                city_priority = self.build_opportunity_priority(
                    df=city_white,
                    entity_col='城市',
                    total_col='城市总量',
                    share_col='目标品牌份额(%)',
                    target_volume_col='目标品牌量',
                    top_n=20,
                )
                if city_priority is not None and not city_priority.empty:
                    results['机会优先级_城市'] = city_priority

        if target_brand and core_dim and core_metric and core_dim in df.columns and core_metric in df.columns:
            hosp_col2 = None
            for cand in ['医院名称', '医疗机构名称']:
                if cand in df.columns:
                    hosp_col2 = cand
                    break
            if hosp_col2:
                hb = df.groupby([hosp_col2, core_dim], as_index=False)[core_metric].sum()
                hosp_total = hb.groupby(hosp_col2, as_index=False)[core_metric].sum().rename(columns={core_metric: '医院总量'})
                target_h = hb[hb[core_dim] == target_brand].rename(columns={core_metric: '目标品牌量'})[[hosp_col2, '目标品牌量']]
                hosp_white = hosp_total.merge(target_h, on=hosp_col2, how='left').fillna({'目标品牌量': 0})
                hosp_white['目标品牌份额(%)'] = hosp_white['目标品牌量'] / hosp_white['医院总量'] * 100
                if '城市' in df.columns:
                    hosp_white = hosp_white.merge(df[[hosp_col2, '城市']].drop_duplicates(), on=hosp_col2, how='left')
                hosp_white = hosp_white.sort_values(['目标品牌份额(%)', '医院总量'], ascending=[True, False])
                results['医院白院'] = hosp_white
                hosp_priority = self.build_opportunity_priority(
                    df=hosp_white,
                    entity_col=hosp_col2,
                    total_col='医院总量',
                    share_col='目标品牌份额(%)',
                    target_volume_col='目标品牌量',
                    top_n=30,
                )
                if hosp_priority is not None and not hosp_priority.empty:
                    results['机会优先级_医院'] = hosp_priority

        # 价格-量分析（若有价格列）
        price_info = None
        price_col = None
        for cand in ['中选价格', '价格', '单价']:
            if cand in df.columns:
                price_col = cand
                break
        if price_col and core_metric:
            df_price = df[[core_dim, core_metric, price_col]].dropna()
            if len(df_price) >= 3:
                corr = df_price[core_metric].corr(df_price[price_col])
                relation = None
                if corr is not None:
                    if abs(corr) < 0.2:
                        relation = "关系较弱"
                    elif abs(corr) < 0.5:
                        relation = "存在一定关系"
                    else:
                        relation = "相关性较强"
                price_med = df_price[price_col].median()
                qty_med = df_price[core_metric].median()
                lphv = df_price[(df_price[price_col] <= price_med) & (df_price[core_metric] >= qty_med)][core_dim].dropna().astype(str).unique().tolist()
                hplv = df_price[(df_price[price_col] > price_med) & (df_price[core_metric] < qty_med)][core_dim].dropna().astype(str).unique().tolist()
                price_info = {
                    '价格列': price_col,
                    '相关系数': corr,
                    '相关描述': relation,
                    '低价高量': lphv[:5],
                    '高价低量': hplv[:5]
                }
        results['价格分析'] = price_info

        # 结构机会（若有产品结构分析结果）
        structure_summary = None
        structure_df = results.get('产品结构')
        if structure_df is not None and not structure_df.empty and '产品类型' in structure_df.columns:
            temp = structure_df.set_index('产品类型')
            type_total = temp.sum(axis=1)
            total_sum = float(type_total.sum())
            if total_sum > 0:
                shares = (type_total / total_sum * 100).sort_values(ascending=False)
                main_type = shares.index[0]
                main_share = shares.iloc[0]
                others = ", ".join([f"{t} {s:.1f}%" for t, s in shares.iloc[1:].items()])
                structure_summary = {
                    '主类型': main_type,
                    '主类型占比': main_share,
                    '其他占比': others
                }
        results['结构概览'] = structure_summary

        # 时间趋势摘要（若存在）
        trend_summary = None
        if time_trend is not None and not time_trend.empty and len(time_trend) >= 2:
            start_v = float(time_trend['数值'].iloc[0])
            end_v = float(time_trend['数值'].iloc[-1])
            if end_v > start_v:
                dir_text = "上升"
            elif end_v < start_v:
                dir_text = "下降"
            else:
                dir_text = "相对平稳"
            trend_summary = {
                '起始': start_v,
                '当前': end_v,
                '方向': dir_text
            }
        results['时间趋势摘要'] = trend_summary

        return results

    def generate_insights(self, analysis_results, company_name=None):
        """
        生成洞察和建议

        Args:
            analysis_results: 分析结果
            company_name: 目标公司名称

        Returns:
            dict: 洞察和建议
        """
        insights = {}

        # 市场份额洞察（若存在）
        brand_share = analysis_results.get('品牌份额')
        if brand_share is not None and not brand_share.empty and company_name:
            company_rank = None
            company_share = 0
            if company_name in brand_share['品牌名称'].values:
                company_rank = brand_share[brand_share['品牌名称'] == company_name].index[0] + 1
                company_share = brand_share[brand_share['品牌名称'] == company_name]['市场份额'].iloc[0]

            insights['市场地位'] = {
                '排名': company_rank,
                '份额': f"{company_share:.2f}%",
                '领先品牌': brand_share.iloc[0]['品牌名称'] if len(brand_share) > 0 else '未知'
            }

        # 机会洞察（若存在）
        opp_cities = analysis_results.get('机会城市')
        if opp_cities is not None and len(opp_cities) > 0:
            top_opp_city = opp_cities.iloc[0]
            insights['机会城市'] = {
                '最佳机会': top_opp_city['城市'],
                '容量': f"{top_opp_city['总容量']:,.0f}",
                '当前份额': f"{top_opp_city['林华份额']:.2f}%"
            }

        # 竞争洞察（若存在）
        coverage = analysis_results.get('覆盖分析')
        if coverage is not None and company_name and company_name in coverage['品牌名称'].values:
            comp_data = coverage[coverage['品牌名称'] == company_name].iloc[0]
            insights['覆盖情况'] = {
                '覆盖医院数': f"{comp_data['覆盖医院数']}",
                '单院均量': f"{comp_data['单院均量']:,.0f}",
                '总业务量': f"{comp_data['总量']:,.0f}"
            }

        # 通用洞察
        general_insights = []
        numeric_stats = analysis_results.get('数值列统计')
        if numeric_stats is not None and not numeric_stats.empty:
            top_metric = numeric_stats.sort_values('均值', ascending=False).iloc[0]
            general_insights.append({
                'icon': 'fas fa-gem',
                'title': '核心指标',
                'content': f"{top_metric['字段']} 平均值 {top_metric['均值']:.2f}，可作为关键指标重点关注。"
            })
            if len(numeric_stats) > 1:
                volatile_metric = numeric_stats.sort_values('标准差', ascending=False).iloc[0]
                general_insights.append({
                    'icon': 'fas fa-wave-square',
                    'title': '波动提醒',
                    'content': f"{volatile_metric['字段']} 波动幅度最大（标准差 {volatile_metric['标准差']:.2f}），建议排查异常波动来源。"
                })

        categorical_summary = analysis_results.get('分类分布')
        if categorical_summary is not None and not categorical_summary.empty:
            top_category = categorical_summary.sort_values('占比(%)', ascending=False).iloc[0]
            general_insights.append({
                'icon': 'fas fa-layer-group',
                'title': '最集中的分类',
                'content': f"{top_category['字段']} 中 {top_category['类别']} 占比 {top_category['占比(%)']:.2f}%，代表当前最主要的构成。"
            })

        corr_matrix = analysis_results.get('相关性矩阵')
        if corr_matrix is not None:
            corr_abs = corr_matrix.abs().copy()
            np.fill_diagonal(corr_abs.values, 0)
            stacked = corr_abs.stack()
            if not stacked.empty and stacked.max() > 0:
                idx = stacked.idxmax()
                value = corr_matrix.loc[idx]
                general_insights.append({
                    'icon': 'fas fa-link',
                    'title': '强相关指标',
                    'content': f"{idx[0]} 与 {idx[1]} 相关系数 {value:.2f}，可用于组合指标或异常监控。"
                })

        time_trend = analysis_results.get('时间趋势')
        opportunity_suggestions = []
        if time_trend is not None and len(time_trend) >= 2:
            start_value = float(time_trend['数值'].iloc[0])
            end_value = float(time_trend['数值'].iloc[-1])
            direction = '上升' if end_value > start_value else ('下降' if end_value < start_value else '平稳')
            general_insights.append({
                'icon': 'fas fa-chart-line',
                'title': '时间趋势',
                'content': f"整体趋势呈现{direction}态势，起始值 {start_value:,.2f} → 当前值 {end_value:,.2f}。"
            })
            opportunity_suggestions.append({
                'icon': 'fas fa-bullseye',
                'title': '趋势建议',
                'description': f"继续跟踪时间序列的{direction}趋势，必要时结合节奏调整策略。"
            })

        if categorical_summary is not None and not categorical_summary.empty:
            tail_category = categorical_summary.sort_values('占比(%)', ascending=True).iloc[0]
            opportunity_suggestions.append({
                'icon': 'fas fa-lightbulb',
                'title': '潜在增量',
                'description': f"{tail_category['字段']} 中 {tail_category['类别']} 占比仅 {tail_category['占比(%)']:.2f}%，可作为差异化突破点。"
            })

        city_priority = analysis_results.get('机会优先级_城市')
        if city_priority is not None and not city_priority.empty and '城市' in city_priority.columns:
            top_city = city_priority.iloc[0]
            opportunity_suggestions.append({
                'icon': 'fas fa-location-dot',
                'title': '优先城市',
                'description': f"优先级最高城市为 {top_city['城市']}（综合分 {top_city['综合优先级分']:.1f}），建议优先配置市场动作。"
            })

        hosp_priority = analysis_results.get('机会优先级_医院')
        if hosp_priority is not None and not hosp_priority.empty:
            name_col = None
            for cand in ['医院名称', '医疗机构名称']:
                if cand in hosp_priority.columns:
                    name_col = cand
                    break
            if name_col:
                top_hosp = hosp_priority.iloc[0]
                opportunity_suggestions.append({
                    'icon': 'fas fa-hospital',
                    'title': '优先机构',
                    'description': f"优先级最高机构为 {top_hosp[name_col]}（综合分 {top_hosp['综合优先级分']:.1f}），可作为近期突破点。"
                })

        if general_insights:
            insights['通用洞察'] = general_insights

        if opportunity_suggestions:
            insights['机会建议'] = opportunity_suggestions

        return insights
