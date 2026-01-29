"""
數據處理與統計分析模塊：合併、分頁、統計計算
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import math
import logging

logger = logging.getLogger(__name__)


class DataProcessor:
    """數據處理和統計分析"""

    def __init__(self, milestones_per_page=50):
        """
        初始化數據處理器

        Args:
            milestones_per_page (int): 每頁最大里程碑數
        """
        self.milestones_per_page = milestones_per_page

    def process_data(self, df):
        """
        處理和驗證數據

        Args:
            df (pd.DataFrame): 包含 'date' 和 'event' 列的 DataFrame

        Returns:
            pd.DataFrame: 處理後的 DataFrame
        """
        if df.empty:
            raise ValueError("數據為空")

        # 移除重複項（保留第一個）
        df = df.drop_duplicates(subset=['date', 'event'], keep='first')

        return df

    def merge_same_date_events(self, df):
        """
        合併同日期的事件

        Args:
            df (pd.DataFrame): 輸入 DataFrame

        Returns:
            pd.DataFrame: 合併後的 DataFrame
        """
        merged_df = df.groupby('date', as_index=False).agg(
            {'event': lambda x: ', '.join(x)}
        )
        return merged_df.sort_values('date').reset_index(drop=True)

    def calculate_statistics(self, df):
        """
        計算統計信息

        Args:
            df (pd.DataFrame): 輸入 DataFrame

        Returns:
            dict: 包含統計信息的字典
        """
        if df.empty:
            return {}

        start_date = df['date'].min()
        end_date = df['date'].max()
        total_days = (end_date - start_date).days

        # 計算月度分佈
        df['year_month'] = df['date'].dt.to_period('M')
        monthly_dist = df.groupby('year_month').size().to_dict()
        monthly_list = [
            {'period': str(period), 'count': count}
            for period, count in sorted(monthly_dist.items())
        ]

        # 計算里程碑密度（每月平均）
        total_months = len(monthly_list)
        milestone_density = len(df) / total_months if total_months > 0 else 0

        stats = {
            'total_milestones': len(df),
            'start_date': start_date,
            'end_date': end_date,
            'total_days': total_days,
            'total_months': total_months,
            'milestone_density': round(milestone_density, 2),
            'monthly_distribution': monthly_list,
        }

        logger.info(
            f"統計完成: {len(df)} 個里程碑，跨度 {total_days} 天，密度 {stats['milestone_density']}/月")
        return stats

    def paginate_data(self, df):
        """
        按指定數量分頁

        Args:
            df (pd.DataFrame): 輸入 DataFrame

        Returns:
            list: 分頁列表，每個元素是一個 DataFrame
        """
        num_pages = math.ceil(len(df) / self.milestones_per_page)
        pages = []

        for i in range(num_pages):
            start_idx = i * self.milestones_per_page
            end_idx = start_idx + self.milestones_per_page
            page_df = df.iloc[start_idx:end_idx].copy()
            pages.append(page_df)

        logger.info(f"分頁完成: 共 {num_pages} 頁")
        return pages

    def process_all(self, df):
        """
        完整處理流程

        Args:
            df (pd.DataFrame): 輸入 DataFrame

        Returns:
            tuple: (合併後的 DataFrame, 統計信息, 分頁列表)
        """
        # 1. 初步數據驗證
        df = self.process_data(df)

        # 2. 合併同日期事件
        merged_df = self.merge_same_date_events(df)

        # 3. 計算統計信息
        stats = self.calculate_statistics(merged_df)

        # 4. 分頁
        pages = self.paginate_data(merged_df)

        return merged_df, stats, pages
