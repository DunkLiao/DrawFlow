"""
可視化模塊：生成彩色水平時間線和統計面板
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
from matplotlib import rcParams
import numpy as np
from datetime import datetime, timedelta
import logging
import platform
from pathlib import Path

logger = logging.getLogger(__name__)

# 配置中文字體


def setup_chinese_fonts():
    """設定中文字體支援"""
    system = platform.system()

    # 禁用字體警告
    import warnings
    warnings.filterwarnings('ignore')

    # Windows 系統配置
    if system == 'Windows':
        # 嘗試尋找系統中的中文字體
        font_names = ['SimHei', 'SimSun', 'DengXian', 'Microsoft YaHei']
        available = []

        for font in font_names:
            try:
                plt.rcParams['font.sans-serif'].insert(0, font)
                available.append(font)
                break
            except:
                pass

        if not available:
            # 如果沒找到系統字體，使用默認
            plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']

    elif system == 'Darwin':  # macOS
        plt.rcParams['font.sans-serif'] = ['STHeiti', 'SimHei', 'DejaVu Sans']

    else:  # Linux
        plt.rcParams['font.sans-serif'] = [
            'SimHei',
            'WenQuanYi Zen Hei',
            'WenQuanYi Micro Hei',
            'DejaVu Sans'
        ]

    # 重要配置
    rcParams['axes.unicode_minus'] = False  # 解決負號顯示問題
    rcParams['pdf.fonttype'] = 42  # 使用 TrueType 字體，支持中文嵌入
    rcParams['ps.fonttype'] = 42   # PostScript 也使用 TrueType


setup_chinese_fonts()


class TimelineVisualizer:
    """時間線可視化生成器"""

    def __init__(self, config=None):
        """
        初始化可視化器

        Args:
            config (dict): 配置參數字典，包含排版和色彩設定
        """
        self.config = config or self._default_config()
        self._setup_colors()

    def _default_config(self):
        """取得預設配置"""
        from config.settings import (
            PAGE_WIDTH, PAGE_HEIGHT, MARGIN_INCH,
            TITLE_FONT_SIZE, STAT_FONT_SIZE, LABEL_FONT_SIZE,
            MONTHLY_CHART_WIDTH_INCH, MONTHLY_CHART_HEIGHT_INCH,
            TIMELINE_AXIS_COLOR, TIMELINE_AXIS_WIDTH,
            TEXT_COLOR, STAT_TEXT_COLOR,
            STAT_PANEL_COLOR, STAT_PANEL_EDGE_COLOR,
            COLOR_GRADIENT_START, COLOR_GRADIENT_MID, COLOR_GRADIENT_END
        )

        return {
            'page_width': PAGE_WIDTH,
            'page_height': PAGE_HEIGHT,
            'margin': MARGIN_INCH,
            'title_font': TITLE_FONT_SIZE,
            'stat_font': STAT_FONT_SIZE,
            'label_font': LABEL_FONT_SIZE,
            'monthly_chart_width': MONTHLY_CHART_WIDTH_INCH,
            'monthly_chart_height': MONTHLY_CHART_HEIGHT_INCH,
            'timeline_axis_color': TIMELINE_AXIS_COLOR,
            'timeline_axis_width': TIMELINE_AXIS_WIDTH,
            'text_color': TEXT_COLOR,
            'stat_text_color': STAT_TEXT_COLOR,
            'stat_panel_color': STAT_PANEL_COLOR,
            'stat_panel_edge_color': STAT_PANEL_EDGE_COLOR,
            'color_gradient': (COLOR_GRADIENT_START, COLOR_GRADIENT_MID, COLOR_GRADIENT_END),
        }

    def _setup_colors(self):
        """設定色彩漸變"""
        colors = self.config['color_gradient']
        # 創建自定義漸變色彩映射
        n_bins = 100
        self.cmap = LinearSegmentedColormap.from_list(
            'milestone_gradient',
            [colors[0], colors[1], colors[2]],
            N=n_bins
        )

    def create_timeline_figure(self, page_df, stats, page_num=1, total_pages=1, title="里程碑時間線"):
        """
        創建單頁時間線圖表

        Args:
            page_df (pd.DataFrame): 當前頁的里程碑數據
            stats (dict): 統計信息字典
            page_num (int): 當前頁碼
            total_pages (int): 總頁數
            title (str): 圖表標題

        Returns:
            matplotlib.figure.Figure: 圖表物件
        """
        cfg = self.config
        fig = plt.figure(
            figsize=(cfg['page_width'], cfg['page_height']), dpi=100)
        fig.patch.set_facecolor('white')

        # 計算版面尺寸
        margin = cfg['margin']
        content_width = cfg['page_width'] - 2 * margin
        content_height = cfg['page_height'] - 2 * margin

        # 標題區域高度
        title_height = 0.5
        # 統計區域高度
        stat_height = 1.5
        # 時間線區域高度
        timeline_height = content_height - title_height - stat_height - 0.3

        # 標題
        ax_title = fig.add_axes(
            [margin, cfg['page_height'] - margin - title_height, content_width, title_height])
        ax_title.axis('off')
        ax_title.text(0.5, 0.5, title, fontsize=cfg['title_font'], weight='bold',
                      ha='center', va='center', color=cfg['text_color'],
                      transform=ax_title.transAxes, family='sans-serif')

        # 統計面板 + 月度圖表
        stat_top = cfg['page_height'] - margin - title_height - stat_height
        self._draw_stat_panel(fig, margin, stat_top,
                              content_width, stat_height, stats, page_df)

        # 時間線
        timeline_top = stat_top - 0.3 - timeline_height
        self._draw_timeline(fig, margin, timeline_top,
                            content_width, timeline_height, page_df)

        # 頁碼
        ax_footer = fig.add_axes([margin, margin - 0.3, content_width, 0.2])
        ax_footer.axis('off')
        page_text = f"第 {page_num} 頁，共 {total_pages} 頁"
        ax_footer.text(0.5, 0.5, page_text, fontsize=cfg['label_font'],
                       ha='center', va='center', color=cfg['text_color'],
                       transform=ax_footer.transAxes, family='sans-serif')

        return fig

    def _draw_stat_panel(self, fig, x, y, width, height, stats, page_df):
        """
        繪製統計面板（文字 + 月度分佈圖）

        Args:
            fig (matplotlib.figure.Figure): 圖表物件
            x, y (float): 位置（英寸）
            width, height (float): 尺寸（英寸）
            stats (dict): 統計信息
            page_df (pd.DataFrame): 當前頁數據
        """
        cfg = self.config

        # 統計面板背景
        ax_stat = fig.add_axes([x, y, width, height])
        ax_stat.axis('off')

        # 繪製背景和邊框
        rect = mpatches.FancyBboxPatch(
            (0, 0), 1, 1, boxstyle="round,pad=0.01",
            transform=ax_stat.transAxes,
            facecolor=cfg['stat_panel_color'],
            edgecolor=cfg['stat_panel_edge_color'],
            linewidth=2
        )
        ax_stat.add_patch(rect)

        # 統計文字區域（左側）
        stat_text_width = 0.55
        stat_lines = [
            f"里程碑總數: {stats.get('total_milestones', 0)}",
            f"開始日期: {stats.get('start_date', '').strftime('%Y-%m-%d') if stats.get('start_date') else 'N/A'}",
            f"結束日期: {stats.get('end_date', '').strftime('%Y-%m-%d') if stats.get('end_date') else 'N/A'}",
            f"時間跨度: {stats.get('total_days', 0)} 天",
            f"里程碑密度: {stats.get('milestone_density', 0):.2f} 個/月",
        ]

        y_pos = 0.85
        for line in stat_lines:
            ax_stat.text(0.05, y_pos, line, fontsize=cfg['stat_font'],
                         va='top', color=cfg['stat_text_color'],
                         transform=ax_stat.transAxes, weight='bold', family='sans-serif')
            y_pos -= 0.16

        # 月度分佈小圖表（右側）
        if stats.get('monthly_distribution'):
            self._draw_monthly_chart(
                fig, x + width * stat_text_width, y + height * 0.1,
                width * (1 - stat_text_width) - 0.05, height * 0.8,
                stats['monthly_distribution']
            )

    def _draw_monthly_chart(self, fig, x, y, width, height, monthly_data):
        """
        繪製月度分佈柱狀圖

        Args:
            fig (matplotlib.figure.Figure): 圖表物件
            x (float): 位置 x (英寸)
            y (float): 位置 y (英寸)
            width (float): 寬度 (英寸)
            height (float): 高度 (英寸)
            monthly_data (list): 月度數據列表
        """
        ax_monthly = fig.add_axes([x, y, width, height])

        months = [item['period'][-2:] for item in monthly_data]  # 取月份
        counts = [item['count'] for item in monthly_data]

        # 柱狀圖
        colors = plt.cm.Blues(np.linspace(0.4, 0.8, len(months)))
        ax_monthly.bar(range(len(months)), counts, color=colors,
                       edgecolor='navy', linewidth=0.5)

        ax_monthly.set_xticks(range(len(months)))
        ax_monthly.set_xticklabels(
            months, fontsize=9, rotation=45, family='sans-serif', weight='bold')
        ax_monthly.set_ylabel('里程碑數', fontsize=10,
                              family='sans-serif', weight='bold')
        ax_monthly.grid(axis='y', alpha=0.3, linestyle='--')
        ax_monthly.set_axisbelow(True)
        ax_monthly.tick_params(labelsize=9)

    def _draw_timeline(self, fig, x, y, width, height, page_df):
        """
        繪製水平時間線

        Args:
            fig(matplotlib.figure.Figure): 圖表物件
            x, y(float): 位置（英寸）
            width, height(float): 尺寸（英寸）
            page_df(pd.DataFrame): 里程碑數據
        """
        cfg = self.config
        ax_timeline = fig.add_axes([x, y, width, height])
        ax_timeline.set_xlim(-0.05, 1.05)
        ax_timeline.set_ylim(-0.5, len(page_df) + 0.5)

        # 時間軸範圍
        min_date = page_df['date'].min()
        max_date = page_df['date'].max()
        date_range = (max_date - min_date).days
        if date_range == 0:
            date_range = 1  # 避免除以零

        # 繪製主軸線
        ax_timeline.plot([0.05, 0.95], [len(page_df) - 1, len(page_df) - 1], '-',
                         color=cfg['timeline_axis_color'], linewidth=cfg['timeline_axis_width'])

        # 色彩映射
        norm = plt.Normalize(0, len(page_df) - 1)
        colors = self.cmap(norm(np.arange(len(page_df))))

        # 繪製里程碑點和標籤
        for idx, (_, row) in enumerate(page_df.iterrows()):
            date = row['date']
            event = row['event']

            # 計算日期在時間線上的位置
            days_from_start = (date - min_date).days
            x_pos = 0.05 + (days_from_start / date_range) * 0.9

            # 繪製圓點
            circle = mpatches.Circle((x_pos, len(page_df) - 1), 0.015, color=colors[idx],
                                     ec='darkgray', linewidth=0.5, zorder=3)
            ax_timeline.add_patch(circle)

            # 繪製垂直虛線
            ax_timeline.plot([x_pos, x_pos], [len(page_df) - 1, len(page_df) - 1.3],
                             '--', color='lightgray', linewidth=0.5, zorder=1)

            # 繪製事件標籤（使用 family 指定字體族）
            label_y = len(page_df) - 2 if idx % 2 == 0 else len(page_df) - 1.6
            ax_timeline.text(x_pos, label_y, event, fontsize=cfg['label_font'] + 1,
                             ha='center', va='top', rotation=45 if len(event) > 8 else 0,
                             color=cfg['text_color'], wrap=True, family='sans-serif', weight='bold')

            # 繪製日期標籤
            date_text = date.strftime('%Y-%m-%d')
            ax_timeline.text(x_pos, len(page_df) - 0.3, date_text, fontsize=cfg['label_font'],
                             ha='center', va='top', color=cfg['text_color'], family='sans-serif')

        # 隱藏軸
        ax_timeline.set_xticks([])
        ax_timeline.set_yticks([])
        ax_timeline.spines['top'].set_visible(False)
        ax_timeline.spines['right'].set_visible(False)
        ax_timeline.spines['left'].set_visible(False)
        ax_timeline.spines['bottom'].set_visible(False)

        logger.info(f"時間線繪製完成，共 {len(page_df)} 個里程碑")

    def generate_pdf(self, pages_data, stats, title="里程碑時間線", output_path=None):
        """
        生成多頁 PDF

        Args:
            pages_data(list): 分頁數據列表（每個元素是一個 DataFrame）
            stats(dict): 統計信息字典
            title(str): 圖表標題
            output_path(str): 輸出路徑

        Returns:
            list: 生成的圖表列表
        """
        figures = []
        total_pages = len(pages_data)

        for page_num, page_df in enumerate(pages_data, 1):
            fig = self.create_timeline_figure(
                page_df, stats, page_num, total_pages, title
            )
            figures.append(fig)
            logger.info(f"生成第 {page_num}/{total_pages} 頁")

        return figures
