"""
配置文件：定義色彩方案、排版參數、常數設定
"""

# ==================== 排版參數 ====================
MARGIN_MM = 20  # 邊距 mm
MARGIN_INCH = MARGIN_MM / 25.4  # 轉換為英寸

# 紙張尺寸 (A4 縱向)
PAGE_WIDTH = 8.27  # 英寸
PAGE_HEIGHT = 11.69  # 英寸

# 字體大小 (pt)
TITLE_FONT_SIZE = 20
STAT_FONT_SIZE = 14
LABEL_FONT_SIZE = 12
SMALL_FONT_SIZE = 10

# 圖表尺寸
MONTHLY_CHART_WIDTH_CM = 8  # cm
MONTHLY_CHART_HEIGHT_CM = 3  # cm
MONTHLY_CHART_WIDTH_INCH = MONTHLY_CHART_WIDTH_CM / 2.54
MONTHLY_CHART_HEIGHT_INCH = MONTHLY_CHART_HEIGHT_CM / 2.54

# ==================== 色彩方案 ====================
# 漸變色彩：藍 -> 綠 -> 紅
COLOR_GRADIENT_START = (0.2, 0.6, 1.0)  # 藍色 RGB
COLOR_GRADIENT_MID = (0.2, 1.0, 0.2)    # 綠色 RGB
COLOR_GRADIENT_END = (1.0, 0.2, 0.2)    # 紅色 RGB

# 統計面板顏色
STAT_PANEL_COLOR = (1.0, 0.95, 0.8)  # 金黃/橙色
STAT_PANEL_EDGE_COLOR = (0.8, 0.7, 0.3)  # 邊框色

# 時間線顏色
TIMELINE_AXIS_COLOR = (0.3, 0.3, 0.3)  # 深灰色
TIMELINE_AXIS_WIDTH = 2

# 文字顏色
TEXT_COLOR = (0.2, 0.2, 0.2)  # 深灰色
STAT_TEXT_COLOR = (0.3, 0.3, 0.3)  # 統計面板文字色

# ==================== 分頁設定 ====================
MILESTONES_PER_PAGE = 50  # 每頁最大里程碑數
DPI = 300  # PDF 解析度

# ==================== 數據驗證 ====================
DATE_FORMATS = [
    '%Y-%m-%d',
    '%Y/%m/%d',
    '%d-%m-%Y',
    '%d/%m/%Y',
    '%m-%d-%Y',
    '%m/%d/%Y',
]
