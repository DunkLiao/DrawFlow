"""
Excel 讀取模塊：讀取日期和事件列，驗證數據有效性
"""

import pandas as pd
from datetime import datetime
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class ExcelReader:
    """讀取 Excel 檔案中的里程碑數據"""

    def __init__(self, file_path):
        """
        初始化 Excel 讀取器

        Args:
            file_path (str): Excel 檔案路徑
        """
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"檔案不存在: {file_path}")

        if self.file_path.suffix.lower() not in ['.xlsx', '.xls']:
            raise ValueError(f"不支援的檔案格式: {self.file_path.suffix}")

    def read_milestone_data(self, sheet_name=0, date_col=0, event_col=1):
        """
        讀取里程碑數據

        Args:
            sheet_name (int or str): 工作表名稱或索引，預設為 0
            date_col (int or str): 日期列索引或列名，預設為 0
            event_col (int or str): 事件列索引或列名，預設為 1

        Returns:
            pd.DataFrame: 包含 'date' 和 'event' 列的 DataFrame
        """
        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            logger.info(f"讀取 Excel 成功，共 {len(df)} 行")

            # 重命名列
            df.columns = ['date', 'event']

            # 移除全空行
            df = df.dropna(how='all')

            # 轉換日期列
            df['date'] = pd.to_datetime(df['date'], errors='coerce')

            # 移除日期無效的行
            invalid_count = df['date'].isna().sum()
            if invalid_count > 0:
                logger.warning(f"發現 {invalid_count} 行無效日期，已移除")
                df = df.dropna(subset=['date'])

            # 確保事件為字符串
            df['event'] = df['event'].astype(str).str.strip()

            # 移除空事件
            df = df[df['event'] != '']
            df = df[df['event'] != 'nan']

            # 按日期排序
            df = df.sort_values('date').reset_index(drop=True)

            logger.info(f"數據驗證完成，有效數據 {len(df)} 行")
            return df

        except Exception as e:
            logger.error(f"讀取檔案失敗: {str(e)}")
            raise
