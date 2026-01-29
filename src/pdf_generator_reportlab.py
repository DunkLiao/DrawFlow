"""
使用 reportlab 的 PDF 生成模塊（改進的中文支持）
"""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm, inch
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import logging
from pathlib import Path
import sys

logger = logging.getLogger(__name__)


class ReportLabPDFGenerator:
    """使用 ReportLab 生成 PDF（改進的中文支援）"""

    def __init__(self, dpi=300):
        """
        初始化 PDF 生成器

        Args:
            dpi (int): PDF 解析度
        """
        self.dpi = dpi
        self._register_fonts()

    def _register_fonts(self):
        """註冊中文字體"""
        try:
            # 嘗試註冊系統中文字體
            font_paths = [
                'C:\\Windows\\Fonts\\SimHei.ttf',
                'C:\\Windows\\Fonts\\simsun.ttc',
                '/System/Library/Fonts/STHeiti Light.ttc',
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
            ]

            for font_path in font_paths:
                if Path(font_path).exists():
                    try:
                        # 提取字體文件名作為名稱
                        font_name = Path(font_path).stem
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        self.font_name = font_name
                        logger.info(f"成功註冊字體: {font_path}")
                        return
                    except Exception as e:
                        logger.debug(f"字體註冊失敗 {font_path}: {e}")

            # 如果沒找到，使用默認字體
            logger.warning("未找到中文字體，使用默認字體")
            self.font_name = 'Helvetica'

        except Exception as e:
            logger.error(f"字體註冊錯誤: {e}")
            self.font_name = 'Helvetica'

    def save_pdf(self, figures, output_path, title="里程碑時間線"):
        """
        將圖表保存為 PDF 檔案（此方法與 matplotlib 相同，保持兼容性）

        Args:
            figures (list): matplotlib Figure 物件列表
            output_path (str): 輸出檔案路徑
            title (str): PDF 標題

        Returns:
            Path: 輸出檔案路徑
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            # 使用 matplotlib 的 PdfPages 保存
            from matplotlib.backends.backend_pdf import PdfPages
            import warnings

            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                with PdfPages(str(output_path)) as pdf:
                    for fig in figures:
                        fig.dpi = self.dpi
                        pdf.savefig(fig, bbox_inches='tight', pad_inches=0.1)

            logger.info(f"PDF 生成成功: {output_path}")
            logger.info(f"總共 {len(figures)} 頁")
            return output_path

        except Exception as e:
            logger.error(f"PDF 生成失敗: {str(e)}")
            raise
