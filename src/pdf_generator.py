"""
PDF 生成模塊：將圖表保存為 PDF 檔案
"""

from matplotlib.backends.backend_pdf import PdfPages
from pathlib import Path
import logging
import warnings

logger = logging.getLogger(__name__)


class PDFGenerator:
    """PDF 檔案生成器"""

    def __init__(self, dpi=300):
        """
        初始化 PDF 生成器

        Args:
            dpi (int): PDF 解析度
        """
        self.dpi = dpi
        # 禁止字體警告
        warnings.filterwarnings('ignore', category=UserWarning)

    def save_pdf(self, figures, output_path, title="里程碑時間線"):
        """
        將圖表保存為 PDF 檔案

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
            # 使用 PdfPages 保存，確保字體正確處理
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')

                with PdfPages(str(output_path)) as pdf:
                    for idx, fig in enumerate(figures, 1):
                        # 設定 DPI
                        fig.dpi = self.dpi

                        # 使用 bbox_inches='tight' 避免裁剪
                        pdf.savefig(fig, bbox_inches='tight', pad_inches=0.1)

            logger.info(f"PDF 生成成功: {output_path}")
            logger.info(f"總共 {len(figures)} 頁")
            return output_path

        except Exception as e:
            logger.error(f"PDF 生成失敗: {str(e)}")
            raise
