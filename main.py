"""
主程序：整合所有模塊，完整的工作流
"""

from src.excel_generator import ExcelGenerator
from src.visualizer import TimelineVisualizer
from src.data_processor import DataProcessor
from src.excel_reader import ExcelReader
import logging
import sys
from pathlib import Path

# 添加項目根目錄到 Python 路徑
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))


# 配置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('milestone_timeline.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


def main(input_excel_path, output_excel_path=None, title="里程碑時間線"):
    """
    主流程：從 Excel 生成里程碑時間線 Excel 報告

    Args:
        input_excel_path (str): 輸入 Excel 檔案路徑
        output_excel_path (str): 輸出 Excel 檔案路徑，預設為 data/output/timeline_report.xlsx
        title (str): 報告標題
    """
    try:
        logger.info("=" * 50)
        logger.info("開始生成里程碑時間線 Excel 報告")
        logger.info("=" * 50)

        # 設定輸出路徑
        if output_excel_path is None:
            output_excel_path = project_root / "data" / "output" / "timeline_report.xlsx"

        # 1. 讀取 Excel
        logger.info(f"步驟 1: 讀取 Excel 檔案: {input_excel_path}")
        reader = ExcelReader(input_excel_path)
        df = reader.read_milestone_data()
        logger.info(f"✓ 成功讀取 {len(df)} 條記錄")

        # 2. 數據處理
        logger.info("步驟 2: 數據處理和統計分析")
        processor = DataProcessor(milestones_per_page=50)
        merged_df, stats, pages = processor.process_all(df)
        logger.info(f"✓ 合併同日期事件，共 {len(merged_df)} 個里程碑")
        logger.info(f"✓ 分頁完成，共 {len(pages)} 頁")

        # 3. 生成可視化
        logger.info("步驟 3: 生成可視化圖表")
        visualizer = TimelineVisualizer()
        figures = visualizer.generate_pdf(pages, stats, title=title)
        logger.info(f"✓ 生成 {len(figures)} 頁圖表")

        # 4. 導出 Excel 檔案
        logger.info("步驟 4: 導出 Excel 檔案")
        excel_gen = ExcelGenerator(dpi=300)
        output_path = excel_gen.save_excel(
            figures, output_excel_path, stats, title=title)
        logger.info(f"✓ Excel 檔案已保存: {output_path}")

        # 關閉所有圖表，釋放內存
        import matplotlib.pyplot as plt
        plt.close('all')

        logger.info("=" * 50)
        logger.info("✓ 報告生成完成!")
        logger.info("=" * 50)

        return output_path

    except Exception as e:
        logger.error(f"✗ 錯誤: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    # 示例：使用 data/input 目錄下的 Excel 檔案
    input_file = project_root / "data" / "input" / "sample_milestones.xlsx"

    if input_file.exists():
        main(str(input_file))
    else:
        print(f"示例檔案不存在: {input_file}")
        print("請先在 data/input 目錄下放置 Excel 檔案")
        print("\n用法示例:")
        print("  python main.py")
        print("\n或在代碼中調用:")
        print("  from main import main")
        print("  main('path/to/your/file.xlsx')")
