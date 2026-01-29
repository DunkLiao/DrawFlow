"""
測試步驟 4：驗證生成的 PDF 質量

此腳本檢查 PDF 的基本信息和內容
"""

import os
from pathlib import Path


def test_step4_pdf_quality():
    """測試 PDF 質量"""

    pdf_path = Path('data/output/timeline.pdf')

    print("\n" + "="*60)
    print("步驟 4 測試：PDF 質量驗證")
    print("="*60)

    # 1. 檢查檔案是否存在
    if not pdf_path.exists():
        print("✗ PDF 檔案不存在")
        return False

    print(f"✓ PDF 檔案存在: {pdf_path}")

    # 2. 檢查檔案大小
    file_size = pdf_path.stat().st_size
    print(f"✓ 檔案大小: {file_size / 1024:.2f} KB")

    if file_size < 1000:
        print("  ⚠ 警告：檔案大小偏小，可能內容不完整")

    # 3. 驗證 PDF 檔案格式
    with open(pdf_path, 'rb') as f:
        header = f.read(4)
        if header == b'%PDF':
            print("✓ PDF 檔案格式正確")
        else:
            print("✗ PDF 檔案格式可能有問題")
            return False

    # 4. 檢查日誌
    log_path = Path('milestone_timeline.log')
    if log_path.exists():
        with open(log_path, 'r', encoding='utf-8') as f:
            log_content = f.read()
            if '報告生成完成' in log_content:
                print("✓ 日誌顯示報告生成完成")
            else:
                print("⚠ 日誌中未找到成功完成標記")

    # 5. 提示後續步驟
    print("\n" + "-"*60)
    print("後續驗證步驟：")
    print("-"*60)
    print("1. 用 PDF 閱讀器打開: data/output/timeline.pdf")
    print("2. 檢查以下項目：")
    print("   ✓ 標題「里程碑時間線」是否正確顯示")
    print("   ✓ 統計面板中的文字是否清晰可讀")
    print("   ✓ 月度分佈柱狀圖是否清晰")
    print("   ✓ 時間線上的日期和事件是否顯示正確")
    print("   ✓ 色彩漸變是否流暢（藍→綠→紅）")
    print("   ✓ 頁碼信息是否顯示正確")
    print("\n" + "="*60)
    print("✓ PDF 檔案驗證完成")
    print("="*60 + "\n")

    return True


if __name__ == "__main__":
    test_step4_pdf_quality()
