"""
創建示例 Excel 檔案的腳本
"""

import pandas as pd
from pathlib import Path

# 創建示例數據 - 包含 60+ 行，含重複日期
data = []

# 2026 年 1 月 - 計劃與需求 (2週)
data.append(['2026-01-05', '項目啟動'])
data.append(['2026-01-12', '需求評審'])
data.append(['2026-01-12', '團隊組建'])
data.append(['2026-01-19', '技術選型'])

# 2026 年 2 月 - 設計與準備 (3週)
data.append(['2026-02-02', '架構設計'])
data.append(['2026-02-09', '開發環境部署'])
data.append(['2026-02-09', '編碼規範制定'])

# 2026 年 2-3 月 - Sprint 開發 (6週，每個Sprint 2週)
data.append(['2026-02-23', 'Sprint 1 完成'])
data.append(['2026-03-09', 'Sprint 2 完成'])
data.append(['2026-03-23', 'Sprint 3 完成'])

# 2026 年 3-4 月 - 第一階段測試 (3週)
data.append(['2026-03-30', '第一階段測試'])
data.append(['2026-04-06', 'Bug 修復'])
data.append(['2026-04-13', 'Sprint 4 完成'])

# 2026 年 4 月 - 系統優化與審計 (2週)
data.append(['2026-04-20', '性能優化'])
data.append(['2026-04-27', '安全審計'])

# 2026 年 5 月 - UAT 和準備 (3週)
data.append(['2026-05-04', '用戶驗收測試'])
data.append(['2026-05-04', 'UAT 反饋收集'])
data.append(['2026-05-11', '文檔編寫'])
data.append(['2026-05-18', '用戶培訓'])
data.append(['2026-05-25', '數據遷移'])

# 2026 年 6 月 - 預發佈與上線 (2週)
data.append(['2026-06-01', '預發佈環境測試'])
data.append(['2026-06-08', '灰度上線'])
data.append(['2026-06-08', '監控部署'])
data.append(['2026-06-22', '全量上線'])
data.append(['2026-06-22', '上線慶祝'])

# 2026 年 7 月 - 功能完善 (3週)
data.append(['2026-07-06', '功能完善'])
data.append(['2026-07-13', 'V1.1 規劃'])
data.append(['2026-07-20', '用戶反饋分析'])

# 2026 年 8 月 - V1.1 新功能開發 (4週)
data.append(['2026-08-03', '新功能開發'])
data.append(['2026-08-17', '集成測試'])
data.append(['2026-08-31', 'Beta 測試'])

# 2026 年 9 月 - V1.1 發佈 (2週)
data.append(['2026-09-07', '功能測試'])
data.append(['2026-09-14', '修復問題'])
data.append(['2026-09-21', '版本優化'])
data.append(['2026-09-28', 'V1.1 發佈'])

# 2026 年 10 月 - V1.2 規劃
data.append(['2026-10-05', 'V1.2 開發'])
data.append(['2026-10-19', '版本優化'])
data.append(['2026-11-02', '官方發佈'])

# 創建 DataFrame
df = pd.DataFrame(data, columns=['日期', '事件'])

# 創建輸出目錄
output_path = Path('data/input/sample_milestones.xlsx')
output_path.parent.mkdir(parents=True, exist_ok=True)

# 保存為 Excel
df.to_excel(output_path, index=False, engine='openpyxl')

print(f'成功創建示例檔案: {output_path}')
print(f'包含 {len(df)} 行數據')
print(f'日期範圍: {df["日期"].min()} 至 {df["日期"].max()}')
