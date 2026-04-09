import pandas as pd
import json

file_path = r'c:\Users\cyt18\桌面\Antigravity\儀表板\LINE OA 加入資料_分析結果.xlsx'
df = pd.read_excel(file_path)

# Show headers and first 5 rows
print("Headers:", df.columns.tolist())
print("Sample Data:\n", df.head().to_string())

# Summary statistics for numeric or categorical analysis
print("\nSummary Statistics:")
print(df.describe(include='all').to_string())
