import json

# 讀取 response.json 文件
with open('response.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# 獲取 issues 陣列
issues = data.get('issues', [])

# 輸出 issues 的數量
print(f"共有 {len(issues)} 筆 issues")