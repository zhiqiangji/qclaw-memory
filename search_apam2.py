import urllib.request, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
port = 19000

# Search 4: Yang anionic polyacrylamide freeze-thaw silty clay (论文已引的文献)
keyword4 = "Yang anionic polyacrylamide freeze-thaw silty clay 2024"
data4 = json.dumps({"keyword": keyword4, "site": "sciencedirect.com"}).encode('utf-8')
req4 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data4,
    headers={"Content-Type": "application/json"}
)
resp4 = urllib.request.urlopen(req4, timeout=15)
result4 = json.loads(resp4.read().decode('utf-8'))
print("=== 搜索4: Yang APAM freeze-thaw (ScienceDirect) ===")
if result4.get('success'):
    print(result4['message'])
else:
    print(f"Error: {result4.get('message','unknown')}")

print("\n")

# Search 5: 聚丙烯酰胺 土壤 导热系数
keyword5 = "聚丙烯酰胺 土壤 导热系数"
data5 = json.dumps({"keyword": keyword5}).encode('utf-8')
req5 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data5,
    headers={"Content-Type": "application/json"}
)
resp5 = urllib.request.urlopen(req5, timeout=15)
result5 = json.loads(resp5.read().decode('utf-8'))
print("=== 搜索5: 聚丙烯酰胺 土壤 导热系数 ===")
if result5.get('success'):
    print(result5['message'])
else:
    print(f"Error: {result5.get('message','unknown')}")

print("\n")

# Search 6: 高分子改良土 温度场 冻结 降温速率
keyword6 = "高分子改良土 温度场 冻结 降温速率"
data6 = json.dumps({"keyword": keyword6}).encode('utf-8')
req6 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data6,
    headers={"Content-Type": "application/json"}
)
resp6 = urllib.request.urlopen(req6, timeout=15)
result6 = json.loads(resp6.read().decode('utf-8'))
print("=== 搜索6: 高分子改良土 温度场 冻结 降温速率 ===")
if result6.get('success'):
    print(result6['message'])
else:
    print(f"Error: {result6.get('message','unknown')}")
