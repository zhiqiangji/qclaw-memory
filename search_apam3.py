import urllib.request, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
port = 19000

# Search: Yang APAM Construction Building Materials 2023 loess
keyword = "Yang anionic polyacrylamide loess physico-mechanical microstructural 2023"
data = json.dumps({"keyword": keyword}).encode('utf-8')
req = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data,
    headers={"Content-Type": "application/json"}
)
resp = urllib.request.urlopen(req, timeout=15)
result = json.loads(resp.read().decode('utf-8'))
print("=== Yang APAM loess 2023 ===")
if result.get('success'):
    print(result['message'])
else:
    print(f"Error: {result.get('message','unknown')}")

print("\n")

# Search: PAM土壤降低导热 温度调节
keyword2 = "PAM聚丙烯酰胺 土壤 降低导热 温度调节 保水"
data2 = json.dumps({"keyword": keyword2}).encode('utf-8')
req2 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data2,
    headers={"Content-Type": "application/json"}
)
resp2 = urllib.request.urlopen(req2, timeout=15)
result2 = json.loads(resp2.read().decode('utf-8'))
print("=== PAM土壤降低导热 温度调节 保水 ===")
if result2.get('success'):
    print(result2['message'])
else:
    print(f"Error: {result2.get('message','unknown')}")

print("\n")

# Search: 水泥土改良 导热系数 增强 冻结 温度场
keyword3 = "改良土 导热系数 增大 冻结温度场 水泥土"
data3 = json.dumps({"keyword": keyword3}).encode('utf-8')
req3 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data3,
    headers={"Content-Type": "application/json"}
)
resp3 = urllib.request.urlopen(req3, timeout=15)
result3 = json.loads(resp3.read().decode('utf-8'))
print("=== 改良土导热系数增大冻结温度场 ===")
if result3.get('success'):
    print(result3['message'])
else:
    print(f"Error: {result3.get('message','unknown')}")
