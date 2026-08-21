import urllib.request, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
port = 19000

# Search 1: APAM改良土导热/降温机理
keyword1 = "聚丙烯酰胺改良土导热系数降温机理冻胀"
data1 = json.dumps({"keyword": keyword1}).encode('utf-8')
req1 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data1,
    headers={"Content-Type": "application/json"}
)
resp1 = urllib.request.urlopen(req1, timeout=15)
result1 = json.loads(resp1.read().decode('utf-8'))
print("=== 搜索1: 聚丙烯酰胺改良土导热系数降温机理冻胀 ===")
if result1.get('success'):
    print(result1['message'])
else:
    print(f"Error: {result1.get('message','unknown')}")

print("\n")

# Search 2: PAM soil thermal conductivity
keyword2 = "polyacrylamide改良土温度场冻结"
data2 = json.dumps({"keyword": keyword2}).encode('utf-8')
req2 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data2,
    headers={"Content-Type": "application/json"}
)
resp2 = urllib.request.urlopen(req2, timeout=15)
result2 = json.loads(resp2.read().decode('utf-8'))
print("=== 搜索2: polyacrylamide改良土温度场冻结 ===")
if result2.get('success'):
    print(result2['message'])
else:
    print(f"Error: {result2.get('message','unknown')}")

print("\n")

# Search 3: 土体改良剂导热性能冻胀抑制
keyword3 = "土壤改良剂导热性能冻胀抑制机理"
data3 = json.dumps({"keyword": keyword3}).encode('utf-8')
req3 = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data3,
    headers={"Content-Type": "application/json"}
)
resp3 = urllib.request.urlopen(req3, timeout=15)
result3 = json.loads(resp3.read().decode('utf-8'))
print("=== 搜索3: 土壤改良剂导热性能冻胀抑制机理 ===")
if result3.get('success'):
    print(result3['message'])
else:
    print(f"Error: {result3.get('message','unknown')}")
