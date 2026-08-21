import urllib.request, json, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
port = 19000
keyword = "APAM polyacrylamide soil thermal conductivity frost heave mechanism"
data = json.dumps({"keyword": keyword}).encode('utf-8')
req = urllib.request.Request(
    f"http://localhost:{port}/proxy/prosearch/search",
    data=data,
    headers={"Content-Type": "application/json"}
)
resp = urllib.request.urlopen(req, timeout=15)
result = json.loads(resp.read().decode('utf-8'))
if result.get('success'):
    print(result['message'])
else:
    print(f"Error: {result.get('message','unknown')}")
