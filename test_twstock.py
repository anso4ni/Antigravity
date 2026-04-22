import sys
sys.path.append('d:/Antigravity')
import data

print("Fetching 2330.TW...")
res = data._fetch_price("2330.TW")
print(res)

print("Fetching 8299.TWO...")
res = data._fetch_price("8299.TWO")
print(res)
