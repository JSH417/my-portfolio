import urllib.request
import re

req = urllib.request.Request('https://finance.naver.com/marketindex/goldDetail.naver', headers={'User-Agent': 'Mozilla/5.0'})
try:
    with urllib.request.urlopen(req) as response:
        html = response.read().decode('euc-kr', errors='ignore')
        
    m = re.search(r'class="no_today"[\s\S]*?<span class="blind">([0-9,]+)</span>', html)
    if m:
        print("Naver 1g Gold price:", m.group(1))
    else:
        m2 = re.search(r'(\d{2,3},\d{3})</em>\s*<span class="blind">원', html)
        if m2:
            print("Found:", m2.group(1))
        else:
            print("Snippet around gold:", html[:1000])
except Exception as e:
    print("Error:", e)

