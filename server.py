#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
개인 투자 포트폴리오 트래커 - 로컬 연동 중계 서버
- 포트폴리오 웹앱 정적 파일 서빙 (http://localhost:8080)
- 토스증권 Open API 연동 (OAuth2 토큰, 환율, 해외/국내 시세)
- 네이버 증권 연동 (국내 종목 현재가, 예상 배당률, KRX 금시세)
- 표준 라이브러리(urllib, http.server, json)만 사용하여 추가 패키지 설치 불필요
"""

import http.server
import socketserver
import urllib.request
import urllib.parse
import json
import os
import re
import ssl
import sys
import time
from datetime import datetime

# Windows 콘솔 인코딩 대응
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

PORT = int(os.environ.get("PORT", 8080))
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
WEB_DIR = os.path.dirname(os.path.abspath(__file__))

# SSL 컨텍스트
ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode = ssl.CERT_NONE

HEADERS_CHROME = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Referer": "https://m.stock.naver.com/"
}

# 토스증권 토큰 캐시
toss_token_cache = {
    "access_token": None,
    "expires_at": 0
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"toss_client_id": "", "toss_client_secret": ""}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# --------------------------------------------------------------------------
# 1. 토스증권 Open API 클라이언트
# --------------------------------------------------------------------------
def get_toss_access_token():
    cfg = load_config()
    client_id = cfg.get("toss_client_id", "").strip()
    client_secret = cfg.get("toss_client_secret", "").strip()

    if not client_id or not client_secret:
        return None

    # 캐시된 토큰 확인
    if toss_token_cache["access_token"] and time.time() < toss_token_cache["expires_at"]:
        return toss_token_cache["access_token"]

    try:
        url = "https://openapi.tossinvest.com/oauth2/token"
        payload = json.dumps({
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret
        }).encode("utf-8")

        req = urllib.request.Request(
            url,
            data=payload,
            headers={"Content-Type": "application/json", "User-Agent": "PortfolioTracker/1.0"}
        )
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            token = data.get("access_token")
            expires_in = data.get("expires_in", 3600)
            if token:
                toss_token_cache["access_token"] = token
                toss_token_cache["expires_at"] = time.time() + expires_in - 60
                return token
    except Exception as e:
        print(f"[Toss API Token Error]: {e}")
    return None

def fetch_toss_exchange_rate():
    token = get_toss_access_token()
    if not token:
        return None

    try:
        url = "https://openapi.tossinvest.com/api/v1/exchange-rate"
        req = urllib.request.Request(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
                "User-Agent": "PortfolioTracker/1.0"
            }
        )
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            # 응답 구조 파싱 (토스 API 명세 기준)
            rate = data.get("rate") or data.get("exchangeRate") or (data.get("result", {}).get("rate") if isinstance(data.get("result"), dict) else None)
            if rate:
                return float(rate)
    except Exception as e:
        print(f"[Toss API FX Error]: {e}")
    return None

def fetch_toss_stock_price(ticker):
    token = get_toss_access_token()
    if not token:
        return None

    try:
        url = f"https://openapi.tossinvest.com/api/v1/prices?symbols={ticker}"
        req = urllib.request.Request(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
                "User-Agent": "PortfolioTracker/1.0"
            }
        )
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            # 시세 정보 추출
            prices = data.get("prices") or data.get("result") or []
            if isinstance(prices, list) and len(prices) > 0:
                p = prices[0].get("price") or prices[0].get("closePrice") or prices[0].get("lastPrice")
                if p:
                    return float(p)
            elif isinstance(data, dict):
                p = data.get("price") or data.get("closePrice")
                if p:
                    return float(p)
    except Exception as e:
        print(f"[Toss API Price Error for {ticker}]: {e}")
    return None

# --------------------------------------------------------------------------
# 2. 네이버 증권 & 실시간 오픈 데이터 클라이언트
# --------------------------------------------------------------------------
def fetch_naver_stock_info(code):
    """
    네이버 증권 모바일 API를 통해 국내 종목명, 현재가, 배당수익률(%) 조회
    """
    code_str = str(code).strip().zfill(6)
    result = {
        "ticker": code_str,
        "name": "",
        "price": 0,
        "dividendYield": 0.0,
        "currency": "KRW",
        "country": "한국",
        "assetClass": "주식",
        "assetType": "개별 주식",
        "source": "네이버 증권"
    }

    # 1) 기본 정보 (이름, 현재가, 종목형태)
    try:
        url_basic = f"https://m.stock.naver.com/api/stock/{code_str}/basic"
        req_b = urllib.request.Request(url_basic, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req_b, context=ssl_ctx, timeout=6) as resp:
            data_b = json.loads(resp.read().decode("utf-8"))
            result["name"] = data_b.get("stockName", "")
            raw_price = data_b.get("closePrice", "0").replace(",", "")
            result["price"] = float(raw_price) if raw_price else 0

            # ETF 여부 판단
            if data_b.get("stockEndType") == "etf":
                result["assetType"] = "ETF"
                # ETF 이름에 '미국' 또는 '글로벌'이 들어간 경우 국가 자동 추천
                if "미국" in result["name"] or "S&P" in result["name"] or "나스닥" in result["name"]:
                    result["country"] = "미국"
                elif "글로벌" in result["name"] or "월드" in result["name"]:
                    result["country"] = "글로벌"
    except Exception as e:
        print(f"[Naver Basic Error {code_str}]: {e}")

    # 2) 종합 정보 (배당수익률 / 분배율)
    try:
        url_int = f"https://m.stock.naver.com/api/stock/{code_str}/integration"
        req_i = urllib.request.Request(url_int, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req_i, context=ssl_ctx, timeout=6) as resp:
            data_i = json.loads(resp.read().decode("utf-8"))

            # A. ETF의 경우: etfKeyIndicator.dividendYieldTtm (최근 12개월 분배율)
            etf_ind = data_i.get("etfKeyIndicator")
            if etf_ind and "dividendYieldTtm" in etf_ind and etf_ind["dividendYieldTtm"] is not None:
                result["dividendYield"] = round(float(etf_ind["dividendYieldTtm"]), 2)

            # B. 일반 주식의 경우: totalInfos의 '배당수익률'
            if result["dividendYield"] == 0.0:
                for item in data_i.get("totalInfos", []):
                    k = item.get("key", "")
                    v = item.get("value", "")
                    if "배당수익률" in k:
                        m = re.search(r"([0-9.]+)", v)
                        if m:
                            result["dividendYield"] = round(float(m.group(1)), 2)
                        break
    except Exception as e:
        print(f"[Naver Integration Error {code_str}]: {e}")

    return result

def fetch_krx_gold():
    """KRX 금현물 1g 현재가 조회 (네이버 증권 금시세)"""
    try:
        url = "https://api.stock.naver.com/marketindex/metals/M04020000"
        req = urllib.request.Request(url, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            raw_price = str(data.get("closePrice", "0")).replace(",", "")
            price = float(raw_price) if raw_price and float(raw_price) > 0 else 177480
            return {
                "ticker": "KRX금",
                "name": "KRX 금 1g",
                "price": price,
                "dividendYield": 0.0,
                "currency": "KRW",
                "country": "한국",
                "assetClass": "금",
                "assetType": "금현물",
                "source": "네이버 금시세"
            }
    except Exception as e:
        print(f"[Gold Error]: {e}")
        return {
            "ticker": "KRX금",
            "name": "KRX 금 1g",
            "price": 177480,
            "dividendYield": 0.0,
            "currency": "KRW",
            "country": "한국",
            "assetClass": "금",
            "assetType": "금현물",
            "source": "네이버 금시세"
        }

def fetch_us_stock_yahoo(ticker):
    """미국 종목 현재가 및 연환산 배당률 조회 (Yahoo Finance)"""
    try:
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1y&events=div"
        req = urllib.request.Request(url, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            res = data["chart"]["result"][0]
            meta = res.get("meta", {})
            price = float(meta.get("regularMarketPrice") or meta.get("previousClose") or 0)
            name = meta.get("shortName") or meta.get("symbol") or ticker

            # 최근 1년 배당금 합산하여 배당률 계산
            div_events = res.get("events", {}).get("dividends", {})
            div_sum = sum(v.get("amount", 0) for v in div_events.values())
            dividend_yield = round((div_sum / price * 100), 2) if price > 0 else 0.0

            return {
                "ticker": ticker.upper(),
                "name": name,
                "price": price,
                "dividendYield": dividend_yield,
                "currency": "USD",
                "country": "미국",
                "assetClass": "주식",
                "assetType": "ETF" if "ETF" in name.upper() or ticker.upper() in ["QQQ", "SPY", "VOO", "IVV", "SCHD"] else "개별 주식",
                "source": "Yahoo Finance"
            }
    except Exception as e:
        print(f"[Yahoo Error {ticker}]: {e}")
    return None

def fetch_exchange_rate():
    """
    USD/KRW 환율 조회 (1순위: 토스증권 API, 2순위: 실시간 오픈 환율 API)
    """
    # 1순위: 토스증권 Open API
    toss_fx = fetch_toss_exchange_rate()
    if toss_fx:
        return {"rate": round(toss_fx, 2), "source": "토스증권 Open API", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

    # 2순위: 실시간 오픈 환율 API (open.er-api.com)
    try:
        url = "https://open.er-api.com/v6/latest/USD"
        req = urllib.request.Request(url, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            krw_rate = data.get("rates", {}).get("KRW")
            if krw_rate:
                return {"rate": round(float(krw_rate), 2), "source": "실시간 환율 피드", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    except Exception as e:
        print(f"[Open FX Error]: {e}")

    # Fallback 기본 환율
    return {"rate": 1380.0, "source": "기본 환율 (조회 실패)", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

# --------------------------------------------------------------------------
# 3. HTTP 요청 핸들러
# --------------------------------------------------------------------------
class PortfolioProxyHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def end_headers(self):
        # CORS 허용
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        params = urllib.parse.parse_qs(parsed.query)

        # 1) 환율 조회 API: /api/fx
        if path == "/api/fx":
            fx_data = fetch_exchange_rate()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(fx_data, ensure_ascii=False).encode("utf-8"))
            return

        # 2) 시세 및 배당률 조회 API: /api/quote?ticker=...
        elif path == "/api/quote":
            ticker = params.get("ticker", [""])[0].strip().upper()
            if not ticker:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b'{"error": "Ticker is required"}')
                return

            res = None

            # A. 금현물 여부
            if ticker in ["금", "금현물", "KRX금", "KRX"]:
                res = fetch_krx_gold()

            # B. 국내 종목 (숫자 6자리)
            elif ticker.isdigit() or (len(ticker) <= 6 and re.match(r"^\d+$", ticker)):
                res = fetch_naver_stock_info(ticker)

            # C. 미국/해외 종목 (알파벳 티커: QQQ, NVDA, AAPL 등)
            else:
                # 1순위: 토스증권 API로 가격 조회 시도
                toss_price = fetch_toss_stock_price(ticker)
                if toss_price:
                    res = {
                        "ticker": ticker,
                        "name": ticker,
                        "price": toss_price,
                        "dividendYield": 0.0,
                        "currency": "USD",
                        "country": "미국",
                        "assetClass": "주식",
                        "assetType": "ETF" if ticker in ["QQQ", "SPY", "VOO", "IVV", "SCHD"] else "개별 주식",
                        "source": "토스증권 Open API"
                    }
                else:
                    # 2순위: Yahoo Finance
                    res = fetch_us_stock_yahoo(ticker)

            if res:
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
            else:
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": f"종목 '{ticker}' 조회 실패"}, ensure_ascii=False).encode("utf-8"))
            return

        # 3) 토스 API 설정 상태 확인 API: /api/config
        elif path == "/api/config":
            cfg = load_config()
            client_id = cfg.get("toss_client_id", "")
            has_secret = bool(cfg.get("toss_client_secret", ""))
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            resp_body = {
                "tossConfigured": bool(client_id and has_secret),
                "clientId": client_id
            }
            self.wfile.write(json.dumps(resp_body, ensure_ascii=False).encode("utf-8"))
            return

        # 기본 정적 파일 서빙 (index.html 등)
        super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/config":
            content_length = int(self.headers.get("Content-Length", 0))
            post_body = self.rfile.read(content_length)
            try:
                data = json.loads(post_body.decode("utf-8"))
                cfg = load_config()
                if "toss_client_id" in data:
                    cfg["toss_client_id"] = data["toss_client_id"].strip()
                if "toss_client_secret" in data:
                    cfg["toss_client_secret"] = data["toss_client_secret"].strip()
                save_config(cfg)
                # 토큰 캐시 리셋
                toss_token_cache["access_token"] = None
                toss_token_cache["expires_at"] = 0

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "message": "설정이 성공적으로 저장되었습니다."}, ensure_ascii=False).encode("utf-8"))
            except Exception as e:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            return

        self.send_response(404)
        self.end_headers()

def run_server():
    # 주소 재사용 설정
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), PortfolioProxyHandler) as httpd:
        print("=" * 60)
        print(f"[PORTFOLIO SERVER] Running at: http://localhost:{PORT}")
        print(f" - Web App: http://localhost:{PORT}/index.html")
        print(" - Naver Finance: Domestic price, dividend yield, KRX gold")
        print(" - Toss Open API: USD/KRW exchange rate, US stock prices")
        print("=" * 60)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n서버를 종료합니다.")

if __name__ == "__main__":
    run_server()
