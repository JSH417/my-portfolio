#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vercel Serverless Function - Portfolio Tracker API
Endpoints:
- /api/fx : 환율 조회 (실시간 환율 피드 / 토스 API)
- /api/quote?ticker=... : 국내(네이버 증권), 미국(Yahoo/토스), 금현물 시세 및 배당률
- /api/config : 토스 API 설정 확인 및 임시 저장
"""

import http.server
import urllib.request
import urllib.parse
import json
import os
import re
import ssl
import time
from datetime import datetime

ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode = ssl.CERT_NONE

HEADERS_CHROME = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Referer": "https://m.stock.naver.com/"
}

CONFIG_FILE = "/tmp/config.json"
memory_config = {
    "toss_client_id": os.environ.get("TOSS_CLIENT_ID", ""),
    "toss_client_secret": os.environ.get("TOSS_CLIENT_SECRET", "")
}

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
    return memory_config

def save_config(cfg):
    global memory_config
    memory_config.update(cfg)
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(memory_config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_toss_access_token():
    cfg = load_config()
    client_id = cfg.get("toss_client_id", "").strip()
    client_secret = cfg.get("toss_client_secret", "").strip()

    if not client_id or not client_secret:
        return None

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
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            token = data.get("access_token")
            expires_in = data.get("expires_in", 3600)
            if token:
                toss_token_cache["access_token"] = token
                toss_token_cache["expires_at"] = time.time() + expires_in - 60
                return token
    except Exception as e:
        print(f"[Toss Token Error]: {e}")
    return None

def fetch_toss_exchange_rate():
    token = get_toss_access_token()
    if not token:
        return None
    try:
        url = "https://openapi.tossinvest.com/api/v1/exchange-rate"
        req = urllib.request.Request(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            rate = data.get("rate") or data.get("exchangeRate")
            if rate:
                return float(rate)
    except Exception:
        pass
    return None

def fetch_toss_stock_price(ticker):
    token = get_toss_access_token()
    if not token:
        return None
    try:
        url = f"https://openapi.tossinvest.com/api/v1/prices?symbols={ticker}"
        req = urllib.request.Request(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            prices = data.get("prices") or data.get("result") or []
            if isinstance(prices, list) and len(prices) > 0:
                p = prices[0].get("price") or prices[0].get("closePrice") or prices[0].get("lastPrice")
                if p:
                    return float(p)
    except Exception:
        pass
    return None

def fetch_naver_stock_info(code):
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

    try:
        url_basic = f"https://m.stock.naver.com/api/stock/{code_str}/basic"
        req_b = urllib.request.Request(url_basic, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req_b, context=ssl_ctx, timeout=6) as resp:
            data_b = json.loads(resp.read().decode("utf-8"))
            result["name"] = data_b.get("stockName", "")
            raw_price = data_b.get("closePrice", "0").replace(",", "")
            result["price"] = float(raw_price) if raw_price else 0
            if data_b.get("stockEndType") == "etf":
                result["assetType"] = "ETF"
                if any(w in result["name"] for w in ["미국", "S&P", "나스닥"]):
                    result["country"] = "미국"
                elif any(w in result["name"] for w in ["글로벌", "월드"]):
                    result["country"] = "글로벌"
    except Exception:
        pass

    try:
        url_int = f"https://m.stock.naver.com/api/stock/{code_str}/integration"
        req_i = urllib.request.Request(url_int, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req_i, context=ssl_ctx, timeout=6) as resp:
            data_i = json.loads(resp.read().decode("utf-8"))
            etf_ind = data_i.get("etfKeyIndicator")
            if etf_ind and "dividendYieldTtm" in etf_ind and etf_ind["dividendYieldTtm"] is not None:
                result["dividendYield"] = round(float(etf_ind["dividendYieldTtm"]), 2)
            if result["dividendYield"] == 0.0:
                for item in data_i.get("totalInfos", []):
                    k = item.get("key", "")
                    v = item.get("value", "")
                    if "배당수익률" in k:
                        m = re.search(r"([0-9.]+)", v)
                        if m:
                            result["dividendYield"] = round(float(m.group(1)), 2)
                            break
    except Exception:
        pass

    return result if result["name"] else None

def fetch_krx_gold():
    try:
        url = "https://m.stock.naver.com/api/item/gold/basic"
        req = urllib.request.Request(url, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            raw_price = str(data.get("closePrice", "0")).replace(",", "")
            return {
                "ticker": "KRX금",
                "name": "KRX 금현물 1g",
                "price": float(raw_price) if raw_price else 105000,
                "dividendYield": 0.0,
                "currency": "KRW",
                "country": "한국",
                "assetClass": "금",
                "assetType": "금현물",
                "source": "네이버 증권 (KRX금)"
            }
    except Exception:
        return {
            "ticker": "KRX금",
            "name": "KRX 금현물 1g",
            "price": 105000,
            "dividendYield": 0.0,
            "currency": "KRW",
            "country": "한국",
            "assetClass": "금",
            "assetType": "금현물",
            "source": "기본 시세"
        }

def fetch_us_stock_yahoo(ticker):
    try:
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=5d"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            result = data.get("chart", {}).get("result", [])[0]
            meta = result.get("meta", {})
            price = meta.get("regularMarketPrice") or 0.0
            name = meta.get("shortName") or meta.get("symbol") or ticker
            dividend_yield = 0.0
            if ticker.upper() == "SCHD":
                dividend_yield = 3.52
            elif ticker.upper() == "JEPI":
                dividend_yield = 7.15
            elif ticker.upper() == "QQQ":
                dividend_yield = 0.58
            elif ticker.upper() in ["SPY", "VOO", "IVV"]:
                dividend_yield = 1.25

            return {
                "ticker": ticker.upper(),
                "name": name,
                "price": float(price),
                "dividendYield": dividend_yield,
                "currency": "USD",
                "country": "미국",
                "assetClass": "주식",
                "assetType": "ETF" if "ETF" in name.upper() or ticker.upper() in ["QQQ", "SPY", "VOO", "IVV", "SCHD"] else "개별 주식",
                "source": "Yahoo Finance"
            }
    except Exception:
        return None

def fetch_exchange_rate():
    toss_fx = fetch_toss_exchange_rate()
    if toss_fx:
        return {"rate": round(toss_fx, 2), "source": "토스증권 Open API", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    try:
        url = "https://open.er-api.com/v6/latest/USD"
        req = urllib.request.Request(url, headers=HEADERS_CHROME)
        with urllib.request.urlopen(req, context=ssl_ctx, timeout=6) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            krw_rate = data.get("rates", {}).get("KRW")
            if krw_rate:
                return {"rate": round(float(krw_rate), 2), "source": "실시간 환율 피드", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    except Exception:
        pass
    return {"rate": 1380.0, "source": "기본 환율 (조회 실패)", "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

class handler(http.server.BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        # Vercel rewrite 지원: __path__ 쿼리 또는 x-matched-path 또는 parsed.path 사용
        path = params.get("__path__", [None])[0] or self.headers.get("x-matched-path") or parsed.path

        # 0) 메인 페이지 및 정적 파일 서빙
        if path in ["/", "/index.html", ""]:
            candidates = [
                os.path.join(os.getcwd(), "index.html"),
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "index.html"),
                "index.html"
            ]
            for c in candidates:
                if os.path.exists(c):
                    with open(c, "rb") as f:
                        data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(data)
                    return

        static_mime = {
            "/manifest.json": "application/json",
            "/sw.js": "application/javascript",
            "/icon-192.png": "image/png",
            "/icon-512.png": "image/png",
            "/apple-touch-icon.png": "image/png",
            "/favicon.png": "image/png"
        }
        if path in static_mime:
            fname = path.lstrip("/")
            candidates = [
                os.path.join(os.getcwd(), fname),
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), fname),
                fname
            ]
            for c in candidates:
                if os.path.exists(c):
                    with open(c, "rb") as f:
                        data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", static_mime[path])
                    self.end_headers()
                    self.wfile.write(data)
                    return

        if path.endswith("/api/fx") or path == "/api/fx":
            data = fetch_exchange_rate()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))
            return

        elif "/api/quote" in path:
            ticker = params.get("ticker", [""])[0].strip().upper()
            if not ticker:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b'{"error": "Ticker is required"}')
                return

            res = None
            if ticker in ["금", "금현물", "KRX금", "KRX"]:
                res = fetch_krx_gold()
            elif ticker.isdigit() or (len(ticker) <= 6 and re.match(r"^\d+$", ticker)):
                res = fetch_naver_stock_info(ticker)
            else:
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

        elif "/api/config" in path:
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

        self.send_response(404)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps({"error": "Endpoint not found", "received_path": self.path, "parsed_path": path}).encode("utf-8"))

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        path = params.get("__path__", [None])[0] or self.headers.get("x-matched-path") or parsed.path

        if "/api/config" in path:
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

