# app.py
from flask import Flask, jsonify
import requests
from bs4 import BeautifulSoup
import time
import json
import os
from datetime import datetime
import math

app = Flask(__name__)

# =====================
# CONFIG
# =====================
PLATFORM = "ps"  # only PS
FUTWIZ_LIST_BASE = "https://www.futwiz.com/en/fc25/players?page={}"
FUTWIZ_PRICE_API = "https://www.futwiz.com/en/fc25/playerPrices?player={}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}
MAX_PLAYERS = 50
PRICE_CALL_DELAY = 0.45
PAGE_DELAY = 0.8
HISTORY_FILE = "history.json"
HISTORY_MAX_POINTS = 500
CACHE_TTL = 60
EA_TAX = 0.05
BUY_BUFFER = 0.07
PROFIT_TARGET = 0.20
CLAMP_TOO_LOW_PCT = 0.10

# =====================
# MEMORY
# =====================
_last_data_time = 0
_cached_response = None
_history = {}

# =====================
# HELPERS
# =====================
def load_history():
    global _history
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                _history = json.load(f)
        except:
            _history = {}
    else:
        _history = {}

def save_history():
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(_history, f)
    except Exception as e:
        print("Warning: failed to save history:", e)

def now_ts():
    return int(time.time())

def filter_window(points, seconds_window):
    cutoff = now_ts() - int(seconds_window)
    return [p for p in points if p[0] >= cutoff]

def stats_from_points(points):
    if not points:
        return {"low": 0, "high": 0, "avg": 0}
    prices = [int(p[1]) for p in points]
    return {"low": min(prices), "high": max(prices), "avg": int(sum(prices)/len(prices))}

# =====================
# FUTWIZ SCRAPE
# =====================
def fetch_player_list(max_players=MAX_PLAYERS):
    players = []
    seen_ids = set()
    page = 1
    while len(players) < max_players and page <= 10:
        url = FUTWIZ_LIST_BASE.format(page)
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, "html.parser")
            cards = soup.select(".player-card, .card-player, .playerRow, .player-item")
            if not cards:
                page += 1
                time.sleep(PAGE_DELAY)
                continue
            for c in cards:
                try:
                    pid = c.get("data-playerid") or c.get("data-player-id")
                    if not pid:
                        a = c.select_one("a[href*='/player/']")
                        if a:
                            href = a.get("href", "")
                            parts = href.strip("/").split("/")
                            if parts and parts[-1].isdigit():
                                pid = parts[-1]
                    if not pid or pid in seen_ids:
                        continue
                    seen_ids.add(pid)
                    name_node = c.select_one(".player-name") or c.select_one("a")
                    name = name_node.get_text(strip=True) if name_node else "Unknown"
                    rating_node = c.select_one(".player-rating") or c.select_one(".rating")
                    rating = int(rating_node.get_text(strip=True)) if rating_node and rating_node.get_text(strip=True).isdigit() else 0
                    position_node = c.select_one(".player-position") or c.select_one(".position")
                    position = position_node.get_text(strip=True) if position_node else ""
                    club_node = c.select_one(".player-club") or c.select_one(".club")
                    club = club_node.get("title") if club_node and club_node.get("title") else (club_node.get_text(strip=True) if club_node else "")
                    img_tag = c.select_one("img")
                    img = ""
                    if img_tag:
                        img = img_tag.get("data-src") or img_tag.get("src") or ""
                        if img and img.startswith("/"):
                            img = "https://www.futwiz.com" + img
                    cardType = "standard"
                    cls = " ".join(c.get("class", []))
                    if "icon" in cls.lower():
                        cardType = "icon"
                    players.append({
                        "id": str(pid),
                        "name": name,
                        "rating": rating,
                        "position": position,
                        "club": club,
                        "image": img,
                        "cardType": cardType
                    })
                    if len(players) >= max_players:
                        break
                except Exception:
                    continue
            page += 1
            time.sleep(PAGE_DELAY)
        except Exception as e:
            print("fetch_player_list error:", e)
            break
    return players

def fetch_player_price(player_id):
    try:
        url = FUTWIZ_PRICE_API.format(player_id)
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code != 200:
            return None
        data = r.json()
        prices = data.get("prices", {}) or {}
        ps_price = prices.get("ps") or prices.get("ps4") or prices.get("ps5")
        if ps_price is None:
            return None
        return int(ps_price)
    except:
        return None

# =====================
# TARGET CALC
# =====================
def compute_targets_from_history(player_id, current_bin, history_points):
    now = now_ts()
    windows = {"6h":6*3600,"12h":12*3600,"24h":24*3600,"7d":7*24*3600}
    stats = {}
    for name, seconds in windows.items():
        pts = filter_window(history_points, seconds) if history_points else []
        stats[name] = stats_from_points(pts)
    recent_lows = [s["low"] for k,s in stats.items() if k in ("6h","12h","24h") and s["low"]>0]
    seven_day_low = stats["7d"]["low"] or (min(recent_lows) if recent_lows else 0)
    raw_target = int(min(recent_lows)*(1-BUY_BUFFER)) if recent_lows else int(seven_day_low*(1-BUY_BUFFER)) if seven_day_low else int(current_bin*(1-BUY_BUFFER))
    low_24h = stats["24h"]["low"] or seven_day_low or current_bin
    clamp_floor = int(math.floor(low_24h*(1-CLAMP_TOO_LOW_PCT)))
    sanitized_target_buy = max(raw_target, clamp_floor)
    if sanitized_target_buy>low_24h:
        sanitized_target_buy = int(low_24h)
    if sanitized_target_buy<=0:
        sanitized_target_buy = max(1,int(low_24h or current_bin or 1))
    targetBuy = sanitized_target_buy
    recent_highs = [s["high"] for k,s in stats.items() if k in ("6h","12h","24h") and s["high"]>0]
    seven_day_high = stats["7d"]["high"] or (max(recent_highs) if recent_highs else 0)
    raw_sell = int(max(recent_highs)*(1-0.03)) if recent_highs else int(seven_day_high*0.97) if seven_day_high else int(current_bin*(1+PROFIT_TARGET))
    targetSell = raw_sell if seven_day_high==0 else min(raw_sell, seven_day_high)
    gross_needed = targetBuy*(1+PROFIT_TARGET)
    targetSellFromBuy = int(math.ceil(gross_needed/(1-EA_TAX))) if (1-EA_TAX)>0 else int(math.ceil(gross_needed))
    finalTargetSell = int(max(targetSell,targetSellFromBuy))
    net_from_sell = finalTargetSell*(1-EA_TAX)
    estimatedProfit = int(round(net_from_sell-targetBuy))
    profitMargin = round((estimatedProfit/targetBuy*100) if targetBuy>0 else 0,2)
    volatility = 0.0
    if stats["24h"]["avg"]>0:
        volatility = (stats["24h"]["high"]-stats["24h"]["low"])/stats["24h"]["avg"]
    if current_bin<=targetBuy:
        classification = "Certified Buy"
        reasoning = f"Current BIN ({current_bin:,}) <= targetBuy ({targetBuy:,})."
    elif current_bin<=int(targetBuy*1.03):
        classification = "Hold Until 6PM"
        reasoning = "Price slightly above targetBuy — wait for lightning rounds."
    elif current_bin>=finalTargetSell or volatility>0.6:
        classification = "High Risk"
        reasoning = "Price above targetSell or market volatile."
    else:
        classification = "Monitor"
        reasoning = "No immediate actionable buy."
    return {
        "targetBuy": targetBuy,
        "targetSell": finalTargetSell,
        "estimatedProfit": estimatedProfit,
        "profitMargin": profitMargin,
        "classification": classification,
        "reasoning": reasoning,
        "volatility": volatility,
        "stats": stats
    }

# =====================
# BUILD DATA
# =====================
def build_data(max_players=MAX_PLAYERS):
    players_meta = fetch_player_list(max_players)
    results = []
    for meta in players_meta:
        pid = meta.get("id")
        if not pid:
            continue
        current_bin = fetch_player_price(pid)
        if current_bin is None:
            continue
        entry = _history.get(pid, [])
        entry.append((now_ts(), int(current_bin)))
        cutoff = now_ts()-8*24*3600
        entry = [p for p in entry if p[0]>=cutoff]
        if len(entry)>HISTORY_MAX_POINTS:
            entry = entry[-HISTORY_MAX_POINTS:]
        _history[pid]=entry
        computed = compute_targets_from_history(pid,int(current_bin),entry)
        results.append({
            "id": pid,
            "name": meta.get("name"),
            "rating": meta.get("rating"),
            "position": meta.get("position"),
            "club": meta.get("club"),
            "image": meta.get("image"),
            "cardType": meta.get("cardType"),
            "currentBIN": int(current_bin),
            "historical": {
                "6h": computed["stats"].get("6h",{}),
                "12h": computed["stats"].get("12h",{}),
                "24h": computed["stats"].get("24h",{}),
                "7d": computed["stats"].get("7d",{})
            },
            "trading": computed,
            "lastUpdated": datetime.utcnow().isoformat()
        })
        time.sleep(PRICE_CALL_DELAY)
    save_history()
    return {"players": results, "lastUpdated": datetime.utcnow().isoformat()}

# =====================
# ROUTES
# =====================
@app.route("/")
def root():
    return "✅ FUT Backend (FC25 PS) - /data for JSON", 200

@app.route("/ping")
def ping():
    return jsonify({"status":"alive","time":datetime.utcnow().isoformat()})

@app.route("/data")
def data_route():
    global _last_data_time, _cached_response
    if time.time()-_last_data_time<CACHE_TTL and _cached_response:
        return jsonify(_cached_response)
    try:
        snapshot = build_data(MAX_PLAYERS)
        _cached_response = snapshot
        _last_data_time = time.time()
        return jsonify(snapshot)
    except Exception as e:
        print("Error building data:",e)
        if _cached_response:
            return jsonify(_cached_response)
        return jsonify({"players":[],"error":str(e)}),500

# =====================
# MAIN
# =====================
if __name__=="__main__":
    load_history()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",10000)))
