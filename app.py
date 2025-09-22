# app.py
from flask import Flask, jsonify
import requests
from bs4 import BeautifulSoup
import time
import json
import os
from datetime import datetime, timedelta
import math

app = Flask(__name__)

# ========== CONFIG ==========
PLATFORM = "ps"  # only PS (user requested PS only)
FUTWIZ_LIST_BASE = "https://www.futwiz.com/en/fc25/players?page={}"   # page list
FUTWIZ_PRICE_API = "https://www.futwiz.com/en/fc25/playerPrices?player={}"  # returns JSON
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}
MAX_PLAYERS = 50                 # dashboard expects 50 players
PRICE_CALL_DELAY = 0.45          # seconds between each player price call (politeness & rate-limit)
PAGE_DELAY = 0.8                 # delay between list page fetches
HISTORY_FILE = "history.json"    # local persistence (not durable across immutable containers)
HISTORY_MAX_POINTS = 500         # keep up to this many points per player
CACHE_TTL = 60                   # seconds to cache /data response
EA_TAX = 0.05
BUY_BUFFER = 0.07                # 7% default buy buffer
PROFIT_TARGET = 0.20             # 20% default target profit
CLAMP_TOO_LOW_PCT = 0.10         # don't set targetBuy more than 10% below 24h low

# ========== in-memory caches ==========
_last_data_time = 0
_cached_response = None
_history = {}  # loaded from HISTORY_FILE

# ========== helpers: persistence ==========
def load_history():
    global _history
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                _history = json.load(f)
        except Exception:
            _history = {}
    else:
        _history = {}

def save_history():
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(_history, f)
    except Exception as e:
        print("Warning: failed to save history:", e)

# ========== time helpers ==========
def now_ts():
    return int(time.time())

def filter_window(points, seconds_window):
    """Return list of (ts, price) inside last seconds_window seconds."""
    cutoff = now_ts() - int(seconds_window)
    return [p for p in points if p[0] >= cutoff]

def stats_from_points(points):
    if not points:
        return {"low": 0, "high": 0, "avg": 0}
    prices = [int(p[1]) for p in points]
    return {"low": min(prices), "high": max(prices), "avg": int(sum(prices) / len(prices))}

# ========== futwiz scraping helpers ==========
def fetch_player_list(max_players=MAX_PLAYERS):
    """
    Scrape Futwiz players list pages until we have max_players unique players.
    Extract: name, slug, player_id (data-playerid or extracted), rating, position, club, nation, image URL, cardType if present (icon detection).
    """
    players = []
    seen_ids = set()
    page = 1
    # Defensive: stop after reasonable number of pages
    while len(players) < max_players and page <= 10:
        url = FUTWIZ_LIST_BASE.format(page)
        try:
            r = requests.get(url, headers=HEADERS, timeout=10)
            if r.status_code != 200:
                print(f"List page {page} returned {r.status_code}")
                break
            soup = BeautifulSoup(r.text, "html.parser")

            # Futwiz layouts usually have player cards with attributes we can parse.
            # We'll search for elements that carry player links / data attributes.
            # Try multiple selectors to be robust.
            cards = soup.select(".player-card, .card-player, .playerRow, .player-item")
            if not cards:
                # fallback: rows in tables
                rows = soup.select("table.players-table tbody tr")
                for rnode in rows:
                    try:
                        # columns parsing (may vary by Futwiz layout)
                        cols = rnode.find_all("td")
                        if len(cols) < 3:
                            continue
                        name_node = cols[1].select_one("a") or cols[1]
                        name = name_node.get_text(strip=True)
                        # find href to extract slug/id
                        href = name_node.get("href") or ""
                        # attempt to parse id from href like /player/slug/12345
                        parts = href.strip("/").split("/")
                        pid = None
                        if parts and parts[-1].isdigit():
                            pid = parts[-1]
                        rating = int(cols[0].get_text(strip=True))
                        position = cols[2].get_text(strip=True)
                        club = cols[3].get_text(strip=True) if len(cols) > 3 else ""
                        nation = ""
                        img = ""
                        cardType = "standard"
                        if pid and pid not in seen_ids:
                            seen_ids.add(pid)
                            players.append({
                                "id": pid,
                                "name": name,
                                "rating": rating,
                                "position": position,
                                "club": club,
                                "nation": nation,
                                "image": img,
                                "cardType": cardType
                            })
                            if len(players) >= max_players:
                                break
                    except Exception:
                        continue
                page += 1
                time.sleep(PAGE_DELAY)
                continue

            for c in cards:
                try:
                    # find a link / data-playerid attribute
                    # Try data-playerid attribute
                    pid = c.get("data-playerid") or c.get("data-player-id")
                    # find anchor with href if no data attribute
                    if not pid:
                        a = c.select_one("a[href*='/player/'], a[href*='/player/']")
                        if a:
                            href = a.get("href", "")
                            parts = href.strip("/").split("/")
                            # last part might be numeric id
                            if parts and parts[-1].isdigit():
                                pid = parts[-1]

                    # name
                    name_node = c.select_one(".player-name") or c.select_one(".name") or c.select_one("a")
                    name = name_node.get_text(strip=True) if name_node else "Unknown"

                    # rating
                    rating_node = c.select_one(".player-rating") or c.select_one(".rating")
                    rating = int(rating_node.get_text(strip=True)) if rating_node and rating_node.get_text(strip=True).isdigit() else 0

                    # position
                    position_node = c.select_one(".player-position") or c.select_one(".position")
                    position = position_node.get_text(strip=True) if position_node else ""

                    club_node = c.select_one(".player-club") or c.select_one(".club")
                    club = club_node.get("title") if club_node and club_node.get("title") else (club_node.get_text(strip=True) if club_node else "")

                    nation_node = c.select_one(".player-nation") or c.select_one(".nation")
                    nation = nation_node.get("title") if nation_node and nation_node.get("title") else (nation_node.get_text(strip=True) if nation_node else "")

                    img_tag = c.select_one("img")
                    img = ""
                    if img_tag:
                        img = img_tag.get("data-src") or img_tag.get("src") or ""
                        if img and img.startswith("/"):
                            img = "https://www.futwiz.com" + img

                    # card type: try detect icons or special classes
                    cardType = "standard"
                    cls = " ".join(c.get("class", []))
                    if "icon" in cls.lower():
                        cardType = "icon"
                    # fallback: try to find text label
                    label = c.select_one(".card-type") or c.select_one(".special")
                    if label:
                        lt = label.get_text(strip=True).lower()
                        if "icon" in lt:
                            cardType = "icon"

                    if not pid:
                        # skip if no id found
                        continue

                    if str(pid) in seen_ids:
                        continue
                    seen_ids.add(str(pid))
                    players.append({
                        "id": str(pid),
                        "name": name,
                        "rating": rating,
                        "position": position,
                        "club": club,
                        "nation": nation,
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
    """Query Futwiz price API for a player id and return PS BIN (or None)."""
    try:
        url = FUTWIZ_PRICE_API.format(player_id)
        r = requests.get(url, headers=HEADERS, timeout=10)
        if r.status_code != 200:
            return None
        data = r.json()
        # Futwiz response may include 'prices' keyed by platform abbreviations
        prices = data.get("prices", {}) or {}
        # only PS
        ps_price = prices.get("ps") or prices.get("ps4") or prices.get("ps5")
        if ps_price is None:
            return None
        # ensure int
        return int(ps_price)
    except Exception as e:
        # print("price fetch error", player_id, e)
        return None

# ========== target calculation & classification ==========
def compute_targets_from_history(player_id, current_bin, history_points):
    """
    history_points: list of [ (ts, price), ... ] sorted ascending oldest->newest
    Uses windows: 6h, 12h, 24h, 7d
    Returns computed dict: targetBuy, targetSell, estimatedProfit, profitMargin, classification, reasoning, lows/highs
    """
    now = now_ts()
    windows = {
        "6h": 6 * 3600,
        "12h": 12 * 3600,
        "24h": 24 * 3600,
        "7d": 7 * 24 * 3600
    }

    # build windowed stats
    stats = {}
    for name, seconds in windows.items():
        pts = filter_window(history_points, seconds) if history_points else []
        stats[name] = stats_from_points(pts)

    # primary lows: prefer recent windows (6h/12h/24h)
    recent_lows = [s["low"] for k,s in stats.items() if k in ("6h","12h","24h") and s["low"]>0]
    # fallback to 7d
    seven_day_low = stats["7d"]["low"] or (min(recent_lows) if recent_lows else 0)

    if recent_lows:
        raw_target = int(min(recent_lows) * (1 - BUY_BUFFER))
    else:
        raw_target = int(seven_day_low * (1 - BUY_BUFFER)) if seven_day_low else int(current_bin * (1 - BUY_BUFFER))

    # clamp: don't go more than CLAMP_TOO_LOW_PCT below 24h low
    low_24h = stats["24h"]["low"] or seven_day_low or current_bin
    clamp_floor = int(math.floor(low_24h * (1 - CLAMP_TOO_LOW_PCT)))
    if raw_target < clamp_floor:
        sanitized_target_buy = clamp_floor
    else:
        sanitized_target_buy = raw_target

    # Ensure targetBuy <= 24h_low (we want buy below or equal to recent low)
    if low_24h > 0 and sanitized_target_buy > low_24h:
        sanitized_target_buy = int(low_24h)

    # Ensure not zero
    if sanitized_target_buy <= 0:
        sanitized_target_buy = max(1, int(low_24h or current_bin or 1))

    targetBuy = sanitized_target_buy

    # SELL: use recent highs (6h/12h/24h), clamp by 7d high as an upper guard
    recent_highs = [s["high"] for k,s in stats.items() if k in ("6h","12h","24h") and s["high"]>0]
    seven_day_high = stats["7d"]["high"] or (max(recent_highs) if recent_highs else 0)
    if recent_highs:
        raw_sell = int(max(recent_highs) * (1 - 0.03))  # list slightly below recent peak
    else:
        raw_sell = int(seven_day_high * 0.97) if seven_day_high else int(current_bin * (1 + PROFIT_TARGET))

    targetSell = raw_sell if seven_day_high == 0 else min(raw_sell, seven_day_high)

    # compute estimated profit net of EA tax
    gross_needed = targetBuy * (1 + PROFIT_TARGET)
    if (1 - EA_TAX) <= 0:
        targetSellFromBuy = int(math.ceil(gross_needed))
    else:
        targetSellFromBuy = int(math.ceil(gross_needed / (1 - EA_TAX)))

    # final targetSell preference: show the more conservative (lower) of computed targetSell and targetSellFromBuy? Keep both
    finalTargetSell = int(max(targetSell, targetSellFromBuy))  # dashboard should display realistic sell target

    net_from_sell = finalTargetSell * (1 - EA_TAX)
    estimatedProfit = int(round(net_from_sell - targetBuy))
    profitMargin = round((estimatedProfit / targetBuy * 100) if targetBuy>0 else 0, 2)

    # volatility simple metric: use 24h range vs 24h avg
    volatility = 0.0
    if stats["24h"]["avg"] > 0:
        volatility = (stats["24h"]["high"] - stats["24h"]["low"]) / stats["24h"]["avg"]

    # classification enforce strict rule
    if current_bin <= targetBuy:
        classification = "Certified Buy"
        reasoning = f"Current BIN ({current_bin:,}) is at or below targetBuy ({targetBuy:,})."
    elif current_bin <= int(targetBuy * 1.03):
        classification = "Hold Until 6PM"
        reasoning = "Price slightly above targetBuy — consider waiting for lightning rounds."
    elif current_bin >= finalTargetSell or volatility > 0.6:
        classification = "High Risk"
        reasoning = "Price above targetSell or market volatile."
    else:
        classification = "Monitor"
        reasoning = "No immediate actionable buy."

    return {
        "targetBuy": int(targetBuy),
        "targetSell": int(finalTargetSell),
        "estimatedProfit": int(estimatedProfit),
        "profitMargin": float(profitMargin),
        "classification": classification,
        "reasoning": reasoning,
        "volatility": float(volatility),
        "stats": stats,
        "raw_target_candidate": int(raw_target),
        "clamp_floor": int(clamp_floor)
    }

# ========== main data assemble ==========
def build_data(max_players=MAX_PLAYERS):
    """
    Returns structure:
    {
      "players": [ { id,name,rating,position,club,nation,image,currentBIN,historical,targetBuy,... }, ... ],
      "lastUpdated": ISO
    }
    """
    # 1) fetch master player list (ids + metadata)
    players_meta = fetch_player_list(max_players)
    results = []
    for meta in players_meta:
        pid = meta.get("id")
        if not pid:
            continue

        # 2) get live PS bin for player
        current_bin = fetch_player_price(pid)
        # skip if no bin (still include metadata with currentBIN=0 so dashboard can show)
        if current_bin is None:
            # still update history with no price? skip to avoid gaps
            continue

        # 3) update persistent history
        entry = _history.get(pid, [])
        entry.append((now_ts(), int(current_bin)))
        # keep max length and also trim older than 7d to keep file small
        cutoff = now_ts() - 8 * 24 * 3600
        entry = [p for p in entry if p[0] >= cutoff]
        if len(entry) > HISTORY_MAX_POINTS:
            entry = entry[-HISTORY_MAX_POINTS:]
        _history[pid] = entry

        # 4) compute windows
        # windows use helper filter_window in compute_targets
        computed = compute_targets_from_history(pid, int(current_bin), entry)

        # 5) create historical stats object to return for dashboard (6h/12h/24h/7d)
        # reconstruct stats; compute_targets already returns stats
        hist_for_dashboard = computed["stats"]

        results.append({
            "id": pid,
            "name": meta.get("name"),
            "rating": meta.get("rating"),
            "position": meta.get("position"),
            "club": meta.get("club"),
            "nation": meta.get("nation"),
            "image": meta.get("image"),
            "cardType": meta.get("cardType"),
            "currentBIN": int(current_bin),
            "historical": {
                "6h": hist_for_dashboard.get("6h", {}),
                "12h": hist_for_dashboard.get("12h", {}),
                "24h": hist_for_dashboard.get("24h", {}),
                "7d": hist_for_dashboard.get("7d", {})
            },
            "trading": {
                "targetBuy": computed["targetBuy"],
                "targetSell": computed["targetSell"],
                "estimatedProfit": computed["estimatedProfit"],
                "profitMargin": computed["profitMargin"],
                "classification": computed["classification"],
                "reasoning": computed["reasoning"],
                "volatility": computed["volatility"],
                "meta": {
                    "raw_target_candidate": computed["raw_target_candidate"],
                    "clamp_floor": computed["clamp_floor"]
                }
            },
            "lastUpdated": datetime.utcnow().isoformat()
        })

        # polite delay between price calls
        time.sleep(PRICE_CALL_DELAY)

    # persist history
    save_history()
    return {"players": results, "lastUpdated": datetime.utcnow().isoformat()}

# ========== routes ==========
@app.route("/")
def root():
    return "✅ FUT Backend (FC25 PS) - /data for JSON", 200

@app.route("/ping")
def ping():
    return jsonify({"status": "alive", "time": datetime.utcnow().isoformat()})

@app.route("/data")
def data_route():
    global _last_data_time, _cached_response
    # simple in-memory cache to avoid spamming Futwiz on each dashboard refresh
    if time.time() - _last_data_time < CACHE_TTL and _cached_response:
        return jsonify(_cached_response)
    # build new snapshot
    try:
        snapshot = build_data(MAX_PLAYERS)
        _cached_response = snapshot
        _last_data_time = time.time()
        return jsonify(snapshot)
    except Exception as e:
        print("Error building data:", e)
        # return best-effort partial response if cache exists
        if _cached_response:
            return jsonify(_cached_response)
        return jsonify({"players": [], "error": str(e)}), 500

if __name__ == "__main__":
    load_history()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
