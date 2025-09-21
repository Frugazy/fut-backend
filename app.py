import threading
import time
import requests
from flask import Flask, jsonify
import json

# ================= CONFIG =================
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1415893898437464165/4Jx8hZy01nwJ9Kxdds2Uw9R4MKzbzFGQV0WU4GqLReceKs8TS6LpqO3k3Dp8DossemuY"
STREAMLIT_URL = "https://h82xwhbwrrvrcozunacphm.streamlit.app"
PING_INTERVAL_MINUTES = 15
MAX_PLAYERS = 50
BUY_BUFFER_PERCENTAGE = 7
SPIKE_THRESHOLD = 15

# ==========================================

app = Flask(__name__)

# ========= KEEP-ALIVE FUNCTION ===========
def keep_alive():
    while True:
        try:
            requests.get(STREAMLIT_URL)
            print("Pinged Streamlit dashboard successfully.")
        except Exception as e:
            print(f"Error pinging Streamlit: {e}")
        time.sleep(PING_INTERVAL_MINUTES * 60)

threading.Thread(target=keep_alive, daemon=True).start()
# ==========================================

# ======= MOCK / LIVE SCRAPE FUNCTION ======
# Replace this with your actual FUT scraping logic
def fetch_and_analyze_fut_data():
    # Example: mock live data
    # In production, replace with live FUT scrapes for EA FC 25 / 26
    players = [
        {
            "player": {
                "id": "fut_birthday_001",
                "name": "N'Golo Kanté",
                "rating": 89,
                "position": "CDM",
                "club": "Chelsea",
                "league": "Premier League",
                "cardType": "fut_birthday"
            },
            "market": {
                "currentBIN": 55000,
                "historical": {
                    "24h": {"low": 55000, "high": 55705},
                    "7d": {"low": 47746, "high": 55705},
                    "14d": {"low": 47746, "high": 61599}
                },
                "dataPoints": 17
            },
            "trading": {
                "targetBuy": 44404,
                "targetSell": 52250,
                "estimatedProfit": 7846,
                "profitMargin": 17.67,
                "classification": "Certified Buy",
                "confidence": "High",
                "reasoning": "Near historical low with consistent profit patterns"
            }
        }
        # Add other players here or loop actual scrape
    ]
    return players[:MAX_PLAYERS]

# ========== DISCORD ALERT FUNCTION ==========
def send_discord_alert(player_data):
    embeds = []
    for player in player_data:
        if player["trading"]["classification"] in ["Certified Buy", "Hold Until 6PM"] or player["trading"]["profitMargin"] > 10:
            embed = {
                "title": f"{player['player']['name']} ({player['player']['rating']}) - {player['player']['cardType']}",
                "description": (
                    f"💰 **Current BIN:** {player['market']['currentBIN']:,} coins\n"
                    f"📉 **24h Low:** {player['market']['historical']['24h']['low']:,} coins\n"
                    f"📈 **7d High:** {player['market']['historical']['7d']['high']:,} coins\n"
                    f"🎯 **Target Buy:** {player['trading']['targetBuy']:,} coins\n"
                    f"🏷️ **Target Sell:** {player['trading']['targetSell']:,} coins\n"
                    f"📊 **Profit:** {player['trading']['estimatedProfit']:,} coins ({player['trading']['profitMargin']}%)\n"
                    f"🛡️ **Risk:** {player['trading']['classification']}\n"
                    f"💡 **Reasoning:** {player['trading']['reasoning']}"
                ),
                "color": 3066993 if player["trading"]["classification"] == "Certified Buy" else 15844367 if player["trading"]["classification"] == "Hold Until 6PM" else 15158332,
                "fields": [
                    {"name": "Position", "value": player["player"].get("position", "N/A"), "inline": True},
                    {"name": "Club", "value": player["player"].get("club", "N/A"), "inline": True},
                    {"name": "League", "value": player["player"].get("league", "N/A"), "inline": True}
                ],
                "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "footer": {"text": f"Confidence: {player['trading']['confidence']} | Data Points: {player['market']['dataPoints']}"}
            }
            embeds.append(embed)
    if embeds:
        payload = {"embeds": embeds}
        try:
            response = requests.post(DISCORD_WEBHOOK_URL, json=payload)
            print(f"Discord alert sent, status: {response.status_code}")
        except Exception as e:
            print(f"Error sending Discord alert: {e}")

# ============ FLASK ROUTES =================
@app.route("/data", methods=["GET"])
def get_player_data():
    try:
        data = fetch_and_analyze_fut_data()
        # Send Discord alerts
        send_discord_alert(data)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ============ RUN BACKEND =================
if __name__ == "__main__":
    # For Render / Replit you might use: gunicorn app:app
    app.run(host="0.0.0.0", port=5000, debug=True)