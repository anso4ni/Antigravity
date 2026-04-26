# -*- coding: utf-8 -*-
"""
配置模块 config.py
通过读写 JSON 配置文件获取所有配置项
"""

import json
import os
from datetime import datetime

CONFIG_FILE = "portfolio_config.json"

# 默认配置模板
DEFAULT_CONFIG = {
    "display_currency": "Default",  # Default / TWD / USD
    "cash_holdings": [
        {"currency": "TWD", "amount": 0.0, "note": "台幣現金"},
        {"currency": "USD", "amount": 0.0, "note": "美元現金"},
    ],
    "stock_holdings": [
        # 範例:
        # {
        #     "symbol": "2330.TW",
        #     "name": "台積電",
        #     "market": "TW",        # TW 或 US
        #     "shares": 1000,
        #     "avg_cost": 580.0,
        #     "currency": "TWD",
        #     "note": ""
        # }
    ],
    "transactions": [
        # 範例:
        # {
        #     "date": "2025-01-15",
        #     "symbol": "2330.TW",
        #     "name": "台積電",
        #     "market": "TW",
        #     "action": "BUY",        # BUY 或 SELL
        #     "shares": 1000,
        #     "price": 580.0,
        #     "fee": 585.0,
        #     "tax": 0.0,
        #     "currency": "TWD",
        #     "note": "首次買入"
        # }
    ],
    "watchlist_indices": [
        {"symbol": "^DJI", "name": "道瓊工業指數"},
        {"symbol": "^IXIC", "name": "那斯達克"},
        {"symbol": "^GSPC", "name": "標普500指數"},
        {"symbol": "USDTWD=X", "name": "美元對台幣"},
        {"symbol": "TWD=X", "name": "台幣匯率"},
    ],
    "google_drive": {
        "oauth_client_secrets_path": "",  # OAuth2 用戶端金鑰 JSON（Desktop app）
        "credentials_path": "",           # 服務帳戶 JSON（舊版，備用）
        "spreadsheet_name": "",           # 匯入來源試算表名稱
        "spreadsheet_id": "",             # 匯入來源試算表 ID（優先使用）
        "worksheet_name": "",             # 工作表名稱（空白=第一個）
        "output_spreadsheet_id": "",      # 備份目標試算表 ID
        "last_synced": "",                # 上次同步時間
    },
    "last_updated": None,
}


def get_config_path():
    """取得當前使用者的配置文件路徑"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "portfolio_config.json"
    
    try:
        import streamlit as st
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        if get_script_run_ctx() is not None:
            email = st.session_state.get("google_email", "")
            if email:
                safe_email = "".join(c if c.isalnum() else "_" for c in email)
                filename = f"portfolio_data_{safe_email}.json"
            else:
                filename = "portfolio_data_anonymous.json"
    except Exception:
        pass
        
    return os.path.join(base_dir, filename)


def get_shared_config_path():
    """取得全域共用的系統配置文件路徑 (用於存儲 OAuth 金鑰路徑等)"""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "portfolio_config.json")


def load_config():
    """讀取配置文件，若不存在則建立默認配置"""
    shared_path = get_shared_config_path()
    shared_config = {}
    if os.path.exists(shared_path):
        try:
            with open(shared_path, "r", encoding="utf-8") as f:
                shared_config = json.load(f)
        except Exception:
            pass

    config_path = get_config_path()
    
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            for key, value in DEFAULT_CONFIG.items():
                if key not in config:
                    config[key] = value
        except (json.JSONDecodeError, IOError) as e:
            print(f"⚠️ 配置文件讀取失敗，使用默認配置: {e}")
            config = DEFAULT_CONFIG.copy()
    else:
        config = DEFAULT_CONFIG.copy()
        if "anonymous" not in config_path:
            save_config(config)

    # 確保全域 OAuth 設定強制套用
    if "google_drive" in shared_config and "oauth_client_secrets_path" in shared_config["google_drive"]:
        config.setdefault("google_drive", {})["oauth_client_secrets_path"] = shared_config["google_drive"]["oauth_client_secrets_path"]

    return config


def save_config(config):
    """保存配置到文件，並同步全域設定到共用設定檔"""
    config_path = get_config_path()
    config["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 儲存使用者專屬檔案 (匿名者不存檔)
    if "anonymous" not in config_path:
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except IOError as e:
            print(f"⚠️ 配置文件保存失敗: {e}")

    # 同步儲存全域設定到 portfolio_config.json
    shared_path = get_shared_config_path()
    shared_config = {}
    if os.path.exists(shared_path):
        try:
            with open(shared_path, "r", encoding="utf-8") as f:
                shared_config = json.load(f)
        except Exception:
            pass
            
    # 只有當設定中有 OAuth path 才更新
    if "google_drive" in config and "oauth_client_secrets_path" in config["google_drive"]:
        shared_config.setdefault("google_drive", {})["oauth_client_secrets_path"] = config["google_drive"]["oauth_client_secrets_path"]
        try:
            with open(shared_path, "w", encoding="utf-8") as f:
                json.dump(shared_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ 共用配置保存失敗: {e}")


def add_stock_holding(config, symbol, name, market, shares, avg_cost, currency="TWD", note=""):
    """
    新增股票持倉
    """
    holding = {
        "symbol": symbol,
        "name": name,
        "market": market,
        "shares": int(shares),
        "avg_cost": float(avg_cost),
        "currency": currency,
        "note": note,
    }
    # 檢查是否已存在，若存在則更新
    for i, h in enumerate(config["stock_holdings"]):
        if h["symbol"] == symbol:
            config["stock_holdings"][i] = holding
            save_config(config)
            return config
    config["stock_holdings"].append(holding)
    save_config(config)
    return config


def remove_stock_holding(config, symbol):
    """
    移除股票持倉
    """
    config["stock_holdings"] = [
        h for h in config["stock_holdings"] if h["symbol"] != symbol
    ]
    save_config(config)
    return config


def update_cash(config, currency, amount, note=""):
    """
    更新現金持倉
    """
    for cash in config["cash_holdings"]:
        if cash["currency"] == currency:
            cash["amount"] = float(amount)
            if note:
                cash["note"] = note
            save_config(config)
            return config
    # 若不存在則新增
    config["cash_holdings"].append({
        "currency": currency,
        "amount": float(amount),
        "note": note or f"{currency}現金",
    })
    save_config(config)
    return config


def add_transaction(config, date, symbol, name, market, action, shares, price, fee=0, tax=0, currency="TWD", note=""):
    """
    新增交易紀錄
    """
    transaction = {
        "date": date,
        "symbol": symbol,
        "name": name,
        "market": market,
        "action": action,
        "shares": int(shares),
        "price": float(price),
        "fee": float(fee),
        "tax": float(tax),
        "currency": currency,
        "note": note,
    }
    config["transactions"].append(transaction)

    # 自動更新持倉
    _update_holdings_from_transaction(config, transaction)
    save_config(config)
    return config


def _update_holdings_from_transaction(config, txn):
    """
    根據交易紀錄自動更新持倉
    """
    symbol = txn["symbol"]
    existing = None
    for h in config["stock_holdings"]:
        if h["symbol"] == symbol:
            existing = h
            break

    if txn["action"] == "BUY":
        if existing:
            total_cost = existing["avg_cost"] * existing["shares"] + txn["price"] * txn["shares"]
            total_shares = existing["shares"] + txn["shares"]
            existing["avg_cost"] = total_cost / total_shares if total_shares > 0 else 0
            existing["shares"] = total_shares
        else:
            config["stock_holdings"].append({
                "symbol": symbol,
                "name": txn["name"],
                "market": txn["market"],
                "shares": txn["shares"],
                "avg_cost": txn["price"],
                "currency": txn["currency"],
                "note": "",
            })
        # 扣除現金
        cost = txn["price"] * txn["shares"] + txn["fee"] + txn["tax"]
        for cash in config["cash_holdings"]:
            if cash["currency"] == txn["currency"]:
                cash["amount"] -= cost
                break

    elif txn["action"] == "SELL":
        if existing:
            existing["shares"] -= txn["shares"]
            if existing["shares"] <= 0:
                config["stock_holdings"] = [
                    h for h in config["stock_holdings"] if h["symbol"] != symbol
                ]
        # 增加現金
        revenue = txn["price"] * txn["shares"] - txn["fee"] - txn["tax"]
        for cash in config["cash_holdings"]:
            if cash["currency"] == txn["currency"]:
                cash["amount"] += revenue
                break


def set_display_currency(config, currency):
    """
    設定顯示幣值 (TWD 或 USD)
    """
    config["display_currency"] = currency
    save_config(config)
    return config


def get_all_symbols(config):
    """
    取得所有持倉的股票代碼列表
    """
    return [h["symbol"] for h in config["stock_holdings"]]


def recalculate_holdings(config):
    """
    根據所有交易紀錄重新計算持倉
    """
    config["stock_holdings"] = []
    # 重設現金
    for cash in config["cash_holdings"]:
        cash["amount"] = 0.0

    transactions_backup = config["transactions"][:]
    config["transactions"] = []

    for txn in transactions_backup:
        config["transactions"].append(txn)
        _update_holdings_from_transaction(config, txn)

    save_config(config)
    return config
