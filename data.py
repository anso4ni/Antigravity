# -*- coding: utf-8 -*-
"""
數據模塊 data.py
通過 yfinance 接口獲取台股與美股的即時與歷史數據
使用 akshare 作為補充數據源
"""

import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import time

# ============================================================
# 快取機制：避免頻繁 API 請求
# ============================================================
_cache = {}
_cache_ttl = 300  # 快取有效時間（秒）


def _get_cached(key):
    """取得快取數據"""
    if key in _cache:
        data, timestamp = _cache[key]
        if time.time() - timestamp < _cache_ttl:
            return data
    return None


def _set_cached(key, data):
    """設定快取數據"""
    _cache[key] = (data, time.time())


def clear_cache(key=None):
    """清除快取數據。若指定 key 則僅清除該筆，否則清除全部。"""
    if key:
        _cache.pop(key, None)
    else:
        _cache.clear()


# ============================================================
# 即時指數數據
# ============================================================

def get_market_indices():
    """
    取得主要市場指數的即時數據
    Returns:
        list[dict]: 指數數據列表
    """
    cache_key = "market_indices"
    cached = _get_cached(cache_key)
    if cached:
        return cached

    indices = [
        {"symbol": "^DJI", "name": "道瓊工業指數"},
        {"symbol": "^IXIC", "name": "那斯達克"},
        {"symbol": "^GSPC", "name": "標普500指數"},
        {"symbol": "USDTWD=X", "name": "美元對台幣"},
    ]

    def _fetch_index(symbol, name):
        try:
            ticker = yf.Ticker(symbol)
            current_price, prev_close = None, None

            # fast_info 優先
            try:
                fi = ticker.fast_info
                p = fi.last_price
                if p is not None and not pd.isna(p) and p > 0:
                    current_price = float(p)
                pc = fi.previous_close
                if pc is not None and not pd.isna(pc) and pc > 0:
                    prev_close = float(pc)
            except Exception:
                pass

            # history 作 sparkline 來源，同時作 fallback
            hist = ticker.history(period="5d")
            valid_hist = hist.dropna(subset=["Close"]) if not hist.empty else hist
            sparkline = [round(float(v), 2) for v in valid_hist["Close"].tolist()] if not valid_hist.empty else []

            if current_price is None and not valid_hist.empty:
                current_price = float(valid_hist["Close"].iloc[-1])
            if prev_close is None and not valid_hist.empty:
                prev_close = float(valid_hist["Close"].iloc[-2]) if len(valid_hist) >= 2 else current_price

            if current_price and not pd.isna(current_price):
                prev_close = prev_close or current_price
                change = current_price - prev_close
                change_pct = (change / prev_close * 100) if prev_close else 0
                return {"symbol": symbol, "name": name,
                        "price": round(current_price, 2), "change": round(change, 2),
                        "change_pct": round(change_pct, 2), "sparkline": sparkline}
        except Exception:
            pass
        return {"symbol": symbol, "name": name, "price": 0, "change": 0, "change_pct": 0}

    import concurrent.futures

    results_dict = {}
    indices.append({"symbol": "^TWII", "name": "台灣加權指數"})

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
        future_to_idx = {
            executor.submit(_fetch_index, idx["symbol"], idx["name"]): idx
            for idx in indices
        }
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results_dict[idx["symbol"]] = future.result()
            except Exception:
                results_dict[idx["symbol"]] = {"symbol": idx["symbol"], "name": idx["name"], "price": 0, "change": 0, "change_pct": 0}

    results = [results_dict[idx["symbol"]] for idx in indices]

    _set_cached(cache_key, results)
    return results


# ============================================================
# 個股數據
# ============================================================

def get_stock_price(symbol):
    """
    取得單一股票的即時價格
    若台股 .TW 查無資料，自動嘗試 .TWO（上櫃）
    Args:
        symbol (str): 股票代碼 (e.g., "2330.TW", "AAPL")
    Returns:
        dict: 股票價格資訊
    """
    cache_key = f"stock_price_{symbol}"
    cached = _get_cached(cache_key)
    if cached:
        return cached

    result = _fetch_price(symbol)
    
    # 若 .TW 查無資料 (0 或 NaN)，嘗試 .TWO
    price = result.get("price", 0)
    if (pd.isna(price) or price == 0) and symbol.endswith(".TW"):
        alt_symbol = symbol.replace(".TW", ".TWO")
        alt_result = _fetch_price(alt_symbol)
        alt_price = alt_result.get("price", 0)
        if not pd.isna(alt_price) and alt_price > 0:
            alt_result["symbol"] = symbol  # 保持原始代碼
            _set_cached(cache_key, alt_result)
            return alt_result

    # 不論有無抓到有效價格，都寫入快取，避免切換頁面時重複等待 Timeout
    _set_cached(cache_key, result)
    return result


def _fetch_price(symbol):
    """
    實際從 yfinance 獲取價格的內部函式，並利用 twstock 補充台股即時報價
    """
    try:
        current_price = None
        prev_close = 0
        high = None
        low = None
        volume = 0

        ticker = yf.Ticker(symbol)

        # 1. 優先用 fast_info 取得即時價格（比 history 的最後一筆更可靠）
        try:
            fi = ticker.fast_info
            p = fi.last_price
            if p is not None and not pd.isna(p) and p > 0:
                current_price = float(p)
            pc = fi.previous_close
            if pc is not None and not pd.isna(pc) and pc > 0:
                prev_close = float(pc)
            dh = getattr(fi, "day_high", None)
            if dh is not None and not pd.isna(dh):
                high = float(dh)
            dl = getattr(fi, "day_low", None)
            if dl is not None and not pd.isna(dl):
                low = float(dl)
        except Exception:
            pass

        # 2. fast_info 取不到時，fallback 到 history（取最後有效的 Close）
        if not current_price or pd.isna(current_price):
            hist = ticker.history(period="5d")
            if not hist.empty:
                valid_hist = hist.dropna(subset=["Close"])
                if not valid_hist.empty:
                    current_price = float(valid_hist["Close"].iloc[-1])
                    if prev_close == 0 and len(valid_hist) >= 2:
                        prev_close = float(valid_hist["Close"].iloc[-2])
                    if high is None and not pd.isna(valid_hist["High"].iloc[-1]):
                        high = float(valid_hist["High"].iloc[-1])
                    if low is None and not pd.isna(valid_hist["Low"].iloc[-1]):
                        low = float(valid_hist["Low"].iloc[-1])
                    vol_val = valid_hist["Volume"].iloc[-1]
                    volume = int(vol_val) if not pd.isna(vol_val) else 0

        # 3. 如果是台股，嘗試使用 twstock 覆蓋即時報價
        if symbol.endswith(".TW") or symbol.endswith(".TWO"):
            try:
                import twstock

                tw_code = symbol.replace(".TW", "").replace(".TWO", "")
                rt = twstock.realtime.get(tw_code)

                if isinstance(rt, dict) and rt.get("success"):
                    rd = rt.get("realtime", {})
                    ltp = rd.get("latest_trade_price")
                    if ltp and ltp != "-":
                        current_price = float(ltp)
                    h_str = rd.get("high")
                    if h_str and h_str != "-":
                        high = float(h_str)
                    l_str = rd.get("low")
                    if l_str and l_str != "-":
                        low = float(l_str)
                    acc_vol = rd.get("accumulate_trade_volume")
                    if acc_vol and acc_vol != "-":
                        volume = int(acc_vol)
            except Exception as e:
                err_msg = f"twstock error for {symbol}: {e}\n"
                print(f"[warn] {err_msg}")
                try:
                    with open("d:/Antigravity/twstock_error.log", "a", encoding="utf-8") as f:
                        f.write(err_msg)
                except Exception:
                    pass

        # 4. 組合回傳資料
        if current_price is not None and not pd.isna(current_price) and current_price > 0:
            if prev_close == 0:
                prev_close = current_price
            change = current_price - prev_close
            change_pct = (change / prev_close * 100) if prev_close != 0 else 0
            return {
                "symbol": symbol,
                "price": round(current_price, 2),
                "prev_close": round(prev_close, 2),
                "change": round(change, 2),
                "change_pct": round(change_pct, 2),
                "high": round(high, 2) if high else round(current_price, 2),
                "low": round(low, 2) if low else round(current_price, 2),
                "volume": volume,
            }
        else:
            return {"symbol": symbol, "price": 0, "prev_close": 0, "change": 0, "change_pct": 0}
    except Exception as e:
        return {"symbol": symbol, "price": 0, "prev_close": 0, "change": 0, "change_pct": 0, "error": str(e)}


def get_stock_history(symbol, start_date=None, end_date=None, period="1y"):
    """
    取得個股歷史數據
    Args:
        symbol (str): 股票代碼
        start_date (str): 開始日期 "YYYY-MM-DD"
        end_date (str): 結束日期 "YYYY-MM-DD"
        period (str): 若不指定日期，使用 period (e.g., "1mo", "3mo", "6mo", "1y", "5y")
    Returns:
        pd.DataFrame: 歷史數據
    """
    cache_key = f"stock_hist_{symbol}_{start_date}_{end_date}_{period}"
    cached = _get_cached(cache_key)
    if cached is not None:
        return cached

    def _fetch_hist(sym):
        ticker = yf.Ticker(sym)
        if start_date and end_date:
            return ticker.history(start=start_date, end=end_date)
        else:
            return ticker.history(period=period)

    try:
        hist = _fetch_hist(symbol)

        if hist.empty and symbol.endswith(".TW"):
            alt_symbol = symbol.replace(".TW", ".TWO")
            alt_hist = _fetch_hist(alt_symbol)
            if not alt_hist.empty:
                hist = alt_hist

        if hist.empty:
            return pd.DataFrame()

        _set_cached(cache_key, hist)
        return hist
    except Exception as e:
        print(f"⚠️ 獲取 {symbol} 歷史數據失敗: {e}")
        return pd.DataFrame()


# ============================================================
# 匯率
# ============================================================

def get_usd_twd_rate():
    """
    取得美元兌台幣匯率
    Returns:
        float: 匯率 (1 USD = ? TWD)
    """
    cache_key = "usd_twd_rate"
    cached = _get_cached(cache_key)
    if cached:
        return cached

    try:
        ticker = yf.Ticker("USDTWD=X")
        hist = ticker.history(period="1d")
        if len(hist) >= 1:
            rate = float(hist["Close"].iloc[-1])
            _set_cached(cache_key, rate)
            return rate
    except Exception:
        pass

    # 備用匯率也寫入快取，避免重複等待 timeout
    _set_cached(cache_key, 32.0)
    return 32.0


# ============================================================
# 資產計算
# ============================================================

def calculate_portfolio_value(config):
    """
    計算投資組合的總資產淨值
    Args:
        config (dict): 配置字典
    Returns:
        dict: 資產明細
    """
    display_currency = config.get("display_currency", "TWD")

    stock_details = []
    total_stock_value_twd = 0.0
    total_stock_value_usd = 0.0

    holdings = config.get("stock_holdings", [])
    
    # 預先並發獲取所有股票價格以加速 (會自動寫入內部 _cache)
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=12) as executor:
        # 並發獲取匯率
        future_rate = executor.submit(get_usd_twd_rate)
        # 並發獲取股票價格
        if holdings:
            symbols_to_fetch = list(set([h["symbol"] for h in holdings]))
            list(executor.map(get_stock_price, symbols_to_fetch))
            
    usd_twd_rate = future_rate.result()

    for holding in holdings:
        symbol = holding["symbol"]
        shares = holding["shares"]
        avg_cost = holding["avg_cost"]
        currency = holding.get("currency", "TWD")

        price_info = get_stock_price(symbol)
        current_price = price_info.get("price", 0)

        market_value = current_price * shares
        cost_basis = avg_cost * shares
        profit_loss = market_value - cost_basis
        profit_loss_pct = (profit_loss / cost_basis * 100) if cost_basis > 0 else 0

        if currency == "TWD":
            market_value_twd = market_value
            market_value_usd = market_value / usd_twd_rate if usd_twd_rate > 0 else 0
        else:
            market_value_usd = market_value
            market_value_twd = market_value * usd_twd_rate

        total_stock_value_twd += market_value_twd
        total_stock_value_usd += market_value_usd

        stock_details.append({
            "symbol": symbol,
            "name": holding.get("name", symbol),
            "market": holding.get("market", ""),
            "shares": shares,
            "avg_cost": avg_cost,
            "current_price": current_price,
            "market_value": market_value,
            "market_value_twd": market_value_twd,
            "market_value_usd": market_value_usd,
            "profit_loss": profit_loss,
            "profit_loss_pct": round(profit_loss_pct, 2),
            "currency": currency,
            "change": price_info.get("change", 0),
            "change_pct": price_info.get("change_pct", 0),
            "category": holding.get("category", ""),
            "source": holding.get("source", ""),
        })

    # 現金
    total_cash_twd = 0.0
    total_cash_usd = 0.0
    for cash in config.get("cash_holdings", []):
        if cash["currency"] == "TWD":
            total_cash_twd += cash["amount"]
            total_cash_usd += cash["amount"] / usd_twd_rate if usd_twd_rate > 0 else 0
        elif cash["currency"] == "USD":
            total_cash_usd += cash["amount"]
            total_cash_twd += cash["amount"] * usd_twd_rate

    total_value_twd = total_stock_value_twd + total_cash_twd
    total_value_usd = total_stock_value_usd + total_cash_usd

    return {
        "display_currency": display_currency,
        "usd_twd_rate": usd_twd_rate,
        "total_value_twd": round(total_value_twd, 2),
        "total_value_usd": round(total_value_usd, 2),
        "total_stock_value_twd": round(total_stock_value_twd, 2),
        "total_stock_value_usd": round(total_stock_value_usd, 2),
        "total_cash_twd": round(total_cash_twd, 2),
        "total_cash_usd": round(total_cash_usd, 2),
        "stock_details": stock_details,
        "cash_holdings": config.get("cash_holdings", []),
    }


def calculate_transaction_returns(config):
    """
    計算交易紀錄的報酬結果
    Args:
        config (dict): 配置字典
    Returns:
        list[dict]: 各筆交易的報酬結果
    """
    usd_twd_rate = get_usd_twd_rate()
    results = []

    # 按股票代碼分組計算
    symbol_transactions = {}
    for txn in config.get("transactions", []):
        symbol = txn["symbol"]
        if symbol not in symbol_transactions:
            symbol_transactions[symbol] = []
        symbol_transactions[symbol].append(txn)

    for symbol, txns in symbol_transactions.items():
        buy_total_cost = 0.0
        buy_total_shares = 0
        sell_total_revenue = 0.0
        sell_total_shares = 0
        total_fees = 0.0
        total_tax = 0.0

        for txn in txns:
            if txn["action"] == "BUY":
                buy_total_cost += txn["price"] * txn["shares"]
                buy_total_shares += txn["shares"]
            elif txn["action"] == "SELL":
                sell_total_revenue += txn["price"] * txn["shares"]
                sell_total_shares += txn["shares"]
            total_fees += txn.get("fee", 0)
            total_tax += txn.get("tax", 0)

        # 目前持有的股票市值
        remaining_shares = buy_total_shares - sell_total_shares
        current_price_info = get_stock_price(symbol)
        current_price = current_price_info.get("price", 0)
        unrealized_value = remaining_shares * current_price

        # 已實現損益
        if sell_total_shares > 0 and buy_total_shares > 0:
            avg_buy_price = buy_total_cost / buy_total_shares
            realized_pl = sell_total_revenue - (avg_buy_price * sell_total_shares) - total_fees - total_tax
        else:
            realized_pl = 0

        # 未實現損益
        if remaining_shares > 0 and buy_total_shares > 0:
            avg_buy_price = buy_total_cost / buy_total_shares
            unrealized_pl = unrealized_value - (avg_buy_price * remaining_shares)
        else:
            unrealized_pl = 0

        total_pl = realized_pl + unrealized_pl
        total_invested = buy_total_cost + total_fees + total_tax
        return_pct = (total_pl / total_invested * 100) if total_invested > 0 else 0

        currency = txns[0].get("currency", "TWD")
        results.append({
            "symbol": symbol,
            "name": txns[0].get("name", symbol),
            "market": txns[0].get("market", ""),
            "currency": currency,
            "buy_total_shares": buy_total_shares,
            "sell_total_shares": sell_total_shares,
            "remaining_shares": remaining_shares,
            "buy_total_cost": round(buy_total_cost, 2),
            "sell_total_revenue": round(sell_total_revenue, 2),
            "total_fees": round(total_fees, 2),
            "total_tax": round(total_tax, 2),
            "current_price": current_price,
            "unrealized_value": round(unrealized_value, 2),
            "realized_pl": round(realized_pl, 2),
            "unrealized_pl": round(unrealized_pl, 2),
            "total_pl": round(total_pl, 2),
            "return_pct": round(return_pct, 2),
        })

    return results
