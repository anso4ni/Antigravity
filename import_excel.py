# -*- coding: utf-8 -*-
"""
匯入 stock record.xlsx 中的資料到 portfolio_config.json
- "個股 0414" 分頁 → 持倉狀態 + 現金
- "交易紀錄" 分頁 → 交易歷史
"""
import pandas as pd
import json
import os
from datetime import datetime


def import_all():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, "stock record.xlsx")
    config_path = os.path.join(base_dir, "portfolio_config.json")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # ========================================
    # 1. 從 "個股 0414" 匯入持倉 (不合併 FT 與複委託)
    # ========================================
    holdings = _parse_holdings_sheet(excel_path)
    config["stock_holdings"] = holdings
    print(f"[持倉] 匯入 {len(holdings)} 檔")
    for h in holdings:
        mkt = "TW" if h["market"] == "TW" else "US"
        print(f"  {mkt:2s} | {h['source']:4s} | {h['symbol']:12s} {h['name']:10s} x {h['shares']:>10} @ {h['avg_cost']:>10.2f} {h['currency']}  [{h.get('category','')}]")

    # ========================================
    # 2. 從 "個股 0414" 匯入現金
    # ========================================
    cash_holdings = _parse_cash(excel_path)
    config["cash_holdings"] = cash_holdings
    print(f"\n[現金] 匯入 {len(cash_holdings)} 筆")
    for c in cash_holdings:
        sym = "NT$" if c["currency"] == "TWD" else "US$"
        print(f"  {c['source']:6s} | {sym} {c['amount']:>12,.2f}  ({c['note']})")

    # ========================================
    # 3. 從 "交易紀錄" 匯入交易歷史
    # ========================================
    transactions = _parse_transactions_sheet(excel_path)
    config["transactions"] = transactions
    print(f"\n[交易] 匯入 {len(transactions)} 筆")

    # ========================================
    # 4. 儲存
    # ========================================
    config["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"\n[OK] 已儲存到 {config_path}")


def _parse_holdings_sheet(excel_path):
    """解析 '個股持倉' — 持倉資料，FT 與複委託分開"""
    df = pd.read_excel(excel_path, sheet_name="個股持倉", header=None)
    holdings = []

    # --- 台股 (rows 3~18) ---
    for idx in range(3, 18):
        row = df.iloc[idx]
        symbol_raw = str(row[1]).strip() if pd.notna(row[1]) else ""
        if not symbol_raw:
            continue
        shares = int(_safe_float(row[4]))
        if shares == 0:  # 跳過標題列或持股為 0 的行
            continue
        category = str(row[0]).strip() if pd.notna(row[0]) else ""
        symbol_yf = symbol_raw + ".TW" if not symbol_raw.endswith((".TW", ".TWO")) else symbol_raw
        holdings.append({
            "symbol": symbol_yf,
            "name": category,
            "market": "TW",
            "shares": shares,
            "avg_cost": round(_safe_float(row[3]), 4),
            "currency": "TWD",
            "category": category,
            "source": "台股",
            "note": "",
        })

    # --- 美股 FT (rows 21~36) ---
    for idx in range(21, 37):
        row = df.iloc[idx]
        symbol_raw = str(row[1]).strip() if pd.notna(row[1]) else ""
        if not symbol_raw or symbol_raw == "NaN":
            continue
        shares = round(_safe_float(row[4]), 5)
        if shares == 0:
            continue
        category = str(row[0]).strip() if pd.notna(row[0]) else ""
        holdings.append({
            "symbol": symbol_raw,
            "name": symbol_raw,
            "market": "US",
            "shares": shares,
            "avg_cost": round(_safe_float(row[3]), 4),
            "currency": "USD",
            "category": category,
            "source": "FT",
            "note": "",
        })

    # --- 複委託 (rows 41~50) ---
    for idx in range(41, 51):
        if idx >= len(df):
            break
        row = df.iloc[idx]
        symbol_raw = str(row[1]).strip() if pd.notna(row[1]) else ""
        if not symbol_raw or symbol_raw == "NaN":
            continue
        shares = round(_safe_float(row[4]), 5)
        if shares == 0:
            continue
        category = str(row[0]).strip() if pd.notna(row[0]) else ""
        holdings.append({
            "symbol": symbol_raw,
            "name": symbol_raw,
            "market": "US",
            "shares": shares,
            "avg_cost": round(_safe_float(row[3]), 4),
            "currency": "USD",
            "category": category,
            "source": "複委託",
            "note": "",
        })

    return holdings


def _parse_cash(excel_path):
    """解析 '個股持倉' — 銀行現金 + 美股現金"""
    df = pd.read_excel(excel_path, sheet_name="個股持倉", header=None)
    cash_list = []

    # 台幣銀行帳戶 (rows 7~10)
    bank_rows = {7: "台新", 8: "元大", 9: "中信", 10: "LineBank"}
    for row_idx, bank_name in bank_rows.items():
        if row_idx < len(df):
            val = _safe_float(df.iloc[row_idx, 15])
            if val > 0:
                cash_list.append({
                    "currency": "TWD",
                    "amount": val,
                    "note": f"{bank_name}銀行",
                    "source": "台股",
                })

    # 美股現金 FT (row 17, col 14)
    ft_cash = _safe_float(df.iloc[17, 14])
    cash_list.append({
        "currency": "USD",
        "amount": ft_cash,
        "note": "Firstrade 現金",
        "source": "FT",
    })

    # 複委託現金 = 合計美股現金 (row 18, col 14) - FT 現金
    total_us_cash = _safe_float(df.iloc[18, 14])
    broker_cash = total_us_cash - ft_cash
    cash_list.append({
        "currency": "USD",
        "amount": round(broker_cash, 2),
        "note": "複委託 現金",
        "source": "複委託",
    })

    return cash_list


def _parse_transactions_sheet(excel_path):
    """解析 '交易紀錄' 分頁"""
    df = pd.read_excel(excel_path, sheet_name="交易紀錄")
    transactions = []
    for _, row in df.iterrows():
        if pd.isna(row.get("買賣日期")):
            continue
        market_raw = str(row["台/美股"]).strip()
        if market_raw in ["複委託", "FT"]:
            market, currency = "US", "USD"
        else:
            market, currency = "TW", "TWD"
        action = "BUY" if row["買賣"] == "買" else "SELL"
        symbol = str(row["代號"]).strip()
        if not symbol or symbol == "nan":
            continue
        if market == "TW" and not symbol.endswith((".TW", ".TWO")):
            symbol_yf = symbol + ".TW"
        else:
            symbol_yf = symbol
        src = "FT" if market_raw == "FT" else ("複委託" if market_raw == "複委託" else "台股")
        transactions.append({
            "date": pd.Timestamp(row["買賣日期"]).strftime("%Y-%m-%d"),
            "symbol": symbol_yf,
            "name": str(row["名 稱"]).strip(),
            "market": market,
            "action": action,
            "shares": int(abs(float(row["股數"]))),
            "price": float(row["單價"]),
            "fee": float(row["手續費"]) if pd.notna(row["手續費"]) else 0.0,
            "tax": float(row["證交稅"]) if pd.notna(row["證交稅"]) else 0.0,
            "currency": currency,
            "source": src,
            "note": f"from excel ({market_raw})",
        })
    return transactions


def _safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def import_all_from_bytes(file_bytes, cfg):
    """
    從 bytes（Google Drive 下載內容）匯入 Excel，邏輯與 import_all() 相同。
    Args:
        file_bytes (bytes): Excel 檔案內容
        cfg (dict): 現有 portfolio config，會直接覆寫持倉/現金/交易
    Returns:
        dict: {"holdings": n, "cash": n, "transactions": n}
    """
    import io
    import config as _config

    buf = io.BytesIO(file_bytes)

    holdings = _parse_holdings_sheet(buf)
    buf.seek(0)
    cash_holdings = _parse_cash(buf)
    buf.seek(0)
    transactions = _parse_transactions_sheet(buf)

    cfg["stock_holdings"] = holdings
    cfg["cash_holdings"] = cash_holdings
    cfg["transactions"] = transactions
    cfg["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    _config.save_config(cfg)

    return {
        "holdings": len(holdings),
        "cash": len(cash_holdings),
        "transactions": len(transactions),
    }


def import_holdings_from_bytes(file_bytes, cfg):
    """
    從 bytes 匯入持倉 + 現金，不動交易紀錄。
    Returns: {"holdings": n, "cash": n}
    """
    import io
    import config as _config

    buf = io.BytesIO(file_bytes)
    holdings = _parse_holdings_sheet(buf)
    buf.seek(0)
    cash_holdings = _parse_cash(buf)

    cfg["stock_holdings"] = holdings
    cfg["cash_holdings"] = cash_holdings
    cfg["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _config.save_config(cfg)

    return {"holdings": len(holdings), "cash": len(cash_holdings)}


def import_transactions_from_bytes(file_bytes, cfg):
    """
    從 bytes 匯入交易紀錄，不動持倉/現金。
    Returns: {"transactions": n, "records": [list of transaction dicts]}
    """
    import io
    import config as _config

    buf = io.BytesIO(file_bytes)
    transactions = _parse_transactions_sheet(buf)

    cfg["transactions"] = transactions
    cfg["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _config.save_config(cfg)

    return {"transactions": len(transactions), "records": transactions}


if __name__ == "__main__":
    import_all()
