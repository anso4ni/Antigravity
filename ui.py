# -*- coding: utf-8 -*-
"""
UI模塊 ui.py
使用 Streamlit 建構使用者介面
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import config
import data
import google_drive


def setup_page():
    """頁面基本設定"""
    st.set_page_config(
        page_title="📈 股票交易紀錄系統",
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(get_custom_css(), unsafe_allow_html=True)


def get_custom_css():
    """自訂 CSS 樣式"""
    return """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * { font-family: 'Inter', sans-serif; }
    
    .main .block-container {
        padding-top: 1.5rem;
        max-width: 1400px;
    }
    
    .app-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    
    .index-ticker-bar {
        display: flex;
        align-items: center;
        background: transparent;
        border: none;
        border-radius: 0;
        padding: 6px 0;
        margin-bottom: 4px;
        border-bottom: 1px solid rgba(102, 126, 234, 0.1);
    }
    .index-item {
        flex: 1;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 4px 18px;
        border-right: 1px solid rgba(102, 126, 234, 0.1);
        gap: 10px;
    }
    .index-item:last-child { border-right: none; }
    .index-info { flex: 1; }
    .index-name { color: #718096; font-size: 0.72rem; font-weight: 500; margin-bottom: 1px; letter-spacing: 0.02em; }
    .index-price-up { color: #00e676; font-size: 1.2rem; font-weight: 700; line-height: 1.2; }
    .index-price-down { color: #ff1744; font-size: 1.2rem; font-weight: 700; line-height: 1.2; }
    .index-change-up { color: #00e676; font-size: 0.72rem; font-weight: 600; margin-top: 1px; }
    .index-change-down { color: #ff1744; font-size: 0.72rem; font-weight: 600; margin-top: 1px; }
    .index-spark { flex-shrink: 0; }
    
    .net-value-area {
        padding: 8px 0 12px 0;
        margin-bottom: 6px;
    }
    .net-value-label { color: #718096; font-size: 0.82rem; font-weight: 500; }
    .net-value-amount { font-size: 2.4rem; font-weight: 800; letter-spacing: -1px; line-height: 1.15; }
    .net-value-sub { color: #888; font-size: 0.82rem; font-weight: 500; margin-left: 14px; }
    .net-value-row { display: flex; align-items: baseline; gap: 0; flex-wrap: wrap; }
    .cat-card {
        background: rgba(245, 245, 250, 0.08);
        border: 1px solid rgba(200, 200, 220, 0.15);
        border-radius: 14px;
        padding: 14px 16px;
        border-left: 3px solid rgba(102, 126, 234, 0.4);
    }
    .cat-card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
    }
    .cat-card-name { color: #888; font-size: 0.78rem; font-weight: 600; }
    .cat-card-pct { font-size: 0.74rem; font-weight: 700; background: rgba(102,126,234,0.12); color: #667eea; padding: 2px 8px; border-radius: 8px; }
    .cat-card-value { font-size: 1.15rem; font-weight: 700; margin-bottom: 4px; }
    .cat-card-detail { color: #888; font-size: 0.72rem; line-height: 1.6; }
    
    .asset-card {
        background: linear-gradient(145deg, #1e1e3a, #252547);
        border-radius: 16px;
        padding: 20px 22px;
        border: 1px solid rgba(255,255,255,0.06);
        box-shadow: 0 4px 24px rgba(0,0,0,0.2);
        margin-bottom: 8px;
    }
    .asset-symbol { color: #e2e8f0; font-size: 1.05rem; font-weight: 700; }
    .asset-name { color: #a0aec0; font-size: 0.8rem; }
    .asset-value { color: #e2e8f0; font-size: 1.1rem; font-weight: 600; }
    .asset-pl-up { color: #48bb78; font-weight: 600; }
    .asset-pl-down { color: #fc8181; font-weight: 600; }
    
    .home-detail-bar {
        display: flex; align-items: baseline; gap: 12px;
        padding: 8px 16px; margin: 20px 0 10px 0;
        border-left: 4px solid #667eea;
        background: rgba(102, 126, 234, 0.05);
        border-radius: 0 8px 8px 0;
    }
    .home-detail-pct { font-size: 0.8rem; font-weight: 700; background: rgba(102,126,234,0.2); color: #667eea; padding: 2px 8px; border-radius: 8px; }
    .home-detail-val { font-size: 1.4rem; font-weight: 800; color: #e2e8f0; }
    .home-detail-inv { font-size: 0.8rem; font-weight: 500; color: #718096; }
    .home-detail-pl { margin-left: auto; font-size: 0.9rem; font-weight: 700; }
    
    .section-title {
        color: #e2e8f0;
        font-size: 1.2rem;
        font-weight: 700;
        margin: 1.5rem 0 1rem 0;
        padding-left: 12px;
        border-left: 4px solid #667eea;
    }
    
    .stMetric > div { background: rgba(30,30,58,0.6); border-radius: 12px; padding: 12px; }
    
    div[data-testid="stForm"] {
        background: rgba(30,30,58,0.4);
        border: 1px solid rgba(102,126,234,0.15);
        border-radius: 16px;
        padding: 24px;
    }

    .google-btn {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(30, 30, 58, 0.7);
        border: 1px solid rgba(102, 126, 234, 0.25);
        border-radius: 12px;
        padding: 8px 18px;
        cursor: pointer;
        transition: all 0.25s;
        text-decoration: none;
    }
    .google-btn:hover {
        background: rgba(66, 133, 244, 0.15);
        border-color: #4285f4;
        transform: translateY(-1px);
        box-shadow: 0 4px 16px rgba(66, 133, 244, 0.2);
    }
    .google-icon {
        width: 20px; height: 20px;
    }
    .google-status-dot {
        width: 8px; height: 8px;
        border-radius: 50%;
        display: inline-block;
        margin-right: 4px;
    }
    .google-status-connected { background: #48bb78; box-shadow: 0 0 6px #48bb78; }
    .google-status-disconnected { background: #fc8181; box-shadow: 0 0 6px #fc8181; }
    </style>
    """


def _is_cloud_mode():
    """判斷是否在 Streamlit Cloud 雲端模式（secrets 中有 google_oauth）"""
    try:
        return "google_oauth" in st.secrets
    except Exception:
        return False


def _get_cloud_oauth_secrets():
    """取得雲端 OAuth 設定，回傳 (client_secrets_dict, redirect_uri)"""
    import json as _json
    sec = st.secrets["google_oauth"]
    raw = sec.get("client_secrets_json", "")
    client_secrets_dict = _json.loads(raw) if isinstance(raw, str) else dict(raw)
    redirect_uri = sec.get("redirect_uri", "")
    return client_secrets_dict, redirect_uri


def _fetch_google_email(creds):
    """從 credentials 取得使用者 email"""
    try:
        from google.auth.transport.requests import AuthorizedSession
        session = AuthorizedSession(creds)
        resp = session.get("https://www.googleapis.com/oauth2/v1/userinfo")
        if resp.status_code == 200:
            return resp.json().get("email", "")
    except Exception:
        pass
    return ""


def render_google_status():
    """渲染 Google 雲端連線狀態（從 session_state 讀取 OAuth 登入狀態）"""
    email = st.session_state.get("google_email", "")
    is_connected = bool(email)

    status_dot = "google-status-connected" if is_connected else "google-status-disconnected"
    short_email = (email[:16] + "…") if len(email) > 17 else email
    status_text = short_email if is_connected else "未登入"
    google_svg = '''<svg class="google-icon" viewBox="0 0 48 48"><path fill="#4285F4" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#34A853" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#EA4335" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg>'''

    return google_svg, status_dot, status_text, is_connected


def _build_sparkline_svg(data_points, color, width=60, height=28):
    """生成迷你走勢圖 SVG（含淺色面積）"""
    if not data_points or len(data_points) < 2:
        return ""
    
    min_val = min(data_points)
    max_val = max(data_points)
    val_range = max_val - min_val if max_val != min_val else 1
    
    padding = 2
    draw_w = width - padding * 2
    draw_h = height - padding * 2
    
    points = []
    for i, val in enumerate(data_points):
        x = padding + (i / (len(data_points) - 1)) * draw_w
        y = padding + draw_h - ((val - min_val) / val_range) * draw_h
        points.append(f"{x:.1f},{y:.1f}")
    
    polyline = " ".join(points)
    # 面積填充：從第一個點底部開始，沿曲線走，再回到最後一個點底部
    first_x = padding
    last_x = padding + draw_w
    bottom_y = height
    area_points = f"{first_x:.1f},{bottom_y} {polyline} {last_x:.1f},{bottom_y}"
    
    # 使用唯一 ID 避免衝突
    grad_id = f"sg{hash(tuple(data_points)) % 99999}"
    
    return f'''<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg">
        <defs><linearGradient id="{grad_id}" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stop-color="{color}" stop-opacity="0.3"/>
            <stop offset="100%" stop-color="{color}" stop-opacity="0.03"/>
        </linearGradient></defs>
        <polygon points="{area_points}" fill="url(#{grad_id})"/>
        <polyline points="{polyline}" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>'''

def render_market_indices(cfg):
    """渲染頂部即時指數區塊（含 Google 登入圖示）"""

    # ── 雲端模式：處理 Google OAuth callback（?code=xxx 回傳） ──────────
    if _is_cloud_mode() and "google_client" not in st.session_state:
        params = st.query_params
        if "code" in params:
            try:
                with st.spinner("正在完成 Google 授權…"):
                    client_secrets_dict, redirect_uri = _get_cloud_oauth_secrets()
                    creds = google_drive.exchange_web_auth_code(
                        client_secrets_dict,
                        code=params["code"],
                        state=params.get("state", ""),
                        redirect_uri=redirect_uri,
                    )
                    client = google_drive.make_gspread_client(creds)
                    email = _fetch_google_email(creds)
                st.session_state.google_client = client
                st.session_state.google_email = email
                st.session_state.google_creds = google_drive.credentials_to_dict(creds)
                st.session_state.pop("pending_auth_url", None)
                st.session_state.pop("oauth_state", None)
                st.query_params.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Google 授權失敗：{e}")
                st.query_params.clear()

    # ── 本地模式：有 token 檔案則自動靜默重新登入 ────────────────────────
    if not _is_cloud_mode() and "google_client" not in st.session_state and google_drive.has_valid_token():
        try:
            secrets_path = cfg.get("google_drive", {}).get("oauth_client_secrets_path", "")
            if secrets_path:
                client, email = google_drive.authenticate_oauth(secrets_path)
                st.session_state.google_client = client
                st.session_state.google_email = email
        except Exception:
            pass

    google_svg, status_dot, status_text, is_connected = render_google_status()

    title_col, google_col = st.columns([4, 2])
    with title_col:
        st.markdown('<p class="app-title">📈 股票交易紀錄系統</p>', unsafe_allow_html=True)
    with google_col:
        st.markdown(f'''
        <div class="google-btn" title="Google 帳戶狀態">
            {google_svg}
            <span style="color:#e2e8f0; font-size:0.82rem; font-weight:500;">
                <span class="google-status-dot {status_dot}"></span>
                {status_text}
            </span>
        </div>
        ''', unsafe_allow_html=True)
        if is_connected:
            if st.button("登出", key="google_logout_btn", use_container_width=True):
                if not _is_cloud_mode():
                    google_drive.logout_oauth()
                for k in list(st.session_state.keys()):
                    if k.startswith("gd_subfolders_"):
                        del st.session_state[k]
                for _k in ("google_client", "google_email", "google_creds", "gd_folder_stack"):
                    st.session_state.pop(_k, None)
                st.rerun()
        else:
            # 雲端模式：若已產生授權 URL，直接顯示連結按鈕
            if _is_cloud_mode() and st.session_state.get("pending_auth_url"):
                st.link_button(
                    "🔗 點此前往 Google 授權",
                    st.session_state["pending_auth_url"],
                    use_container_width=True,
                    type="primary",
                )
            elif st.button("登入 Google", key="google_login_btn", use_container_width=True, type="primary"):
                if _is_cloud_mode():
                    # 雲端：產生授權 URL，存入 session_state 後 rerun 顯示連結
                    try:
                        client_secrets_dict, redirect_uri = _get_cloud_oauth_secrets()
                        auth_url, state = google_drive.get_web_auth_url(
                            client_secrets_dict, redirect_uri
                        )
                        st.session_state["pending_auth_url"] = auth_url
                        st.session_state["oauth_state"] = state
                        st.rerun()
                    except Exception as e:
                        st.error(f"無法取得授權 URL：{e}")
                else:
                    # 本地：開啟瀏覽器視窗完成授權
                    secrets_path = cfg.get("google_drive", {}).get("oauth_client_secrets_path", "")
                    if not secrets_path:
                        st.session_state["show_google_setup"] = True
                    else:
                        try:
                            with st.spinner("請在瀏覽器完成 Google 授權…"):
                                client, email = google_drive.authenticate_oauth(secrets_path)
                            st.session_state.google_client = client
                            st.session_state.google_email = email
                            st.session_state.pop("show_google_setup", None)
                            st.rerun()
                        except Exception as e:
                            st.error(f"登入失敗：{e}")

    if not is_connected and not _is_cloud_mode() and st.session_state.get("show_google_setup"):
        with st.expander("🔑 設定 Google OAuth 金鑰", expanded=True):
            _render_google_quick_connect(cfg)

    with st.spinner("正在載入市場數據..."):
        indices = data.get_market_indices()

    # 組裝 ticker bar HTML
    items_html = ""
    for idx in indices:
        price = f"{idx['price']:,.2f}"
        change = idx["change"]
        change_pct = idx["change_pct"]
        sparkline = idx.get("sparkline", [])
        if change >= 0:
            arrow = "▲"
            change_class = "index-change-up"
            price_class = "index-price-up"
            sign = "+"
            spark_color = "#00e676"
        else:
            arrow = "▼"
            change_class = "index-change-down"
            price_class = "index-price-down"
            sign = ""
            spark_color = "#ff1744"

        # 生成 SVG sparkline
        spark_svg = _build_sparkline_svg(sparkline, spark_color)

        items_html += f'''
        <div class="index-item">
            <div class="index-info">
                <div class="index-name">{idx['name']}</div>
                <div class="{price_class}">{price}</div>
                <div class="{change_class}">{arrow} {sign}{change:,.2f} ({sign}{change_pct:.2f}%)</div>
            </div>
            <div class="index-spark">{spark_svg}</div>
        </div>'''

    st.markdown(f'<div class="index-ticker-bar">{items_html}</div>', unsafe_allow_html=True)


def _render_google_quick_connect(cfg):
    """OAuth 快速設定引導"""
    st.markdown('''
    > **設定步驟：**
    > 1. 前往 [Google Cloud Console](https://console.cloud.google.com/) 建立專案
    > 2. 啟用 **Google Sheets API** 和 **Google Drive API**
    > 3. 建立 **OAuth 2.0 用戶端 ID**（類型選「**桌面應用程式**」），下載 JSON
    > 4. 將 JSON 路徑填入下方，再按右上角「登入 Google」
    ''')

    gd_cfg = cfg.get("google_drive", {})
    with st.form("quick_connect_form"):
        secrets_path = st.text_input(
            "🔑 OAuth 用戶端金鑰路徑（JSON）",
            value=gd_cfg.get("oauth_client_secrets_path", ""),
            placeholder="C:/path/to/client_secrets.json",
            key="qc_secrets",
        )
        submitted = st.form_submit_button("💾 儲存路徑", use_container_width=True, type="primary")
        if submitted and secrets_path:
            cfg.setdefault("google_drive", {})["oauth_client_secrets_path"] = secrets_path
            config.save_config(cfg)
            st.success("✅ 已儲存！請按右上角「登入 Google」完成授權。")
            st.rerun()


def render_net_value(portfolio):
    """渲染總資產淨值區塊 + 分類占比小方塊"""
    display_currency = portfolio["display_currency"]

    col_val, col_switch, col_refresh = st.columns([5, 1, 1.2])
    with col_switch:
        _cur_options = ["Default", "TWD", "USD"]
        _cur_idx = _cur_options.index(display_currency) if display_currency in _cur_options else 0
        new_currency = st.selectbox(
            "幣值", _cur_options,
            index=_cur_idx,
            key="currency_select",
            label_visibility="collapsed",
        )
        if new_currency != display_currency:
            cfg = config.load_config()
            config.set_display_currency(cfg, new_currency)
            st.rerun()
            
    with col_refresh:
        st.markdown('''
            <style>
            div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child(3) button {
                border-radius: 20px !important;
                border: 1px solid #d1d5db !important;
                color: #888 !important;
                background-color: transparent !important;
                transition: all 0.2s;
            }
            div[data-testid="stHorizontalBlock"]:nth-of-type(2) div[data-testid="column"]:nth-child(3) button:hover {
                border-color: #9ca3af !important;
                color: #555 !important;
                background-color: #f3f4f6 !important;
            }
            </style>
        ''', unsafe_allow_html=True)
        if st.button("↻ Refresh", help="重新整理報價", use_container_width=True):
            data.clear_cache()
            st.session_state['last_update'] = datetime.now()
            st.rerun()
            
        last_update = st.session_state.get('last_update', datetime.now())
        diff_mins = int((datetime.now() - last_update).total_seconds() / 60)
        time_str = "剛剛更新" if diff_mins == 0 else f"{diff_mins} 分鐘前更新"
        st.markdown(f'<div style="text-align:center; color:#888; font-size:0.8rem; margin-top:-10px;">{time_str}</div>', unsafe_allow_html=True)

    with col_val:
        def safe_val(v): return 0 if pd.isna(v) else v

        if display_currency == "USD":
            total = safe_val(portfolio["total_value_usd"])
            sym = "US$"
            stock_val = safe_val(portfolio["total_stock_value_usd"])
            cash_val = safe_val(portfolio["total_cash_usd"])
        else:  # "TWD" or "Default" → aggregate in TWD
            total = safe_val(portfolio["total_value_twd"])
            sym = "NT$"
            stock_val = safe_val(portfolio["total_stock_value_twd"])
            cash_val = safe_val(portfolio["total_cash_twd"])

        rate = portfolio['usd_twd_rate']
        st.markdown(f'''
        <div class="net-value-area">
            <div class="net-value-label">總淨值</div>
            <div class="net-value-row">
                <div class="net-value-amount">{sym} {total:,.0f}</div>
                <div class="net-value-sub">投資 {sym} {stock_val:,.0f}</div>
                <div class="net-value-sub">現金 {sym} {cash_val:,.0f}</div>
                <div class="net-value-sub">匯率 1 USD = {rate:.2f} TWD</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    # --- 分類占比小方塊 ---
    stocks = portfolio.get("stock_details", [])
    cash_holdings = portfolio.get("cash_holdings", [])
    rate = portfolio['usd_twd_rate']

    # 依 source 分組計算市值
    cat_data = {}
    disp_mode = display_currency  # "Default", "TWD", "USD"

    def _eff_twd_stock(s):
        if disp_mode == "TWD": return True
        if disp_mode == "USD": return False
        return s["currency"] == "TWD"  # Default: native currency

    def _src_sym(src):
        if disp_mode == "TWD": return "NT$"
        if disp_mode == "USD": return "US$"
        return "NT$" if src == "台股" else "US$"  # Default: native

    for s in stocks:
        src = s.get("source", "其他")
        eff_twd = _eff_twd_stock(s)
        mv = s.get("market_value_twd", 0) if eff_twd else s.get("market_value_usd", 0)
        if pd.isna(mv): mv = 0

        if src not in cat_data:
            cat_data[src] = {"market_value": 0, "cost": 0, "count": 0, "day_change": 0, "market_value_twd": 0}

        cat_data[src]["market_value"] += mv
        cat_data[src]["market_value_twd"] += (s.get("market_value_twd", 0) or 0)

        cost_orig = s["avg_cost"] * s["shares"]
        if pd.isna(cost_orig): cost_orig = 0
        if eff_twd:
            cost = cost_orig if s["currency"] == "TWD" else cost_orig * rate
        else:
            cost = cost_orig if s["currency"] == "USD" else cost_orig / rate
        cat_data[src]["cost"] += cost

        chg = s.get("change", 0)
        if pd.isna(chg): chg = 0
        day_change_orig = chg * s["shares"]
        if eff_twd:
            day_change = day_change_orig if s["currency"] == "TWD" else day_change_orig * rate
        else:
            day_change = day_change_orig if s["currency"] == "USD" else day_change_orig / rate
        cat_data[src]["day_change"] += day_change

        cat_data[src]["count"] += 1

    # 加入現金
    cash_by_source = {}
    cash_by_source_twd = {}
    for c in cash_holdings:
        src = c.get("source", "其他")
        amt_twd = c["amount"] if c["currency"] == "TWD" else c["amount"] * rate
        if disp_mode == "TWD":
            amt = amt_twd
        elif disp_mode == "USD":
            amt = c["amount"] if c["currency"] == "USD" else c["amount"] / rate
        else:  # Default: native
            amt = c["amount"]
        cash_by_source[src] = cash_by_source.get(src, 0) + amt
        cash_by_source_twd[src] = cash_by_source_twd.get(src, 0) + amt_twd

    total_cleaned_mv = sum(cd["market_value"] for cd in cat_data.values())
    total_cleaned_cash = sum(cash_by_source.values())
    total_value_twd = safe_val(portfolio["total_value_twd"])
    if disp_mode == "Default":
        total_for_pct = total_value_twd if total_value_twd > 0 else 1
    else:
        total_for_pct = total_cleaned_mv + total_cleaned_cash
        if total_for_pct <= 0: total_for_pct = 1

    # 定義顯示順序和圖示（含 anchor id）
    cat_order = [
        ("台股", "🇹🇼", "cat-tw"),
        ("FT", "🇺🇸", "cat-ft"),
        ("複委託", "🏦", "cat-sub"),
    ]

    # 準備各分類數據
    cat_cards_data = []
    for src, icon, anchor in cat_order:
        cd = cat_data.get(src, {"market_value": 0, "cost": 0, "count": 0, "day_change": 0, "market_value_twd": 0})
        cash_amt = cash_by_source.get(src, 0)
        total_cat = cd["market_value"] + cash_amt

        # pct = stocks only (cash card accounts for cash separately → sum = 100%)
        if disp_mode == "Default":
            pct = (cd.get("market_value_twd", 0) / total_for_pct * 100)
        else:
            pct = (cd["market_value"] / total_for_pct * 100)

        pl = cd["market_value"] - cd["cost"] if cd["cost"] > 0 else 0
        pl_pct = (pl / cd["cost"] * 100) if cd["cost"] > 0 else 0

        pl_color = "#ff1744" if pl >= 0 else "#00e676"
        pl_sign = "+" if pl >= 0 else ""

        dc = cd.get("day_change", 0)
        prev_mv = cd["market_value"] - dc
        dc_pct = (dc / prev_mv * 100) if prev_mv > 0 else 0
        dc_color = "#ff1744" if dc >= 0 else "#00e676"
        dc_sign = "+" if dc >= 0 else ""

        cat_cards_data.append({
            "icon": icon, "src": src, "anchor": anchor, "total": total_cat, "pct": pct,
            "sym": _src_sym(src),
            "count": cd["count"], "mv": cd["market_value"],
            "pl": pl, "pl_pct": pl_pct, "pl_color": pl_color, "pl_sign": pl_sign,
            "dc": dc, "dc_pct": dc_pct, "dc_color": dc_color, "dc_sign": dc_sign,
        })

    # 現金卡片
    if disp_mode == "USD":
        total_cash_disp = safe_val(portfolio["total_cash_usd"])
        cash_sym = "US$"
    else:
        total_cash_disp = safe_val(portfolio["total_cash_twd"])
        cash_sym = "NT$"
    cash_pct = (safe_val(portfolio["total_cash_twd"]) / total_for_pct * 100) if disp_mode == "Default" else (total_cash_disp / total_for_pct * 100)

    # 使用 st.columns 渲染
    cols = st.columns(len(cat_cards_data) + 1)
    for i, cd in enumerate(cat_cards_data):
        with cols[i]:
            card_sym = cd["sym"]
            st.markdown(f'''
            <a href="#{cd['anchor']}" style="text-decoration: none; color: inherit; display: block; height: 100%;">
            <div class="cat-card" style="transition: transform 0.2s; cursor: pointer;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
                <div class="cat-card-header">
                    <span class="cat-card-name">{cd["icon"]} {cd["src"]}</span>
                    <span class="cat-card-pct">{cd["pct"]:.1f}%</span>
                </div>
                <div class="cat-card-value">{card_sym} {cd["total"]:,.0f}</div>
                <div class="cat-card-detail">
                    持倉 {cd["count"]} 檔 · {card_sym} {cd["mv"]:,.0f}<br>
                    <span style="color:{cd["pl_color"]}">損益 {cd["pl_sign"]}{cd["pl"]:,.0f} ({cd["pl_sign"]}{cd["pl_pct"]:.1f}%)</span><br>
                    <span style="color:{cd["dc_color"]}">今日 {cd["dc_sign"]}{cd["dc"]:,.0f} ({cd["dc_sign"]}{cd["dc_pct"]:.2f}%)</span>
                </div>
            </div>
            </a>
            ''', unsafe_allow_html=True)

    with cols[-1]:
        st.markdown(f'''
        <a href="#cat-cash" style="text-decoration: none; color: inherit; display: block; height: 100%;">
        <div class="cat-card" style="transition: transform 0.2s; cursor: pointer;" onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
            <div class="cat-card-header">
                <span class="cat-card-name">💵 現金</span>
                <span class="cat-card-pct">{cash_pct:.1f}%</span>
            </div>
            <div class="cat-card-value">{cash_sym} {total_cash_disp:,.0f}</div>
            <div class="cat-card-detail">
                {len(cash_holdings)} 個帳戶
            </div>
        </div>
        </a>
        ''', unsafe_allow_html=True)



def render_asset_cards(portfolio):
    """渲染資產方塊卡片（支援圖示化/條列式切換 + 台美股/複委託篩選）"""
    # 標題列 + 篩選 + 切換按鈕
    title_col, filter_col, toggle_col = st.columns([3, 1.5, 1])
    with title_col:
        st.markdown('<div class="section-title">📦 持有資產</div>', unsafe_allow_html=True)
    with filter_col:
        market_filter = st.segmented_control(
            "市場篩選",
            options=["🌐 全部", "🇹🇼 台股", "🇺🇸 美股", "🏦 複委託"],
            default="🌐 全部",
            key="asset_market_filter",
            label_visibility="collapsed",
        )
    with toggle_col:
        view_mode = st.segmented_control(
            "顯示方式",
            options=["🗂️ 圖示", "📋 條列"],
            default="🗂️ 圖示",
            key="asset_view_mode",
            label_visibility="collapsed",
        )

    stocks = portfolio.get("stock_details", [])
    cash_holdings = portfolio.get("cash_holdings", [])

    # 根據市場/來源篩選
    if market_filter == "🇹🇼 台股":
        stocks = [s for s in stocks if s.get("source") == "台股"]
        cash_holdings = [c for c in cash_holdings if c.get("source") == "台股"]
    elif market_filter == "🇺🇸 美股":
        stocks = [s for s in stocks if s.get("source") == "FT"]
        cash_holdings = [c for c in cash_holdings if c.get("source") == "FT"]
    elif market_filter == "🏦 複委託":
        stocks = [s for s in stocks if s.get("source") == "複委託"]
        cash_holdings = [c for c in cash_holdings if c.get("source") == "複委託"]

    if not stocks and not cash_holdings:
        st.info("🔍 該市場目前沒有持倉。")
        return

    # 顯示篩選後的統計
    total_count = len(stocks)
    st.caption(f"共 {total_count} 檔持倉")

    if view_mode == "📋 條列":
        _render_assets_list_view(stocks, cash_holdings)
    else:
        _render_assets_card_view(stocks, cash_holdings)


def _render_assets_card_view(stocks, cash_holdings):
    """圖示化模式：卡片方格"""
    if stocks:
        cols = st.columns(3)
        for i, stock in enumerate(stocks):
            with cols[i % 3]:
                pl = stock["profit_loss"]
                pl_pct = stock["profit_loss_pct"]
                pl_class = "asset-pl-up" if pl >= 0 else "asset-pl-down"
                pl_sign = "+" if pl >= 0 else ""
                cur = "NT$" if stock["currency"] == "TWD" else "US$"
                mkt = "🇹🇼" if stock["market"] == "TW" else "🇺🇸"

                st.markdown(f"""
                <div class="asset-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <span class="asset-symbol">{mkt} {stock['symbol']}</span>
                            <div class="asset-name">{stock['name']}</div>
                        </div>
                        <div style="text-align:right;">
                            <div class="asset-value">{cur} {stock['current_price']:,.2f}</div>
                            <div class="{pl_class}" style="font-size:0.8rem;">
                                {pl_sign}{pl:,.0f} ({pl_sign}{pl_pct:.2f}%)
                            </div>
                        </div>
                    </div>
                    <div style="margin-top:10px; color:#a0aec0; font-size:0.78rem; display:flex; justify-content:space-between;">
                        <span>持有 {stock['shares']:,} 股</span>
                        <span>成本 {cur}{stock['avg_cost']:,.2f}</span>
                        <span>市值 {cur}{stock['market_value']:,.0f}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

    # 現金卡片
    if cash_holdings:
        cash_cols = st.columns(min(len(cash_holdings), 3))
        for i, cash in enumerate(cash_holdings):
            with cash_cols[i % 3]:
                cur_sym = "NT$" if cash["currency"] == "TWD" else "US$"
                src_icon = "🏦" if cash.get("source") == "複委託" else ("🇺🇸" if cash.get("source") == "FT" else "🏧")
                st.markdown(f"""
                <div class="asset-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <span class="asset-symbol">{src_icon} {cash.get('note', '現金')}</span>
                        </div>
                        <div class="asset-value">{cur_sym} {cash['amount']:,.2f}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)


def _render_assets_list_view(stocks, cash_holdings):
    """條列式模式：表格清單"""
    if stocks:
        # 建立帶有顏色標記的資料
        rows = []
        for s in stocks:
            pl = s["profit_loss"]
            pl_pct = s["profit_loss_pct"]
            cur = "NT$" if s["currency"] == "TWD" else "US$"
            rows.append({
                "市場": "🇹🇼" if s["market"] == "TW" else "🇺🇸",
                "分類": s.get("category", ""),
                "代碼": s["symbol"],
                "名稱": s["name"],
                "現價": s["current_price"],
                "持股數": s["shares"],
                "均價": round(s["avg_cost"], 2),
                "市值": round(s["market_value"], 0),
                "損益": round(pl, 0),
                "報酬率%": pl_pct,
                "來源": s.get("source", ""),
            })

        df = pd.DataFrame(rows)

        # 用 Streamlit 原生 column_config 做彩色顯示
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "現價": st.column_config.NumberColumn("現價", format="%.2f"),
                "均價": st.column_config.NumberColumn("均價", format="%.2f"),
                "市值": st.column_config.NumberColumn("市值", format="%,.0f"),
                "損益": st.column_config.NumberColumn("損益", format="%,.0f"),
                "報酬率%": st.column_config.NumberColumn("報酬率%", format="%.2f%%"),
            },
        )

    # 現金條列
    if cash_holdings:
        cash_rows = []
        for cash in cash_holdings:
            cash_rows.append({
                "類型": "💵 現金",
                "幣別": cash["currency"],
                "備註": cash.get("note", ""),
                "金額": cash["amount"],
            })
        cash_df = pd.DataFrame(cash_rows)
        st.dataframe(
            cash_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "金額": st.column_config.NumberColumn("金額", format="%,.2f"),
            },
        )



def render_asset_table(portfolio):
    """渲染下方資產明細表格"""
    st.markdown('<div class="section-title">📋 資產明細</div>', unsafe_allow_html=True)

    stocks = portfolio.get("stock_details", [])
    if stocks:
        df = pd.DataFrame([{
            "市場": "🇹🇼 台股" if s["market"] == "TW" else "🇺🇸 美股",
            "代碼": s["symbol"],
            "名稱": s["name"],
            "持股數": s["shares"],
            "均價": s["avg_cost"],
            "現價": s["current_price"],
            "市值": round(s["market_value"], 0),
            "損益": round(s["profit_loss"], 0),
            "報酬率%": s["profit_loss_pct"],
            "幣別": s["currency"],
        } for s in stocks])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("目前沒有股票持倉。")


def render_transaction_history(cfg):
    """渲染交易紀錄"""
    st.markdown('<div class="section-title">📜 交易紀錄</div>', unsafe_allow_html=True)

    txns = cfg.get("transactions", [])
    if txns:
        df = pd.DataFrame([{
            "日期": t["date"],
            "代碼": t["symbol"],
            "名稱": t["name"],
            "操作": "🟢 買入" if t["action"] == "BUY" else "🔴 賣出",
            "股數": t["shares"],
            "價格": t["price"],
            "手續費": t.get("fee", 0),
            "稅金": t.get("tax", 0),
            "幣別": t["currency"],
            "備註": t.get("note", ""),
        } for t in txns])
        df = df.sort_values("日期", ascending=False)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("尚無交易紀錄。")

    # ── XIRR 分析 ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">📈 XIRR 年化報酬率（現有持股）</div>', unsafe_allow_html=True)

    with st.spinner("計算 XIRR…"):
        xirr_rows, overall_xirr = data.calculate_xirr_data(cfg)

    if not xirr_rows:
        st.info("無足夠交易資料計算 XIRR。")
        return

    # 整體 XIRR 大卡片
    if overall_xirr is not None:
        color = "#00e676" if overall_xirr >= 0 else "#ff1744"
        st.markdown(f'''
        <div style="background:rgba(255,255,255,0.04);border-radius:12px;padding:16px 24px;
                    display:inline-block;margin-bottom:16px;border:1px solid rgba(255,255,255,0.1);">
            <div style="color:#888;font-size:0.85rem;">整體持股 XIRR（年化）</div>
            <div style="font-size:2rem;font-weight:700;color:{color};">{overall_xirr*100:+.2f}%</div>
        </div>
        ''', unsafe_allow_html=True)

    # 逐檔明細表
    def _fmt_xirr(v):
        if v is None:
            return "—"
        color = "#00e676" if v >= 0 else "#ff1744"
        return f'<span style="color:{color};font-weight:600;">{v*100:+.2f}%</span>'

    def _sym_label(src):
        if src == "台股": return "NT$"
        return "US$"

    rows_display = []
    for r in xirr_rows:
        sym_label = _sym_label(r["source"])
        rows_display.append({
            "代碼": r["symbol"],
            "名稱": r["name"],
            "來源": r["source"],
            "現值": f'{sym_label} {r["current_mv"]:,.0f}',
            "交易筆數": r["txn_count"],
            "XIRR（年化）": _fmt_xirr(r["xirr"]),
        })

    df_xirr = pd.DataFrame(rows_display)
    st.write(
        df_xirr.to_html(escape=False, index=False),
        unsafe_allow_html=True,
    )


def render_returns_analysis(cfg):
    """渲染報酬分析"""
    st.markdown('<div class="section-title">📊 報酬分析</div>', unsafe_allow_html=True)

    results = data.calculate_transaction_returns(cfg)
    if not results:
        st.info("沒有交易紀錄可供分析。")
        return

    df = pd.DataFrame([{
        "代碼": r["symbol"],
        "名稱": r["name"],
        "買入總股數": r["buy_total_shares"],
        "賣出總股數": r["sell_total_shares"],
        "剩餘股數": r["remaining_shares"],
        "買入總成本": r["buy_total_cost"],
        "賣出總收入": r["sell_total_revenue"],
        "已實現損益": r["realized_pl"],
        "未實現損益": r["unrealized_pl"],
        "總損益": r["total_pl"],
        "報酬率%": r["return_pct"],
    } for r in results])
    st.dataframe(df, use_container_width=True, hide_index=True)

    # 報酬率圖表
    if len(results) > 0:
        fig = go.Figure()
        colors = ["#48bb78" if r["total_pl"] >= 0 else "#fc8181" for r in results]
        fig.add_trace(go.Bar(
            x=[r["name"] for r in results],
            y=[r["return_pct"] for r in results],
            marker_color=colors,
            text=[f"{r['return_pct']:.1f}%" for r in results],
            textposition="outside",
        ))
        fig.update_layout(
            title="各標的報酬率",
            yaxis_title="報酬率 (%)",
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=350,
            margin=dict(t=50, b=30),
        )
        st.plotly_chart(fig, use_container_width=True)


def render_stock_chart(cfg):
    """渲染個股走勢圖"""
    symbols = config.get_all_symbols(cfg)
    if not symbols:
        return

    st.markdown('<div class="section-title">📈 個股走勢</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])
    with col1:
        selected = st.selectbox("選擇股票", symbols, key="chart_stock")
    with col2:
        period = st.selectbox("時間範圍", ["1mo", "3mo", "6mo", "1y", "2y"], index=3, key="chart_period")

    if selected:
        hist = data.get_stock_history(selected, period=period)
        if not hist.empty:
            fig = go.Figure()
            fig.add_trace(go.Candlestick(
                x=hist.index,
                open=hist["Open"], high=hist["High"],
                low=hist["Low"], close=hist["Close"],
                name=selected,
            ))
            fig.update_layout(
                title=f"{selected} K線圖",
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                height=420,
                xaxis_rangeslider_visible=False,
                margin=dict(t=50, b=30),
            )
            st.plotly_chart(fig, use_container_width=True)


def render_operation_panel(cfg):
    """渲染操作面板（新增交易、管理資產等）"""
    st.markdown('<div class="section-title">⚙️ 操作面板</div>', unsafe_allow_html=True)
    tab1, tab2, tab3, tab4 = st.tabs(["📝 新增交易", "💵 現金管理", "🗑️ 管理持倉", "☁️ Google 雲端"])

    with tab1:
        _render_add_transaction_form(cfg)
    with tab2:
        _render_cash_form(cfg)
    with tab3:
        _render_manage_holdings(cfg)
    with tab4:
        _render_google_drive_import(cfg)


def _render_add_transaction_form(cfg):
    """新增交易表單"""
    with st.form("add_txn_form", clear_on_submit=True):
        st.markdown("#### 新增交易紀錄")
        
        market = st.selectbox("市場", ["TW", "US"], key="txn_market")
        symbol = st.text_input("股票代碼", placeholder="例: 2330.TW 或 AAPL", key="txn_symbol")
        name = st.text_input("股票名稱", placeholder="例: 台積電", key="txn_name")
        action = st.selectbox("操作", ["BUY", "SELL"], key="txn_action")
        
        col1, col2 = st.columns(2)
        with col1:
            shares = st.number_input("股數", min_value=1, value=1000, key="txn_shares")
        with col2:
            price = st.number_input("價格", min_value=0.01, value=100.0, format="%.2f", key="txn_price")
        
        col3, col4 = st.columns(2)
        with col3:
            fee = st.number_input("手續費", min_value=0.0, value=0.0, format="%.2f", key="txn_fee")
        with col4:
            tax = st.number_input("交易稅", min_value=0.0, value=0.0, format="%.2f", key="txn_tax")
        
        currency = "TWD" if market == "TW" else "USD"
        txn_date = st.date_input("日期", value=datetime.now(), key="txn_date")
        note = st.text_input("備註", key="txn_note")
        
        submitted = st.form_submit_button("✅ 新增交易", use_container_width=True)
        if submitted and symbol:
            config.add_transaction(
                cfg, txn_date.strftime("%Y-%m-%d"),
                symbol, name, market, action,
                shares, price, fee, tax, currency, note,
            )
            st.success(f"已新增 {action} {symbol} x {shares}")
            st.rerun()


def _render_cash_form(cfg):
    """現金管理表單"""
    st.markdown("#### 目前現金")
    for cash in cfg.get("cash_holdings", []):
        cur_sym = "NT$" if cash["currency"] == "TWD" else "US$"
        st.metric(cash["currency"], f"{cur_sym} {cash['amount']:,.2f}")

    with st.form("cash_form", clear_on_submit=True):
        st.markdown("#### 更新現金")
        currency = st.selectbox("幣別", ["TWD", "USD"], key="cash_currency")
        amount = st.number_input("金額", value=0.0, format="%.2f", key="cash_amount")
        submitted = st.form_submit_button("💰 更新現金", use_container_width=True)
        if submitted:
            config.update_cash(cfg, currency, amount)
            st.success(f"已更新 {currency} 現金為 {amount:,.2f}")
            st.rerun()


def _render_manage_holdings(cfg):
    """管理持倉"""
    st.markdown("#### 目前持倉")
    holdings = cfg.get("stock_holdings", [])
    if holdings:
        for i, h in enumerate(holdings):
            src = h.get("source", "")
            col1, col2 = st.columns([3, 1])
            with col1:
                src_label = f"[{src}]" if src else ""
                st.write(f"**{h['symbol']}** {h['name']} - {h['shares']}股 {src_label}")
            with col2:
                if st.button("🗑️", key=f"del_{i}_{h['symbol']}_{src}"):
                    holdings.pop(i)
                    config.save_config(cfg)
                    st.rerun()
    else:
        st.info("目前無持倉")


def _render_drive_picker(oauth_client, creds, key_prefix):
    """
    可重用的 Google Drive 資料夾瀏覽器 + Excel 檔案選擇器。
    key_prefix: session_state 的命名空間前綴 (e.g. "gd_hold" / "gd_txn")
    Returns: list of excel_file dicts in current folder
    """
    stack_key = f"{key_prefix}_folder_stack"
    if stack_key not in st.session_state:
        st.session_state[stack_key] = []

    folder_stack = st.session_state[stack_key]
    parent_id = folder_stack[-1]["id"] if folder_stack else None
    crumbs = ["📁 My Drive"] + [f["name"] for f in folder_stack]
    st.caption("資料夾路徑：" + " / ".join(crumbs))

    sf_key = f"{key_prefix}_sf_{parent_id}"
    xl_key = f"{key_prefix}_xl_{parent_id}"
    if sf_key not in st.session_state:
        with st.spinner("載入資料夾…"):
            st.session_state[sf_key] = google_drive.list_drive_folders(
                oauth_client, parent_id=parent_id, creds_dict=creds
            )
    if xl_key not in st.session_state:
        with st.spinner("載入 Excel 檔案…"):
            st.session_state[xl_key] = google_drive.list_drive_files(
                oauth_client, folder_id=parent_id, extensions=[".xlsx", ".xls"], creds_dict=creds
            )
    subfolders = st.session_state.get(sf_key) or []
    excel_files = st.session_state.get(xl_key) or []

    col_fsel, col_finto, col_fup, col_fref = st.columns([5, 1, 1, 1])
    with col_fsel:
        folder_options = ["📁 此資料夾"] + [("📂 " + f["name"]) for f in subfolders]
        folder_nav_idx = st.selectbox(
            "子資料夾", range(len(folder_options)),
            format_func=lambda i: folder_options[i],
            key=f"{key_prefix}_folder_nav",
            label_visibility="collapsed",
        )
    with col_finto:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if folder_nav_idx > 0:
            if st.button("📂 進入", key=f"{key_prefix}_into"):
                sel = subfolders[folder_nav_idx - 1]
                st.session_state[stack_key].append({"id": sel["id"], "name": sel["name"]})
                st.rerun()
    with col_fup:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if folder_stack:
            if st.button("⬆️ 返回", key=f"{key_prefix}_up"):
                st.session_state[stack_key].pop()
                st.rerun()
    with col_fref:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("🔄", key=f"{key_prefix}_refresh", help="重新整理"):
            for k in list(st.session_state.keys()):
                if k.startswith(f"{key_prefix}_sf_") or k.startswith(f"{key_prefix}_xl_"):
                    del st.session_state[k]
            st.rerun()

    return excel_files


def _render_google_drive_import(cfg):
    """Google 雲端硬碟匯入功能"""
    st.markdown("#### ☁️ Google 雲端匯入")

    # 檢查套件
    ok, msg = google_drive.check_dependencies()
    if not ok:
        st.error(msg)
        return

    gd_cfg = cfg.get("google_drive", {})
    oauth_client = st.session_state.get("google_client")

    # 連線狀態提示
    if oauth_client:
        st.success(f"✅ 已登入：{st.session_state.get('google_email', '')}")
    else:
        st.warning("尚未登入 Google，請點選右上角「登入 Google」")

    # 本地模式才顯示 OAuth 金鑰路徑設定
    if not _is_cloud_mode():
        with st.form("gd_oauth_form"):
            st.markdown("##### 🔑 OAuth 金鑰路徑（本地）")
            new_path = st.text_input(
                "OAuth 用戶端金鑰路徑（JSON）",
                value=gd_cfg.get("oauth_client_secrets_path", ""),
                placeholder="C:/keys/client_secrets.json",
                key="gd_oauth_secrets",
            )
            if st.form_submit_button("💾 儲存") and new_path:
                cfg.setdefault("google_drive", {})["oauth_client_secrets_path"] = new_path
                config.save_config(cfg)
                st.success("✅ 已儲存")

    if not oauth_client:
        st.info("登入 Google 後即可使用匯入與備份功能")
        return

    # ── 從 Google Drive 匯入 ─────────────────────────────────────
    st.markdown("---")
    st.markdown("##### 📂 從 Google Drive 匯入")

    _creds = st.session_state.get("google_creds")
    tab_hold, tab_txn = st.tabs(["📊 股票持倉", "🧾 交易紀錄"])

    # ── Tab 1: 股票持倉 ──────────────────────────────────────────
    with tab_hold:
        st.caption("讀取 Excel 的「個股持倉」工作表，更新持倉與現金資料。")
        excel_files_hold = _render_drive_picker(oauth_client, _creds, "gd_hold")

        if not excel_files_hold:
            st.info("此資料夾沒有 Excel 檔案（.xlsx / .xls），請進入其他子資料夾")
        else:
            file_opts_hold = [f["name"] for f in excel_files_hold]
            sel_h = st.selectbox(
                "選擇 Excel 檔案", range(len(file_opts_hold)),
                format_func=lambda i: file_opts_hold[i],
                key="gd_hold_file_select",
            )
            sel_file_hold = excel_files_hold[sel_h]
            st.caption(f"修改日期：{sel_file_hold.get('modifiedTime', '')[:10]}")

            st.warning("⚠️ 警告：系統將清除先前持倉/現金資料，以此檔完全覆寫！")
            confirm_hold = st.checkbox("我已了解，將覆寫持倉與現金資料", key="gd_hold_confirm")

            col_prev_h, col_imp_h = st.columns(2)
            with col_prev_h:
                if st.button("👁️ 預覽工作表", use_container_width=True, key="gd_hold_preview"):
                    try:
                        with st.spinner(f"下載 {sel_file_hold['name']}…"):
                            fb = google_drive.download_drive_file(
                                oauth_client, sel_file_hold["id"], creds_dict=_creds
                            )
                        import io as _io
                        xl = pd.ExcelFile(_io.BytesIO(fb))
                        st.success(f"工作表：{', '.join(xl.sheet_names)}")
                        sheet = "個股持倉" if "個股持倉" in xl.sheet_names else xl.sheet_names[0]
                        df_prev = xl.parse(sheet, header=None, nrows=10)
                        st.dataframe(df_prev, use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.error(f"❌ 預覽失敗: {e}")

            with col_imp_h:
                if st.button("📥 匯入持倉", use_container_width=True,
                             key="gd_hold_run", type="primary", disabled=not confirm_hold):
                    try:
                        with st.spinner(f"匯入 {sel_file_hold['name']}…"):
                            fb = google_drive.download_drive_file(
                                oauth_client, sel_file_hold["id"], creds_dict=_creds
                            )
                            import import_excel as _ie
                            cfg = config.load_config()
                            result = _ie.import_holdings_from_bytes(fb, cfg)
                        st.success(f"✅ 持倉匯入完成！持倉 {result['holdings']} 檔、現金 {result['cash']} 筆")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 匯入失敗: {e}")

    # ── Tab 2: 交易紀錄 ──────────────────────────────────────────
    with tab_txn:
        st.caption("讀取 Excel 的「交易紀錄」工作表，更新交易歷史（不影響持倉）。")
        excel_files_txn = _render_drive_picker(oauth_client, _creds, "gd_txn")

        if not excel_files_txn:
            st.info("此資料夾沒有 Excel 檔案（.xlsx / .xls），請進入其他子資料夾")
        else:
            file_opts_txn = [f["name"] for f in excel_files_txn]
            sel_t = st.selectbox(
                "選擇 Excel 檔案", range(len(file_opts_txn)),
                format_func=lambda i: file_opts_txn[i],
                key="gd_txn_file_select",
            )
            sel_file_txn = excel_files_txn[sel_t]
            st.caption(f"修改日期：{sel_file_txn.get('modifiedTime', '')[:10]}")

            st.warning("⚠️ 警告：系統將清除先前所有交易紀錄，以此檔完全覆寫！")
            confirm_txn = st.checkbox("我已了解，將覆寫所有交易紀錄", key="gd_txn_confirm")

            col_prev_t, col_imp_t = st.columns(2)
            with col_prev_t:
                if st.button("👁️ 預覽工作表", use_container_width=True, key="gd_txn_preview"):
                    try:
                        with st.spinner(f"下載 {sel_file_txn['name']}…"):
                            fb = google_drive.download_drive_file(
                                oauth_client, sel_file_txn["id"], creds_dict=_creds
                            )
                        import io as _io
                        xl = pd.ExcelFile(_io.BytesIO(fb))
                        st.success(f"工作表：{', '.join(xl.sheet_names)}")
                        sheet = "交易紀錄" if "交易紀錄" in xl.sheet_names else xl.sheet_names[0]
                        df_prev = xl.parse(sheet, nrows=10)
                        st.dataframe(df_prev, use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.error(f"❌ 預覽失敗: {e}")

            with col_imp_t:
                if st.button("📥 匯入交易紀錄", use_container_width=True,
                             key="gd_txn_run", type="primary", disabled=not confirm_txn):
                    try:
                        with st.spinner(f"匯入 {sel_file_txn['name']}…"):
                            fb = google_drive.download_drive_file(
                                oauth_client, sel_file_txn["id"], creds_dict=_creds
                            )
                            import import_excel as _ie
                            cfg = config.load_config()
                            result = _ie.import_transactions_from_bytes(fb, cfg)
                        st.success(f"✅ 交易紀錄匯入完成！共 {result['transactions']} 筆")

                        # 預覽匯入結果
                        if result["records"]:
                            df_txn = pd.DataFrame([{
                                "日期": t["date"],
                                "代碼": t["symbol"],
                                "名稱": t["name"],
                                "操作": "🟢 買入" if t["action"] == "BUY" else "🔴 賣出",
                                "股數": t["shares"],
                                "價格": t["price"],
                                "幣別": t["currency"],
                            } for t in result["records"]])
                            df_txn = df_txn.sort_values("日期", ascending=False)
                            st.dataframe(df_txn, use_container_width=True, hide_index=True)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 匯入失敗: {e}")

    # 備份到 Google 雲端硬碟
    st.markdown("---")
    st.markdown("##### 💾 備份投資組合到 Google 雲端硬碟")

    gd_cfg = cfg.get("google_drive", {})
    last_synced = gd_cfg.get("last_synced", "")
    if last_synced:
        st.caption(f"上次備份：{last_synced}")

    # ── 資料夾瀏覽器 ──────────────────────────────────────────
    if "gd_folder_stack" not in st.session_state:
        st.session_state.gd_folder_stack = []

    current_parent_id = (
        st.session_state.gd_folder_stack[-1]["id"]
        if st.session_state.gd_folder_stack else None
    )
    crumbs = ["📁 My Drive"] + [f["name"] for f in st.session_state.gd_folder_stack]
    st.caption("目標路徑：" + " / ".join(crumbs))

    cache_key = f"gd_subfolders_{current_parent_id}"
    if cache_key not in st.session_state:
        with st.spinner("載入資料夾清單…"):
            st.session_state[cache_key] = google_drive.list_drive_folders(
                oauth_client, parent_id=current_parent_id,
                creds_dict=st.session_state.get("google_creds")
            )
    subfolders = st.session_state.get(cache_key) or []

    folder_nav_options = ["📁 存入此處"] + [("📂 " + f["name"]) for f in subfolders]
    col_fsel, col_finto, col_fup, col_fref = st.columns([5, 1, 1, 1])
    with col_fsel:
        folder_nav_idx = st.selectbox(
            "目標資料夾",
            range(len(folder_nav_options)),
            format_func=lambda i: folder_nav_options[i],
            key="gd_folder_nav_select",
        )
    with col_finto:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if folder_nav_idx > 0:
            if st.button("📂 進入", key="gd_folder_into", help="進入此子資料夾"):
                sel = subfolders[folder_nav_idx - 1]
                st.session_state.gd_folder_stack.append({"id": sel["id"], "name": sel["name"]})
                st.rerun()
    with col_fup:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.session_state.gd_folder_stack:
            if st.button("⬆️ 返回", key="gd_folder_up", help="返回上一層"):
                st.session_state.gd_folder_stack.pop()
                st.rerun()
    with col_fref:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("🔄", key="gd_folder_refresh", help="重新整理資料夾清單"):
            for k in list(st.session_state.keys()):
                if k.startswith("gd_subfolders_"):
                    del st.session_state[k]
            st.rerun()

    if folder_nav_idx == 0:
        target_folder_id = current_parent_id
    else:
        target_folder_id = subfolders[folder_nav_idx - 1]["id"]

    # ── 格式選擇 + 備份按鈕 ──────────────────────────────────
    col_fmt, col_backup = st.columns([2, 3])
    with col_fmt:
        backup_fmt = st.selectbox(
            "備份格式",
            ["JSON（完整設定）", "CSV ZIP（持倉 + 交易）"],
            key="gd_backup_fmt",
        )
    with col_backup:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("☁️ 立即備份到 Google Drive", use_container_width=True, key="gd_drive_upload", type="primary"):
            from datetime import datetime as _dt
            now_str = _dt.now().strftime("%Y%m%d_%H%M%S")
            try:
                with st.spinner("正在上傳到 Google 雲端硬碟…"):
                    if backup_fmt.startswith("JSON"):
                        file_bytes = google_drive.export_portfolio_as_json(cfg)
                        filename = f"portfolio_backup_{now_str}.json"
                        mime_type = "application/json"
                    else:
                        file_bytes = google_drive.export_portfolio_as_csv_zip(cfg)
                        filename = f"portfolio_backup_{now_str}.zip"
                        mime_type = "application/zip"

                    result = google_drive.upload_file_to_drive(
                        oauth_client,
                        file_bytes=file_bytes,
                        filename=filename,
                        mime_type=mime_type,
                        folder_id=target_folder_id,
                        creds_dict=st.session_state.get("google_creds"),
                    )
                synced_at = _dt.now().strftime("%Y-%m-%d %H:%M:%S")
                cfg["google_drive"]["last_synced"] = synced_at
                config.save_config(cfg)
                link = result.get("webViewLink", "")
                st.success(f"✅ 已上傳：{filename}")
                if link:
                    st.markdown(f"[📄 在 Google Drive 開啟]({link})")
            except Exception as e:
                st.error(f"❌ 備份失敗: {e}")


_SORT_COLS = [
    ("代碼",    "symbol"),
    ("名稱",    "name"),
    ("數量",    "shares"),
    ("均價",    "avg_cost"),
    ("投入金額", "invested"),
    ("現價",    "current_price"),
    ("日漲跌",  "change"),
    ("市值",    "market_value"),
    ("損益",    "pl"),
    ("報酬率%", "pl_pct"),
    ("操作",    None),   # not sortable
]
_COL_WIDTHS = [1.2, 1.2, 0.8, 1, 1.2, 1, 1, 1.2, 1.2, 1, 0.7]


def render_home_details(portfolio):
    """渲染首頁下方各分類的詳細列表資訊（台股、美股、複委託、現金）"""
    st.markdown("<br><br>", unsafe_allow_html=True)

    stocks = portfolio.get("stock_details", [])
    cash_holdings = portfolio.get("cash_holdings", [])
    rate = portfolio['usd_twd_rate']
    disp_mode = portfolio.get("display_currency", "Default")
    is_twd = disp_mode == "TWD"
    sym = "NT$" if is_twd else ("US$" if disp_mode == "USD" else "NT$")
    # For pct base: Default and TWD both use TWD total
    total_disp = (portfolio["total_value_usd"] if disp_mode == "USD" else portfolio["total_value_twd"])
    total_disp = total_disp if total_disp > 0 else 1

    def _stock_sym(stock_currency):
        if disp_mode == "Default":
            return "NT$" if stock_currency == "TWD" else "US$"
        return sym

    def _to_disp(amount, stock_currency):
        if disp_mode == "Default":
            return amount  # keep native currency, no conversion
        if is_twd:
            return amount if stock_currency == "TWD" else amount * rate
        else:
            return amount if stock_currency == "USD" else amount / rate

    def _to_pct_base(amount, stock_currency):
        """Always convert to TWD for pct calculation in Default mode."""
        if disp_mode == "Default":
            return amount if stock_currency == "TWD" else amount * rate
        return _to_disp(amount, stock_currency)

    # 預計算各 source 的現金（用於對齊首頁卡片的 pct 分母）
    cash_by_source_twd = {}
    cash_by_source_disp = {}
    for c in cash_holdings:
        src = c.get("source", "其他")
        amt_twd = c["amount"] if c["currency"] == "TWD" else c["amount"] * rate
        cash_by_source_twd[src] = cash_by_source_twd.get(src, 0) + amt_twd
        if disp_mode == "USD":
            amt_d = c["amount"] / rate if c["currency"] == "TWD" else c["amount"]
        elif disp_mode == "TWD":
            amt_d = amt_twd
        else:  # Default: native
            amt_d = c["amount"]
        cash_by_source_disp[src] = cash_by_source_disp.get(src, 0) + amt_d

    # 定義分類與其對應的顏色、圖示
    categories = [
        {"id": "cat-tw", "name": "🇹🇼 台股", "source": "台股", "color": "#00e676"},
        {"id": "cat-ft", "name": "🇺🇸 美股 (FT)", "source": "FT", "color": "#2196f3"},
        {"id": "cat-sub", "name": "🏦 複委託", "source": "複委託", "color": "#9c27b0"},
    ]

    for cat in categories:
        # 過濾此分類股票
        cat_stocks = [s for s in stocks if s.get("source") == cat["source"]]
        if not cat_stocks:
            continue


        # 計算此分類匯總（依顯示幣別），pct = 股票市值 / 總資產（不含現金，現金由現金卡片計算）
        if disp_mode == "Default":
            cat_native_currency = "TWD" if cat["source"] == "台股" else "USD"
            cat_mv = sum(s["market_value_twd"] if cat_native_currency == "TWD" else s["market_value_usd"] for s in cat_stocks)
            cat_cost = sum(_to_disp(s["avg_cost"] * s["shares"], s["currency"]) for s in cat_stocks)
            cat_mv_twd = sum(s["market_value_twd"] for s in cat_stocks)
            pct = (cat_mv_twd / total_disp) * 100
        else:
            cat_mv = sum(s["market_value_twd"] if is_twd else s["market_value_usd"] for s in cat_stocks)
            cat_cost = sum(_to_disp(s["avg_cost"] * s["shares"], s["currency"]) for s in cat_stocks)
            pct = (cat_mv / total_disp) * 100
        cat_sym = _stock_sym(cat_stocks[0]["currency"]) if cat_stocks else sym

        pl = cat_mv - cat_cost
        pl_pct = (pl / cat_cost * 100) if cat_cost > 0 else 0
        pl_color = "#00e676" if pl >= 0 else "#ff1744"
        pl_sign = "+" if pl >= 0 else ""

        # 渲染標題和 Summary Bar
        st.markdown(f'''
        <div id="{cat["id"]}">
            <h3 style="margin-bottom:0;">{cat["name"]}</h3>
            <div class="home-detail-bar" style="border-left-color: {cat["color"]}">
                <span class="home-detail-pct">{pct:.1f}%</span>
                <span class="home-detail-val">{cat_sym} {cat_mv:,.0f}</span>
                <span class="home-detail-inv">投入: {cat_sym} {cat_cost:,.0f}</span>
                <span class="home-detail-pl" style="color: {pl_color}">
                    P&L: {pl_sign}{cat_sym} {abs(pl):,.0f} ({pl_sign}{pl_pct:.2f}%)
                </span>
            </div>
        </div>
        ''', unsafe_allow_html=True)

        # 渲染表格（含操作欄位）
        rows = []
        for s in cat_stocks:
            chg = s.get("change", 0)
            if pd.isna(chg): chg = 0

            # 所有金額依顯示幣別換算
            disp_avg_cost   = _to_disp(s["avg_cost"], s["currency"])
            disp_invested   = _to_disp(s["avg_cost"] * s["shares"], s["currency"])
            disp_price      = _to_disp(s["current_price"], s["currency"])
            disp_change     = _to_disp(chg, s["currency"])
            disp_mv         = _to_disp(s["market_value"], s["currency"])
            disp_pl         = _to_disp(s["profit_loss"], s["currency"])
            s_pl_pct        = s["profit_loss_pct"]

            # 顏色判定
            price_color  = "#ff1744" if chg > 0 else ("#00e676" if chg < 0 else "inherit")
            pl_val_color = "#ff1744" if disp_pl > 0 else ("#00e676" if disp_pl < 0 else "inherit")

            rows.append({
                "symbol": s["symbol"],
                "name": s["name"],
                "shares": s["shares"],
                "avg_cost": round(disp_avg_cost, 2),
                "invested": round(disp_invested, 2),
                "current_price": disp_price,
                "change": disp_change,
                "market_value": round(disp_mv, 2),
                "pl": round(disp_pl, 2),
                "pl_pct": s_pl_pct,
                "price_color": price_color,
                "pl_color": pl_val_color,
                "s_sym": _stock_sym(s["currency"]),
                "currency": s["currency"],
                "market": s.get("market", ""),
                "source": s.get("source", ""),
            })
        
        # ── 排序狀態 ──────────────────────────────────────────────
        _sk = f"sort_col_{cat['source']}"
        _sd = f"sort_asc_{cat['source']}"
        if _sk not in st.session_state:
            st.session_state[_sk] = None
            st.session_state[_sd] = True
        _cur_col = st.session_state[_sk]
        _cur_asc = st.session_state[_sd]

        if _cur_col:
            rows.sort(
                key=lambda r: (r[_cur_col] is None, r[_cur_col] if r[_cur_col] is not None else 0),
                reverse=not _cur_asc,
            )

        # 表頭（可點擊排序）
        st.markdown('<div class="sort-hdr"></div>', unsafe_allow_html=True)
        header_cols = st.columns(_COL_WIDTHS)
        for i, (h_label, h_key) in enumerate(_SORT_COLS):
            with header_cols[i]:
                if h_key is None:
                    st.markdown(
                        f'<div style="font-weight:600;font-size:0.82rem;color:#888;'
                        f'padding:4px 0;border-bottom:1px solid #333;">{h_label}</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    is_active = _cur_col == h_key
                    arrow = (" ▲" if _cur_asc else " ▼") if is_active else ""
                    if st.button(
                        f"{h_label}{arrow}",
                        key=f"sort_{cat['source']}_{h_key}",
                        use_container_width=True,
                    ):
                        if _cur_col == h_key:
                            st.session_state[_sd] = not _cur_asc
                        else:
                            st.session_state[_sk] = h_key
                            st.session_state[_sd] = True
                        st.rerun()

        # 資料列
        for row_idx, r in enumerate(rows):
            data_cols = st.columns([1.2, 1.2, 0.8, 1, 1.2, 1, 1, 1.2, 1.2, 1, 0.7])
            
            data_cols[0].markdown(f'<div style="font-size:0.85rem; padding:6px 0;">{r["symbol"]}</div>', unsafe_allow_html=True)
            data_cols[1].markdown(f'<div style="font-size:0.85rem; padding:6px 0;">{r["name"]}</div>', unsafe_allow_html=True)
            data_cols[2].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right;">{r["shares"]:,}</div>', unsafe_allow_html=True)
            data_cols[3].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right;">{r["avg_cost"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[4].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right;">{r["invested"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[5].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right; color:{r["price_color"]};">{r["current_price"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[6].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right; color:{r["price_color"]};">{r["change"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[7].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right;">{r["market_value"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[8].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right; color:{r["pl_color"]};">{r["pl"]:,.2f}</div>', unsafe_allow_html=True)
            data_cols[9].markdown(f'<div style="font-size:0.85rem; padding:6px 0; text-align:right; color:{r["pl_color"]};">{r["pl_pct"]:.2f}%</div>', unsafe_allow_html=True)
            
            # 操作按鈕（編輯 + 刪除）
            with data_cols[10]:
                btn_cols = st.columns(2)
                edit_key = f"edit_{cat['source']}_{r['symbol']}_{row_idx}"
                del_key = f"del_{cat['source']}_{r['symbol']}_{row_idx}"
                
                if btn_cols[0].button("✏️", key=edit_key, help=f"編輯 {r['name']}"):
                    st.session_state[f"editing_{r['symbol']}"] = True
                    
                if btn_cols[1].button("🗑️", key=del_key, help=f"刪除 {r['name']}"):
                    st.session_state[f"confirm_delete_{r['symbol']}"] = True
            
            # 分隔線
            st.markdown('<div style="border-bottom:1px solid rgba(255,255,255,0.05); margin:0;"></div>', unsafe_allow_html=True)
            
            # 編輯表單（展開在該行下方）
            if st.session_state.get(f"editing_{r['symbol']}", False):
                with st.container():
                    st.markdown(f'<div style="background:rgba(255,255,255,0.03); border-radius:8px; padding:12px; margin:8px 0 16px 0; border:1px solid rgba(255,255,255,0.1);">', unsafe_allow_html=True)
                    st.markdown(f"**✏️ 編輯 {r['name']} ({r['symbol']})**")
                    
                    edit_form_cols = st.columns(5)
                    new_symbol = edit_form_cols[0].text_input("股票代碼", value=r["symbol"], key=f"edit_symbol_{r['symbol']}")
                    new_name = edit_form_cols[1].text_input("名稱", value=r["name"], key=f"edit_name_{r['symbol']}")
                    new_shares = edit_form_cols[2].number_input("持股數", value=r["shares"], min_value=0, step=1, key=f"edit_shares_{r['symbol']}")
                    new_avg_cost = edit_form_cols[3].number_input("均價", value=r["avg_cost"], min_value=0.0, step=0.01, format="%.2f", key=f"edit_cost_{r['symbol']}")
                    new_source = edit_form_cols[4].text_input("來源", value=r.get("source", ""), key=f"edit_source_{r['symbol']}")
                    
                    save_col, cancel_col, _ = st.columns([1, 1, 6])
                    if save_col.button("💾 儲存", key=f"save_{r['symbol']}", type="primary"):
                        cfg = config.load_config()
                        old_symbol = r["symbol"]
                        new_symbol_clean = new_symbol.strip()
                        
                        for h in cfg["stock_holdings"]:
                            if h["symbol"] == old_symbol:
                                # 更新代碼（若有變更）
                                h["symbol"] = new_symbol_clean
                                h["name"] = new_name
                                h["shares"] = int(new_shares)
                                h["avg_cost"] = float(new_avg_cost)
                                h["source"] = new_source
                                # 根據新代碼自動判斷 market
                                if new_symbol_clean.endswith(".TW") or new_symbol_clean.endswith(".TWO"):
                                    h["market"] = "TW"
                                    h["currency"] = "TWD"
                                else:
                                    h["market"] = "US"
                                    h["currency"] = "USD"
                                break
                        
                        config.save_config(cfg)
                        
                        # 清除舊報價快取，讓系統用新代碼重新抓取
                        data.clear_cache(f"stock_price_{old_symbol}")
                        if new_symbol_clean != old_symbol:
                            data.clear_cache(f"stock_price_{new_symbol_clean}")
                        
                        st.session_state[f"editing_{old_symbol}"] = False
                        st.rerun()
                    
                    if cancel_col.button("❌ 取消", key=f"cancel_{r['symbol']}"):
                        st.session_state[f"editing_{r['symbol']}"] = False
                        st.rerun()
                    
                    st.markdown('</div>', unsafe_allow_html=True)
            
            # 刪除確認
            if st.session_state.get(f"confirm_delete_{r['symbol']}", False):
                with st.container():
                    st.warning(f"⚠️ 確定要刪除 **{r['name']} ({r['symbol']})** 嗎？此操作無法復原。")
                    confirm_col, cancel_del_col, _ = st.columns([1, 1, 6])
                    if confirm_col.button("🗑️ 確定刪除", key=f"confirm_del_{r['symbol']}", type="primary"):
                        cfg = config.load_config()
                        config.remove_stock_holding(cfg, r["symbol"])
                        st.session_state[f"confirm_delete_{r['symbol']}"] = False
                        st.rerun()
                    
                    if cancel_del_col.button("取消", key=f"cancel_del_{r['symbol']}"):
                        st.session_state[f"confirm_delete_{r['symbol']}"] = False
                        st.rerun()

    # 現金區塊
    if cash_holdings:
        if disp_mode == "USD":
            total_cash_disp = portfolio["total_cash_usd"]
            cash_sym_h = "US$"
        else:
            total_cash_disp = portfolio["total_cash_twd"]
            cash_sym_h = "NT$"
        cash_pct = (portfolio["total_cash_twd"] / total_disp) * 100

        st.markdown(f'''
        <div id="cat-cash">
            <h3 style="margin-bottom:0;">💵 現金部位</h3>
            <div class="home-detail-bar" style="border-left-color: #fbc02d">
                <span class="home-detail-pct">{cash_pct:.1f}%</span>
                <span class="home-detail-val">{cash_sym_h} {total_cash_disp:,.0f}</span>
                <span class="home-detail-inv">總現金</span>
            </div>
        </div>
        ''', unsafe_allow_html=True)

        cash_rows = []
        for c in cash_holdings:
            if disp_mode == "USD":
                amt_disp = c["amount"] / rate if c["currency"] == "TWD" else c["amount"]
                disp_label = "折合美元"
            elif disp_mode == "TWD":
                amt_disp = c["amount"] if c["currency"] == "TWD" else c["amount"] * rate
                disp_label = "折合台幣"
            else:  # Default: show native, no conversion column
                amt_disp = c["amount"]
                disp_label = "金額"
            cash_rows.append({
                "帳戶/來源": c.get("source", "現金"),
                "備註": c.get("note", ""),
                "幣別": c["currency"],
                disp_label: round(amt_disp, 2),
            })

        df_cash = pd.DataFrame(cash_rows)
        st.dataframe(
            df_cash,
            use_container_width=True,
            hide_index=True,
            column_config={
                disp_label: st.column_config.NumberColumn(format="%,.2f"),
            }
        )
