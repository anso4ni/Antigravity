# -*- coding: utf-8 -*-
"""
主模塊 astock_single.py
股票交易紀錄程序 - 支援台股與美股
使用方式: streamlit run astock_single.py
"""

import streamlit as st
import config
import data
import ui


@st.fragment(run_every=30)
def _home_portfolio(cfg):
    """首頁投資組合區塊：每 30 秒自動重新抓取最新報價並重繪"""
    from datetime import datetime
    portfolio = data.calculate_portfolio_value(cfg)
    ui.render_net_value(portfolio)
    ui.render_home_details(portfolio)
    st.caption(f"⏱ 報價自動更新中，最後更新：{datetime.now().strftime('%H:%M:%S')}")


def main():
    """程序主入口"""
    # 1. 頁面設定
    ui.setup_page()

    # 2. 讀取配置
    cfg = config.load_config()

    # 3. 側邊欄導航
    with st.sidebar:
        st.markdown("## 🧭 導覽")
        page = st.radio(
            "選擇頁面", 
            ["🏠 首頁", "⚙️ 操作面板", "📦 持有資產", "📝 資產明細", "🧾 交易紀錄", "📊 報酬分析"],
            label_visibility="collapsed"
        )
        st.markdown("---")
        st.caption("📈 股票交易紀錄系統 v2.0")

    # 4. 重新載入配置（確保最新）
    cfg = config.load_config()

    if page == "🏠 首頁":
        # 頂部即時指數 + Google 登入圖示
        ui.render_market_indices(cfg)
        
        if not st.session_state.get("google_email"):
            st.info("🔒 您目前處於未登入狀態。為了保護您的資產資料，本系統已實施帳號隔離。請點擊右上角「登入 Google」以載入專屬於您的投資組合。")
            return
            
        # 計算並顯示投資組合（每 30 秒自動刷新報價）
        _home_portfolio(cfg)

    elif page == "⚙️ 操作面板":
        if not st.session_state.get("google_email"):
            st.warning("請先登入 Google 帳號以使用操作面板。")
            return
        ui.render_operation_panel(cfg)

    elif page == "📦 持有資產":
        if not st.session_state.get("google_email"):
            st.warning("請先登入 Google 帳號以查看持有資產。")
            return
        portfolio = data.calculate_portfolio_value(cfg)
        ui.render_asset_cards(portfolio, cfg)
        ui.render_stock_chart(cfg)

    elif page == "📝 資產明細":
        if not st.session_state.get("google_email"):
            st.warning("請先登入 Google 帳號以查看資產明細。")
            return
        portfolio = data.calculate_portfolio_value(cfg)
        ui.render_asset_table(portfolio)

    elif page == "🧾 交易紀錄":
        if not st.session_state.get("google_email"):
            st.warning("請先登入 Google 帳號以查看交易紀錄。")
            return
        ui.render_transaction_history(cfg)

    elif page == "📊 報酬分析":
        if not st.session_state.get("google_email"):
            st.warning("請先登入 Google 帳號以查看報酬分析。")
            return
        ui.render_returns_analysis(cfg)


if __name__ == "__main__":
    main()
