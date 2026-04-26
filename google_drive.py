# -*- coding: utf-8 -*-
"""
Google 雲端硬碟模塊 google_drive.py
串接 Google Sheets / Google Drive 讀取雲端上的 Excel 交易紀錄

使用方式:
1. 前往 Google Cloud Console 建立專案
2. 啟用 Google Sheets API 和 Google Drive API
3. 建立服務帳戶 (Service Account)，下載 JSON 金鑰檔
4. 將試算表分享給服務帳戶的 email (xxx@xxx.iam.gserviceaccount.com)
5. 在程式中設定金鑰路徑與試算表名稱
"""

import os
import pandas as pd

# 嘗試導入 gspread，若未安裝則標記
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# 嘗試導入 OAuth2 套件
try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as OAuthCredentials
    OAUTH_AVAILABLE = True
except ImportError:
    OAUTH_AVAILABLE = False


# Google API 範圍（唯讀，用於匯入）
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Google API 範圍（讀寫，用於匯出備份）
WRITE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 預設欄位對應（Google Sheet 欄位名稱 → 系統內部欄位名稱）
DEFAULT_COLUMN_MAP = {
    "日期": "date",
    "date": "date",
    "Date": "date",
    "股票代碼": "symbol",
    "代碼": "symbol",
    "symbol": "symbol",
    "Symbol": "symbol",
    "股票名稱": "name",
    "名稱": "name",
    "name": "name",
    "Name": "name",
    "市場": "market",
    "market": "market",
    "Market": "market",
    "操作": "action",
    "action": "action",
    "Action": "action",
    "股數": "shares",
    "shares": "shares",
    "Shares": "shares",
    "價格": "price",
    "price": "price",
    "Price": "price",
    "手續費": "fee",
    "fee": "fee",
    "Fee": "fee",
    "交易稅": "tax",
    "稅金": "tax",
    "tax": "tax",
    "Tax": "tax",
    "幣別": "currency",
    "currency": "currency",
    "Currency": "currency",
    "備註": "note",
    "note": "note",
    "Note": "note",
}


OAUTH_SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

OAUTH_TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "google_token.json")


def has_valid_token(token_path=None):
    """檢查本地是否有有效（或可刷新）的 OAuth token"""
    if not OAUTH_AVAILABLE:
        return False
    if token_path is None:
        token_path = OAUTH_TOKEN_FILE
    if not os.path.exists(token_path):
        return False
    try:
        creds = OAuthCredentials.from_authorized_user_file(token_path, OAUTH_SCOPES)
        return creds.valid or bool(creds.expired and creds.refresh_token)
    except Exception:
        return False


def get_oauth_credentials(client_secrets_path, token_path=None):
    """
    本地端：取得 OAuth2 憑證，自動處理 token 快取與刷新。
    若無有效 token，會開啟瀏覽器讓使用者登入。
    """
    if not OAUTH_AVAILABLE:
        raise ImportError("請先安裝: pip install google-auth-oauthlib")

    if token_path is None:
        token_path = OAUTH_TOKEN_FILE

    creds = None

    if os.path.exists(token_path):
        try:
            creds = OAuthCredentials.from_authorized_user_file(token_path, OAUTH_SCOPES)
        except Exception:
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open(token_path, "w", encoding="utf-8") as f:
                f.write(creds.to_json())
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if not os.path.exists(client_secrets_path):
            raise FileNotFoundError(f"找不到 OAuth 用戶端金鑰: {client_secrets_path}")
        flow = InstalledAppFlow.from_client_secrets_file(client_secrets_path, OAUTH_SCOPES)
        creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return creds


# ── 雲端 Web OAuth 流程 ──────────────────────────────────────────────────────

def get_web_auth_url(client_secrets_dict, redirect_uri):
    """
    雲端：產生 Google OAuth 授權 URL（含 PKCE）。
    code_verifier 被編碼進 state 參數，redirect 回來後可自動取出，
    不依賴 session_state 跨請求保留。
    Returns:
        (auth_url, encoded_state)
    """
    if not OAUTH_AVAILABLE:
        raise ImportError("請先安裝: pip install google-auth-oauthlib")

    import base64 as _b64
    import hashlib as _hl
    import json as _json
    import secrets as _sec
    from google_auth_oauthlib.flow import Flow

    # 手動產生 PKCE pair
    code_verifier = _b64.urlsafe_b64encode(_sec.token_bytes(32)).decode().rstrip("=")
    code_challenge = _b64.urlsafe_b64encode(
        _hl.sha256(code_verifier.encode()).digest()
    ).decode().rstrip("=")

    # 把 code_verifier 編進 state，redirect 回來後從 state 解碼取出
    raw_state = _sec.token_urlsafe(16)
    state_payload = _json.dumps({"s": raw_state, "cv": code_verifier})
    encoded_state = _b64.urlsafe_b64encode(state_payload.encode()).decode().rstrip("=")

    flow = Flow.from_client_config(
        client_secrets_dict,
        scopes=OAUTH_SCOPES,
        redirect_uri=redirect_uri,
    )
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
        state=encoded_state,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )
    return auth_url, encoded_state


def exchange_web_auth_code(client_secrets_dict, code, state, redirect_uri, code_verifier=None):
    """
    雲端：用授權碼換取 credentials。
    code_verifier 優先從參數取；若未提供，嘗試從 state 解碼取出。
    Returns:
        google.oauth2.credentials.Credentials
    """
    if not OAUTH_AVAILABLE:
        raise ImportError("請先安裝: pip install google-auth-oauthlib")

    import base64 as _b64
    import json as _json
    import os as _os
    from google_auth_oauthlib.flow import Flow

    _os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

    # 從 state 解碼取出 code_verifier
    if code_verifier is None and state:
        try:
            padding = (4 - len(state) % 4) % 4
            decoded = _b64.urlsafe_b64decode(state + "=" * padding).decode()
            code_verifier = _json.loads(decoded).get("cv")
        except Exception:
            pass

    flow = Flow.from_client_config(
        client_secrets_dict,
        scopes=OAUTH_SCOPES,
        redirect_uri=redirect_uri,
        state=state,
    )
    flow.fetch_token(code=code, code_verifier=code_verifier)
    return flow.credentials


def credentials_from_dict(creds_dict):
    """從 dict（token JSON）還原 OAuthCredentials"""
    import json as _json
    if isinstance(creds_dict, str):
        creds_dict = _json.loads(creds_dict)
    return OAuthCredentials.from_authorized_user_info(creds_dict, OAUTH_SCOPES)


def credentials_to_dict(creds):
    """將 OAuthCredentials 序列化為 dict，方便存入 session_state"""
    import json as _json
    return _json.loads(creds.to_json())


def make_gspread_client(creds):
    """從 credentials 建立 gspread.Client"""
    if not GSPREAD_AVAILABLE:
        raise ImportError("請先安裝: pip install gspread")
    return gspread.authorize(creds)


def authenticate_oauth(client_secrets_path, token_path=None):
    """
    以 OAuth2 使用者身份登入，回傳 (gspread.Client, user_email)。
    首次登入會開啟瀏覽器完成 Google 授權。
    """
    if not GSPREAD_AVAILABLE:
        raise ImportError("請先安裝: pip install gspread")

    creds = get_oauth_credentials(client_secrets_path, token_path)
    client = gspread.authorize(creds)

    email = ""
    try:
        import google.auth.transport.requests as _gtr
        session = _gtr.AuthorizedSession(creds)
        resp = session.get("https://www.googleapis.com/oauth2/v1/userinfo")
        if resp.status_code == 200:
            email = resp.json().get("email", "")
    except Exception:
        pass

    return client, email


def logout_oauth(token_path=None):
    """刪除本地 OAuth token，完成登出"""
    if token_path is None:
        token_path = OAUTH_TOKEN_FILE
    if os.path.exists(token_path):
        os.remove(token_path)


def authenticate_write(credentials_path):
    """
    使用讀寫權限認證（用於匯出備份）
    Args:
        credentials_path (str): 服務帳戶 JSON 金鑰檔路徑
    Returns:
        gspread.Client: 認證後的 gspread 客戶端
    """
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(f"找不到金鑰檔案: {credentials_path}")
    creds = Credentials.from_service_account_file(credentials_path, scopes=WRITE_SCOPES)
    client = gspread.authorize(creds)
    return client


def _get_or_create_worksheet(sh, title):
    """取得或建立指定名稱的工作表"""
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=2000, cols=20)


def export_portfolio_to_sheet(cfg, spreadsheet_id=None, spreadsheet_name=None, credentials_path=None, client=None):
    """
    將當前投資組合資料寫入 Google Sheet
    Args:
        credentials_path (str): 金鑰檔路徑
        cfg (dict): 投資組合設定 dict
        spreadsheet_id (str): 試算表 ID（優先使用）
        spreadsheet_name (str): 試算表名稱
    Returns:
        str: 寫入完成的時間戳記
    """
    from datetime import datetime as _dt

    ok, msg = check_dependencies()
    if not ok:
        raise ImportError(msg)

    if client is None:
        if credentials_path:
            client = authenticate_write(credentials_path)
        else:
            raise ValueError("需提供 client 或 credentials_path")

    if spreadsheet_id:
        sh = client.open_by_key(spreadsheet_id)
    elif spreadsheet_name:
        sh = client.open(spreadsheet_name)
    else:
        raise ValueError("需提供試算表 ID 或名稱")

    now = _dt.now().strftime("%Y-%m-%d %H:%M:%S")

    # 持倉
    holdings = cfg.get("stock_holdings", [])
    ws = _get_or_create_worksheet(sh, "持倉")
    ws.clear()
    rows = [["股票代碼", "股票名稱", "市場", "股數", "平均成本", "幣別", "備註"]]
    for h in holdings:
        rows.append([
            h.get("symbol", ""), h.get("name", ""), h.get("market", ""),
            h.get("shares", 0), h.get("avg_cost", 0), h.get("currency", ""), h.get("note", ""),
        ])
    ws.update(rows)

    # 交易紀錄
    transactions = sorted(cfg.get("transactions", []), key=lambda x: x.get("date", ""), reverse=True)
    ws = _get_or_create_worksheet(sh, "交易紀錄")
    ws.clear()
    rows = [["日期", "股票代碼", "股票名稱", "市場", "操作", "股數", "價格", "手續費", "交易稅", "幣別", "備註"]]
    for t in transactions:
        rows.append([
            t.get("date", ""), t.get("symbol", ""), t.get("name", ""), t.get("market", ""),
            t.get("action", ""), t.get("shares", 0), t.get("price", 0),
            t.get("fee", 0), t.get("tax", 0), t.get("currency", ""), t.get("note", ""),
        ])
    ws.update(rows)

    # 現金
    ws = _get_or_create_worksheet(sh, "現金")
    ws.clear()
    rows = [["幣別", "金額", "備註"]]
    for c in cfg.get("cash_holdings", []):
        rows.append([c.get("currency", ""), c.get("amount", 0), c.get("note", "")])
    ws.update(rows)

    # 總覽
    ws = _get_or_create_worksheet(sh, "總覽")
    ws.clear()
    rows = [
        ["最後同步時間", now],
        ["持股檔數", len(holdings)],
        ["交易筆數", len(cfg.get("transactions", []))],
    ]
    for c in cfg.get("cash_holdings", []):
        rows.append([f"{c.get('currency', '')} 現金", c.get("amount", 0)])
    ws.update(rows)

    # 刪除預設的空白 Sheet1（避免使用者打開看到空白頁）
    try:
        default_ws = sh.worksheet("Sheet1")
        if not default_ws.get_all_values():
            sh.del_worksheet(default_ws)
    except Exception:
        pass

    return now


def _drive_session(client):
    """從 gspread Client 取得可呼叫 Drive API 的 HTTP session"""
    from google.auth.transport.requests import AuthorizedSession as _AS

    # gspread 6.x: client.session 是帶授權的 HTTP session，可直接使用
    session = getattr(client, "session", None)
    if session is not None and hasattr(session, "get") and hasattr(session, "post"):
        return session

    # gspread 5.x: client.auth 是 raw credentials
    creds = getattr(client, "auth", None)
    if creds is not None:
        return _AS(creds)

    # 最後手段：從已快取的 OAuth token 檔重新建立 session
    if OAUTH_AVAILABLE and os.path.exists(OAUTH_TOKEN_FILE):
        from google.auth.transport.requests import Request
        creds = OAuthCredentials.from_authorized_user_file(OAUTH_TOKEN_FILE, OAUTH_SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return _AS(creds)

    raise RuntimeError("無法從 gspread Client 取得憑證，請確認 gspread 版本")


def list_drive_folders(client, parent_id=None):
    """
    列出 Google Drive 中指定層級的資料夾
    Args:
        client: 已認證的 gspread.Client
        parent_id (str): 父資料夾 ID，None 表示根目錄
    Returns:
        list[dict]: [{"id": ..., "name": ...}, ...]
    """
    session = _drive_session(client)
    parent = parent_id if parent_id else "root"
    folders = []
    page_token = None
    while True:
        params = {
            "q": (
                f"mimeType='application/vnd.google-apps.folder' "
                f"and '{parent}' in parents and trashed=false"
            ),
            "fields": "nextPageToken,files(id,name)",
            "orderBy": "name",
            "pageSize": 100,
        }
        if page_token:
            params["pageToken"] = page_token
        resp = session.get("https://www.googleapis.com/drive/v3/files", params=params)
        if resp.status_code != 200:
            raise Exception(f"Drive API 錯誤 ({resp.status_code}): {resp.text}")
        data = resp.json()
        folders.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return folders


def list_drive_files(client, folder_id=None, extensions=None):
    """
    列出 Google Drive 資料夾中的檔案（非資料夾）
    Args:
        extensions: [".xlsx", ".xls"] 等副檔名過濾，None = 全部
    Returns:
        list[dict]: [{"id", "name", "mimeType", "modifiedTime"}, ...]
    """
    session = _drive_session(client)
    parent = folder_id if folder_id else "root"

    q_parts = [f"'{parent}' in parents", "trashed=false",
               "mimeType != 'application/vnd.google-apps.folder'"]

    if extensions:
        ext_lower = [e.lower() for e in extensions]
        mime_map = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls": "application/vnd.ms-excel",
        }
        mime_filters = [
            f"mimeType='{mime_map[e]}'" for e in ext_lower if e in mime_map
        ]
        if mime_filters:
            q_parts.append(f"({' or '.join(mime_filters)})")

    files = []
    page_token = None
    while True:
        params = {
            "q": " and ".join(q_parts),
            "fields": "nextPageToken,files(id,name,mimeType,modifiedTime)",
            "orderBy": "name",
            "pageSize": 100,
        }
        if page_token:
            params["pageToken"] = page_token
        resp = session.get("https://www.googleapis.com/drive/v3/files", params=params)
        if resp.status_code != 200:
            raise Exception(f"Drive API 錯誤 ({resp.status_code}): {resp.text}")
        data = resp.json()
        files.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return files


def download_drive_file(client, file_id):
    """
    下載 Google Drive 檔案內容，回傳 bytes
    """
    session = _drive_session(client)
    resp = session.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"alt": "media"},
    )
    if resp.status_code != 200:
        raise Exception(f"下載失敗 ({resp.status_code}): {resp.text}")
    return resp.content


def export_portfolio_as_json(cfg):
    """將投資組合匯出為 JSON bytes"""
    import json as _json
    return _json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8")


def export_portfolio_as_csv_zip(cfg):
    """將投資組合的持倉、交易、現金匯出為 ZIP 壓縮檔（含三個 CSV）"""
    import io
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        holdings = cfg.get("stock_holdings", [])
        if holdings:
            df = pd.DataFrame(holdings)
            zf.writestr("holdings.csv", df.to_csv(index=False, encoding="utf-8-sig"))

        transactions = cfg.get("transactions", [])
        if transactions:
            df = pd.DataFrame(transactions)
            zf.writestr("transactions.csv", df.to_csv(index=False, encoding="utf-8-sig"))

        cash = cfg.get("cash_holdings", [])
        if cash:
            df = pd.DataFrame(cash)
            zf.writestr("cash.csv", df.to_csv(index=False, encoding="utf-8-sig"))

    buf.seek(0)
    return buf.read()


def upload_file_to_drive(client, file_bytes, filename, mime_type="application/octet-stream", folder_id=None):
    """
    上傳檔案到 Google Drive（multipart upload）
    Args:
        client: 已認證的 gspread.Client
        file_bytes (bytes): 檔案內容
        filename (str): 儲存在 Drive 的檔名
        mime_type (str): MIME type
        folder_id (str|None): 目標資料夾 ID，None = 根目錄
    Returns:
        dict: {"id": ..., "name": ..., "webViewLink": ...}
    """
    import json as _json

    session = _drive_session(client)
    metadata = {"name": filename}
    if folder_id:
        metadata["parents"] = [folder_id]

    boundary = "portfolio_drive_boundary_xyz"
    meta_bytes = _json.dumps(metadata).encode("utf-8")

    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n".encode()
        + meta_bytes
        + f"\r\n--{boundary}\r\nContent-Type: {mime_type}\r\n\r\n".encode()
        + file_bytes
        + f"\r\n--{boundary}--".encode()
    )

    resp = session.post(
        "https://www.googleapis.com/upload/drive/v3/files"
        "?uploadType=multipart&fields=id,name,webViewLink",
        headers={"Content-Type": f"multipart/related; boundary={boundary}"},
        data=body,
    )

    if resp.status_code in (200, 201):
        return resp.json()
    raise Exception(f"上傳失敗 ({resp.status_code}): {resp.text}")


def check_dependencies():
    """檢查是否已安裝必要的套件"""
    if not GSPREAD_AVAILABLE:
        return False, "請先安裝 gspread 和 google-auth: pip install gspread google-auth"
    return True, "OK"


def authenticate(credentials_path):
    """
    使用服務帳戶金鑰認證
    Args:
        credentials_path (str): 服務帳戶 JSON 金鑰檔路徑
    Returns:
        gspread.Client: 認證後的 gspread 客戶端
    """
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(f"找不到金鑰檔案: {credentials_path}")

    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client


def list_spreadsheets(client):
    """
    列出使用者可存取的所有 Google 試算表
    Args:
        client: 已認證的 gspread.Client
    Returns:
        list[dict]: [{"id": ..., "title": ...}, ...]
    """
    try:
        files = client.list_spreadsheet_files()
        return [{"id": f["id"], "title": f["name"]} for f in files]
    except Exception as e:
        return [{"error": str(e)}]


def list_worksheets(client, spreadsheet_name=None, spreadsheet_id=None):
    """
    列出試算表中的所有工作表
    Args:
        client: gspread 客戶端
        spreadsheet_name (str): 試算表名稱
        spreadsheet_id (str): 試算表 ID（優先使用）
    Returns:
        list[str]: 工作表名稱列表
    """
    try:
        if spreadsheet_id:
            sh = client.open_by_key(spreadsheet_id)
        elif spreadsheet_name:
            sh = client.open(spreadsheet_name)
        else:
            return []
        return [ws.title for ws in sh.worksheets()]
    except Exception as e:
        return []


def read_sheet_as_dataframe(client, spreadsheet_name=None, spreadsheet_id=None, worksheet_name=None):
    """
    讀取 Google Sheet 中的數據為 DataFrame
    Args:
        client: gspread 客戶端
        spreadsheet_name (str): 試算表名稱
        spreadsheet_id (str): 試算表 ID（優先使用）
        worksheet_name (str): 工作表名稱（預設為第一個）
    Returns:
        pd.DataFrame: 試算表數據
    """
    try:
        if spreadsheet_id:
            sh = client.open_by_key(spreadsheet_id)
        elif spreadsheet_name:
            sh = client.open(spreadsheet_name)
        else:
            raise ValueError("需提供試算表名稱或 ID")

        if worksheet_name:
            ws = sh.worksheet(worksheet_name)
        else:
            ws = sh.sheet1

        records = ws.get_all_records()
        if not records:
            return pd.DataFrame()

        df = pd.DataFrame(records)
        return df

    except Exception as e:
        raise Exception(f"讀取 Google Sheet 失敗: {e}")


def map_columns(df, column_map=None):
    """
    將 DataFrame 欄位名稱對應到系統內部名稱
    Args:
        df (pd.DataFrame): 原始數據
        column_map (dict): 自訂欄位對應（可選）
    Returns:
        pd.DataFrame: 對應後的數據
    """
    if column_map is None:
        column_map = DEFAULT_COLUMN_MAP

    rename_dict = {}
    for col in df.columns:
        if col in column_map:
            rename_dict[col] = column_map[col]

    if rename_dict:
        df = df.rename(columns=rename_dict)

    return df


def parse_transactions(df):
    """
    將 DataFrame 解析為交易紀錄列表
    Args:
        df (pd.DataFrame): 已對應欄位的數據
    Returns:
        list[dict]: 交易紀錄列表
    """
    required_cols = ["date", "symbol", "action", "shares", "price"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位: {missing}")

    transactions = []
    for _, row in df.iterrows():
        # 解析操作類型
        action = str(row.get("action", "")).strip().upper()
        if action in ["買入", "買", "BUY", "B"]:
            action = "BUY"
        elif action in ["賣出", "賣", "SELL", "S"]:
            action = "SELL"
        else:
            continue  # 跳過無法辨識的操作

        # 解析市場
        market = str(row.get("market", "")).strip().upper()
        if market not in ["TW", "US"]:
            symbol = str(row.get("symbol", ""))
            market = "TW" if ".TW" in symbol or ".TWO" in symbol else "US"

        # 解析幣別
        currency = str(row.get("currency", "")).strip().upper()
        if currency not in ["TWD", "USD"]:
            currency = "TWD" if market == "TW" else "USD"

        txn = {
            "date": str(row.get("date", "")),
            "symbol": str(row.get("symbol", "")).strip(),
            "name": str(row.get("name", row.get("symbol", ""))).strip(),
            "market": market,
            "action": action,
            "shares": int(float(row.get("shares", 0))),
            "price": float(row.get("price", 0)),
            "fee": float(row.get("fee", 0)),
            "tax": float(row.get("tax", 0)),
            "currency": currency,
            "note": str(row.get("note", "")),
        }
        transactions.append(txn)

    return transactions


def import_transactions_from_google(
    credentials_path=None,
    spreadsheet_name=None,
    spreadsheet_id=None,
    worksheet_name=None,
    column_map=None,
    client=None,
):
    """
    一站式從 Google Sheets 匯入交易紀錄
    Args:
        credentials_path (str): 服務帳戶金鑰檔路徑（與 client 二擇一）
        spreadsheet_name (str): 試算表名稱
        spreadsheet_id (str): 試算表 ID
        worksheet_name (str): 工作表名稱
        column_map (dict): 自訂欄位對應
        client: 已認證的 gspread.Client（優先使用）
    Returns:
        tuple: (transactions_list, raw_dataframe)
    """
    ok, msg = check_dependencies()
    if not ok:
        raise ImportError(msg)

    if client is None:
        if not credentials_path:
            raise ValueError("需提供 client 或 credentials_path")
        client = authenticate(credentials_path)
    df = read_sheet_as_dataframe(client, spreadsheet_name, spreadsheet_id, worksheet_name)

    if df.empty:
        return [], df

    df = map_columns(df, column_map)
    transactions = parse_transactions(df)

    return transactions, df
