import base64
import hashlib
import hmac
import html
import io
import json
import os
import secrets as token_secrets
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openai import OpenAI
from PIL import Image

load_dotenv()

APP_TITLE = "AI仕訳アシスタント（日本の中小企業・個人事業主向け）"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")
CUSTOMER_DB_PATH = Path(os.getenv("CUSTOMER_DB_PATH", "customer_accounts.json"))
LEDGER_STORAGE_DIR = Path(os.getenv("LEDGER_STORAGE_DIR", "customer_ledgers"))
PLAN_CONFIG = {
    "free": {"label": "無料プラン", "limit": 1},
    "starter": {"label": "スタータープラン", "limit": 5},
    "business": {"label": "ビジネスプラン", "limit": 20},
    "accounting_firm": {"label": "会計事務所プラン", "limit": 50},
}

PLAN_TEXT_EXAMPLES = {
    "free": "例：2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。キーボードとマウスを購入。消費税10%、適格請求書取得済み。",
    "starter": "\n".join([
        "例：1行に1取引ずつ入力できます。",
        "2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。キーボードとマウスを購入。消費税10%。",
        "2026年6月11日、現金でローソンへ680円支払い。会議用のお茶を購入。",
        "2026年6月12日、普通預金からNTTへ8,800円支払い。会社携帯料金。",
    ]),
    "business": "\n".join([
        "例：月次の経費をまとめて入力できます。1行1取引で記載してください。",
        "2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。PC周辺機器を購入。消費税10%。",
        "2026年6月11日、JRへ1,240円支払い。営業訪問の交通費。",
        "2026年6月12日、Zoomへ2,200円支払い。オンライン会議ツール利用料。",
        "2026年6月13日、カフェで1,500円支払い。取引先との打合せ。",
    ]),
    "accounting_firm": "\n".join([
        "例：顧問先ごとの明細をまとめて入力できます。1行1取引で記載してください。",
        "A社 2026年6月10日、Amazon Japan、11,000円、法人カード、PC周辺機器、消費税10%。",
        "A社 2026年6月11日、JR、1,240円、営業交通費。",
        "B社 2026年6月12日、NTT、8,800円、普通預金、通信費。",
        "B社 2026年6月13日、飲食店、9,900円、法人カード、取引先接待。",
    ]),
}

PAGE_CSS = """
<style>
    .stApp {
        background: #0b1014;
        color: #e5edf5;
    }
    .block-container {
        max-width: 1040px;
        padding-top: 4.25rem;
        padding-bottom: 3rem;
    }
    .app-header {
        border: 1px solid #263340;
        border-radius: 8px;
        background: #101820;
        padding: 0.8rem 1rem 0.85rem 1rem;
        margin: 0.2rem auto 0.9rem auto;
        max-width: 760px;
        text-align: center;
    }
    .app-title {
        color: #f8fafc;
        font-size: 1.34rem;
        line-height: 1.55;
        font-weight: 760;
        margin: 0;
        padding: 0;
    }
    .app-subtitle {
        color: #a8b3c2;
        font-size: 0.9rem;
        line-height: 1.45;
        margin: 0.1rem 0 0 0;
    }
    .login-shell {
        max-width: 980px;
        margin: 1.25rem auto 0 auto;
    }
    h1, h2, h3, h4, h5, h6,
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li {
        color: #e5edf5;
    }
    div[data-testid="stMetric"] {
        background: #121a21;
        border: 1px solid #2a3642;
        border-radius: 8px;
        padding: 0.55rem 0.75rem;
        box-shadow: none;
        min-height: 72px;
    }
    div[data-testid="stMetric"] label {
        color: #a8b3c2 !important;
        font-weight: 650;
        font-size: 0.78rem !important;
    }
    div[data-testid="stMetricValue"] {
        color: #f8fafc !important;
        font-weight: 750;
        font-size: 1.35rem !important;
    }
    div[data-testid="stMetricValue"] > div {
        color: #f8fafc !important;
        font-size: 1.35rem !important;
    }
    section[data-testid="stSidebar"] {
        background: #10161d;
        border-right: 1px solid #24303a;
    }
    section[data-testid="stSidebar"] * {
        color: #e5edf5;
    }
    .app-panel {
        border: 1px solid #2a3642;
        border-radius: 8px;
        padding: 0.85rem 1rem;
        background: #121a21;
        margin-bottom: 1rem;
    }
    .app-panel-muted {
        border: 1px solid #263340;
        border-radius: 8px;
        padding: 0.85rem 1rem;
        background: #0f171e;
        margin-bottom: 1rem;
    }
    .eyebrow {
        color: #7dd3fc;
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0;
        margin-bottom: 0.25rem;
    }
    .panel-title {
        color: #f8fafc;
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .panel-copy {
        color: #cbd5e1;
        font-size: 0.94rem;
        line-height: 1.55;
        margin: 0;
    }
    .plan-pill {
        display: inline-block;
        border: 1px solid #2dd4bf;
        border-radius: 999px;
        padding: 0.2rem 0.55rem;
        color: #ccfbf1;
        background: #102a2a;
        font-size: 0.82rem;
        font-weight: 600;
    }
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="stMetric"]) {
        gap: 0.75rem;
    }
    div[data-testid="stAlert"] {
        border-radius: 8px;
    }
    div[data-testid="stDataFrame"],
    div[data-testid="stDataEditor"] {
        border: 1px solid #2a3642;
        border-radius: 8px;
        overflow: hidden;
    }
    textarea,
    input {
        color: #f8fafc !important;
    }
</style>
"""

EXCEL_COLUMNS = [
    "取引日", "証憑日付", "取引先", "摘要", "借方勘定科目", "借方補助科目", "借方金額", "借方税区分",
    "貸方勘定科目", "貸方補助科目", "貸方金額", "貸方税区分", "税込金額", "税抜金額", "消費税率",
    "消費税額", "支払方法", "インボイス登録番号", "証憑種類", "ステータス", "信頼度", "確認事項", "元ファイル名",
]

MONTHLY_CHECKLIST_ITEMS = [
    ("証憑回収", "領収書・請求書・カード明細・銀行明細が揃っているか確認"),
    ("預金照合", "普通預金・カード・電子マネー残高と帳簿残高を照合"),
    ("売掛金/買掛金", "未回収・未払の残高と入金/支払予定を確認"),
    ("未払/前払", "家賃・通信費・サブスクなど期間対応が必要な費用を確認"),
    ("給与/源泉", "給与・社会保険・源泉所得税の計上漏れを確認"),
    ("消費税区分", "課税・非課税・不課税・対象外とインボイス番号を確認"),
    ("固定資産候補", "10万円以上または資産性のある支出を確認"),
    ("確認事項", "AIが出した確認事項を原本・税理士へ確認"),
]

YEARLY_CHECKLIST_ITEMS = [
    ("年間証憑整理", "年間の証憑・契約書・明細を保管し、不足を確認"),
    ("棚卸", "商品・材料・仕掛品がある場合は期末棚卸を確認"),
    ("減価償却", "固定資産台帳、取得・売却・除却、償却費を確認"),
    ("決算整理", "未払費用・前払費用・未収収益・前受収益を確認"),
    ("売掛/買掛残高", "期末残高と入金/支払予定の整合性を確認"),
    ("貸倒/不良債権", "長期未回収の債権がある場合は回収可能性を確認"),
    ("消費税", "課税売上割合、インボイス、簡易/原則など申告方式を確認"),
    ("税務申告", "法人税・所得税・地方税・源泉所得税などを専門家へ確認"),
]

ACCOUNTING_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "transactions": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "transaction_date": {"type": "string"},
                    "evidence_date": {"type": "string"},
                    "vendor": {"type": "string"},
                    "description": {"type": "string"},
                    "gross_amount": {"type": "integer"},
                    "net_amount": {"type": "integer"},
                    "tax_amount": {"type": "integer"},
                    "tax_rate": {"type": "string"},
                    "debit_account": {"type": "string"},
                    "debit_sub_account": {"type": "string"},
                    "debit_amount": {"type": "integer"},
                    "debit_tax_category": {"type": "string"},
                    "credit_account": {"type": "string"},
                    "credit_sub_account": {"type": "string"},
                    "credit_amount": {"type": "integer"},
                    "credit_tax_category": {"type": "string"},
                    "payment_method": {"type": "string"},
                    "invoice_registration_number": {"type": "string"},
                    "evidence_type": {"type": "string"},
                    "status": {"type": "string"},
                    "confidence": {"type": "number"},
                    "review_notes": {"type": "array", "items": {"type": "string"}},
                },
                "required": [
                    "transaction_date", "evidence_date", "vendor", "description", "gross_amount", "net_amount", "tax_amount", "tax_rate",
                    "debit_account", "debit_sub_account", "debit_amount", "debit_tax_category", "credit_account", "credit_sub_account",
                    "credit_amount", "credit_tax_category", "payment_method", "invoice_registration_number", "evidence_type",
                    "status", "confidence", "review_notes",
                ],
            },
        }
    },
    "required": ["transactions"],
}

SYSTEM_PROMPT = """
あなたは日本の中小企業・個人事業主向けのAI経理・仕訳アシスタントです。

目的：
領収書・請求書・レシート・クレジットカード明細・銀行明細の画像またはテキストから、日本の小規模事業者が確認しやすい仕訳候補を作成してください。

重要ルール：
1. 出力は指定JSONスキーマに厳密に従うこと。
2. 不明な項目は必ず「要確認」とすること。推測で断定しないこと。
3. 金額は円の整数で出力すること。
4. 借方金額と貸方金額は必ず一致させること。
5. 税込金額・税抜金額・消費税額を可能な範囲で分離すること。
6. 日本の消費税区分を意識すること。
7. インボイス登録番号（T+13桁）が見える場合は必ず抽出すること。
8. 画像の読み取りが不鮮明な場合は review_notes にその旨を書くこと。
9. 最終的な税務判断は税理士確認が必要であることを review_notes に含めること。
10. 複数取引が明確に分かれる場合は transactions に複数行で出力すること。

よく使う勘定科目：
現金、普通預金、売掛金、買掛金、未払金、未収入金、前払費用、売上高、仕入高、消耗品費、
事務用品費、旅費交通費、通信費、水道光熱費、広告宣伝費、接待交際費、会議費、福利厚生費、
外注費、支払手数料、支払報酬料、地代家賃、租税公課、保険料、給料手当、法定福利費、
工具器具備品、ソフトウェア、減価償却費、雑費、仮払消費税等、仮受消費税等。

税区分候補：
課税仕入10%、課税仕入8%、課税売上10%、課税売上8%、非課税、不課税、免税、対象外、要確認。

判断方針：
- 事務用品・PC周辺機器・少額備品：消耗品費
- 10万円以上など固定資産の可能性があるもの：工具器具備品またはソフトウェアを検討し、review_notes に確認事項を書く
- 電車・タクシー・出張関連：旅費交通費
- 顧客との会食：接待交際費
- 社内会議・打合せの飲食：会議費
- 従業員向け飲食・福利厚生：福利厚生費
- 支払方法がクレジットカードの場合、貸方は通常「未払金」
- 現金払いの場合、貸方は「現金」
"""

def get_api_key() -> str:
    try:
        key = st.secrets.get("OPENAI_API_KEY")
        if key:
            return key
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "")


def get_config_value(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def escape_html(value: str) -> str:
    return html.escape(str(value), quote=True)


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or token_secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000)
    return f"pbkdf2_sha256${salt}${digest.hex()}"


def verify_password(password: str, profile: dict) -> bool:
    password_hash = str(profile.get("password_hash", ""))
    if password_hash.startswith("pbkdf2_sha256$"):
        try:
            _, salt, expected = password_hash.split("$", 2)
        except ValueError:
            return False
        actual = hash_password(password, salt).split("$", 2)[2]
        return hmac.compare_digest(actual, expected)

    plain_password = str(profile.get("password", ""))
    return bool(plain_password) and hmac.compare_digest(password, plain_password)


def load_registered_customer_accounts() -> dict[str, dict]:
    if not CUSTOMER_DB_PATH.exists():
        return {}

    try:
        with CUSTOMER_DB_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
    except (OSError, json.JSONDecodeError):
        return {}

    return data if isinstance(data, dict) else {}


def save_registered_customer_accounts(customers: dict[str, dict]) -> None:
    CUSTOMER_DB_PATH.write_text(
        json.dumps(customers, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )


def load_customer_accounts() -> dict[str, dict]:
    customers: dict[str, dict] = load_registered_customer_accounts()

    raw_customers = get_config_value("CUSTOMER_ACCOUNTS")
    if raw_customers:
        try:
            parsed = json.loads(raw_customers)
            if isinstance(parsed, dict):
                customers.update(parsed)
        except json.JSONDecodeError:
            pass

    try:
        secret_customers = st.secrets.get("customers", {})
        customers.update({str(username): dict(profile) for username, profile in secret_customers.items()})
    except Exception:
        pass

    normalized: dict[str, dict] = {}
    for username, profile in customers.items():
        if not isinstance(profile, dict):
            continue

        plan_key = str(profile.get("plan", "free")).strip()
        password = str(profile.get("password", ""))
        password_hash = str(profile.get("password_hash", ""))
        if not username or plan_key not in PLAN_CONFIG:
            continue
        if not password and not password_hash:
            continue

        normalized[str(username).strip()] = {
            "password": password,
            "password_hash": password_hash,
            "plan": plan_key,
            "name": str(profile.get("name", username)).strip() or str(username),
        }

    return normalized


def authenticate_customer(username: str, password: str) -> dict | None:
    username = (username or "").strip()
    password = password or ""
    profile = load_customer_accounts().get(username)

    if not profile:
        return None
    if not verify_password(password, profile):
        return None

    plan_key = profile["plan"]
    return {
        "username": username,
        "name": profile["name"],
        "plan_key": plan_key,
        "plan_label": PLAN_CONFIG[plan_key]["label"],
        "transaction_limit": PLAN_CONFIG[plan_key]["limit"],
    }


def register_free_customer(username: str, password: str, name: str) -> tuple[bool, str]:
    username = (username or "").strip().lower()
    name = (name or "").strip() or username
    password = password or ""

    if not username or "@" not in username:
        return False, "メールアドレスを入力してください。"
    if len(password) < 8:
        return False, "パスワードは8文字以上で入力してください。"
    if username in load_customer_accounts():
        return False, "このメールアドレスはすでに登録されています。"

    registered_customers = load_registered_customer_accounts()
    registered_customers[username] = {
        "password_hash": hash_password(password),
        "plan": "free",
        "name": name,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }

    try:
        save_registered_customer_accounts(registered_customers)
    except OSError:
        return False, "登録情報を保存できませんでした。管理者にお問い合わせください。"

    return True, "無料プランで登録しました。"


def get_logged_in_customer() -> dict | None:
    customer = st.session_state.get("customer")
    if isinstance(customer, dict) and customer.get("username") and customer.get("plan_key") in PLAN_CONFIG:
        return customer
    return None


def logout_customer() -> None:
    for key in ["customer", "df", "result"]:
        st.session_state.pop(key, None)
    st.rerun()


def render_login() -> None:
    upgrade_contact = get_config_value("UPGRADE_CONTACT", "上位プランをご希望の場合は管理者までお問い合わせください。")
    safe_upgrade_contact = escape_html(upgrade_contact)

    st.markdown("<div class=\"login-shell\">", unsafe_allow_html=True)
    left, right = st.columns([1, 1], gap="large")

    with left:
        st.markdown(
            """
            <div class="app-panel-muted">
                <div class="eyebrow">SMALL BUSINESS ACCOUNTING</div>
                <div class="panel-title">領収書から仕訳候補まで、確認しやすく。</div>
                <p class="panel-copy">
                    日本の中小企業・個人事業主向けに、証憑画像や取引メモから仕訳候補を作成します。
                    無料登録後は1取引ずつ試せます。
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="app-panel">
                <div class="panel-title">上位プラン</div>
                <p class="panel-copy">{safe_upgrade_contact}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        login_tab, register_tab = st.tabs(["ログイン", "無料登録"])

        with login_tab:
            with st.form("customer_login_form"):
                username = st.text_input("メールアドレス / ユーザーID")
                password = st.text_input("パスワード", type="password")
                submitted = st.form_submit_button("ログイン", type="primary", width="stretch")

            if submitted:
                customer = authenticate_customer(username, password)
                if customer:
                    st.session_state["customer"] = customer
                    st.rerun()
                st.error("ユーザーIDまたはパスワードが正しくありません。")

        with register_tab:
            st.caption("登録直後は無料プランです。")
            with st.form("customer_register_form"):
                register_name = st.text_input("会社名 / お名前")
                register_email = st.text_input("メールアドレス")
                register_password = st.text_input("パスワード（8文字以上）", type="password")
                register_password_confirm = st.text_input("パスワード確認", type="password")
                registered = st.form_submit_button("無料登録して始める", type="primary", width="stretch")

            if registered:
                if register_password != register_password_confirm:
                    st.error("確認用パスワードが一致しません。")
                else:
                    success, message = register_free_customer(register_email, register_password, register_name)
                    if not success:
                        st.error(message)
                        st.stop()
                    customer = authenticate_customer(register_email, register_password)
                    if customer:
                        st.session_state["customer"] = customer
                        st.success(message)
                        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


def get_client() -> OpenAI:
    api_key = get_api_key()
    if not api_key:
        st.error("OPENAI_API_KEY が設定されていません。Streamlit Secrets または .env に設定してください。")
        st.stop()
    return OpenAI(api_key=api_key)


def image_to_data_url(uploaded_file) -> str:
    image = Image.open(uploaded_file).convert("RGB")
    image.thumbnail((1800, 1800))
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=88)
    encoded = base64.b64encode(buffer.getvalue()).decode("utf-8")
    return f"data:image/jpeg;base64,{encoded}"


def parse_json_safely(text: str) -> dict:
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        cleaned = cleaned.replace("json\n", "", 1).replace("JSON\n", "", 1)
    return json.loads(cleaned)


def analyze_with_ai_vision(client: OpenAI, model: str, text_input: str, uploaded_file) -> dict:
    content = [
        {
            "type": "input_text",
            "text": "以下の証憑画像または取引内容から、日本会計向けの仕訳候補JSONを作成してください。\n\nユーザー入力：\n" + (text_input or "なし"),
        }
    ]
    if uploaded_file is not None:
        content.append({"type": "input_image", "image_url": image_to_data_url(uploaded_file)})

    response = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": [{"type": "input_text", "text": SYSTEM_PROMPT}]},
            {"role": "user", "content": content},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "japanese_accounting_entries",
                "schema": ACCOUNTING_SCHEMA,
                "strict": True,
            }
        },
    )
    return parse_json_safely(response.output_text)


def transactions_to_dataframe(result: dict, file_name: str = "") -> pd.DataFrame:
    rows = []
    for item in result.get("transactions", []):
        rows.append(
            {
                "取引日": item.get("transaction_date", "要確認"),
                "証憑日付": item.get("evidence_date", "要確認"),
                "取引先": item.get("vendor", "要確認"),
                "摘要": item.get("description", "要確認"),
                "借方勘定科目": item.get("debit_account", "要確認"),
                "借方補助科目": item.get("debit_sub_account", ""),
                "借方金額": item.get("debit_amount", 0),
                "借方税区分": item.get("debit_tax_category", "要確認"),
                "貸方勘定科目": item.get("credit_account", "要確認"),
                "貸方補助科目": item.get("credit_sub_account", ""),
                "貸方金額": item.get("credit_amount", 0),
                "貸方税区分": item.get("credit_tax_category", "要確認"),
                "税込金額": item.get("gross_amount", 0),
                "税抜金額": item.get("net_amount", 0),
                "消費税率": item.get("tax_rate", "要確認"),
                "消費税額": item.get("tax_amount", 0),
                "支払方法": item.get("payment_method", "要確認"),
                "インボイス登録番号": item.get("invoice_registration_number", "要確認"),
                "証憑種類": item.get("evidence_type", "要確認"),
                "ステータス": item.get("status", "確認待ち"),
                "信頼度": item.get("confidence", 0),
                "確認事項": " / ".join(item.get("review_notes", [])),
                "元ファイル名": file_name,
            }
        )
    return pd.DataFrame(rows, columns=EXCEL_COLUMNS)


def prepare_closing_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    closing_df = df.copy()
    if closing_df.empty:
        return closing_df

    closing_df["取引日_dt"] = pd.to_datetime(closing_df["取引日"], errors="coerce")
    closing_df["月"] = closing_df["取引日_dt"].dt.strftime("%Y-%m").fillna("日付要確認")
    closing_df["年"] = closing_df["取引日_dt"].dt.strftime("%Y").fillna("日付要確認")

    for column in ["借方金額", "貸方金額", "税込金額", "税抜金額", "消費税額"]:
        closing_df[column] = pd.to_numeric(closing_df[column], errors="coerce").fillna(0).astype(int)

    closing_df["確認フラグ"] = closing_df["確認事項"].astype(str).str.len() > 0
    return closing_df


def build_monthly_closing_report(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    closing_df = prepare_closing_dataframe(df)
    if closing_df.empty:
        empty_summary = pd.DataFrame(columns=["月", "取引件数", "税込合計", "税抜合計", "消費税合計", "確認待ち件数"])
        empty_accounts = pd.DataFrame(columns=["月", "借方勘定科目", "税込合計", "消費税合計", "取引件数"])
        empty_taxes = pd.DataFrame(columns=["月", "借方税区分", "税込合計", "消費税合計", "取引件数"])
        return empty_summary, empty_accounts, empty_taxes

    monthly_summary = (
        closing_df.groupby("月", dropna=False)
        .agg(
            取引件数=("税込金額", "count"),
            税込合計=("税込金額", "sum"),
            税抜合計=("税抜金額", "sum"),
            消費税合計=("消費税額", "sum"),
            確認待ち件数=("確認フラグ", "sum"),
        )
        .reset_index()
        .sort_values("月")
    )

    account_summary = (
        closing_df.groupby(["月", "借方勘定科目"], dropna=False)
        .agg(
            税込合計=("税込金額", "sum"),
            消費税合計=("消費税額", "sum"),
            取引件数=("税込金額", "count"),
        )
        .reset_index()
        .sort_values(["月", "税込合計"], ascending=[True, False])
    )

    tax_summary = (
        closing_df.groupby(["月", "借方税区分"], dropna=False)
        .agg(
            税込合計=("税込金額", "sum"),
            消費税合計=("消費税額", "sum"),
            取引件数=("税込金額", "count"),
        )
        .reset_index()
        .sort_values(["月", "税込合計"], ascending=[True, False])
    )

    return monthly_summary, account_summary, tax_summary


def build_yearly_closing_report(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    closing_df = prepare_closing_dataframe(df)
    if closing_df.empty:
        empty_summary = pd.DataFrame(columns=["年", "取引件数", "税込合計", "税抜合計", "消費税合計", "確認待ち件数"])
        empty_accounts = pd.DataFrame(columns=["年", "借方勘定科目", "税込合計", "消費税合計", "取引件数"])
        empty_pending = pd.DataFrame(columns=EXCEL_COLUMNS)
        return empty_summary, empty_accounts, empty_pending

    yearly_summary = (
        closing_df.groupby("年", dropna=False)
        .agg(
            取引件数=("税込金額", "count"),
            税込合計=("税込金額", "sum"),
            税抜合計=("税抜金額", "sum"),
            消費税合計=("消費税額", "sum"),
            確認待ち件数=("確認フラグ", "sum"),
        )
        .reset_index()
        .sort_values("年")
    )

    account_summary = (
        closing_df.groupby(["年", "借方勘定科目"], dropna=False)
        .agg(
            税込合計=("税込金額", "sum"),
            消費税合計=("消費税額", "sum"),
            取引件数=("税込金額", "count"),
        )
        .reset_index()
        .sort_values(["年", "税込合計"], ascending=[True, False])
    )

    pending_items = closing_df[closing_df["確認フラグ"]].copy()
    pending_items = pending_items[[column for column in EXCEL_COLUMNS if column in pending_items.columns]]
    return yearly_summary, account_summary, pending_items


def build_checklist(items: list[tuple[str, str]], period_label: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "区分": period_label,
                "チェック項目": title,
                "確認内容": description,
                "ステータス": "未確認",
                "メモ": "",
            }
            for title, description in items
        ]
    )


def build_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    monthly_summary, monthly_accounts, monthly_taxes = build_monthly_closing_report(df)
    yearly_summary, yearly_accounts, yearly_pending = build_yearly_closing_report(df)
    monthly_checklist = build_checklist(MONTHLY_CHECKLIST_ITEMS, "月次決算")
    yearly_checklist = build_checklist(YEARLY_CHECKLIST_ITEMS, "年次決算")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="仕訳一覧")
        pending = df[df["確認事項"].astype(str).str.len() > 0].copy()
        pending.to_excel(writer, index=False, sheet_name="確認待ち")
        monthly_summary.to_excel(writer, index=False, sheet_name="月次サマリー")
        monthly_accounts.to_excel(writer, index=False, sheet_name="月次科目別")
        monthly_taxes.to_excel(writer, index=False, sheet_name="月次税区分別")
        yearly_summary.to_excel(writer, index=False, sheet_name="年次サマリー")
        yearly_accounts.to_excel(writer, index=False, sheet_name="年次科目別")
        yearly_pending.to_excel(writer, index=False, sheet_name="年次確認事項")
        monthly_checklist.to_excel(writer, index=False, sheet_name="月次チェックリスト")
        yearly_checklist.to_excel(writer, index=False, sheet_name="年次チェックリスト")
        summary = pd.DataFrame([
            {"項目": "処理日時", "内容": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"項目": "取引件数", "内容": len(df)},
            {"項目": "税込合計", "内容": int(pd.to_numeric(df["税込金額"], errors="coerce").fillna(0).sum())},
            {"項目": "出力内容", "内容": "仕訳一覧、月次決算、年次決算、確認チェックリスト"},
            {"項目": "注意", "内容": "AIによる参考判定です。決算・申告の最終判断は税理士へ確認してください。"},
        ])
        summary.to_excel(writer, index=False, sheet_name="集計")
        workbook = writer.book
        add_cover_sheet(workbook, "基本記帳・決算補助")
        add_financial_summary_sheet(workbook, "基本記帳・決算補助")
        apply_financial_report_format(workbook)
        for sheet in workbook.worksheets:
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 38)
    return output.getvalue()


def estimate_text_transaction_lines(text: str) -> int:
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    return len(lines)


def classify_account(account: str) -> str:
    account = str(account or "")
    asset_keywords = ["現金", "預金", "売掛", "未収", "前払", "仮払", "棚卸", "商品", "建物", "車両", "工具", "備品", "ソフトウェア"]
    liability_keywords = ["買掛", "未払", "未払金", "借入", "預り", "前受", "仮受", "未払消費税"]
    equity_keywords = ["資本金", "元入金", "利益剰余", "繰越利益", "事業主借", "事業主貸"]
    revenue_keywords = ["売上", "収益", "雑収入", "受取利息"]
    expense_keywords = [
        "仕入", "費", "損", "給料", "賃金", "旅費", "交通", "通信", "水道", "光熱", "広告", "接待",
        "会議", "福利", "外注", "手数料", "家賃", "租税", "保険", "消耗", "減価償却", "雑費",
        "支払報酬", "報酬", "荷造運賃", "運賃", "発送", "配送", "送料", "支払",
    ]

    if any(keyword in account for keyword in revenue_keywords):
        return "収益"
    if any(keyword in account for keyword in expense_keywords):
        return "費用"
    if any(keyword in account for keyword in liability_keywords):
        return "負債"
    if any(keyword in account for keyword in equity_keywords):
        return "純資産"
    if any(keyword in account for keyword in asset_keywords):
        return "資産"
    return "未分類"


def read_accounting_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            raw_df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            raw_df = pd.read_csv(uploaded_file, encoding="cp932")
    else:
        raw_df = pd.read_excel(uploaded_file)

    return normalize_uploaded_ledger(raw_df)


def normalize_uploaded_ledger(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty:
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    column_aliases = {
        "取引日": ["取引日", "日付", "仕訳日", "伝票日付", "発生日"],
        "証憑日付": ["証憑日付", "証憑日", "日付", "取引日"],
        "取引先": ["取引先", "相手先", "支払先", "得意先", "摘要取引先"],
        "摘要": ["摘要", "内容", "説明", "メモ", "取引内容"],
        "借方勘定科目": ["借方勘定科目", "借方科目", "借方", "勘定科目", "科目"],
        "借方補助科目": ["借方補助科目", "借方補助", "補助科目"],
        "借方金額": ["借方金額", "借方額", "金額", "税込金額", "支出", "出金"],
        "借方税区分": ["借方税区分", "税区分", "消費税区分"],
        "貸方勘定科目": ["貸方勘定科目", "貸方科目", "貸方", "相手勘定科目"],
        "貸方補助科目": ["貸方補助科目", "貸方補助"],
        "貸方金額": ["貸方金額", "貸方額", "金額", "税込金額", "収入", "入金"],
        "貸方税区分": ["貸方税区分"],
        "税込金額": ["税込金額", "金額", "借方金額", "貸方金額"],
        "税抜金額": ["税抜金額", "本体金額"],
        "消費税率": ["消費税率", "税率"],
        "消費税額": ["消費税額", "税額"],
        "支払方法": ["支払方法", "決済方法", "口座"],
        "インボイス登録番号": ["インボイス登録番号", "登録番号"],
        "証憑種類": ["証憑種類", "証憑", "書類種別"],
        "ステータス": ["ステータス"],
        "信頼度": ["信頼度"],
        "確認事項": ["確認事項", "備考", "メモ"],
        "元ファイル名": ["元ファイル名"],
    }

    normalized = pd.DataFrame()
    for target, aliases in column_aliases.items():
        source = next((column for column in aliases if column in raw_df.columns), None)
        normalized[target] = raw_df[source] if source else ""

    for column in ["借方金額", "貸方金額", "税込金額", "税抜金額", "消費税額"]:
        normalized[column] = pd.to_numeric(normalized[column], errors="coerce").fillna(0).astype(int)

    normalized["取引日"] = normalized["取引日"].replace("", "要確認")
    missing_evidence_date = normalized["証憑日付"].astype(str).str.len() == 0
    normalized.loc[missing_evidence_date, "証憑日付"] = normalized.loc[missing_evidence_date, "取引日"]
    normalized["ステータス"] = normalized["ステータス"].replace("", "取込データ")
    normalized["確認事項"] = normalized["確認事項"].fillna("")

    missing_gross = normalized["税込金額"] == 0
    normalized.loc[missing_gross, "税込金額"] = normalized.loc[missing_gross, ["借方金額", "貸方金額"]].max(axis=1)

    missing_debit = normalized["借方金額"] == 0
    normalized.loc[missing_debit, "借方金額"] = normalized.loc[missing_debit, "税込金額"]
    missing_credit = normalized["貸方金額"] == 0
    normalized.loc[missing_credit, "貸方金額"] = normalized.loc[missing_credit, "税込金額"]

    return normalized[EXCEL_COLUMNS]


def get_customer_ledger_path(customer: dict) -> Path:
    username = str(customer.get("username", "anonymous"))
    digest = hashlib.sha256(username.encode("utf-8")).hexdigest()[:24]
    return LEDGER_STORAGE_DIR / f"{digest}.json"


def load_customer_ledger(customer: dict) -> pd.DataFrame:
    path = get_customer_ledger_path(customer)
    if not path.exists():
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    if not isinstance(data, list):
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    df = pd.DataFrame(data)
    for column in EXCEL_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    return df[EXCEL_COLUMNS]


def save_customer_ledger(customer: dict, df: pd.DataFrame) -> None:
    LEDGER_STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    clean_df = normalize_uploaded_ledger(df)
    path = get_customer_ledger_path(customer)
    path.write_text(clean_df.to_json(orient="records", force_ascii=False, indent=2), encoding="utf-8")


def append_customer_ledger(customer: dict, df: pd.DataFrame) -> int:
    if df.empty:
        return len(load_customer_ledger(customer))

    existing_df = load_customer_ledger(customer)
    merged_df = pd.concat([existing_df, normalize_uploaded_ledger(df)], ignore_index=True)
    save_customer_ledger(customer, merged_df)
    return len(merged_df)


def clear_customer_ledger(customer: dict) -> None:
    path = get_customer_ledger_path(customer)
    if path.exists():
        path.unlink()


def build_trial_balance(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["区分", "勘定科目", "借方合計", "貸方合計", "残高"])

    debit = df.groupby("借方勘定科目", dropna=False)["借方金額"].sum().rename("借方合計")
    credit = df.groupby("貸方勘定科目", dropna=False)["貸方金額"].sum().rename("貸方合計")
    trial = pd.concat([debit, credit], axis=1).fillna(0).reset_index().rename(columns={"index": "勘定科目"})
    trial = trial[trial["勘定科目"].astype(str).str.len() > 0]
    trial["借方合計"] = trial["借方合計"].astype(int)
    trial["貸方合計"] = trial["貸方合計"].astype(int)
    trial["残高"] = trial["借方合計"] - trial["貸方合計"]
    trial["区分"] = trial["勘定科目"].apply(classify_account)
    return trial[["区分", "勘定科目", "借方合計", "貸方合計", "残高"]].sort_values(["区分", "勘定科目"])


def build_profit_and_loss(trial_balance: pd.DataFrame) -> pd.DataFrame:
    rows = []
    revenue = trial_balance[trial_balance["区分"] == "収益"].copy()
    expense = trial_balance[trial_balance["区分"] == "費用"].copy()

    for _, item in revenue.iterrows():
        rows.append({"区分": "収益", "科目": item["勘定科目"], "金額": int(-item["残高"])})
    for _, item in expense.iterrows():
        rows.append({"区分": "費用", "科目": item["勘定科目"], "金額": int(item["残高"])})

    statement = pd.DataFrame(rows, columns=["区分", "科目", "金額"])
    revenue_total = int(statement[statement["区分"] == "収益"]["金額"].sum()) if not statement.empty else 0
    expense_total = int(statement[statement["区分"] == "費用"]["金額"].sum()) if not statement.empty else 0
    net_income = revenue_total - expense_total
    totals = pd.DataFrame(
        [
            {"区分": "合計", "科目": "収益合計", "金額": revenue_total},
            {"区分": "合計", "科目": "費用合計", "金額": expense_total},
            {"区分": "合計", "科目": "当期利益", "金額": net_income},
        ]
    )
    return pd.concat([statement, totals], ignore_index=True)


def build_balance_sheet(trial_balance: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for category in ["資産", "負債", "純資産", "未分類"]:
        category_df = trial_balance[trial_balance["区分"] == category]
        for _, item in category_df.iterrows():
            amount = int(item["残高"])
            if category in ["負債", "純資産"]:
                amount = -amount
            rows.append({"区分": category, "科目": item["勘定科目"], "金額": amount})

    statement = pd.DataFrame(rows, columns=["区分", "科目", "金額"])
    totals = []
    for category in ["資産", "負債", "純資産", "未分類"]:
        total = int(statement[statement["区分"] == category]["金額"].sum()) if not statement.empty else 0
        totals.append({"区分": "合計", "科目": f"{category}合計", "金額": total})
    return pd.concat([statement, pd.DataFrame(totals)], ignore_index=True)


def style_excel_sheet(sheet, freeze: str = "A2") -> None:
    header_fill = PatternFill("solid", fgColor="0F172A")
    header_font = Font(color="FFFFFF", bold=True)
    border_color = "CBD5E1"
    thin_border = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color),
    )

    sheet.sheet_view.showGridLines = False
    if freeze:
        sheet.freeze_panes = freeze

    max_row = sheet.max_row or 1
    max_column = sheet.max_column or 1

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row in sheet.iter_rows(min_row=2, max_row=max_row, max_col=max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0;[Red]-#,##0;-'

    for column_idx in range(1, max_column + 1):
        column_letter = get_column_letter(column_idx)
        max_length = 0
        for cell in sheet[column_letter]:
            max_length = max(max_length, len(str(cell.value or "")))
        sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 42)


def add_cover_sheet(workbook, period_label: str) -> None:
    sheet = workbook.create_sheet("表紙", 0)
    sheet.sheet_view.showGridLines = False
    sheet["B2"] = "財務報告書"
    sheet["B3"] = period_label
    sheet["B5"] = "作成日時"
    sheet["C5"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet["B7"] = "含まれる帳票"
    reports = ["財務サマリー", "試算表", "損益計算書", "貸借対照表", "チェックリスト", "取込仕訳"]
    for idx, report in enumerate(reports, start=8):
        sheet[f"B{idx}"] = report
    sheet["B16"] = "注意"
    sheet["C16"] = "AIとルールによる参考資料です。決算・申告の最終判断は税理士へ確認してください。"

    sheet["B2"].font = Font(size=22, bold=True, color="0F172A")
    sheet["B3"].font = Font(size=14, bold=True, color="2563EB")
    for cell_ref in ["B5", "B7", "B16"]:
        sheet[cell_ref].font = Font(bold=True, color="334155")
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 72


def add_financial_summary_sheet(workbook, period_label: str) -> None:
    sheet = workbook.create_sheet("財務サマリー", 1)
    sheet.sheet_view.showGridLines = False
    sheet["B2"] = "財務サマリー"
    sheet["C2"] = period_label

    rows = [
        ("売上高", '=SUMIF(損益計算書!A:A,"収益",損益計算書!C:C)'),
        ("費用合計", '=SUMIF(損益計算書!B:B,"費用合計",損益計算書!C:C)'),
        ("当期利益", '=SUMIF(損益計算書!B:B,"当期利益",損益計算書!C:C)'),
        ("資産合計", '=SUMIF(貸借対照表!B:B,"資産合計",貸借対照表!C:C)'),
        ("負債合計", '=SUMIF(貸借対照表!B:B,"負債合計",貸借対照表!C:C)'),
        ("純資産合計", '=SUMIF(貸借対照表!B:B,"純資産合計",貸借対照表!C:C)'),
        ("B/S差額チェック", '=C6-C7-C8'),
    ]
    sheet["B4"] = "指標"
    sheet["C4"] = "金額"
    for row_idx, (label, formula) in enumerate(rows, start=5):
        sheet[f"B{row_idx}"] = label
        sheet[f"C{row_idx}"] = formula

    sheet["E4"] = "確認ポイント"
    notes = [
        "B/S差額チェックが0に近いか確認",
        "未分類科目がある場合は科目分類を確認",
        "確認事項シートと原本証憑を照合",
        "税区分・インボイス登録番号は税理士確認",
    ]
    for row_idx, note in enumerate(notes, start=5):
        sheet[f"E{row_idx}"] = note

    sheet["B2"].font = Font(size=18, bold=True, color="0F172A")
    sheet["C2"].font = Font(size=12, bold=True, color="2563EB")
    for cell_ref in ["B4", "C4", "E4"]:
        sheet[cell_ref].fill = PatternFill("solid", fgColor="0F172A")
        sheet[cell_ref].font = Font(color="FFFFFF", bold=True)
    for row in range(5, 12):
        sheet[f"C{row}"].number_format = '#,##0;[Red]-#,##0;-'
    sheet.column_dimensions["B"].width = 22
    sheet.column_dimensions["C"].width = 18
    sheet.column_dimensions["E"].width = 48


def apply_financial_report_format(workbook) -> None:
    for sheet in workbook.worksheets:
        if sheet.title not in ["表紙", "財務サマリー"]:
            style_excel_sheet(sheet)

    for sheet_name in ["損益計算書", "貸借対照表", "試算表"]:
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows(min_row=2):
            first_value = str(row[0].value or "")
            second_value = str(row[1].value or "") if len(row) > 1 else ""
            if first_value == "合計" or second_value.endswith("合計") or second_value == "当期利益":
                for cell in row:
                    cell.font = Font(bold=True, color="0F172A")
                    cell.fill = PatternFill("solid", fgColor="E2E8F0")


def build_financial_statement_excel(
    ledger_df: pd.DataFrame,
    trial_balance: pd.DataFrame,
    profit_and_loss: pd.DataFrame,
    balance_sheet: pd.DataFrame,
    period_label: str,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ledger_df.to_excel(writer, index=False, sheet_name="取込仕訳")
        trial_balance.to_excel(writer, index=False, sheet_name="試算表")
        profit_and_loss.to_excel(writer, index=False, sheet_name="損益計算書")
        balance_sheet.to_excel(writer, index=False, sheet_name="貸借対照表")
        checklist_items = MONTHLY_CHECKLIST_ITEMS if period_label == "月次決算" else YEARLY_CHECKLIST_ITEMS
        build_checklist(checklist_items, period_label).to_excel(writer, index=False, sheet_name="チェックリスト")

        workbook = writer.book
        add_cover_sheet(workbook, period_label)
        add_financial_summary_sheet(workbook, period_label)
        apply_financial_report_format(workbook)

        for sheet in workbook.worksheets:
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 38)
    return output.getvalue()


def render_bookkeeping_workspace(customer: dict, model: str, transaction_limit: int) -> None:
    overview_cols = st.columns(4)
    overview_cols[0].metric("現在のプラン", customer["plan_label"])
    overview_cols[1].metric("処理上限", f"{transaction_limit} 取引/回")
    overview_cols[2].metric("今回の結果", f"{len(st.session_state.get('df', []))} 件")
    overview_cols[3].metric("出力形式", "Excel")

    left, right = st.columns([0.82, 1.18], gap="large")

    with left:
        st.markdown(
            """
            <div class="app-panel">
                <div class="eyebrow">BASIC BOOKKEEPING</div>
                <div class="panel-title">証憑または取引内容から記帳</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        input_mode = st.radio("入力方式", ["画像アップロード", "テキスト入力"], horizontal=True)
        uploaded_files = []
        text_input = ""

        if input_mode == "画像アップロード":
            uploaded_files = st.file_uploader(
                "領収書・請求書・明細画像",
                type=["png", "jpg", "jpeg", "webp"],
                accept_multiple_files=True,
                key="bookkeeping_images",
            )
            if uploaded_files:
                if len(uploaded_files) > transaction_limit:
                    st.error(f"現在のプランでは一度に処理できる証憑画像は{transaction_limit}件までです。")
                    st.stop()
                st.success(f"{len(uploaded_files)}件の証憑をアップロードしました")
                preview_cols = st.columns(2)
                for idx, file in enumerate(uploaded_files[:4]):
                    preview_cols[idx % 2].image(file, caption=file.name, use_container_width=True)
                if len(uploaded_files) > 4:
                    st.info(f"ほか {len(uploaded_files) - 4} 件のプレビューを省略しています。")
        else:
            text_input = st.text_area(
                "取引内容",
                height=210,
                placeholder=PLAN_TEXT_EXAMPLES.get(customer["plan_key"], PLAN_TEXT_EXAMPLES["free"]),
                key="bookkeeping_text",
            )

        run = st.button("仕訳候補を作成", type="primary", width="stretch")

    if run:
        if not uploaded_files and not text_input.strip():
            st.warning("画像をアップロードするか、取引内容を入力してください。")
            return
        if not uploaded_files:
            estimated_lines = estimate_text_transaction_lines(text_input)
            if estimated_lines > transaction_limit:
                st.error(
                    f"現在のプランでは一度に処理できる取引は{transaction_limit}件までです。"
                    f"入力は{estimated_lines}行あります。取引数を減らすか、上位プランをご利用ください。"
                )
                return
        client = get_client()
        with st.spinner("AI Visionで証憑を解析しています..."):
            try:
                all_results = []
                all_dfs = []
                if uploaded_files:
                    progress = st.progress(0)
                    for idx, file in enumerate(uploaded_files, start=1):
                        result = analyze_with_ai_vision(client, model, "", file)
                        df = transactions_to_dataframe(result, file.name)
                        all_results.append({"file_name": file.name, "result": result})
                        all_dfs.append(df)
                        progress.progress(idx / len(uploaded_files))
                else:
                    result = analyze_with_ai_vision(client, model, text_input, None)
                    transaction_count = len(result.get("transactions", []))
                    if transaction_count > transaction_limit:
                        st.error(
                            f"現在のプランでは一度に処理できる取引は{transaction_limit}件までです。"
                            f"入力内容は{transaction_count}件として解析されました。取引数を減らすか、上位プランをご利用ください。"
                        )
                        return
                    df = transactions_to_dataframe(result, "テキスト入力")
                    all_results.append({"file_name": "テキスト入力", "result": result})
                    all_dfs.append(df)
                merged_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame(columns=EXCEL_COLUMNS)
                if len(merged_df) > transaction_limit:
                    st.error(
                        f"現在のプランでは一度に処理できる取引は{transaction_limit}件までです。"
                        f"解析結果は{len(merged_df)}件でした。証憑または取引数を減らすか、上位プランをご利用ください。"
                    )
                    return
                st.session_state["result"] = all_results
                st.session_state["df"] = merged_df
            except Exception as exc:
                st.error(f"エラーが発生しました：{exc}")
                return

    with right:
        if "df" not in st.session_state:
            st.markdown(
                """
                <div class="app-panel-muted">
                    <div class="eyebrow">RESULT</div>
                    <div class="panel-title">仕訳候補はここに表示されます</div>
                    <p class="panel-copy">証憑画像または取引内容を入力して、基本記帳を始めてください。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            return
        df = st.session_state["df"]
        total_amount = int(pd.to_numeric(df["税込金額"], errors="coerce").fillna(0).sum()) if not df.empty else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("取引件数", f"{len(df)} 件")
        c2.metric("税込合計", f"{total_amount:,} 円")
        c3.metric("確認待ち", f"{len(df[df['ステータス'].astype(str).str.contains('確認', na=False)])} 件")

        edited_df = st.data_editor(df, num_rows="fixed", width="stretch", hide_index=True)
        st.session_state["df"] = edited_df

        saved_ledger = load_customer_ledger(customer)
        s1, s2 = st.columns(2)
        s1.metric("保存済み帳簿", f"{len(saved_ledger)} 件")
        if s2.button("この仕訳を保存済み帳簿に追加", width="stretch"):
            total_saved = append_customer_ledger(customer, st.session_state["df"])
            st.success(f"保存しました。保存済み帳簿は合計 {total_saved} 件です。月次決算・年次決算で利用できます。")

        with st.expander("保存済み帳簿の管理"):
            saved_ledger = load_customer_ledger(customer)
            st.dataframe(saved_ledger, width="stretch", hide_index=True)
            ledger_excel = build_excel(saved_ledger) if not saved_ledger.empty else b""
            if not saved_ledger.empty:
                st.download_button(
                    "保存済み帳簿をExcelでダウンロード",
                    data=ledger_excel,
                    file_name=f"saved_ledger_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
            if st.button("保存済み帳簿をクリア", width="stretch"):
                clear_customer_ledger(customer)
                st.success("保存済み帳簿をクリアしました。")
                st.rerun()

        excel_bytes = build_excel(st.session_state["df"])
        filename = f"bookkeeping_entries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "記帳・決算Excelをダウンロード",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )


def render_closing_workspace(period_label: str, transaction_limit: int, customer: dict) -> None:
    st.markdown(
        f"""
        <div class="app-panel">
            <div class="eyebrow">{period_label.upper()}</div>
            <div class="panel-title">顧客の帳簿データから決算表・財務諸表を作成</div>
            <p class="panel-copy">CSV / Excel の仕訳帳をアップロードするか、基本記帳で保存した帳簿を利用できます。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    source = st.radio("データ元", ["帳簿ファイルをアップロード", "保存済み帳簿を使う"], horizontal=True, key=f"{period_label}_source")
    ledger_df = pd.DataFrame(columns=EXCEL_COLUMNS)

    if source == "帳簿ファイルをアップロード":
        uploaded_files = st.file_uploader(
            "仕訳帳 CSV / Excel（複数ファイル可）",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
            key=f"{period_label}_ledger_upload",
        )
        if not uploaded_files:
            st.info("顧客の仕訳帳ファイルをアップロードしてください。年次決算では月別ファイルを複数選択できます。")
            return
        try:
            ledger_dfs = []
            for uploaded_file in uploaded_files:
                uploaded_df = read_accounting_upload(uploaded_file)
                if not uploaded_df.empty:
                    uploaded_df["元ファイル名"] = uploaded_file.name
                    ledger_dfs.append(uploaded_df)
            ledger_df = pd.concat(ledger_dfs, ignore_index=True) if ledger_dfs else pd.DataFrame(columns=EXCEL_COLUMNS)
            st.success(f"{len(uploaded_files)}ファイルを読み込みました。合計 {len(ledger_df)} 件の仕訳を集計します。")
        except Exception as exc:
            st.error(f"ファイルを読み取れませんでした：{exc}")
            return
    else:
        ledger_df = load_customer_ledger(customer)
        if ledger_df.empty:
            st.info("保存済み帳簿がありません。基本記帳で仕訳を保存するか、帳簿ファイルをアップロードしてください。")
            return
        st.success(f"保存済み帳簿 {len(ledger_df)} 件を読み込みました。")

    if ledger_df.empty:
        st.warning("決算表を作成できるデータがありません。")
        return
    if len(ledger_df) > transaction_limit and period_label == "月次決算":
        st.warning(f"現在のプランの1回処理目安は{transaction_limit}取引です。大きな帳簿は上位プランをご検討ください。")

    prepared_df = prepare_closing_dataframe(ledger_df)
    selectable_periods = sorted(prepared_df["月"].dropna().unique()) if period_label == "月次決算" else sorted(prepared_df["年"].dropna().unique())
    if selectable_periods:
        selected_period = st.selectbox("対象期間", selectable_periods, index=len(selectable_periods) - 1)
        if period_label == "月次決算":
            ledger_df = prepared_df[prepared_df["月"] == selected_period][EXCEL_COLUMNS].copy()
        else:
            ledger_df = prepared_df[prepared_df["年"] == selected_period][EXCEL_COLUMNS].copy()

    trial_balance = build_trial_balance(ledger_df)
    profit_and_loss = build_profit_and_loss(trial_balance)
    balance_sheet = build_balance_sheet(trial_balance)

    total_debit = int(trial_balance["借方合計"].sum()) if not trial_balance.empty else 0
    total_credit = int(trial_balance["貸方合計"].sum()) if not trial_balance.empty else 0
    net_income_row = profit_and_loss[profit_and_loss["科目"] == "当期利益"]
    net_income = int(net_income_row["金額"].iloc[0]) if not net_income_row.empty else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("取引件数", f"{len(ledger_df)} 件")
    m2.metric("借方合計", f"{total_debit:,} 円")
    m3.metric("貸方合計", f"{total_credit:,} 円")
    m4.metric("当期利益", f"{net_income:,} 円")

    trial_tab, pl_tab, bs_tab, check_tab = st.tabs(["試算表", "損益計算書", "貸借対照表", "チェックリスト"])
    with trial_tab:
        st.dataframe(trial_balance, width="stretch", hide_index=True)
    with pl_tab:
        st.dataframe(profit_and_loss, width="stretch", hide_index=True)
    with bs_tab:
        st.dataframe(balance_sheet, width="stretch", hide_index=True)
    with check_tab:
        checklist = build_checklist(MONTHLY_CHECKLIST_ITEMS if period_label == "月次決算" else YEARLY_CHECKLIST_ITEMS, period_label)
        st.dataframe(checklist, width="stretch", hide_index=True)

    excel_bytes = build_financial_statement_excel(ledger_df, trial_balance, profit_and_loss, balance_sheet, period_label)
    filename = f"{period_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "決算表・財務諸表Excelをダウンロード",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    st.markdown(PAGE_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="app-header">
            <div class="app-title">AI仕訳アシスタント</div>
            <p class="app-subtitle">日本の中小企業・個人事業主向け。証憑画像または取引メモから仕訳候補を作成します。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    customer = get_logged_in_customer()
    if not customer:
        render_login()

    with st.sidebar:
        st.header("アカウント")
        st.markdown(f"**{customer['name']}**")
        st.caption(customer["username"])
        if st.button("ログアウト", width="stretch"):
            logout_customer()

        st.divider()
        model = DEFAULT_MODEL
        transaction_limit = int(customer["transaction_limit"])
        st.markdown(f"<span class=\"plan-pill\">{customer['plan_label']}</span>", unsafe_allow_html=True)
        st.caption(f"上限：{transaction_limit} 取引/回")
        st.caption(f"AIモデル：{model}")
        st.divider()
        st.info(get_config_value("UPGRADE_CONTACT", "上位プランをご希望の場合は管理者までお問い合わせください。"))

    bookkeeping_tab, monthly_tab, yearly_tab = st.tabs(["基本記帳", "月次決算", "年次決算"])
    with bookkeeping_tab:
        render_bookkeeping_workspace(customer, model, transaction_limit)
    with monthly_tab:
        render_closing_workspace("月次決算", transaction_limit, customer)
    with yearly_tab:
        render_closing_workspace("年次決算", transaction_limit, customer)


if __name__ == "__main__":
    main()
