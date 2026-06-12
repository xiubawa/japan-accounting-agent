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
from openai import OpenAI
from PIL import Image

load_dotenv()

APP_TITLE = "AI仕訳アシスタント（日本の中小企業・個人事業主向け）"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")
CUSTOMER_DB_PATH = Path(os.getenv("CUSTOMER_DB_PATH", "customer_accounts.json"))
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
    .block-container {
        max-width: 1180px;
        padding-top: 2.2rem;
        padding-bottom: 3rem;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 0.85rem 1rem;
    }
    div[data-testid="stMetric"] label {
        color: #64748b;
    }
    section[data-testid="stSidebar"] {
        border-right: 1px solid #e5e7eb;
    }
    .app-panel {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 1rem 1.1rem;
        background: #ffffff;
        margin-bottom: 1rem;
    }
    .app-panel-muted {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 1rem 1.1rem;
        background: #f8fafc;
        margin-bottom: 1rem;
    }
    .eyebrow {
        color: #64748b;
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0;
        margin-bottom: 0.25rem;
    }
    .panel-title {
        color: #0f172a;
        font-size: 1.08rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .panel-copy {
        color: #475569;
        font-size: 0.94rem;
        line-height: 1.55;
        margin: 0;
    }
    .plan-pill {
        display: inline-block;
        border: 1px solid #cbd5e1;
        border-radius: 999px;
        padding: 0.2rem 0.55rem;
        color: #334155;
        background: #f8fafc;
        font-size: 0.82rem;
        font-weight: 600;
    }
</style>
"""

EXCEL_COLUMNS = [
    "取引日", "証憑日付", "取引先", "摘要", "借方勘定科目", "借方補助科目", "借方金額", "借方税区分",
    "貸方勘定科目", "貸方補助科目", "貸方金額", "貸方税区分", "税込金額", "税抜金額", "消費税率",
    "消費税額", "支払方法", "インボイス登録番号", "証憑種類", "ステータス", "信頼度", "確認事項", "元ファイル名",
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

    left, right = st.columns([0.95, 1.05], gap="large")

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


def build_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="仕訳一覧")
        pending = df[df["確認事項"].astype(str).str.len() > 0].copy()
        pending.to_excel(writer, index=False, sheet_name="確認待ち")
        summary = pd.DataFrame([
            {"項目": "処理日時", "内容": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"項目": "取引件数", "内容": len(df)},
            {"項目": "税込合計", "内容": int(pd.to_numeric(df["税込金額"], errors="coerce").fillna(0).sum())},
            {"項目": "注意", "内容": "AIによる参考判定です。最終判断は税理士へ確認してください。"},
        ])
        summary.to_excel(writer, index=False, sheet_name="集計")
        workbook = writer.book
        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 38)
    return output.getvalue()


def estimate_text_transaction_lines(text: str) -> int:
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    return len(lines)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    st.markdown(PAGE_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE)
    st.caption("証憑画像または取引メモから、日本会計向けの仕訳候補を作成します。")

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
                <div class="eyebrow">INPUT</div>
                <div class="panel-title">証憑または取引内容</div>
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
                    <div class="panel-title">結果はここに表示されます</div>
                    <p class="panel-copy">証憑画像または取引内容を入力して、仕訳候補を作成してください。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            return
        df = st.session_state["df"]
        total_amount = int(pd.to_numeric(df["税込金額"], errors="coerce").fillna(0).sum()) if not df.empty else 0
        st.markdown(
            """
            <div class="app-panel">
                <div class="eyebrow">RESULT</div>
                <div class="panel-title">仕訳候補</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns(3)
        c1.metric("取引件数", f"{len(df)} 件")
        c2.metric("税込合計", f"{total_amount:,} 円")
        c3.metric("確認待ち", f"{len(df[df['ステータス'].astype(str).str.contains('確認', na=False)])} 件")

        edited_df = st.data_editor(df, num_rows="fixed", width="stretch", hide_index=True)
        st.session_state["df"] = edited_df
        excel_bytes = build_excel(edited_df)
        filename = f"ai_accounting_entries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with st.expander("出力", expanded=True):
            st.download_button(
                "Excelをダウンロード",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
        with st.expander("AI JSONデータ"):
            st.json(st.session_state.get("result", {}))


if __name__ == "__main__":
    main()
