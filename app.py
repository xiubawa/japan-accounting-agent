import base64
import io
import json
import os
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image

load_dotenv()

APP_TITLE = "AI仕訳アシスタント（日本の中小企業・個人事業主向け）"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")
PLAN_LIMITS = {
    "無料プラン（1件/回）": 1,
    "スタータープラン（5件/回）": 5,
    "ビジネスプラン（20件/回）": 20,
    "会計事務所プラン（50件/回）": 50,
}

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


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("証憑画像または取引内容入力 → AI Vision解析 → 仕訳候補生成 → 確認 → Excel出力")

    with st.sidebar:
        st.header("設定")
        st.success("現在モード：AI Vision API版")
        model = st.text_input("OpenAI Model", value=DEFAULT_MODEL)
        plan_name = st.selectbox("料金プラン", options=list(PLAN_LIMITS.keys()), index=1)
        max_files = PLAN_LIMITS[plan_name]
        st.caption(f"このプランでは一度に最大 {max_files} 件まで処理できます。")
        st.warning("AIによる参考判定です。最終的な会計・税務判断は税理士へ確認してください。")

    left, right = st.columns([0.85, 1.15])

    with left:
        uploaded_files = st.file_uploader(
            "領収書・請求書・明細画像をアップロード",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
        )
        if uploaded_files:
            if len(uploaded_files) > max_files:
                st.error(f"一度に処理できるのは{max_files}件までです。")
                st.stop()
            st.success(f"{len(uploaded_files)}件の証憑をアップロードしました")
            for file in uploaded_files[:5]:
                st.image(file, caption=file.name, width=260)
            if len(uploaded_files) > 5:
                st.info(f"ほか {len(uploaded_files) - 5} 件のプレビューを省略しています。")

        text_input = st.text_area(
            "取引内容を入力",
            height=160,
            placeholder="例：2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。キーボードとマウスを購入。消費税10%、適格請求書取得済み。",
        )
        run = st.button("AI仕訳生成", type="primary", width="stretch")

    if run:
        if not uploaded_files and not text_input.strip():
            st.warning("画像をアップロードするか、取引内容を入力してください。")
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
                    df = transactions_to_dataframe(result, "テキスト入力")
                    all_results.append({"file_name": "テキスト入力", "result": result})
                    all_dfs.append(df)
                merged_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame(columns=EXCEL_COLUMNS)
                st.session_state["result"] = all_results
                st.session_state["df"] = merged_df
            except Exception as exc:
                st.error(f"エラーが発生しました：{exc}")
                return

    with right:
        if "df" not in st.session_state:
            st.subheader("入力待機中")
            st.write("証憑画像または取引内容を入力し、『AI仕訳生成』をクリックしてください。")
            return
        df = st.session_state["df"]
        st.subheader("処理結果")
        total_amount = int(pd.to_numeric(df["税込金額"], errors="coerce").fillna(0).sum()) if not df.empty else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("取引件数", f"{len(df)} 件")
        c2.metric("税込合計", f"{total_amount:,} 円")
        c3.metric("確認待ち", f"{len(df[df['ステータス'].astype(str).str.contains('確認', na=False)])} 件")
        st.subheader("仕訳一覧")
        edited_df = st.data_editor(df, num_rows="fixed", width="stretch", hide_index=True)
        st.session_state["df"] = edited_df
        excel_bytes = build_excel(edited_df)
        filename = f"ai_accounting_entries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "Excel出力",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        with st.expander("AI JSONデータ表示"):
            st.json(st.session_state.get("result", {}))


if __name__ == "__main__":
    main()
