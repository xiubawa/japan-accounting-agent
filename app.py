import io
import os
import re
import tempfile
from datetime import datetime

import easyocr
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from PIL import Image


load_dotenv()

APP_TITLE = "AI経理アシスタント（OCRオフライン版）"

EXCEL_COLUMNS = [
    "取引日",
    "証憑日付",
    "取引先",
    "摘要",
    "借方勘定科目",
    "借方補助科目",
    "借方金額",
    "借方税区分",
    "貸方勘定科目",
    "貸方補助科目",
    "貸方金額",
    "貸方税区分",
    "税込金額",
    "税抜金額",
    "消費税率",
    "消費税額",
    "支払方法",
    "インボイス登録番号",
    "証憑種類",
    "ステータス",
    "信頼度",
    "確認事項",
]


@st.cache_resource
def get_ocr_reader():
    # 初回起動時はEasyOCRモデルのダウンロードに数分かかる場合があります。
    return easyocr.Reader(["ja", "en"], gpu=False)


def extract_text_from_image(uploaded_file) -> str:
    if uploaded_file is None:
        return ""

    image = Image.open(uploaded_file).convert("RGB")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
        image.save(tmp.name)
        tmp_path = tmp.name

    try:
        reader = get_ocr_reader()
        results = reader.readtext(tmp_path, detail=0)
        return "\n".join(str(x) for x in results)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def normalize_text(text: str) -> str:
    return (text or "").replace("￥", "¥").replace(",", "")


def extract_amount(text: str) -> int:
    text = normalize_text(text)

    # 「合計」「お支払」「TOTAL」などの近くにある金額を優先して抽出します。
    patterns = [
        r"(?:合計|総合計|税込合計|お支払|支払|total|TOTAL)[^\d¥￥]{0,10}[¥￥]?\s*(\d{2,8})",
        r"[¥￥]\s*(\d{2,8})",
    ]

    for pattern in patterns:
        matches = re.findall(pattern, text, flags=re.IGNORECASE)
        if matches:
            nums = [int(x) for x in matches if x.isdigit()]
            if nums:
                return max(nums)

    # 最後の補助判定：本文中の大きめの数字を金額候補として扱います。
    nums = [int(x) for x in re.findall(r"\b\d{2,8}\b", text)]
    nums = [n for n in nums if n > 50]
    return max(nums) if nums else 0


def extract_date(text: str) -> str:
    text = normalize_text(text)

    patterns = [
        r"(20\d{2})[/-年.](\d{1,2})[/-月.](\d{1,2})",
        r"(\d{4})(\d{2})(\d{2})",
    ]

    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            y, mo, d = m.groups()
            try:
                return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
            except ValueError:
                pass

    return datetime.now().strftime("%Y-%m-%d")


def guess_vendor_and_account(text: str):
    t = text.lower()

    rules = [
        (
            ["amazon", "アマゾン"],
            "Amazon Japan",
            "消耗品費",
            "事務用品・PC周辺機器は通常『消耗品費』として処理します。",
        ),
        (
            ["ローソン", "lawson"],
            "ローソン",
            "福利厚生費",
            "コンビニでの購入が事業経費に該当するか確認が必要です。",
        ),
        (
            ["セブン", "7-eleven", "seven"],
            "セブン-イレブン",
            "福利厚生費",
            "コンビニでの購入が事業経費に該当するか確認が必要です。",
        ),
        (
            ["ファミリーマート", "familymart"],
            "ファミリーマート",
            "福利厚生費",
            "コンビニでの購入が事業経費に該当するか確認が必要です。",
        ),
        (
            ["jr", "電車", "きっぷ", "切符", "suica", "pasmo"],
            "交通機関",
            "旅費交通費",
            "交通費の業務目的を確認してください。",
        ),
        (
            ["タクシー", "taxi"],
            "タクシー",
            "旅費交通費",
            "タクシー利用の目的を確認してください。",
        ),
        (
            ["ガソリン", "eneos", "出光", "shell"],
            "ガソリンスタンド",
            "旅費交通費",
            "車両の業務利用状況を確認してください。",
        ),
        (
            ["会食", "接待", "居酒屋", "レストラン", "飲食"],
            "飲食店",
            "接待交際費",
            "参加者・目的・人数の確認が必要です。",
        ),
        (
            ["会議", "打合せ", "カフェ", "喫茶"],
            "飲食店",
            "会議費",
            "会議関連支出に該当するか確認してください。",
        ),
        (
            ["家賃", "賃料", "オフィス"],
            "賃貸先",
            "地代家賃",
            "住宅用・事業用の区分および消費税区分を確認してください。",
        ),
        (
            ["電気", "ガス", "水道"],
            "公共料金",
            "水道光熱費",
            "事業利用割合を確認してください。",
        ),
        (
            ["通信", "携帯", "docomo", "au", "softbank"],
            "通信会社",
            "通信費",
            "事業利用割合を確認してください。",
        ),
    ]

    for keywords, vendor, account, note in rules:
        if any(k.lower() in t for k in keywords):
            return vendor, account, note

    return "要確認", "雑費", "OCR読取結果が不十分なため、勘定科目の確認が必要です。"


def guess_payment_method(text: str) -> str:
    t = text.lower()
    if "visa" in t or "master" in t or "カード" in t or "credit" in t or "クレジット" in t:
        return "クレジットカード"
    if "現金" in t or "cash" in t:
        return "現金"
    if "paypay" in t or "交通系" in t or "suica" in t or "pasmo" in t:
        return "電子マネー"
    return "要確認"


def make_transaction(text_input: str, uploaded_file) -> dict:
    ocr_text = extract_text_from_image(uploaded_file)
    full_text = f"{text_input or ''}\n{ocr_text or ''}".strip()

    if not full_text:
        full_text = "要確認"

    date = extract_date(full_text)
    amount = extract_amount(full_text)
    vendor, debit_account, note = guess_vendor_and_account(full_text)
    payment_method = guess_payment_method(full_text)

    tax_rate = "10%"
    tax_amount = round(amount / 11) if amount else 0
    net_amount = amount - tax_amount if amount else 0

    credit_account = "要確認"
    if payment_method == "クレジットカード":
        credit_account = "未払金"
    elif payment_method == "現金":
        credit_account = "現金"

    result = {
        "transactions": [
            {
                "transaction_date": date,
                "evidence_date": date,
                "vendor": vendor,
                "description": "OCR・取引内容入力から自動生成",
                "gross_amount": amount,
                "net_amount": net_amount,
                "tax_amount": tax_amount,
                "tax_rate": tax_rate if amount else "要確認",
                "debit_account": debit_account,
                "debit_sub_account": "",
                "debit_amount": amount,
                "debit_tax_category": "課税仕入10%" if amount else "要確認",
                "credit_account": credit_account,
                "credit_sub_account": "",
                "credit_amount": amount,
                "credit_tax_category": "対象外",
                "payment_method": payment_method,
                "invoice_registration_number": "要確認",
                "evidence_type": "レシート" if uploaded_file is not None else "テキスト",
                "status": "確認待ち",
                "confidence": 0.65 if ocr_text else 0.45,
                "review_notes": [
                    note,
                    "本システムはルールベースの参考判定です。",
                    "消費税・インボイス制度・勘定科目については税理士へご確認ください。",
                ],
            }
        ],
        "ocr_text": ocr_text,
    }

    return result


def transactions_to_dataframe(result: dict) -> pd.DataFrame:
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
            }
        )
    return pd.DataFrame(rows, columns=EXCEL_COLUMNS)


def build_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="仕訳一覧")

        pending = df[df["確認事項"].astype(str).str.len() > 0].copy()
        pending.to_excel(writer, index=False, sheet_name="確認待ち")

        rules = pd.DataFrame(
            [
                {
                    "キーワード": "Amazon / 文房具 / PC周辺機器",
                    "推奨勘定科目": "消耗品費",
                    "税区分": "課税仕入10%",
                    "備考": "高額な場合は工具器具備品を検討",
                },
                {
                    "キーワード": "電車 / タクシー / 出張",
                    "推奨勘定科目": "旅費交通費",
                    "税区分": "課税仕入10% または 要確認",
                    "備考": "海外交通費は対象外の可能性",
                },
                {
                    "キーワード": "会食 / 飲食 / 顧客",
                    "推奨勘定科目": "接待交際費",
                    "税区分": "課税仕入10%",
                    "備考": "参加者・目的を確認",
                },
                {
                    "キーワード": "会議 / 打合せ / 軽食",
                    "推奨勘定科目": "会議費",
                    "税区分": "課税仕入10%",
                    "備考": "会議実態を確認",
                },
                {
                    "キーワード": "家賃 / オフィス賃料",
                    "推奨勘定科目": "地代家賃",
                    "税区分": "課税仕入10% または 非課税",
                    "備考": "住宅家賃は非課税の可能性",
                },
                {
                    "キーワード": "売上 / 請求",
                    "推奨勘定科目": "売上高",
                    "税区分": "課税売上10% または 要確認",
                    "備考": "輸出・海外売上は免税/対象外を確認",
                },
            ]
        )
        rules.to_excel(writer, index=False, sheet_name="科目ルール")

        workbook = writer.book
        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 36
                )

    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    st.caption("証憑画像または取引内容入力 → OCR読取 → 仕訳判定 → 確認 → Excel出力")

    with st.sidebar:
        st.header("設定")
        st.success("現在モード：オフラインOCR版（OpenAI API未使用）")
        st.info("初回起動時にEasyOCRモデルをダウンロードします。数分かかる場合があります。")

    left, right = st.columns([0.9, 1.1])

    with left:
        uploaded_file = st.file_uploader(
            "領収書・請求書・明細画像をアップロード",
            type=["png", "jpg", "jpeg", "webp"],
        )
        text_input = st.text_area(
            "取引内容を入力",
            height=180,
            placeholder=(
                "例：2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。"
                "キーボードとマウスを購入。消費税10%、適格請求書取得済み。"
            ),
        )

        if uploaded_file is not None:
            st.image(uploaded_file, caption="アップロード済み証憑", width="stretch")

        run = st.button("仕訳生成", type="primary", width="stretch")

    if run:
        if uploaded_file is None and not text_input.strip():
            st.warning("画像をアップロードするか、取引内容を入力してください。")
            return

        with st.spinner("OCR解析および仕訳生成中..."):
            try:
                result = make_transaction(text_input, uploaded_file)
                st.session_state["result"] = result
                st.session_state["df"] = transactions_to_dataframe(result)
            except Exception as exc:
                st.error(f"エラーが発生しました：{exc}")
                return

    with right:
        if "df" not in st.session_state:
            st.subheader("入力待機中")
            st.write("証憑画像または取引内容を入力し、「仕訳生成」をクリックしてください。")
            return

        st.subheader("OCR結果・仕訳一覧")
        edited_df = st.data_editor(
            st.session_state["df"],
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
        )
        st.session_state["df"] = edited_df

        excel_bytes = build_excel(edited_df)
        filename = f"japan_accounting_entries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "Excel出力",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

        with st.expander("OCR読取結果"):
            st.text(st.session_state.get("result", {}).get("ocr_text", ""))

        with st.expander("JSONデータ表示"):
            st.json(st.session_state.get("result", {}))


if __name__ == "__main__":
    main()
