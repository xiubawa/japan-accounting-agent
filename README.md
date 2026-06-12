# Japan Small Business Accounting Agent

日本の中小企業・個人事業主向けの Streamlit 会計仕訳 Agent です。

- 領収書、請求書、レシート、カード明細、銀行明細の画像をアップロード
- 手入力の取引内容にも対応
- OpenAI Vision で証憑を読み取り、仕訳候補を JSON 形式で生成
- 日本の勘定科目、消費税区分、インボイス登録番号を意識した確認リストを作成
- 画面上で人間が修正してから Excel に出力

## セットアップ

```powershell
cd C:\Users\fuzil\Documents\Codex\2026-06-12\agent\outputs\japan-accounting-agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
```

`.env` に OpenAI API Key を設定してください。

```text
OPENAI_API_KEY=your_api_key
OPENAI_MODEL=gpt-4o
PLAN_ACCESS_CODES={"starter-demo":"starter","business-demo":"business","firm-demo":"accounting_firm"}
```

本番では `PLAN_ACCESS_CODES` を Streamlit Cloud の Secrets に設定してください。例：

```toml
OPENAI_API_KEY = "your_api_key"
OPENAI_MODEL = "gpt-4o"
PLAN_ACCESS_CODES = '{"starter-xxxx":"starter","business-xxxx":"business","firm-xxxx":"accounting_firm"}'
```

または次の形式でも設定できます。

```toml
[plan_codes]
starter-xxxx = "starter"
business-xxxx = "business"
firm-xxxx = "accounting_firm"
```

有料ユーザーには対応するアクセスコードだけを渡してください。コードが未入力または無効な場合は、自動的に無料プランとして扱われます。

## 起動

```powershell
streamlit run app.py
```

ブラウザでローカル画面が開きます。画像をアップロードするか、次のような取引内容を入力してください。

```text
2026年6月10日、法人カードでAmazon Japanへ11,000円支払い。
キーボードとマウスを購入。消費税10%、適格請求書取得済み。
```

## Excel 出力

出力ファイルには次のシートが含まれます。

- `仕訳一覧`: AI が生成した仕訳候補と確認事項
- `確認待ち`: 確認事項がある取引の一覧
- `集計`: 処理日時、取引件数、税込合計、注意事項

## 注意

このアプリは会計処理の補助ツールです。法人税、所得税、消費税申告、インボイス制度、源泉所得税、固定資産判定などの最終判断は、必ず税理士または専門家に確認してください。
