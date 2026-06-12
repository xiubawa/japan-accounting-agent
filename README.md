# Japan Small Business Accounting Agent

日本の中小企業・個人事業主向けの Streamlit 会計仕訳 Agent です。

- 領収書、請求書、レシート、カード明細、銀行明細の画像をアップロード
- 手入力の取引内容にも対応
- OpenAI Vision で証憑を読み取り、仕訳候補を JSON 形式で生成
- 日本の勘定科目、消費税区分、インボイス登録番号を意識した確認リストを作成
- `基本記帳`、`月次決算`、`年次決算` を分けて利用
- 月次決算・年次決算では顧客の帳簿 CSV / Excel をアップロードして、試算表・損益計算書・貸借対照表を作成
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
UPGRADE_CONTACT=上位プランをご希望の場合は example@example.com までお問い合わせください。
CUSTOMER_ACCOUNTS={"demo@example.com":{"password":"demo-pass","plan":"starter","name":"Demo Company"}}
```

ユーザーはログイン画面から自助で無料登録できます。登録直後のプランは必ず `free` です。上位プランに変更したいユーザーには、画面に表示される `UPGRADE_CONTACT` の連絡先から問い合わせてもらいます。

本番では OpenAI Key と問い合わせ先を Streamlit Cloud の Secrets に設定してください。必要に応じて、有料顧客アカウントも Secrets に設定できます。例：

```toml
OPENAI_API_KEY = "your_api_key"
OPENAI_MODEL = "gpt-4o"
UPGRADE_CONTACT = "上位プランをご希望の場合は your-email@example.com までお問い合わせください。"
CUSTOMER_ACCOUNTS = '{"demo@example.com":{"password":"demo-pass","plan":"starter","name":"Demo Company"}}'
```

または次の形式でも設定できます。

```toml
[customers.demo]
password = "demo-pass"
plan = "starter"
name = "Demo Company"

[customers.firm]
password = "firm-pass"
plan = "accounting_firm"
name = "Demo Accounting Firm"
```

顧客にはユーザーIDとパスワードを渡してください。ログインに成功した顧客だけがアプリを利用でき、顧客ごとの `plan` に応じて一度に処理できる取引件数が制限されます。画像アップロードだけでなく、テキスト入力で複数取引を書いた場合も同じ上限が適用されます。

自助登録ユーザーは `customer_accounts.json` に保存されます。このファイルは `.gitignore` に含まれているため、GitHub にはコミットされません。Streamlit Cloud の無料環境では、再デプロイや環境リセット時にローカルファイルが失われる可能性があります。本格運用ではデータベース連携を推奨します。

利用できる `plan` は次の通りです。

- `free`: 1取引/回
- `starter`: 5取引/回
- `business`: 20取引/回
- `accounting_firm`: 50取引/回

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
- `試算表`: アップロード帳簿から作成した試算表
- `損益計算書`: 収益・費用・当期利益
- `貸借対照表`: 資産・負債・純資産
- `月次サマリー`: 月別の取引件数、税込合計、税抜合計、消費税合計
- `月次科目別`: 月別・勘定科目別の集計
- `月次税区分別`: 月別・税区分別の集計
- `年次サマリー`: 年別の取引件数、税込合計、税抜合計、消費税合計
- `年次科目別`: 年別・勘定科目別の集計
- `年次確認事項`: 決算前に確認すべき取引
- `月次チェックリスト`: 月次決算の確認項目
- `年次チェックリスト`: 年次決算の確認項目
- `集計`: 処理日時、取引件数、税込合計、注意事項

## 注意

このアプリは会計処理の補助ツールです。法人税、所得税、消費税申告、インボイス制度、源泉所得税、固定資産判定などの最終判断は、必ず税理士または専門家に確認してください。
