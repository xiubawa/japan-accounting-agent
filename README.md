# Japan Small Business Accounting Agent

日本の中小企業・個人事業主向けの Streamlit 会計仕訳 Agent です。

- 領収書、請求書、レシート、カード明細、銀行明細の画像をアップロード
- 手入力の取引内容にも対応
- OpenAI Vision で証憑を読み取り、仕訳候補を JSON 形式で生成
- 日本の勘定科目、消費税区分、インボイス登録番号を意識した確認リストを作成
- `基本記帳`、`月次決算`、`年次決算` を分けて利用
- 月次決算・年次決算では顧客の帳簿 CSV / Excel をアップロードして、試算表・損益計算書・貸借対照表を作成
- 年次決算では月別ファイルを複数アップロードして、1年分に合算可能
- 基本記帳で作成した仕訳を顧客ごとの保存済み帳簿に追加し、月次決算・年次決算で再利用可能
- Excel 出力には表紙、財務サマリー、公式付きKPI、整形済み財務諸表を含める
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

自助登録ユーザーは `customer_accounts.json` に保存されます。Supabase を設定していない場合、顧客ごとの保存済み帳簿と開始残高は `customer_ledgers/` に保存されます。どちらも `.gitignore` に含まれているため、GitHub にはコミットされません。Streamlit Cloud の無料環境では、再デプロイや環境リセット時にローカルファイルが失われる可能性があります。本格運用では Supabase 連携を設定してください。

Supabase に保存する場合は、Supabase の SQL Editor で次のテーブルを作成します。

```sql
create table if not exists public.customer_storage (
  storage_key text primary key,
  payload jsonb not null,
  updated_at timestamptz default now()
);
```

その後、Streamlit Cloud の Secrets に次を追加してください。

```toml
SUPABASE_URL = "https://your-project.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "your-service-role-key"
SUPABASE_STORAGE_TABLE = "customer_storage"
```

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

画面の `開始残高設定` では、このアプリを使い始める前のBS残高や当期途中までのPL累計額を顧客ごとに保存できます。保存した開始残高は、基本記帳・月次決算・年次決算の `試算表`、`PL`、`BS` に反映されます。

開始残高のアップロードは、`科目 / 残高` の標準表だけでなく、タイトル行付きの貸借対照表・損益計算書や、資産と負債が左右に並ぶ表も自動判定します。読み取り後は画面上の表で確認・修正して保存してください。

出力ファイルには次のシートが含まれます。

- `仕訳帳`: 日々の仕訳を入力するメインシート。借方科目・貸方科目は科目マスタから選択できます。
- `科目マスタ`: 科目コード、科目名、分類、PL/BS区分、開始残高を管理する設定シート
- `試算表`: 開始残高と仕訳帳から `SUMIF` で借方合計・貸方合計・現在残高を自動集計
- `PL`: 試算表の収益・費用から損益計算書を自動作成
- `BS`: 試算表の資産・負債・純資産とPLの当期純利益から貸借対照表を自動作成
- `精算表`: 年次決算の出力時のみ追加される、修正記入・PL・BSを並べたワークシート

## 注意

このアプリは会計処理の補助ツールです。法人税、所得税、消費税申告、インボイス制度、源泉所得税、固定資産判定などの最終判断は、必ず税理士または専門家に確認してください。
