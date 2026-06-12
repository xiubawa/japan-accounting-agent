# Japan Accounting Excel Agent

一个最小可运行的日本会计仕訳 Agent：

- 支持上传收据/发票图片
- 支持手动输入交易描述
- 自动提取交易信息
- 自动推荐日本会计勘定科目、税区分、借贷分录
- 支持人工修改
- 导出 Excel

## 1. 安装

```powershell
cd C:\Users\fuzil\Documents\Codex\2026-06-12\agent\outputs\japan-accounting-agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
```

然后打开 `.env`，填入：

```text
OPENAI_API_KEY=你的 OpenAI API Key
```

## 2. 运行

```powershell
streamlit run app.py
```

浏览器会打开本地页面。你可以上传图片，或者直接输入：

```text
2026年6月10日，公司信用卡支付 Amazon Japan 11000日元，购买办公键盘和鼠标，含10%消费税，取得合格发票。
```

## 3. Excel 输出字段

导出的 Excel 包含：

- 取引日
- 取引先
- 摘要
- 借方勘定科目
- 借方金額
- 借方税区分
- 貸方勘定科目
- 貸方金額
- 貸方税区分
- 消費税率
- 消費税額
- 支払方法
- インボイス登録番号
- 証憑種類
- ステータス
- 確認事項

## 4. 注意

这个 MVP 用于会计处理辅助。涉及法人税、消费税申报、インボイス制度、源泉所得税等最终判断时，请让税理士确认。
