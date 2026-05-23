# mf-categorizer ── マネーフォワード CSV を AI ルールで再分類

マネーフォワード ME の CSV エクスポートを、自分用の分類ルールで **再分類** して Excel レポートを出します。MF の「未分類」「明らかな誤分類」を AI ルールで潰す月 1 回のワークフロー用。

## 使い方

### 1. マネーフォワードから CSV をエクスポート

MF ME → 該当月 → 「ダウンロード」→ CSV を選択。

### 2. 初回だけ:ルールファイルを作る

```
cp rules.example.json rules.json
```

`rules.example.json` は全国チェーン中心の汎用サンプル。`rules.json` の方を自分の使う店名で育てていく(=こちらは `.gitignore` で除外済み、commit されない)。

### 3. スクリプトを実行

```
python mf_categorize.py ~/Downloads/収入・支出詳細_2026-05-01_2026-05-31.csv
```

### 4. 結果を確認

`output/<元のファイル名>_categorized.xlsx` が出る。中身は 3 シート:

| シート | 内容 |
|---|---|
| 取引明細 | 全取引。MFの分類と AI の分類を並列表示。判定が違うものは **黄色**、ルール未マッチは **赤**、振替は **灰** |
| 月次サマリー | カテゴリ別の合計 |
| 要確認 | ルールに無くて自動分類できなかった取引(ここを潰していけば来月から自動化が進む) |

## 学習のさせ方

`rules.json` に、覚えさせたい店名とカテゴリを追加するだけ。

### 単純な店名ルール

```json
{"match": "新しい店名の一部", "category": "食費", "subcategory": "食料品"}
```

`match` は **部分一致**。`"マルキョウ"` と書けば `"マルキョウ 松田店 福岡県 福岡市東区"` にもマッチします。

### 同じ店でも金額帯で振り分けたい場合(ららぽーと方式)

```json
{
  "match": "ららぽーと",
  "splits": [
    {"max_amount": 1200, "category": "食費", "subcategory": "外食"},
    {                    "category": "衣服・美容", "subcategory": "子供服"}
  ]
}
```

`amount_rules` 配列に追加。`max_amount` 以下なら 1 つ目、超えたら 2 つ目(`max_amount` 無しがデフォルト)。

## ファイル構成

```
mf-categorizer/
├── mf_categorize.py     # メインスクリプト
├── rules.example.json   # 分類ルールのサンプル(これを rules.json にコピーして使う)
├── rules.json           # あなた専用の分類ルール(.gitignore で除外)
├── requirements.txt     # 依存ライブラリ
├── LICENSE              # MIT
├── README.md            # このファイル
└── output/              # 再分類済み Excel が出るところ(.gitignore で除外)
```

## 必要なもの

- Python 3.9 以上
- `pip install -r requirements.txt`(openpyxl)

## プライバシー設計

- `rules.json`(あなた専用のルール)・`output/`(取引データを含むExcel)・`*.csv`(MFエクスポート) はすべて `.gitignore` で除外
- 配布されるのは `rules.example.json`(全国チェーンの汎用例のみ)
- 自分用に育てた `rules.json` は手元にしか存在しない設計

## 限界(正直)

- **クイックペイ / Apple Pay** の取引は決済方法名しか来ず、店名が落ちている。デフォルトは「食費/コンビニ」にしているが、コンビニ以外で使った場合はズレる。
- **Amazon / 楽天** は店名はあるが品目不明。**Phase B(Gmail 連携)** で注文確認メールと突き合わせれば品目までわかる(将来拡張)。
- **イオン九州** などスーパー兼日用品店は、レシート分解しない限り「食費 vs 日用品」の比率は分からない。デフォルトは食費。

## ルール変更後の再実行

`rules.json` を編集したら、同じコマンドで再実行するだけで新しいルールが反映される。CSV は同じものを使い回せる。
