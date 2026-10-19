# officeskill

コーディングエージェントで Excel / Word / PowerPoint / PDF を読み書きするためのスキル 2 つ。
SKILL.md 形式を読むエージェントならどれでも使え、スクリプト単体でも動く。
外部のスキルやサービスに依存せず、MIT / BSD ライセンスの Python ライブラリだけで動く。

```
officeskill/
├── office-read/        読む。xlsx docx pptx pdf csv → Markdown / JSON
│   ├── SKILL.md        エージェント向けの手順（英語）
│   ├── scripts/        office2md（読む）/ xlsx_deps（数式の依存）/ office_diff（2 版の差分）
│   │                   / xlsx_check（ルール検査）/ office_review（変更履歴とコメント）
│   └── profiles/       組織固有の定型ファイルの知識を置く場所（配布時は空）
├── office-edit/        書く。置換、セル更新、再計算、PDF の結合・分割・しおり
│   ├── SKILL.md
│   ├── scripts/        docx_replace / pptx_replace / xlsx_set / recalc / pdf_tool
│   └── references/pitfalls.md   ライブラリで壊れやすい点の一覧
├── tests/              smoke_test.py（全スクリプトの通し確認）
├── install.ps1 / install.sh
└── LICENSE
```

## インストール

```powershell
pip install -r office-read/scripts/requirements.txt -r office-edit/scripts/requirements.txt
.\install.ps1          # スキルフォルダにコピー（mac/Linux は ./install.sh）
python tests/smoke_test.py
```

2 つはセットで入れる。office-edit は office-read のスクリプトと profiles を参照する。

コピー先は既定で `~/.claude/skills`。別のエージェントを使う場合は引数でそのスキルフォルダを渡す。

```powershell
.\install.ps1 C:\path\to\your-agent\skills
./install.sh ~/.your-agent/skills
```

再計算と PDF 出力は Windows なら Office 本体（COM）、なければ LibreOffice を使う。
どちらも無い環境では該当機能だけがスキップされ、他は動く。

## 日常の使い方

スキルは自動で発動する。「この xlsx 読んで」「この docx の社名を置換して」で良い。
手動で実行したいとき:

```bash
python office-read/scripts/office2md.py 売上.xlsx --outline          # 構造だけ
python office-read/scripts/office2md.py 売上.xlsx --sheet 明細 --max-rows 0
python office-read/scripts/office2md.py 売上.xlsx --sheet メイン --cells   # メイン!C4 = 12 の形で 1 セル 1 行
python office-read/scripts/office2md.py 申請書.xlsx --sheet 申請書 --form  # ラベル = 値 の対で出す
python office-read/scripts/xlsx_deps.py 売上.xlsx --cell "メイン!D10"      # D10 が何から計算されているか
python office-read/scripts/xlsx_deps.py 売上.xlsx --cell "メイン!B3" --dependents   # B3 を変えると何に影響するか
python office-read/scripts/office_diff.py 報告書_v1.xlsx 報告書_v2.xlsx        # 2 版の差分だけをセル番地付きで
python office-read/scripts/xlsx_check.py 報告書.xlsx --profile office-read/profiles/report.md   # プロファイルのルールで検査
python office-read/scripts/office_review.py 報告書.docx                      # 変更履歴とコメントを誰が・いつ・どこで付きで
python office-edit/scripts/docx_replace.py 報告書.docx --replace "2025年度=2026年度" --dry-run
python office-edit/scripts/recalc.py 売上.xlsx                       # 数式再計算と #REF! 検出
python office-edit/scripts/pdf_tool.py merge out.pdf a.pdf b.pdf
```

各スクリプトは `--help` で全オプションが出る。

## 拡張のしかた（よくある順）

### 1. 組織固有の定型ファイルの知識を足す

`office-read/profiles/` に 1 ファイル種類につき 1 つ Markdown を置く。雛形は自動生成できる。

```bash
python office-read/scripts/office2md.py 売上報告_202610.xlsx --profile --out office-read/profiles/sales-report.md
```

構造（シート、ヘッダ行、数式セルの範囲）は埋まった状態で出るので、`TODO` の行だけ埋める。
書く内容の優先順位と例は `office-read/profiles/README.md` にある。
これはコードを触らない唯一の拡張で、現場で一番使う。

### 2. 壊れやすい点を学んだら書き足す

`office-edit/references/pitfalls.md` の該当形式の節に 1 項目足す。エージェントはコードを書く前に
この節を読むので、同じ事故を二度起こさなくなる。

### 3. 定型の編集操作を増やす

`office-edit/scripts/` にスクリプトを 1 本足し、`office-edit/SKILL.md` の表に 1 行足す。
既存スクリプトの型に合わせる: `--dry-run` があり、既定でバックアップを取り、変更点を
出力し、Windows でも UTF-8 で出す（各スクリプト冒頭の sys.stdout.reconfigure の 4 行をコピー）。
Word / PowerPoint のテキスト置換が必要なら `_runs.py` の `replace_in_runs` を使う。

### 4. 読める形式を増やす

`office-read/scripts/office2md.py` に `Handler` のサブクラスを 1 つ書き、末尾の `HANDLERS`
辞書に拡張子を登録する。`outline()` `dump()` `to_markdown()` の 3 つを実装すれば、
CLI・打ち切り・JSON・`--out` はそのまま使える。既存の `CsvHandler` が最小の例。

### 変更したら

```bash
python tests/smoke_test.py && .\install.ps1
```

テストに 1 行足すのは `tests/smoke_test.py` の `check(...)` を真似る。

## 設計上の決めごと

- 読みと書きを分けているのは、スキルの発動条件を明確にするためと、読み取り専用で使いたい
  環境に office-read だけ渡せるようにするため。
- SKILL.md と pitfalls.md はエージェントが読む文書なので英語。README と profiles は人が書く
  文書なので日本語。
- 旧形式（xls / doc / ppt）は非対応。LibreOffice か Office で変換してから扱う。
- PDF 本文の書き換えは対応しない。元の docx / pptx を直して `pdf_tool.py topdf` で出し直す。
