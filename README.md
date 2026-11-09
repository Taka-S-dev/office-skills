# officeskill

コーディングエージェントで Excel / Word / PowerPoint / PDF、そしてフォルダに保存したメール（.eml / .msg）を読み書きするためのスキル 2 つ。
SKILL.md 形式を読むエージェントならどれでも使え、スクリプト単体でも動く。
外部のスキルやサービスに依存せず、MIT / BSD ライセンスの Python ライブラリだけで動く。

```
officeskill/
├── office-read/        読む。xlsx docx pptx pdf csv eml msg zip 7z mpp vsdx（と旧形式 xls doc ppt）→ Markdown / JSON
│   ├── SKILL.md        エージェント向けの手順（英語）
│   ├── scripts/        office2md（読む）/ xlsx_deps（数式の依存）/ office_diff（2 版の差分）
│   │                   / xlsx_check（ルール検査）/ office_review（変更履歴とコメント）/ office_find（横断検索）
│   │                   / mail_digest（フォルダ内メールの時系列）/ unpack（zip・7z の展開）
│   │                   / md_crosscheck（Word・Excel → md の照合と目視確認）/ diagram_lint（mermaid・PlantUML の構文検査）
│   │                   （検査の部品は review_checks、図形の読み取りは xlsx_drawing）
│   │                   / visio_reader（Visio の業務フロー）/ office_com（Office を COM で安全に動かす）
│   │                   / mpp_reader（MS Project の計画。単体でも動く）/ mpp_history（2 版の差分、1 年分の推移）
│   └── profiles/       組織固有の定型ファイルの知識を置く場所（配布時は空）
├── office-edit/        書く。置換、セル更新、再計算、PDF の結合・分割・しおり
│   ├── SKILL.md
│   ├── scripts/        docx_replace / pptx_replace / xlsx_set / recalc / pdf_tool / pdf_headings（しおり用の章立て抽出）
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

MS Project の .mpp を読むには、次のどちらかが要る（無くても、Project で「XML 形式で保存」した .xml は読める）。

- Microsoft Project が入っている PC: 追加のインストールは要らない。Project 自身に XML を書き出させる（最も忠実な経路）。
- 入っていない PC: `pip install mpxj`（Java が必要。MPXJ は LGPL のライブラリで、取り込まずに呼ぶだけ）。

どちらで読んだかは出力の `Read with:` 行に出る。

コピー先は既定で `~/.claude/skills` と、GitHub Copilot のフォルダ（`~/.copilot`）があれば `~/.copilot/skills`。
それ以外のエージェントを使う場合は引数でそのスキルフォルダを渡す（そのときはそこだけにコピーする）。

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
python office-read/scripts/office2md.py 仕様書.docx --out 仕様書.md --images ./fig   # 番号付き見出し・脚注・図を保った md
python office-read/scripts/md_crosscheck.py 仕様書.docx --out-dir ./md             # pandoc 等と突き合わせ、差分を review.md に
python office-read/scripts/office2md.py 業務フロー.vsdx                         # レーンごとの流れを mermaid と一覧で
python office-read/scripts/office2md.py 統計.pdf --layout                       # 表の多い PDF は列をそろえて
python office-read/scripts/office2md.py 仕様書.pdf --columns                    # 狭い列の表で文が細切れになる PDF
python office-read/scripts/office2md.py 計画.mpp --gantt                          # 遅れ・予定進捗との差・計画の品質・終了日を決める連なり・gantt 図
python office-read/scripts/mpp_history.py 旧.mpp 新.mpp                         # 追加・削除・日付のずれ・進捗の差分
python office-read/scripts/mpp_history.py ./計画の履歴 --cache ./mpp_cache          # 1 年分: 悪化した時期、後から増えたタスク、マイルストーンの推移
python office-read/scripts/office2md.py 売上.xlsx --outline          # 構造だけ
python office-read/scripts/office2md.py 売上.xlsx --sheet 明細 --max-rows 0
python office-read/scripts/office2md.py 売上.xlsx --sheet メイン --cells   # メイン!C4 = 12 の形で 1 セル 1 行
python office-read/scripts/office2md.py 申請書.xlsx --sheet 申請書 --form  # ラベル = 値 の対で出す
python office-read/scripts/xlsx_deps.py 売上.xlsx --cell "メイン!D10"      # D10 が何から計算されているか
python office-read/scripts/xlsx_deps.py 売上.xlsx --cell "メイン!B3" --dependents   # B3 を変えると何に影響するか
python office-read/scripts/office_diff.py 報告書_v1.xlsx 報告書_v2.xlsx        # 2 版の差分だけをセル番地付きで
python office-read/scripts/xlsx_check.py 報告書.xlsx --profile office-read/profiles/report.md   # プロファイルのルールで検査
python office-read/scripts/office_review.py 報告書.docx                      # 変更履歴とコメントを誰が・いつ・どこで付きで
python office-read/scripts/office_find.py "XYZ-100" ./reports              # フォルダ内の全ファイルから場所付きで検索
python office-read/scripts/mail_digest.py ./案件 --threads --bodies 300     # .eml / .msg を時系列に並べて経緯を掴む
python office-read/scripts/office2md.py ./案件/依頼.msg --extract-attachments ./att   # メール 1 通と添付
python office-read/scripts/mail_digest.py ./案件 --extract ./att --deep       # 添付と zip/7z の中身まで一覧
python office-read/scripts/unpack.py 添付資料.zip --out ./extracted            # 文字化けせず安全に展開
python office-read/scripts/unpack.py 添付.7z --out ./ --password-file ~/passwords.txt   # 候補を順に試す
python office-edit/scripts/docx_replace.py 報告書.docx --replace "2025年度=2026年度" --dry-run
python office-edit/scripts/recalc.py 売上.xlsx                       # 数式再計算と #REF! 検出
python office-edit/scripts/pdf_tool.py merge out.pdf a.pdf b.pdf
python office-edit/scripts/pdf_headings.py 仕様書.pdf --out toc.txt   # 全ページから章・節を拾ってしおり用リストに
python office-edit/scripts/pdf_tool.py bookmarks 仕様書.pdf --set toc.txt --out 仕様書_bm.pdf
```

各スクリプトは `--help` で全オプションが出る。

### Word の仕様書をエージェント向けの md にする

変換そのものは `office2md.py` で済むが、「最も正確な 1 本」を作るときは `md_crosscheck.py` を使う。
確かめる深さは 3 段階（時間は 48 ページの文書で測った目安）。

| 段階 | コマンド | 確かめること | 時間 |
|---|---|---|---|
| 本文だけ | `office2md.py 仕様書.docx` | なし（変換のみ） | 1 秒未満 |
| 標準 | `md_crosscheck.py 仕様書.docx --no-render` | pandoc / markitdown との照合、リンク、画像、取り消し線 | 約 3 秒 |
| 忠実 | `md_crosscheck.py 仕様書.docx` | 加えて Word で印刷したページ、Word の段落番号との照合 | 約 14 秒（2 回目からは約 2 秒） |


1. `md_crosscheck.py 仕様書.docx --out-dir ./md` を実行。`仕様書.md`（自前変換。見出しの自動番号、脚注、
   変更履歴の印（`{++追加++}` `{--削除--}`）とコメントの位置、
   図の alt、テキストボックス、リンク、段落の中の太字・斜体・下線・取り消し線・蛍光ペン・上付きと下付き（m<sup>2</sup>、H<sub>2</sub>O）、
   ヘッダーの中の表やテキストボックスを保持。図形で描いた図（キャンバス、グループ、線）は置かれた位置に
   `[drawing: …]` と印を出し、Word があれば Word の表示どおりに PNG にして画像リンクにする。EMF / WMF / TIFF の
   画像も PNG にする。Markdown の 6 段を超える見出しは番号付きの太字行になる）と、pandoc / markitdown が入っていればそれらの変換結果、
   そして照合結果ができる。出力先に残るのは `仕様書.md` と図のフォルダ（`仕様書_figures/`）だけで、
   pandoc などの変換結果、印刷した PDF、照合結果の `review.md`、確認の記録は `仕様書_review/` にまとまる
   （以前の版が md の隣に置いたファイルも、次の実行でここに移る）。
2. 数式の LaTeX 化、太字、リンクのうち、一意に当たるものはツールが自動で `仕様書.md` に取り込む。
   Word で印刷できたときは、段落番号が Word の表示と違う行（その文が 1 行だけにあるもの）を Word の番号に直す。
   `review.md` の先頭に `Status: READY` か `NEEDS REVIEW` と、検査ごとの結果（ok / FAIL / not run）が出る。
   検査は、md がどの版のファイルから作られたか、参照側にあって自前に無い文（同じ文が何回出るかも数える）、
   リンク、見出し、表の形、目視確認（中身のある図、図形で描いた図、SmartArt、グラフ、埋め込みオブジェクト）。
   元ファイルとも直接比べる。画像リンクの先が実在するか、ビューアで出ない形式（EMF / WMF / TIFF）へのリンクが無いか、
   取り消し線の文字数、画像と図の数、図の並び順（pandoc と比べる）、Word が表示する段落番号（ListString）が md の同じ行にあるか。
   別の見出しの下に移った文や、判断が要る太字や数式は、止めずに一覧だけ出す。エージェントはそれを見て手直しする。
   加えて、Word で印刷した PDF（`<stem>_review/pages/document.pdf`）を変換器とは独立した照合相手にする。
   各ページの文字が md に全部あるかを見て（どの変換器も落とす文字もここで見つかる）、
   ページと md の行の対応から、散らばったページと一致の低いページを目視確認 `page<n>` に出す。
3. 正しいままでよい項目は `<stem>_review/checked.txt` に `accept <id> <理由>` と書いて閉じる（理由の無い行は無効）。
   実行ごとの件数は `<stem>_review/rounds.jsonl` に残り、review.md に回ごとの表が出るので、直して増えた回が分かる。
   サブエージェントが使えるなら、文書ごとに新しいエージェントに md・印刷 PDF・review.md を渡して
   変換の不具合（見出しの階層、表の形、行の構造、PDF にあって md に無い内容）を探させ、直して再実行し、
   新しい指摘が出なくなるまで別のエージェントで繰り返す。
4. 直したら再実行し、READY になるか、残るのが意図した差分（畳んだ目次、pandoc が落とす表題、自前だけの番号）
   だけになれば完成。最終の Status と残した差分の理由を報告する。

READY は「ここに挙げた検査の範囲で未解決が無い」という意味で、元の文書と意味が同じことの証明ではない。
md の先頭の `_Source:` 行に元ファイルのハッシュが入り、ファイルが変わると古い md は FAIL になる。

pandoc 単体の出力を使わないのは、見出しの自動番号（「3.2.1 を参照」の辿り先）を pandoc が落とすため。
pandoc が無い環境でも 1 の md は作られ、review に「照合相手なし」と出る。

### Excel の設計書（シーケンス図・レイアウト表）を md にする

同じ `md_crosscheck.py` を xlsx に使う。Excel には自前変換を超える参照変換がないので、照合相手は
「ブック自身の全テキスト」と「シートの描画結果」になる。

- 出力 md はシートごとに表や注記のまとまり（ブロック）に分け、結合セルに `[B3:E3, 4 cols]` と幅を付ける
  （レイアウト表の項目の幅が読める）。
- 図形で描いたシーケンス図は、生存線・参加者・矢印の向き・ラベルを位置から復元し、mermaid と
  「どのセルの矢印か」付きの番号リストで出す。フロー図はコネクタの接続で出す。グラフは系列の値、
  入力規則は選択肢、画像は寸法と代替テキストまで出す。
- シーケンス図は戻り矢印（点線）、自分宛ての矢印、活性区間、alt / loop などの枠と else、ノートまで復元する。
  `--uml plantuml|both` で PlantUML でも出せる。図そのものは `<名前>.drawio` に位置と線をそのまま写した
  編集可能な形でも残る。md の中の mermaid / PlantUML は構文を検査し、壊れていれば READY にしない
  （`diagram_lint.py` 単体でも使える。`--render-diagrams` を付け、PlantUML か mermaid-cli が入っていれば
  実際に描画して確かめる。PlantUML は砂箱設定で動かす）。
- SmartArt の文字（階層付き）、古い形式（.xls 由来）のテキストボックス、グラフだけのシート、印刷用の
  ヘッダー・フッター（`&P` などはページ番号と分かる形で）、取り消し線（セルの一部の語も `~~削除~~`）も出す。
  数値は書式どおりに出す（負の数のかっこ、書式に書いた単位、円記号）。
- 各シートを `<名前>_review/sheets/sheetNN.pdf` に描画し、review.md に「ブックの文字で md に無いもの」と
  「目視で確認すべき図・グラフ・画像・SmartArt・読めない部品」を出す。文字はシートごとに、出てくる回数まで
  照合する。1 や ○ のような短い値は、同じ行の項目名と並んでいるかで確かめる。
- エージェントは PDF を見て md を直し、確認したものの ID（`シーケンス#diagram@1a2b3c4d` のように、内容の指紋付き）
  を `<名前>_review/checked.txt` に書く。元ファイルの図が変わると ID が変わり、その図だけ確認し直しになる。
  全部そろえば READY。描画は同じ版のファイルの間だけ使い回す。

### パスワード付きの zip / 7z

チームの定番パスワードが数個あるなら、1 行 1 個のテキストに並べて `--password-file` で渡す。
上から順に試し、最初に開けたものを使う（候補はこのファイルの中身だけ。生成や総当たりはしない）。
開けたパスワード自体は画面に出ない。1 個だけなら `--password` を直接付ける。

```
# 例: ~/passwords.txt（リポジトリの外に置く）
2026sales
project-x
```

置き場所はリポジトリの**外**（ホームディレクトリなど）にする。.gitignore は `passwords.txt` や
`*password*.txt` を止めるが、名前を変えられると守れないので、中に置かない運用が前提。
`mail_digest.py --deep --password-file` でも同じファイルが使える。

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

手元に評価用の文書の束（リポジトリには入れない）があるなら、変換を変えた前後で回して比べる。
検査が ok から FAIL に変わった文書や、欠落などの件数が増えた文書を一覧にし、1 件でもあれば終了コード 1 を返す。

```bash
python tests/corpus_check.py D:\文書の束 --save before.json            # 変更前に記録
python tests/corpus_check.py D:\文書の束 --save after.json --baseline before.json
```

## 設計上の決めごと

- 読みと書きを分けているのは、スキルの発動条件を明確にするためと、読み取り専用で使いたい
  環境に office-read だけ渡せるようにするため。
- SKILL.md と pitfalls.md はエージェントが読む文書なので英語。README と profiles は人が書く
  文書なので日本語。
- 旧形式（xls / doc / ppt）は非対応。LibreOffice か Office で変換してから扱う。
- PDF 本文の書き換えは対応しない。元の docx / pptx を直して `pdf_tool.py topdf` で出し直す。
