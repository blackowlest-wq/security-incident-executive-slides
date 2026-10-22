# 経営層向けセキュリティインシデント・改善投資資料

国内の重大事例から、事業継続・復旧とセキュリティ改善投資を説明する日本語の資料です。本編14枚、出典4枚。調査基準日は2026年10月22日です。

## 資料

- [追加提供資料：security_dopamine.pptx](output/security_dopamine.pptx)（提供されたファイルをそのまま収録）
- [追加提供資料：security_incident_executive_2026.pptx](output/security_incident_executive_2026.pptx)（提供されたファイルをそのまま収録）
- [編集可能なPowerPoint](output/経営層向け_セキュリティ改善投資_図解改訂版.pptx)
- [確認用PDF](output/経営層向け_セキュリティ改善投資_図解改訂版.pdf)
- [制作時の変更・検証記録](output/図解改訂版_変更・検証.md)
- [スライド構成](docs/構成案.md)
- [出典・重要事実](src/sources.json)

図・テキスト・グラフは可能な限りPowerPointのネイティブ要素で構成しています。詳細な説明と出典URLは発表者ノートにも記載しています。公表値・会社推定・提案を区別し、自社の予算額や効果を仮の数値で補っていません。

PPTXは直接編集できます。PDFと検証記録は制作時点の書き出しであり、その後のPPTXの手動編集が自動で反映されるものではありません。

## 生成用ソース

Node.jsとPython 3を使用します。仕上げの処理にはWindows版Microsoft PowerPointが必要です。

```powershell
npm ci
npm run build:draft
python src/normalize.py
powershell -ExecutionPolicy Bypass -File src/native_save.ps1
powershell -ExecutionPolicy Bypass -File src/polish.ps1
```

生成物は`src/candidate_polished.pptx`です。公開中の資料は上書きしません。PowerPointで確認後、必要に応じて配布用ファイルに差し替え、PDFも再出力してください。手動編集済みの配布版と生成結果は一致しない場合があります。

- `src/build.cjs`：PptxGenJSによるスライド生成
- `src/previous_notes.json`：詳細説明の元データ
- `src/sources.json`：出典台帳（ページ番号は別途、構成と照合してください）
- `src/normalize.py`：PPTX内の互換性調整
- `src/native_save.ps1`：PowerPointでの保存
- `src/polish.ps1`：因果図・ラベル・矢印の仕上げ

参照資料の著作権は各権利者に帰属します。このリポジトリに第三者の原文資料は同梱せず、出典URLを記載しています。
