const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "4472C4" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 32, font: "メイリオ", color: "2E4057" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "メイリオ", color: "4472C4" })]
  });
}

function body(text, { bold = false, indent = false } = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: 22, font: "メイリオ", bold })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "メイリオ" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 100, after: 100 }, children: [new TextRun("")] });
}

function headerCell(text, width) {
  return new TableCell({
    borders: headerBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "2E4057", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 20, font: "メイリオ", color: "FFFFFF" })]
    })]
  });
}

function cell(text, width, { shade = "FFFFFF", align = AlignmentType.LEFT, bold = false } = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, size: 20, font: "メイリオ", bold })]
    })]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "メイリオ", size: 22 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 240 },
        children: [new TextRun({ text: "ビッグファイブ カードゲーム", bold: true, size: 48, font: "メイリオ", color: "2E4057" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "企画コンセプト書  ver.0.1", size: 24, font: "メイリオ", color: "888888" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 480 },
        children: [new TextRun({ text: "2026年3月", size: 22, font: "メイリオ", color: "888888" })]
      }),

      // ─── 1. ゲームコンセプト ───
      h1("1. ゲームコンセプト"),
      body("このゲームは、心理学の「ビッグファイブ理論」をベースにした教育型カードゲームです。"),
      body("プレイヤーは自分のビッグファイブ診断結果をそのままキャラクターのステータスとして使い、"),
      body("選んだ「バトルフィールド（職種・環境）」での成功を目指します。"),
      spacer(),
      body("カードを使って自分の特性を鍛え、相手の特性を崩しながら、"),
      body("フィールドで求められる理想のビッグファイブ像に近づいていくゲームです。"),
      spacer(),
      body("ゲームをやるうちに、自然と心理学の知識が身につき、", { bold: true }),
      body("「現実の自分はどう行動すればいいか」に気づける設計になっています。", { bold: true }),

      spacer(),

      // ─── 2. ビッグファイブとは ───
      h1("2. ビッグファイブとは"),
      body("人間のパーソナリティを5つの軸で数値化した心理学モデルです。"),
      body("（頭文字をとって「OCEAN」とも呼ばれます）"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [800, 1800, 2913, 2913, 600],
        rows: [
          new TableRow({
            children: [
              headerCell("略称", 800),
              headerCell("特性名", 1800),
              headerCell("高い場合", 2913),
              headerCell("低い場合", 2913),
              headerCell("記号", 600),
            ]
          }),
          new TableRow({
            children: [
              cell("O", 800, { align: AlignmentType.CENTER, bold: true }),
              cell("開放性", 1800, { shade: "EAF2FB" }),
              cell("好奇心旺盛・創造的・新しいことが好き", 2913),
              cell("慎重・現実的・安定志向", 2913),
              cell("↑↓", 600, { align: AlignmentType.CENTER }),
            ]
          }),
          new TableRow({
            children: [
              cell("C", 800, { align: AlignmentType.CENTER, bold: true }),
              cell("誠実性", 1800, { shade: "EAF2FB" }),
              cell("計画的・真面目・自己管理が高い", 2913),
              cell("自由・柔軟・その場対応が得意", 2913),
              cell("↑↓", 600, { align: AlignmentType.CENTER }),
            ]
          }),
          new TableRow({
            children: [
              cell("E", 800, { align: AlignmentType.CENTER, bold: true }),
              cell("外向性", 1800, { shade: "EAF2FB" }),
              cell("社交的・活発・人と話すのが好き", 2913),
              cell("内向的・じっくり型・一人時間が好き", 2913),
              cell("↑↓", 600, { align: AlignmentType.CENTER }),
            ]
          }),
          new TableRow({
            children: [
              cell("A", 800, { align: AlignmentType.CENTER, bold: true }),
              cell("協調性", 1800, { shade: "EAF2FB" }),
              cell("思いやり・共感力が高い・協力的", 2913),
              cell("自己主張・独立心・競争心が強い", 2913),
              cell("↑↓", 600, { align: AlignmentType.CENTER }),
            ]
          }),
          new TableRow({
            children: [
              cell("N", 800, { align: AlignmentType.CENTER, bold: true }),
              cell("神経症傾向", 1800, { shade: "EAF2FB" }),
              cell("繊細・感じやすい・感情が豊か", 2913),
              cell("おおらか・安定・動じにくい", 2913),
              cell("↑↓", 600, { align: AlignmentType.CENTER }),
            ]
          }),
        ]
      }),
      spacer(),
      body("※ どの特性が「良い・悪い」ということはありません。場面や環境によって求められる特性が変わります。", { indent: true }),

      spacer(),

      // ─── 3. ゲームの流れ ───
      h1("3. ゲームの流れ"),

      h2("STEP 1｜ビッグファイブ診断"),
      bullet("ゲーム開始前に全プレイヤーが診断テストを受ける（約30問）"),
      bullet("5つの特性がそれぞれ 0〜100 のスコアで表示される"),
      bullet("これがそのままキャラクターの「初期ステータス」になる"),

      spacer(),
      h2("STEP 2｜バトルフィールドを選ぶ"),
      bullet("今回の「舞台」となる職種・環境を全員で選ぶ"),
      bullet("各フィールドには「勝利に必要な理想ステータス」が設定されている"),
      bullet("例：営業職 → E(外向性)とA(協調性)が高い人が有利"),

      spacer(),
      h2("STEP 3｜カードで特性を変化させる"),
      bullet("手番ごとにカードを引き、自分または相手に使う"),
      bullet("【自己強化カード】自分の特性を上げる（例：挨拶カード → E +2）"),
      bullet("【妨害カード】相手の特性を下げる（例：ジャンクフードカード → 相手のC -2）"),

      spacer(),
      h2("STEP 4｜勝利判定"),
      bullet("一定ターン後、各プレイヤーのステータスとフィールドの理想値を比較"),
      bullet("理想値に最も近いプレイヤーが勝利"),

      spacer(),

      // ─── 4. バトルフィールド ───
      h1("4. バトルフィールド（職種別 理想ステータス）"),
      body("※ 数値はイメージ。今後プレイテストで調整予定。"),
      spacer(),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 1405, 1405, 1405, 1405, 1406],
        rows: [
          new TableRow({
            children: [
              headerCell("フィールド", 2000),
              headerCell("O 開放性", 1405),
              headerCell("C 誠実性", 1405),
              headerCell("E 外向性", 1405),
              headerCell("A 協調性", 1405),
              headerCell("N 神経症傾向", 1406),
            ]
          }),
          new TableRow({ children: [cell("起業家・資本主義", 2000, { shade: "FFF2CC" }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("低", 1405, { align: AlignmentType.CENTER }), cell("低", 1406, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("営業職", 2000, { shade: "E2EFDA" }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("低", 1406, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("研究者・博士", 2000, { shade: "DDEEFF" }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("低", 1405, { align: AlignmentType.CENTER }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("低", 1406, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("カウンセラー", 2000, { shade: "FCE4D6" }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("高", 1406, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("クリエイター・芸術家", 2000, { shade: "EAD1F5" }), cell("高", 1405, { align: AlignmentType.CENTER }), cell("低", 1405, { align: AlignmentType.CENTER }), cell("低", 1405, { align: AlignmentType.CENTER }), cell("中", 1405, { align: AlignmentType.CENTER }), cell("中", 1406, { align: AlignmentType.CENTER })] }),
        ]
      }),

      spacer(),

      // ─── 5. カードの種類 ───
      h1("5. カードの種類（案）"),

      h2("自己強化カード（自分の特性を上げる）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 2500, 2013, 2013],
        rows: [
          new TableRow({ children: [headerCell("カード名", 2500), headerCell("効果", 2500), headerCell("対象特性", 2013), headerCell("変化量", 2013)] }),
          new TableRow({ children: [cell("挨拶カード", 2500), cell("毎日挨拶を続ける行動を取る", 2500), cell("E（外向性）", 2013), cell("+2", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("読書カード", 2500), cell("毎日30分読書する", 2500), cell("O（開放性）", 2013), cell("+2", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("瞑想カード", 2500), cell("毎朝10分瞑想する", 2500), cell("N（神経症傾向）", 2013), cell("-3", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("手帳カード", 2500), cell("スケジュール管理を始める", 2500), cell("C（誠実性）", 2013), cell("+3", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("ボランティアカード", 2500), cell("地域活動に参加する", 2500), cell("A（協調性）", 2013), cell("+2", 2013, { align: AlignmentType.CENTER })] }),
        ]
      }),

      spacer(),

      h2("妨害カード（相手の特性を下げる）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2500, 2500, 2013, 2013],
        rows: [
          new TableRow({ children: [headerCell("カード名", 2500), headerCell("効果", 2500), headerCell("対象特性", 2013), headerCell("変化量", 2013)] }),
          new TableRow({ children: [cell("ジャンクフードカード", 2500), cell("体調管理ができなくなる", 2500), cell("C（誠実性）", 2013), cell("-2", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("夜更かしカード", 2500), cell("睡眠不足で感情が不安定に", 2500), cell("N（神経症傾向）", 2013), cell("+3", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("SNS依存カード", 2500), cell("集中力が下がり誠実性・外向性が低下", 2500), cell("C・E", 2013), cell("-2/-1", 2013, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("孤立カード", 2500), cell("人間関係が薄れて協調性が下がる", 2500), cell("A（協調性）", 2013), cell("-2", 2013, { align: AlignmentType.CENTER })] }),
        ]
      }),

      spacer(),

      // ─── 6. 教育的意義 ───
      h1("6. このゲームが目指すもの"),
      bullet("ビッグファイブ理論を「楽しみながら」自然に覚えられる"),
      bullet("自分の強み・弱みを数値で客観視できるようになる"),
      bullet("カードの行動 ＝ 現実の行動と紐づいているため、ゲーム後に現実でも実践しやすい"),
      bullet("「どの特性が良い・悪い」ではなく、場面によって求められる特性が違うと理解できる"),
      bullet("子供から大人まで楽しめる教育的ゲームとして展開できる"),

      spacer(),

      // ─── 7. 今後の開発ステップ ───
      h1("7. 今後の開発ステップ"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1200, 4413, 3413],
        rows: [
          new TableRow({ children: [headerCell("フェーズ", 1200), headerCell("内容", 4413), headerCell("成果物", 3413)] }),
          new TableRow({ children: [cell("Phase 1", 1200, { shade: "EAF2FB" }), cell("カードリストの作成（自己強化・妨害の全カード）", 4413), cell("カードリスト一覧表", 3413)] }),
          new TableRow({ children: [cell("Phase 2", 1200, { shade: "EAF2FB" }), cell("ビッグファイブ診断テストの設計（子ども向け）", 4413), cell("診断テスト問題集", 3413)] }),
          new TableRow({ children: [cell("Phase 3", 1200, { shade: "EAF2FB" }), cell("バトルフィールドの詳細設計（勝利条件の数値化）", 4413), cell("フィールド設計書", 3413)] }),
          new TableRow({ children: [cell("Phase 4", 1200, { shade: "EAF2FB" }), cell("プロトタイプ作成・テストプレイ", 4413), cell("紙製カードセット", 3413)] }),
          new TableRow({ children: [cell("Phase 5", 1200, { shade: "EAF2FB" }), cell("フィードバックをもとに改善・デジタル化検討", 4413), cell("改訂版 / アプリ企画", 3413)] }),
        ]
      }),

      spacer(),
      spacer(),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "（このドキュメントは随時更新予定）", size: 18, font: "メイリオ", color: "AAAAAA" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "C:\\Users\\user\\Desktop\\Claude Code\\ビッグファイブカードゲーム\\ビッグファイブカードゲーム_企画書v0.1.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("作成完了:", outPath);
});
