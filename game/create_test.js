const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ── 共通スタイル ──────────────────────────────
const bdr = (c) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c) => ({ top: bdr(c), bottom: bdr(c), left: bdr(c), right: bdr(c) });
const margins = { top: 100, bottom: 100, left: 140, right: 140 };

function title(text, color = "2E4057") {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, bold: true, size: 52, font: "メイリオ", color })]
  });
}
function subtitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text, size: 24, font: "メイリオ", color: "888888" })]
  });
}
function h1(text, color = "2E4057") {
  return new Paragraph({
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, font: "メイリオ", color })]
  });
}
function h2(text, color = "4472C4") {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "メイリオ", color })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: opts.indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: opts.size || 22, font: "メイリオ", bold: opts.bold, color: opts.color })]
  });
}
function spacer(n = 120) {
  return new Paragraph({ spacing: { before: n, after: n }, children: [new TextRun("")] });
}
function cell(text, width, opts = {}) {
  return new TableCell({
    borders: borders(opts.bc || "CCCCCC"),
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text, size: opts.size || 20, font: "メイリオ", bold: opts.bold, color: opts.color || "000000" })]
    })]
  });
}
function hcell(text, width, fill = "2E4057") {
  return new TableCell({
    borders: borders(fill),
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, size: 20, font: "メイリオ", bold: true, color: "FFFFFF" })]
    })]
  });
}

// ── 大人向け質問（20問 / 特性ごとに4問）─────────────────
// ※ BFI-10・TIPI-Jを参照し、学術的根拠のある内容をもとに独自設計
// （R）= 逆転項目（採点時に 6-点数 で換算）
const adultQuestions = [
  // O 開放性
  { no: 1,  trait: "O", text: "新しいアイデアや考え方に興味がある",           rev: false },
  { no: 2,  trait: "O", text: "芸術・音楽・文学など創造的なものに引かれる",   rev: false },
  { no: 3,  trait: "O", text: "複雑な問題を深く考えることが好きだ",           rev: false },
  { no: 4,  trait: "O", text: "平凡で型にはまったやり方の方が安心できる",     rev: true  },
  // C 誠実性
  { no: 5,  trait: "C", text: "計画を立て、それに沿って物事を進める方だ",     rev: false },
  { no: 6,  trait: "C", text: "責任感があり、任されたことはやり遂げる",       rev: false },
  { no: 7,  trait: "C", text: "身の回りの整理整頓をきちんとするほうだ",       rev: false },
  { no: 8,  trait: "C", text: "締め切りや約束を守ることが苦手だ",             rev: true  },
  // E 外向性
  { no: 9,  trait: "E", text: "人と話したり、交流したりするのが好きだ",       rev: false },
  { no: 10, trait: "E", text: "初対面の人にも自分から話しかけられる",         rev: false },
  { no: 11, trait: "E", text: "グループの中でリーダー的な役割をとりやすい",   rev: false },
  { no: 12, trait: "E", text: "大勢でいるより、一人でいる方が気楽だ",         rev: true  },
  // A 協調性
  { no: 13, trait: "A", text: "他人の気持ちに共感しやすい方だ",               rev: false },
  { no: 14, trait: "A", text: "困っている人がいたら、進んで助けようとする",   rev: false },
  { no: 15, trait: "A", text: "意見が対立しても、穏やかに話し合える",         rev: false },
  { no: 16, trait: "A", text: "自分の意見を通すために、強引になることがある", rev: true  },
  // N 神経症傾向
  { no: 17, trait: "N", text: "些細なことでも不安になりやすい",               rev: false },
  { no: 18, trait: "N", text: "ストレスがかかるとパニックになりやすい",       rev: false },
  { no: 19, trait: "N", text: "気分の波が激しい方だ",                         rev: false },
  { no: 20, trait: "N", text: "プレッシャーがかかってもどっしりと構えていられる", rev: true },
];

// ── 子ども向け質問（15問 / 特性ごとに3問）─────────────────
const childQuestions = [
  { no: 1,  trait: "O", text: "知らないことを調べたり学んだりするのが好き",                 rev: false },
  { no: 2,  trait: "O", text: "自分で考えた遊びやアイデアを友達に話したくなる",             rev: false },
  { no: 3,  trait: "O", text: "新しいことより、いつも同じことの方が安心する",               rev: true  },
  { no: 4,  trait: "C", text: "宿題や約束したことは、きちんとやり遂げる",                   rev: false },
  { no: 5,  trait: "C", text: "部屋や持ち物を整理するのが得意だ",                           rev: false },
  { no: 6,  trait: "C", text: "やりたいことがあると、他のことをつい忘れてしまう",           rev: true  },
  { no: 7,  trait: "E", text: "友達と話したり一緒に遊んだりするのが好き",                   rev: false },
  { no: 8,  trait: "E", text: "知らない人にも、自分から話しかけることができる",             rev: false },
  { no: 9,  trait: "E", text: "大人数よりも、一人か少人数でいる方が落ち着く",               rev: true  },
  { no: 10, trait: "A", text: "友達が悲しんでいると、自分もつらい気持ちになる",             rev: false },
  { no: 11, trait: "A", text: "困っている人がいたら、助けたいと思う",                       rev: false },
  { no: 12, trait: "A", text: "自分の意見が通らないと、カッとなることがある",               rev: true  },
  { no: 13, trait: "N", text: "失敗したり怒られたりすると、長い間気にしてしまう",           rev: false },
  { no: 14, trait: "N", text: "ちょっとしたことでも不安になったり心配になったりする",       rev: false },
  { no: 15, trait: "N", text: "嫌なことがあっても、すぐに気持ちを切り替えられる",           rev: true  },
];

const traitColor = { O: "5B9BD5", C: "70AD47", E: "ED7D31", A: "FFC000", N: "7030A0" };
const traitName  = { O: "O：開放性", C: "C：誠実性", E: "E：外向性", A: "A：協調性", N: "N：神経症傾向" };

// ── 大人向け質問テーブル ─────────────────────────────────
function adultTable() {
  const rows = [];
  // ヘッダー行
  rows.push(new TableRow({
    tableHeader: true,
    children: [
      hcell("No.", 600),
      hcell("特性", 1400),
      hcell("質問文", 4500),
      hcell("1\n全くそう思わない", 720),
      hcell("2\nそう思わない", 720),
      hcell("3\nどちらとも\nいえない", 720),
      hcell("4\nそう思う", 720),
      hcell("5\n非常にそう思う", 726),
    ]
  }));

  adultQuestions.forEach((q, i) => {
    const isNew = i === 0 || q.trait !== adultQuestions[i - 1].trait;
    const shade = i % 2 === 0 ? "F8F8F8" : "FFFFFF";
    rows.push(new TableRow({
      children: [
        cell(String(q.no), 600, { align: AlignmentType.CENTER, fill: shade, size: 19 }),
        cell(traitName[q.trait] + (q.rev ? "  ＊" : ""), 1400, {
          fill: shade, size: 18, color: traitColor[q.trait], bold: isNew
        }),
        cell(q.text + (q.rev ? "　＊" : ""), 4500, { fill: shade, size: 20 }),
        cell("", 720, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 720, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 720, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 720, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 726, { align: AlignmentType.CENTER, fill: shade }),
      ]
    }));
  });
  return new Table({ width: { size: 9106, type: WidthType.DXA }, columnWidths: [600, 1400, 4500, 720, 720, 720, 720, 726], rows });
}

// ── 子ども向け質問テーブル ───────────────────────────────
function childTable() {
  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: [
      hcell("No.", 600, "2E6099"),
      hcell("しつもん文", 6000, "2E6099"),
      hcell("ちがう\n😟", 840),
      hcell("どちらでも\nない 😐", 840),
      hcell("そう思う\n😊", 826),
    ]
  }));

  childQuestions.forEach((q, i) => {
    const isNew = i === 0 || q.trait !== childQuestions[i - 1].trait;
    const shade = i % 2 === 0 ? "F0F7FF" : "FFFFFF";
    rows.push(new TableRow({
      children: [
        cell(String(q.no), 600, { align: AlignmentType.CENTER, fill: shade, bold: true }),
        cell(q.text + (q.rev ? "　＊" : ""), 6000, { fill: shade, size: 21 }),
        cell("", 840, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 840, { align: AlignmentType.CENTER, fill: shade }),
        cell("", 826, { align: AlignmentType.CENTER, fill: shade }),
      ]
    }));
  });
  return new Table({ width: { size: 9106, type: WidthType.DXA }, columnWidths: [600, 6000, 840, 840, 826], rows });
}

// ── 採点表（大人）────────────────────────────────────────
function adultScoreTable() {
  const rows = [
    new TableRow({ children: [hcell("特性", 1200), hcell("正方向の問", 2000), hcell("逆転項目（6-点数）", 2000), hcell("スコア計算式", 3906)] })
  ];
  const data = [
    { t: "O 開放性",    pos: "Q1・Q2・Q3", rev: "Q4", formula: "（Q1+Q2+Q3+(6-Q4)）÷ 4" },
    { t: "C 誠実性",    pos: "Q5・Q6・Q7", rev: "Q8", formula: "（Q5+Q6+Q7+(6-Q8)）÷ 4" },
    { t: "E 外向性",    pos: "Q9・Q10・Q11", rev: "Q12", formula: "（Q9+Q10+Q11+(6-Q12)）÷ 4" },
    { t: "A 協調性",    pos: "Q13・Q14・Q15", rev: "Q16", formula: "（Q13+Q14+Q15+(6-Q16)）÷ 4" },
    { t: "N 神経症傾向", pos: "Q17・Q18・Q19", rev: "Q20", formula: "（Q17+Q18+Q19+(6-Q20)）÷ 4" },
  ];
  data.forEach((d, i) => {
    const shade = i % 2 === 0 ? "F8F8F8" : "FFFFFF";
    rows.push(new TableRow({ children: [
      cell(d.t, 1200, { fill: shade, bold: true }),
      cell(d.pos, 2000, { fill: shade }),
      cell(d.rev, 2000, { fill: shade, color: "C00000" }),
      cell(d.formula, 3906, { fill: shade, size: 19 }),
    ]}));
  });
  return new Table({ width: { size: 9106, type: WidthType.DXA }, columnWidths: [1200, 2000, 2000, 3906], rows });
}

// ── 採点表（子ども）──────────────────────────────────────
function childScoreTable() {
  const rows = [
    new TableRow({ children: [hcell("特性", 1400, "2E6099"), hcell("そう思う問", 2000, "2E6099"), hcell("逆転項目（4-点数）", 2000, "2E6099"), hcell("スコア計算式", 3706, "2E6099")] })
  ];
  const data = [
    { t: "O 開放性",    pos: "Q1・Q2", rev: "Q3",  formula: "（Q1+Q2+(4-Q3)）÷ 3" },
    { t: "C 誠実性",    pos: "Q4・Q5", rev: "Q6",  formula: "（Q4+Q5+(4-Q6)）÷ 3" },
    { t: "E 外向性",    pos: "Q7・Q8", rev: "Q9",  formula: "（Q7+Q8+(4-Q9)）÷ 3" },
    { t: "A 協調性",    pos: "Q10・Q11", rev: "Q12", formula: "（Q10+Q11+(4-Q12)）÷ 3" },
    { t: "N 神経症傾向", pos: "Q13・Q14", rev: "Q15", formula: "（Q13+Q14+(4-Q15)）÷ 3" },
  ];
  data.forEach((d, i) => {
    const shade = i % 2 === 0 ? "F0F7FF" : "FFFFFF";
    rows.push(new TableRow({ children: [
      cell(d.t, 1400, { fill: shade, bold: true }),
      cell(d.pos, 2000, { fill: shade }),
      cell(d.rev, 2000, { fill: shade, color: "C00000" }),
      cell(d.formula, 3706, { fill: shade, size: 19 }),
    ]}));
  });
  return new Table({ width: { size: 9106, type: WidthType.DXA }, columnWidths: [1400, 2000, 2000, 3706], rows });
}

// ── スコアシート（自己記入欄）────────────────────────────
function scoreSheet() {
  const rows = [
    new TableRow({ children: [hcell("特性", 1400), hcell("スコア（1〜5）", 2000), hcell("バー（塗りつぶし）", 5706)] })
  ];
  ["O 開放性", "C 誠実性", "E 外向性", "A 協調性", "N 神経症傾向"].forEach((t, i) => {
    const shade = i % 2 === 0 ? "F8F8F8" : "FFFFFF";
    rows.push(new TableRow({ children: [
      cell(t, 1400, { fill: shade, bold: true }),
      cell("", 2000, { fill: shade }),
      cell("", 5706, { fill: shade }),
    ]}));
  });
  return new Table({ width: { size: 9106, type: WidthType.DXA }, columnWidths: [1400, 2000, 5706], rows });
}

// ── ドキュメント本体 ─────────────────────────────────────
const doc = new Document({
  styles: { default: { document: { run: { font: "メイリオ", size: 22 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    children: [

      // ── 表紙ブロック ───────────────────────────────
      spacer(200),
      title("ビッグファイブ診断テスト"),
      subtitle("Big Five Personality Inventory"),
      spacer(80),
      body("このテストはBFI-10・TIPI-Jなどの学術的に検証された", { align: AlignmentType.CENTER }),
      body("ビッグファイブ尺度をもとに設計した診断テストです。", { align: AlignmentType.CENTER }),
      spacer(240),

      // ── 大人向け ─────────────────────────────────
      h1("▶ 大人向け診断テスト（20問）"),
      body("【回答方法】各質問を読み、自分にどのくらいあてはまるかを 1〜5 で○をつけてください。", { bold: true }),
      body("　1 = 全くそう思わない　　2 = そう思わない　　3 = どちらともいえない　　4 = そう思う　　5 = 非常にそう思う"),
      body("　※ ＊マークの質問は採点時に逆転します（後ろの採点表を参照）。"),
      spacer(60),
      adultTable(),
      spacer(80),

      // 採点方法
      h2("採点方法（大人向け）"),
      body("各特性のスコア＝ 正方向の点数の合計 ＋ 逆転項目（6－点数）の合計  ÷  4"),
      body("スコア範囲：1.0〜5.0　　数値が高いほどその特性が強い傾向を示します。"),
      spacer(40),
      adultScoreTable(),
      spacer(80),

      // スコアシート
      h2("マイスコアシート"),
      body("計算したスコアを記入してください（小数第1位まで）。"),
      spacer(40),
      scoreSheet(),
      spacer(200),

      // ── ページ区切り → 子ども向け ────────────────
      new Paragraph({ children: [new PageBreak()] }),

      h1("▶ 子ども向け診断テスト（15問）", "2E6099"),
      body("【答え方】それぞれの文を読んで、自分にあてはまる顔に○をつけてね！", { bold: true }),
      body("　😊 そう思う（3点）　　😐 どちらでもない（2点）　　😟 ちがう（1点）"),
      body("　※ ＊マークの問題は、採点するときにひっくり返すよ（後ろの採点表を見てね）。"),
      spacer(60),
      childTable(),
      spacer(80),

      // 採点方法（子ども）
      h2("採点方法（子ども向け）", "2E6099"),
      body("それぞれの特性のスコア＝ 正方向の点数 ＋ 逆転項目（4－点数）  ÷  3"),
      body("スコア範囲：1.0〜3.0　　数値が高いほどその特性が強い傾向があるよ。"),
      spacer(40),
      childScoreTable(),
      spacer(80),

      // ── 特性の解説 ───────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),

      h1("▶ 5つの特性の意味"),
      new Table({
        width: { size: 9106, type: WidthType.DXA },
        columnWidths: [500, 1500, 2553, 2553, 1000],
        rows: [
          new TableRow({ children: [hcell("略", 500), hcell("特性名", 1500), hcell("スコアが高いと", 2553), hcell("スコアが低いと", 2553), hcell("ゲームでの影響", 1000)] }),
          ...[
            ["O", "開放性", "好奇心旺盛・創造的・新しいものへの関心が高い", "慎重・現実的・安定志向・慣れたことが好き", "研究者・クリエイターで有利"],
            ["C", "誠実性", "計画的・真面目・自己管理が高い・責任感がある", "自由・柔軟・その場対応が得意・ルールに縛られない", "起業家・営業職で有利"],
            ["E", "外向性", "社交的・活発・人と話すのが好き・自己主張が強い", "内向的・じっくり型・一人時間が好き・深い思考力", "営業職・リーダー職で有利"],
            ["A", "協調性", "思いやり・共感力が高い・協力的・和を大切にする", "自己主張・独立心・競争心が強い・交渉が得意", "カウンセラー・チームで有利"],
            ["N", "神経症傾向", "繊細・感じやすい・感情が豊か・リスクに敏感", "おおらか・安定・動じにくい・プレッシャーに強い", "高いと起業家で不利"],
          ].map(([ab, name, hi, lo, g], i) => new TableRow({ children: [
            cell(ab, 500, { align: AlignmentType.CENTER, fill: i % 2 === 0 ? "F8F8F8" : "FFFFFF", bold: true }),
            cell(name, 1500, { fill: i % 2 === 0 ? "F8F8F8" : "FFFFFF", bold: true }),
            cell(hi, 2553, { fill: i % 2 === 0 ? "F8F8F8" : "FFFFFF", size: 19 }),
            cell(lo, 2553, { fill: i % 2 === 0 ? "F8F8F8" : "FFFFFF", size: 19 }),
            cell(g, 1000, { fill: i % 2 === 0 ? "F8F8F8" : "FFFFFF", size: 17, color: "444444" }),
          ]}))
        ]
      }),
      spacer(80),

      body("※ どの特性が「良い・悪い」ということはありません。場面によって求められる特性が変わります。", { indent: true, color: "888888" }),
      body("※ このテストはゲーム用に設計されたものです。臨床的な診断には使用しないでください。", { indent: true, color: "888888" }),
      spacer(80),

      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "ビッグファイブカードゲーム 診断テスト  ver.0.1  ／  2026年3月", size: 18, font: "メイリオ", color: "AAAAAA" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = "C:\\Users\\user\\Desktop\\Claude Code\\ビッグファイブカードゲーム\\ビッグファイブ_診断テスト_v0.1.docx";
  fs.writeFileSync(out, buf);
  console.log("作成完了:", out);
});
