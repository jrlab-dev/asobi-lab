const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak
} = require('docx');
const fs = require('fs');

const bdr = (c) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c = "CCCCCC") => ({ top: bdr(c), bottom: bdr(c), left: bdr(c), right: bdr(c) });
const pad = { top: 100, bottom: 100, left: 120, right: 120 };

function h1(text, color = "2E4057") {
  return new Paragraph({
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 34, font: "メイリオ", color })]
  });
}
function h2(text, color = "4472C4") {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "メイリオ", color })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: opts.indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: opts.size || 21, font: "メイリオ", bold: opts.bold, color: opts.color || "000000" })]
  });
}
function spacer(n = 100) {
  return new Paragraph({ spacing: { before: n, after: n }, children: [new TextRun("")] });
}

// スコアバーを文字で表現（●●●○○）
function bar(val, max = 5) {
  const filled = "●".repeat(val);
  const empty = "○".repeat(max - val);
  return filled + empty;
}

// スコアのセル色（1=赤系 ～ 5=緑系）
function scoreColor(val) {
  if (val <= 1) return "FFCCCC";
  if (val === 2) return "FFE5CC";
  if (val === 3) return "FFFFCC";
  if (val === 4) return "CCFFCC";
  return "99FF99";
}

function cell(text, width, opts = {}) {
  return new TableCell({
    borders: borders(opts.bc || "CCCCCC"),
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: pad,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text, size: opts.size || 20, font: "メイリオ",
        bold: opts.bold, color: opts.color || "000000"
      })]
    })]
  });
}
function hcell(text, width, fill = "2E4057") {
  return new TableCell({
    borders: borders(fill),
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: pad,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, size: 19, font: "メイリオ", bold: true, color: "FFFFFF" })]
    })]
  });
}

// ─── フィールドデータ ───────────────────────────────────
// O=開放性 C=誠実性 E=外向性 A=協調性 N=情緒安定性（Nは高い=安定）
// 1〜5スケール。フィールドの「理想値」に近いほど有利
const fields = [
  // ライトフィールド
  {
    category: "ライト", icon: "🔬", name: "研究者・博士",
    O: 5, C: 5, E: 2, A: 3, N: 2,
    winner: "内向的で好奇心旺盛な人",
    loser: "社交的すぎる人・飽きっぽい人",
    note: "一人で深く掘り下げる力が最大の武器。外向性が高いと気が散る。",
    color: "1F4E79"
  },
  {
    category: "ライト", icon: "🧘", name: "カウンセラー・福祉",
    O: 4, C: 3, E: 3, A: 5, N: 3,
    winner: "共感力が高く、感受性が豊かな人",
    loser: "競争心が強い・冷たい人",
    note: "相手の気持ちに寄り添う協調性が核心。感受性も必要なのでN中程度が◎。",
    color: "375623"
  },
  {
    category: "ライト", icon: "🌾", name: "農家・職人・伝統技術",
    O: 2, C: 5, E: 2, A: 3, N: 1,
    winner: "黙々とやり遂げる・ルーティンが好きな人",
    loser: "飽きっぽい・感情の波が激しい人",
    note: "毎日同じ作業を完璧にやり続ける誠実性と情緒安定性が命。新しさは不要。",
    color: "375623"
  },
  {
    category: "ライト", icon: "🎨", name: "芸術家・クリエイター",
    O: 5, C: 2, E: 3, A: 3, N: 4,
    winner: "感受性が豊か・ルールに縛られない人",
    loser: "真面目すぎる・型にはまった人",
    note: "感情の豊かさ（N高）が創造力の源。誠実性が高いと枠に収まりすぎる。",
    color: "7030A0"
  },
  {
    category: "ライト", icon: "👩‍⚕️", name: "看護師・介護士",
    O: 3, C: 4, E: 3, A: 5, N: 3,
    winner: "思いやりがあり、責任感がある人",
    loser: "自己中心的・感情的すぎる人",
    note: "協調性と誠実性のバランスが大事。感受性はあるが流されすぎない安定感も必要。",
    color: "375623"
  },
  // ミドルフィールド
  {
    category: "ミドル", icon: "🚀", name: "起業家・経営者",
    O: 4, C: 4, E: 4, A: 2, N: 2,
    winner: "行動力・競争心・実行力を持つ人",
    loser: "人の顔色ばかり気にする・慎重すぎる人",
    note: "協調性が低くても自分の判断で動ける強さが必要。競争に勝つには情緒安定性も重要。",
    color: "C55A11"
  },
  {
    category: "ミドル", icon: "💼", name: "営業職・セールス",
    O: 3, C: 3, E: 5, A: 4, N: 2,
    winner: "社交的で人に好かれる・断られても立ち直れる人",
    loser: "内向的・すぐ落ち込む人",
    note: "外向性と情緒安定性のバランスが勝負。協調性も必要だが高すぎると押しが弱くなる。",
    color: "C55A11"
  },
  {
    category: "ミドル", icon: "🎭", name: "俳優・芸人・エンタメ",
    O: 4, C: 2, E: 5, A: 3, N: 4,
    winner: "感情豊か・目立ちたがり・感受性が強い人",
    loser: "真面目すぎる・おとなしすぎる人",
    note: "感情の波（N高）が表現力になる。誠実性が低いほど型にはまらない演技ができる。",
    color: "7030A0"
  },
  {
    category: "ミドル", icon: "📱", name: "インフルエンサー・配信者",
    O: 4, C: 3, E: 5, A: 3, N: 4,
    winner: "感情を表現できる・共感を呼ぶ人",
    loser: "感情を表に出さない・目立つのが嫌いな人",
    note: "感情の豊かさと外向性が視聴者を引きつける。芸術家に近いプロフィール。",
    color: "7030A0"
  },
  {
    category: "ミドル", icon: "⚽", name: "プロスポーツ選手",
    O: 3, C: 5, E: 4, A: 3, N: 1,
    winner: "メンタルが強く・ストイックに練習できる人",
    loser: "プレッシャーに弱い・練習をさぼる人",
    note: "情緒安定性（プレッシャーに強い）と誠実性（自己管理）が最も重要なフィールド。",
    color: "C55A11"
  },
  // グレーフィールド（⚠️注記あり）
  {
    category: "グレー", icon: "🗳️", name: "政治家・官僚",
    O: 3, C: 3, E: 5, A: 2, N: 1,
    winner: "カリスマ性があり・感情に左右されない人",
    loser: "人の気持ちを優先しすぎる・感情的になる人",
    note: "⚠️ 協調性が低いほど「冷酷な決断」ができる。これが政治の現実でもある。",
    color: "843C0C", warn: true
  },
  {
    category: "グレー", icon: "⚖️", name: "弁護士・検察官",
    O: 4, C: 5, E: 4, A: 2, N: 2,
    winner: "論理的・精密・感情より事実で動く人",
    loser: "感情移入しすぎる・論理より感情の人",
    note: "⚠️ 相手に共感しすぎると弱くなる。感情を切り離す力が勝負を決める。",
    color: "843C0C", warn: true
  },
  {
    category: "グレー", icon: "🏥", name: "外科医・救急医",
    O: 3, C: 5, E: 3, A: 2, N: 1,
    winner: "精密で・感情を切り離して判断できる人",
    loser: "患者に感情移入しすぎる・プレッシャーに弱い人",
    note: "⚠️ 手術中に感情が入ると命取り。協調性・情緒安定性のバランスが問われる。",
    color: "843C0C", warn: true
  },
  {
    category: "グレー", icon: "📈", name: "証券トレーダー・投資家",
    O: 3, C: 4, E: 3, A: 2, N: 1,
    winner: "感情ゼロで判断できる・リスクを数字で見られる人",
    loser: "損失に動揺する・他人の意見に流される人",
    note: "⚠️ 人への共感は不要、むしろ邪魔。感情に動かされない人が市場で勝つ。",
    color: "843C0C", warn: true
  },
  {
    category: "グレー", icon: "🤝", name: "ネットワークビジネス",
    O: 3, C: 3, E: 5, A: 2, N: 1,
    winner: "断られても動じない・押しが強い人",
    loser: "相手の気持ちを考えすぎる人",
    note: "⚠️ このフィールドで勝てる特性は、現実では人間関係を壊すリスクがある。\n    「勝てる≠幸せになれる」を考えるきっかけに。",
    color: "843C0C", warn: true
  },
];

// ─── メインテーブル ──────────────────────────────────────
function mainTable() {
  const colW = [260, 1600, 620, 620, 620, 620, 620, 2946];
  const total = colW.reduce((a, b) => a + b, 0); // 7906
  const rows = [];

  rows.push(new TableRow({
    tableHeader: true,
    children: [
      hcell("", 260),
      hcell("フィールド名", 1600),
      hcell("O\n開放性", 620),
      hcell("C\n誠実性", 620),
      hcell("E\n外向性", 620),
      hcell("A\n協調性", 620),
      hcell("N\n情緒安定性", 620),
      hcell("有利な人・不利な人", 2946),
    ]
  }));

  fields.forEach((f, i) => {
    const shade = i % 2 === 0 ? "F8F8F8" : "FFFFFF";
    rows.push(new TableRow({
      children: [
        cell(f.icon, 260, { align: AlignmentType.CENTER, fill: shade, size: 22 }),
        new TableCell({
          borders: borders(),
          width: { size: 1600, type: WidthType.DXA },
          shading: { fill: shade, type: ShadingType.CLEAR },
          margins: pad,
          children: [
            new Paragraph({ children: [new TextRun({ text: f.name, size: 20, font: "メイリオ", bold: true })] }),
            new Paragraph({ children: [new TextRun({ text: f.category === "グレー" ? "⚠️ グレーゾーン" : f.category + "フィールド", size: 16, font: "メイリオ", color: f.category === "グレー" ? "C55A11" : f.category === "ミドル" ? "2E6099" : "375623" })] }),
          ]
        }),
        cell(String(f.O) + "\n" + bar(f.O), 620, { align: AlignmentType.CENTER, fill: scoreColor(f.O), size: 16 }),
        cell(String(f.C) + "\n" + bar(f.C), 620, { align: AlignmentType.CENTER, fill: scoreColor(f.C), size: 16 }),
        cell(String(f.E) + "\n" + bar(f.E), 620, { align: AlignmentType.CENTER, fill: scoreColor(f.E), size: 16 }),
        cell(String(f.A) + "\n" + bar(f.A), 620, { align: AlignmentType.CENTER, fill: scoreColor(f.A), size: 16 }),
        cell(String(f.N) + "\n" + bar(f.N), 620, { align: AlignmentType.CENTER, fill: scoreColor(f.N), size: 16 }),
        new TableCell({
          borders: borders(),
          width: { size: 2946, type: WidthType.DXA },
          shading: { fill: shade, type: ShadingType.CLEAR },
          margins: pad,
          children: [
            new Paragraph({ children: [new TextRun({ text: "✅ 有利：" + f.winner, size: 18, font: "メイリオ", color: "375623" })] }),
            new Paragraph({ children: [new TextRun({ text: "❌ 不利：" + f.loser, size: 18, font: "メイリオ", color: "C00000" })] }),
          ]
        }),
      ]
    }));
  });

  return new Table({ width: { size: 7906, type: WidthType.DXA }, columnWidths: colW, rows });
}

// ─── 個別フィールドカード（詳細）────────────────────────
function fieldCard(f) {
  const catColor = f.category === "グレー" ? "843C0C" : f.category === "ミドル" ? "2E4057" : "375623";
  const catLabel = f.category === "グレー" ? "⚠️ グレーゾーンフィールド" : f.category + "フィールド";
  return [
    new Paragraph({
      spacing: { before: 280, after: 80 },
      children: [
        new TextRun({ text: f.icon + " " + f.name + "　", bold: true, size: 28, font: "メイリオ", color: catColor }),
        new TextRun({ text: "【" + catLabel + "】", size: 20, font: "メイリオ", color: catColor }),
      ]
    }),
    new Table({
      width: { size: 7906, type: WidthType.DXA },
      columnWidths: [900, 900, 900, 900, 900, 3406],
      rows: [
        new TableRow({ children: [
          hcell("O 開放性", 900, "#" + "5B9BD5".replace("#","")),
          hcell("C 誠実性", 900, "70AD47"),
          hcell("E 外向性", 900, "ED7D31"),
          hcell("A 協調性", 900, "FFC000"),
          hcell("N 情緒安定性", 900, "7030A0"),
          hcell("解説", 3406, catColor),
        ]}),
        new TableRow({ children: [
          cell(String(f.O) + "  " + bar(f.O), 900, { align: AlignmentType.CENTER, fill: scoreColor(f.O), bold: true }),
          cell(String(f.C) + "  " + bar(f.C), 900, { align: AlignmentType.CENTER, fill: scoreColor(f.C), bold: true }),
          cell(String(f.E) + "  " + bar(f.E), 900, { align: AlignmentType.CENTER, fill: scoreColor(f.E), bold: true }),
          cell(String(f.A) + "  " + bar(f.A), 900, { align: AlignmentType.CENTER, fill: scoreColor(f.A), bold: true }),
          cell(String(f.N) + "  " + bar(f.N), 900, { align: AlignmentType.CENTER, fill: scoreColor(f.N), bold: true }),
          new TableCell({
            borders: borders(),
            width: { size: 3406, type: WidthType.DXA },
            shading: { fill: "FAFAFA", type: ShadingType.CLEAR },
            margins: pad,
            children: [
              new Paragraph({ children: [new TextRun({ text: "✅ " + f.winner, size: 18, font: "メイリオ", color: "375623" })] }),
              new Paragraph({ children: [new TextRun({ text: "❌ " + f.loser, size: 18, font: "メイリオ", color: "C00000" })] }),
              new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: f.note, size: 17, font: "メイリオ", color: f.warn ? "C55A11" : "444444" })] }),
            ]
          }),
        ]}),
      ]
    }),
  ];
}

// ─── 特性×フィールド マトリクス ─────────────────────────
// 「この特性が高い人が有利なフィールド一覧」
function matrixTable() {
  const traits = ["O 開放性", "C 誠実性", "E 外向性", "A 協調性", "N 情緒安定性"];
  const traitKeys = ["O", "C", "E", "A", "N"];
  const traitColors = ["5B9BD5", "70AD47", "ED7D31", "FFC000", "7030A0"];

  const rows = [];
  rows.push(new TableRow({ children: [
    hcell("特性", 1200),
    hcell("スコア5（高い）が有利なフィールド", 4353),
    hcell("スコア1〜2（低い）が有利なフィールド", 4353),
  ]}));

  traits.forEach((t, i) => {
    const key = traitKeys[i];
    const highFields = fields.filter(f => f[key] >= 4).map(f => f.icon + f.name).join("　");
    const lowFields  = fields.filter(f => f[key] <= 2).map(f => f.icon + f.name).join("　");
    rows.push(new TableRow({ children: [
      cell(t, 1200, { bold: true, fill: "EEF4FB", color: traitColors[i] }),
      cell(highFields || "（なし）", 4353, { size: 18 }),
      cell(lowFields  || "（なし）", 4353, { size: 18, color: "888888" }),
    ]}));
  });

  return new Table({ width: { size: 9906, type: WidthType.DXA }, columnWidths: [1200, 4353, 4353], rows });
}

// ─── ドキュメント ────────────────────────────────────────
const children = [
  // タイトル
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text: "フィールド設計書", bold: true, size: 52, font: "メイリオ", color: "2E4057" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "ビッグファイブ カードゲーム  ver.0.1　／　2026年3月", size: 22, font: "メイリオ", color: "888888" })]
  }),
  spacer(80),

  // 設計原則
  h1("設計原則"),
  body("① どの性格タイプでも「有利なフィールド」が必ず存在する", { bold: true }),
  body("② 一つのフィールドで全員に有利な性格は存在しない（バランス設計）"),
  body("③ グレーゾーンフィールドでは「勝てる≠幸せ・正しい」という気づきを与える"),
  body("④ 情緒安定性はN（神経症傾向）をポジティブに言い換えたもの（高い=安定・低い=敏感）"),
  spacer(60),
  body("スコアの見方：1〜5　　理想値に近いほど有利　　●＝その値　○＝空き", { color: "888888", size: 19 }),
  spacer(100),

  // 全フィールド一覧
  h1("全フィールド一覧（15種）"),
  mainTable(),
  spacer(100),

  // ページ区切り
  new Paragraph({ children: [new PageBreak()] }),

  // ライトフィールド詳細
  h1("ライトフィールド詳細", "375623"),
  body("誰もがめざしやすい・社会的に明確に良いとされる職業・環境。", { color: "375623" }),
  spacer(40),
  ...fields.filter(f => f.category === "ライト").flatMap(f => [...fieldCard(f), spacer(40)]),

  new Paragraph({ children: [new PageBreak()] }),

  // ミドルフィールド詳細
  h1("ミドルフィールド詳細", "2E4057"),
  body("実力主義・競争のある世界。ダークな特性も部分的に有効になるフィールド。"),
  spacer(40),
  ...fields.filter(f => f.category === "ミドル").flatMap(f => [...fieldCard(f), spacer(40)]),

  new Paragraph({ children: [new PageBreak()] }),

  // グレーフィールド詳細
  h1("⚠️ グレーゾーンフィールド詳細", "843C0C"),
  body("実在する職業だが、ダークトライアド的な特性が有利に働く場合がある。", { color: "843C0C" }),
  body("「勝てる特性＝現実でも使っていい」ではないことを伝えるフィールド。", { color: "843C0C", bold: true }),
  spacer(40),
  ...fields.filter(f => f.category === "グレー").flatMap(f => [...fieldCard(f), spacer(40)]),

  new Paragraph({ children: [new PageBreak()] }),

  // 特性×フィールド マトリクス
  h1("特性別 有利フィールドマトリクス"),
  body("「自分の強みはどのフィールドで活きるか」が一目でわかる一覧表。"),
  spacer(40),
  matrixTable(),
  spacer(120),

  new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text: "（このドキュメントは随時更新予定）", size: 18, font: "メイリオ", color: "AAAAAA" })]
  }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: "メイリオ", size: 21 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 16838, height: 11906 },
        margin: { top: 1100, right: 1100, bottom: 1100, left: 1100 },
        orientation: "landscape"
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = "C:\\Users\\user\\Desktop\\Claude Code\\ビッグファイブカードゲーム\\フィールド設計書_v0.1.docx";
  fs.writeFileSync(out, buf);
  console.log("作成完了:", out);
});
