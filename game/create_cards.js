const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak
} = require('docx');
const fs = require('fs');

const bdr = (c) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c = "CCCCCC") => ({ top: bdr(c), bottom: bdr(c), left: bdr(c), right: bdr(c) });
const pad = { top: 90, bottom: 90, left: 120, right: 120 };

function h1(text, color = "2E4057") {
  return new Paragraph({
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 34, font: "メイリオ", color })]
  });
}
function h2(text, color = "444444") {
  return new Paragraph({
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, bold: true, size: 26, font: "メイリオ", color })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 50, after: 50 },
    indent: opts.indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: opts.size || 21, font: "メイリオ", bold: opts.bold, color: opts.color || "333333" })]
  });
}
function spacer(n = 80) {
  return new Paragraph({ spacing: { before: n, after: n }, children: [new TextRun("")] });
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
      children: [new TextRun({ text, size: opts.size || 20, font: "メイリオ", bold: opts.bold, color: opts.color || "000000" })]
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

// 変化値を ±表示＋色で表現
function effectCell(val, width) {
  const text = val === 0 ? "－" : (val > 0 ? "+" + val : String(val));
  const fill = val > 0 ? (val >= 3 ? "CCFFCC" : "E8FFE8") :
               val < 0 ? (val <= -3 ? "FFCCCC" : "FFE8E8") : "F8F8F8";
  const color = val > 0 ? "006400" : val < 0 ? "C00000" : "888888";
  return new TableCell({
    borders: borders(),
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: pad,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, size: 22, font: "メイリオ", bold: val !== 0, color })]
    })]
  });
}

// ─── カードデータ ─────────────────────────────────────────
// O=開放性 C=誠実性 E=外向性 A=協調性 N=情緒安定性
// 攻撃カード：相手に使う（マイナス効果）
const attackCards = [
  {
    icon: "🍟", name: "ポテトチップス",
    O: 0, C: -2, E: 0, A: 0, N: -1,
    desc: "食べ始めると止まらない。自己管理が崩れ、気分も不安定になる。",
    real: "糖質・脂質の過剰摂取は前頭葉の機能を低下させ、自己制御力(C)に影響。"
  },
  {
    icon: "🍔", name: "ジャンクフード",
    O: 0, C: -2, E: -1, A: 0, N: -2,
    desc: "生活習慣が乱れ、エネルギーが低下。活力・意欲が失われていく。",
    real: "腸内環境の悪化→セロトニン減少→情緒不安定・無気力(E↓N↓)に繋がる研究あり。"
  },
  {
    icon: "🍰", name: "デザート三昧",
    O: 0, C: -1, E: 0, A: 0, N: -1,
    desc: "甘いものへの依存が高まり、衝動コントロールが弱くなる。",
    real: "血糖値の乱高下はイライラ・集中力低下を引き起こす。"
  },
  {
    icon: "📱", name: "ソーシャルゲーム",
    O: -1, C: -3, E: 0, A: 0, N: -2,
    desc: "ガチャ依存・時間浪費。自己管理が崩壊し、現実から遠ざかる。",
    real: "ゲーム依存はドーパミン回路を乱し、計画性(C)・感情調整(N)に悪影響。"
  },
  {
    icon: "📲", name: "SNS閲覧",
    O: 0, C: -2, E: 0, A: -1, N: -2,
    desc: "他人と比べて落ち込む。集中力も下がり、浅い人間関係が増える。",
    real: "SNS過剰利用→比較不安・自己肯定感低下(N↓)・表面的関係(A↓)の研究多数。"
  },
  {
    icon: "😴", name: "睡眠不足",
    O: 0, C: -2, E: -1, A: 0, N: -3,
    desc: "感情のコントロールが大幅に低下。判断力も鈍る。",
    real: "睡眠不足は扁桃体の過活動を引き起こし、感情調整(N)を最も強く損なう。"
  },
  {
    icon: "🛋️", name: "運動不足",
    O: 0, C: -1, E: -1, A: 0, N: -2,
    desc: "体も気分も重くなる。やる気が出ず、人付き合いも億劫になる。",
    real: "運動不足は脳内BDNFを低下させ、気分調整・認知機能に悪影響。"
  },
  {
    icon: "💸", name: "金銭欲求（強欲）",
    O: 0, C: -1, E: 0, A: -3, N: -1,
    desc: "お金のことしか考えられなくなり、他者への思いやりが薄れる。",
    real: "金銭動機が高まると協調性(A)が低下するという実験結果あり。"
  },
  {
    icon: "🎰", name: "ギャンブル",
    O: 0, C: -3, E: 0, A: 0, N: -2,
    desc: "衝動が抑えられなくなり、自己管理が最低レベルまで落ちる。",
    real: "依存性賭博は報酬系を乱し、前頭前野の制御機能(C)を著しく低下させる。"
  },
  {
    icon: "😤", name: "過労・残業強要",
    O: 0, C: -1, E: 0, A: -1, N: -3,
    desc: "心身が限界を超え、感情が不安定になり他者への余裕がなくなる。",
    real: "慢性過労はコルチゾール過剰分泌→情緒安定性(N)と協調性(A)が著しく低下。"
  },
  {
    icon: "🍺", name: "アルコール依存",
    O: 0, C: -2, E: 0, A: -1, N: -2,
    desc: "感情のコントロールが効かなくなり、人間関係にも影響が出る。",
    real: "アルコールは前頭前野の働きを弱め、自制心(C)・共感力(A)を低下させる。"
  },
  {
    icon: "🗣️", name: "批判・否定（受け続ける）",
    O: -1, C: -1, E: 0, A: -1, N: -3,
    desc: "自己肯定感が下がり、何をやっても意味がないと感じるようになる。",
    real: "慢性的な否定的フィードバックは学習性無力感を引き起こし(N↓↓)、開放性も失われる。"
  },
];

// 防御カード：自分に使う（プラス効果）
const defenseCards = [
  {
    icon: "🧘", name: "瞑想",
    O: 0, C: +1, E: 0, A: 0, N: +3,
    desc: "心が落ち着き、感情のコントロールが格段に上がる。",
    real: "8週間の瞑想でストレス反応が減少・前頭前野が厚くなるという研究あり。"
  },
  {
    icon: "🌿", name: "マインドフルネス",
    O: +1, C: 0, E: 0, A: +1, N: +2,
    desc: "今この瞬間に集中することで、物事への視野が広がり共感力が増す。",
    real: "マインドフルネスは扁桃体の反応を抑え、共感・開放性の向上が報告されている。"
  },
  {
    icon: "🏃", name: "有酸素運動",
    O: 0, C: +1, E: +1, A: 0, N: +2,
    desc: "気分が上がり、エネルギーと行動力が増す。継続で自信もつく。",
    real: "運動はセロトニン・BDNF増加→気分安定(N+)・活動性(E+)向上の科学的根拠あり。"
  },
  {
    icon: "🏋️", name: "ウエイトトレーニング",
    O: 0, C: +3, E: +1, A: 0, N: +2,
    desc: "ストイックに追い込む習慣が誠実性を大きく鍛える。自信もつく。",
    real: "筋トレの継続はストレス耐性・自己規律(C)を高め、テストステロン増加で情緒安定。"
  },
  {
    icon: "❤️", name: "心拍トレーニング（HIIT）",
    O: 0, C: +2, E: 0, A: 0, N: +2,
    desc: "高強度の運動は心肺機能とメンタルの強さを同時に高める。",
    real: "高強度インターバルがコルチゾール耐性を高め情緒安定(N+)・規律(C+)に効果。"
  },
  {
    icon: "💻", name: "プログラミング",
    O: +2, C: +3, E: 0, A: 0, N: +1,
    desc: "論理的思考と問題解決力が鍛えられ、粘り強さが身につく。",
    real: "プログラミング学習は計画性・論理思考(C+)と創造的問題解決(O+)を同時に強化。"
  },
  {
    icon: "🤝", name: "人助け・ボランティア",
    O: 0, C: 0, E: +1, A: +3, N: +1,
    desc: "人の役に立つ喜びが協調性と生きがいを高める。",
    real: "利他的行動はオキシトシン分泌→共感・協調性(A+)と幸福感(N+)が向上する研究多数。"
  },
  {
    icon: "📚", name: "読書",
    O: +3, C: +1, E: 0, A: 0, N: +1,
    desc: "知識と視野が広がり、好奇心と集中力が鍛えられる。",
    real: "読書（特に小説）は共感力・開放性(O+)を高め、集中習慣で誠実性(C+)にも寄与。"
  },
  {
    icon: "👋", name: "挨拶・コミュ習慣",
    O: 0, C: 0, E: +2, A: +1, N: 0,
    desc: "毎日の小さな交流が外向性と人間関係を自然に育てる。",
    real: "日常的な挨拶習慣は社会的接触を増やし、外向性(E+)・協調性(A+)を徐々に高める。"
  },
  {
    icon: "📓", name: "日記・自己分析",
    O: +2, C: +1, E: 0, A: 0, N: +1,
    desc: "自分の感情・思考を整理することで自己理解が深まる。",
    real: "感情日記は感情調整能力(N+)・自己洞察・開放性(O+)を高めることが示されている。"
  },
  {
    icon: "🙏", name: "感謝を伝える",
    O: 0, C: 0, E: 0, A: +2, N: +2,
    desc: "感謝の気持ちを表すと、心が豊かになり人間関係も温かくなる。",
    real: "グラティチュード実践は幸福感(N+)・社会的絆(A+)を高める実験結果が豊富。"
  },
  {
    icon: "🌅", name: "早起き・規則正しい生活",
    O: 0, C: +2, E: 0, A: 0, N: +1,
    desc: "生活リズムが整うと、計画的に行動できるようになる。",
    real: "概日リズムの安定は前頭前野機能を高め、自己管理(C+)と感情調整(N+)を改善。"
  },
  {
    icon: "🎯", name: "目標設定",
    O: +1, C: +3, E: 0, A: 0, N: 0,
    desc: "明確な目標を持つと行動が変わり、誠実性が大きく育つ。",
    real: "SMART目標設定は実行意図を形成し、誠実性(C+)の最も強い強化要因の一つ。"
  },
];

// ─── テーブル生成 ─────────────────────────────────────────
const COL = [400, 1500, 460, 460, 460, 460, 460, 2700, 2600];
const TOTAL = COL.reduce((a,b)=>a+b,0); // 9500

function cardTable(cards, headerColor) {
  const rows = [];
  rows.push(new TableRow({ tableHeader: true, children: [
    hcell("", 400, headerColor),
    hcell("カード名", 1500, headerColor),
    hcell("O\n開放性", 460, headerColor),
    hcell("C\n誠実性", 460, headerColor),
    hcell("E\n外向性", 460, headerColor),
    hcell("A\n協調性", 460, headerColor),
    hcell("N\n情緒\n安定性", 460, headerColor),
    hcell("ゲームでの説明文", 2700, headerColor),
    hcell("心理学的根拠（解説用）", 2600, headerColor),
  ]}));

  cards.forEach((c, i) => {
    const shade = i % 2 === 0 ? "F9F9F9" : "FFFFFF";
    rows.push(new TableRow({ children: [
      cell(c.icon, 400, { align: AlignmentType.CENTER, fill: shade, size: 22 }),
      cell(c.name, 1500, { fill: shade, bold: true, size: 20 }),
      effectCell(c.O, 460),
      effectCell(c.C, 460),
      effectCell(c.E, 460),
      effectCell(c.A, 460),
      effectCell(c.N, 460),
      cell(c.desc, 2700, { fill: shade, size: 18 }),
      cell(c.real, 2600, { fill: shade, size: 17, color: "555555" }),
    ]}));
  });

  return new Table({ width: { size: TOTAL, type: WidthType.DXA }, columnWidths: COL, rows });
}

// ─── バランスチェック表 ────────────────────────────────────
function balanceTable() {
  const traits = ["O", "C", "E", "A", "N"];
  const names  = ["O 開放性","C 誠実性","E 外向性","A 協調性","N 情緒安定性"];
  const cols = [1400, 1000, 1000, 1000, 1000, 3600];
  const total = cols.reduce((a,b)=>a+b,0);

  const rows = [new TableRow({ children: [
    hcell("特性", 1400), hcell("攻撃合計", 1000), hcell("防御合計", 1000),
    hcell("差し引き", 1000), hcell("バランス", 1000), hcell("コメント", 3600),
  ]})];

  const comments = {
    O: "開放性への攻撃は少なめ（批判とソーシャルゲームのみ）。読書・プログラミングで回復しやすい。",
    C: "誠実性は攻撃が多い（ジャンク食・ゲーム・睡眠不足など）。防御も多い。鍛えがいがある特性。",
    E: "外向性への影響は小さい。挨拶・運動で上げ、睡眠不足で下がる程度。",
    A: "協調性への攻撃は強烈（金銭欲求-3・批判-1など）。人助けで大きく回復できる。",
    N: "情緒安定性が最も攻撃を受けやすい。最も重要な防御ポイント。",
  };

  traits.forEach((t, i) => {
    const atk = attackCards.reduce((s, c) => s + (c[t] || 0), 0);
    const def = defenseCards.reduce((s, c) => s + (c[t] || 0), 0);
    const diff = atk + def;
    const balance = diff >= 0 ? "◎ 均衡" : diff >= -3 ? "○ やや攻撃多" : "△ 攻撃多め";
    const diffColor = diff >= 0 ? "006400" : diff >= -3 ? "C55A11" : "C00000";
    const shade = i % 2 === 0 ? "F9F9F9" : "FFFFFF";
    rows.push(new TableRow({ children: [
      cell(names[i], 1400, { fill: shade, bold: true }),
      cell(String(atk), 1000, { align: AlignmentType.CENTER, fill: "FFE8E8", color: "C00000", bold: true }),
      cell("+" + def, 1000, { align: AlignmentType.CENTER, fill: "E8FFE8", color: "006400", bold: true }),
      cell((diff >= 0 ? "+" : "") + diff, 1000, { align: AlignmentType.CENTER, fill: shade, color: diffColor, bold: true }),
      cell(balance, 1000, { align: AlignmentType.CENTER, fill: shade, color: diffColor }),
      cell(comments[t], 3600, { fill: shade, size: 18 }),
    ]}));
  });

  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: cols, rows });
}

// ─── ドキュメント ────────────────────────────────────────────
const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text: "攻撃・防御カード 効果値設計書", bold: true, size: 48, font: "メイリオ", color: "2E4057" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "ビッグファイブ カードゲーム  ver.0.1　／　2026年3月", size: 22, font: "メイリオ", color: "888888" })]
  }),
  spacer(80),

  body("【効果値の見方】　+3=大きく上昇　+2=上昇　+1=少し上昇　－=変化なし　-1=少し低下　-2=低下　-3=大きく低下", { bold: true }),
  body("【攻撃カード】相手に使う。相手の特性値が変化する。　【防御カード】自分に使う。自分の特性値が変化する。", { color: "666666" }),
  spacer(80),

  // ─ 攻撃カード ─
  h1("🔴 攻撃カード（12種）", "C00000"),
  body("相手に使うことで、相手のビッグファイブ特性値を下げるカード。", { color: "C00000" }),
  body("現実の「悪い習慣・環境」を題材にしている。心理学的根拠に基づいた効果値。"),
  spacer(40),
  cardTable(attackCards, "C00000"),
  spacer(100),

  new Paragraph({ children: [new PageBreak()] }),

  // ─ 防御カード ─
  h1("🔵 防御カード（13種）", "1F4E79"),
  body("自分に使うことで、自分のビッグファイブ特性値を上げるカード。", { color: "1F4E79" }),
  body("現実の「良い習慣・行動」を題材にしている。現実でも実践できる内容。"),
  spacer(40),
  cardTable(defenseCards, "1F4E79"),
  spacer(100),

  new Paragraph({ children: [new PageBreak()] }),

  // ─ バランス表 ─
  h1("⚖️ 攻撃・防御バランスチェック"),
  body("各特性への「攻撃の合計値」と「防御の合計値」を比較。ゲームバランスの確認用。"),
  spacer(40),
  balanceTable(),
  spacer(80),

  body("【調整メモ】", { bold: true }),
  body("・N（情緒安定性）が最も攻撃されやすく、防御も最重要。睡眠・瞑想・運動が鍵。", { indent: true }),
  body("・C（誠実性）はゲームの核心。多くの攻撃で崩れ、多くの防御で鍛えられる。", { indent: true }),
  body("・A（協調性）は「金銭欲求-3」が突出して強い。人助けカードで一発回復できるバランス。", { indent: true }),
  body("・E（外向性）とO（開放性）は変動が少なく、専用カードで特化できる特性。", { indent: true }),
  spacer(100),

  // ─ 試作PDFからの変更点メモ ─
  h1("📝 試作PDFからの変更点・検討事項"),
  new Table({
    width: { size: 9500, type: WidthType.DXA },
    columnWidths: [2000, 3750, 3750],
    rows: [
      new TableRow({ children: [hcell("項目", 2000), hcell("試作版", 3750), hcell("変更・理由", 3750)] }),
      ...[
        ["宗教家カード", "攻撃カードに配置", "➡ 状態カードへ移動を検討。宗教は特性変化の直接原因にしにくい。"],
        ["占い師カード", "攻撃カードに配置", "➡ 状態カードへ移動を検討。同上。"],
        ["ギャンブル", "なし", "➡ 新規追加。C-3の最強攻撃カードとして機能。"],
        ["アルコール依存", "なし", "➡ 新規追加。C・N・Aに影響する複合攻撃カード。"],
        ["批判・否定", "なし", "➡ 新規追加。教育的メッセージ（否定的環境の影響）として重要。"],
        ["過労・残業強要", "なし", "➡ 新規追加。社会問題（ブラック職場）と連動した攻撃カード。"],
        ["目標設定", "なし", "➡ 新規追加。C+3の最強防御カードとして機能。"],
        ["感謝を伝える", "なし", "➡ 新規追加。A+2・N+2の教育的に重要なカード。"],
      ].map(([a,b,c],i) => new TableRow({ children: [
        cell(a, 2000, { fill: i%2===0?"F9F9F9":"FFFFFF", bold: true }),
        cell(b, 3750, { fill: i%2===0?"F9F9F9":"FFFFFF" }),
        cell(c, 3750, { fill: i%2===0?"F9F9F9":"FFFFFF", size: 18 }),
      ]}))
    ]
  }),
  spacer(100),

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
        margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = "C:\\Users\\user\\Desktop\\Claude Code\\ビッグファイブカードゲーム\\攻撃防御カード設計書_v0.1.docx";
  fs.writeFileSync(out, buf);
  console.log("作成完了:", out);
});
