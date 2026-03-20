const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak, LevelFormat
} = require('docx');
const fs = require('fs');

const bdr = (c) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c = "CCCCCC") => ({ top: bdr(c), bottom: bdr(c), left: bdr(c), right: bdr(c) });
const pad = { top: 100, bottom: 100, left: 140, right: 140 };

function h1(text, color = "2E4057") {
  return new Paragraph({
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 36, font: "メイリオ", color })]
  });
}
function h2(text, color = "1F4E79") {
  return new Paragraph({
    spacing: { before: 260, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, font: "メイリオ", color })]
  });
}
function h3(text, color = "444444") {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, font: "メイリオ", color })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: opts.indent ? { left: 480 } : undefined,
    children: [new TextRun({ text, size: opts.size || 22, font: "メイリオ", bold: opts.bold, color: opts.color || "222222" })]
  });
}
function bullet(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 50, after: 50 },
    indent: { left: 480, hanging: 280 },
    children: [
      new TextRun({ text: "・", size: opts.size || 22, font: "メイリオ", color: opts.color || "444444" }),
      new TextRun({ text, size: opts.size || 22, font: "メイリオ", bold: opts.bold, color: opts.color || "222222" })
    ]
  });
}
function num(n, text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 60 },
    indent: { left: 500, hanging: 300 },
    children: [
      new TextRun({ text: String(n) + "．", size: opts.size || 22, font: "メイリオ", bold: true, color: opts.color || "1F4E79" }),
      new TextRun({ text, size: opts.size || 22, font: "メイリオ", color: "222222" })
    ]
  });
}
function spacer(n = 100) {
  return new Paragraph({ spacing: { before: n, after: n }, children: [new TextRun("")] });
}
function divider(color = "CCCCCC") {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color } }
  });
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

// ─── ボックス（重要事項を囲む）─────────────────────────────
function box(lines, fillColor = "EAF4FF", borderColor = "4472C4") {
  return new Table({
    width: { size: 8800, type: WidthType.DXA },
    columnWidths: [8800],
    rows: [new TableRow({ children: [new TableCell({
      borders: borders(borderColor),
      shading: { fill: fillColor, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: lines.map(l => new Paragraph({
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: l, size: 21, font: "メイリオ", color: "222222" })]
      }))
    }) ]})]
  });
}

const doc = new Document({
  styles: { default: { document: { run: { font: "メイリオ", size: 22 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1300, right: 1300, bottom: 1300, left: 1300 }
      }
    },
    children: [

      // ══ 表紙 ══════════════════════════════════════════════
      spacer(300),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "ビッグファイブ カードゲーム", bold: true, size: 56, font: "メイリオ", color: "2E4057" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "〜 自分を知り、自分を育てるゲーム 〜", size: 28, font: "メイリオ", color: "4472C4" })]
      }),
      spacer(200),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "ル ー ル ブ ッ ク", bold: true, size: 40, font: "メイリオ", color: "2E4057" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "ver. 0.1　／　2026年3月", size: 22, font: "メイリオ", color: "888888" })]
      }),
      spacer(200),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "対象年齢：10歳以上　　プレイ人数：2〜4人　　プレイ時間：30〜60分", size: 24, font: "メイリオ", color: "444444" })]
      }),
      spacer(60),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "（無料配布・印刷してご使用ください）", size: 20, font: "メイリオ", color: "888888" })]
      }),
      spacer(300),

      // ══ このゲームについて ═══════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("このゲームについて"),
      body("このゲームは「ビッグファイブ理論」という心理学の知識をもとに作られた、自己理解と成長のためのカードゲームです。"),
      spacer(40),
      body("ビッグファイブ理論とは、人間の性格を5つの特性で表す、世界中の心理学者が使っている理論です。"),
      spacer(60),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [440, 1400, 3480, 3480],
        rows: [
          new TableRow({ children: [hcell("略", 440), hcell("特性名", 1400), hcell("高いと…", 3480), hcell("低いと…", 3480)] }),
          ...[
            ["O", "開放性", "好奇心旺盛・創造的・新しいことが好き", "慎重・安定志向・いつものことが好き"],
            ["C", "誠実性", "計画的・真面目・自己管理が高い", "自由・柔軟・ルールに縛られない"],
            ["E", "外向性", "社交的・活発・人と話すのが好き", "内向的・じっくり型・一人時間が好き"],
            ["A", "協調性", "思いやり・共感力が高い・協力的", "自己主張・独立心・競争心が強い"],
            ["N", "情緒安定性", "落ち着いている・プレッシャーに強い", "繊細・感じやすい・感情が豊か"],
          ].map(([ab, name, hi, lo], i) => new TableRow({ children: [
            cell(ab, 440, { align: AlignmentType.CENTER, fill: i%2===0?"F0F7FF":"FFFFFF", bold: true, color: "2E4057" }),
            cell(name, 1400, { fill: i%2===0?"F0F7FF":"FFFFFF", bold: true }),
            cell(hi, 3480, { fill: i%2===0?"F0F7FF":"FFFFFF", size: 20 }),
            cell(lo, 3480, { fill: i%2===0?"F0F7FF":"FFFFFF", size: 20, color: "666666" }),
          ]}))
        ]
      }),
      spacer(80),
      box([
        "💡 大切なこと：どの特性が「良い・悪い」ということはありません。",
        "　　場面や職種によって、求められる特性が変わります。",
        "　　あなたの性格はあなたの強み。育て方次第で輝ける場所が見つかります。"
      ]),
      spacer(80),

      // ══ ゲームの目的 ═══════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("ゲームの目的"),
      body("ゲームの「バトルフィールド（舞台）」が持つ理想のビッグファイブ値に、", { bold: false }),
      body("自分のキャラクターの値をできるだけ近づけたプレイヤーが勝利します。"),
      spacer(60),
      box([
        "🎯 勝利条件：ゲーム終了時に、フィールドの理想値との差が最も小さいプレイヤーが勝ち",
        "（5つの特性すべての差の合計が少ないほど有利）"
      ], "FFFBE6", "FFC000"),
      spacer(80),

      // ══ 準備するもの ════════════════════════════════════════
      h1("準備するもの"),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [3000, 1200, 4600],
        rows: [
          new TableRow({ children: [hcell("アイテム", 3000), hcell("枚数", 1200), hcell("内容", 4600)] }),
          ...[
            ["ビッグファイブ診断テスト", "人数分", "ゲーム開始前に各自が回答する。スコアが自分のキャラクター初期値になる"],
            ["フィールドカード", "15枚", "「研究者」「営業職」「政治家」など、ゲームの舞台となるカード"],
            ["攻撃カード（赤）", "12種×2枚=24枚", "相手のビッグファイブ値を下げるカード"],
            ["防御カード（青）", "13種×2枚=26枚", "自分のビッグファイブ値を上げるカード"],
            ["状態カード（茶）", "各種1枚", "ビッグファイブ値が特定のパターンになったときに受け取るカード"],
            ["キャラクターシート", "人数分", "自分の現在のビッグファイブ値を記録するシート"],
            ["鉛筆・消しゴム", "人数分", "キャラクターシートに記入するため"],
          ].map(([a,b,c],i) => new TableRow({ children: [
            cell(a, 3000, { fill: i%2===0?"F9F9F9":"FFFFFF", bold: true }),
            cell(b, 1200, { fill: i%2===0?"F9F9F9":"FFFFFF", align: AlignmentType.CENTER }),
            cell(c, 4600, { fill: i%2===0?"F9F9F9":"FFFFFF", size: 20 }),
          ]}))
        ]
      }),
      spacer(80),

      // ══ ゲームの流れ ════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("ゲームの流れ"),

      h2("STEP 1｜診断テストを受ける"),
      body("ゲームを始める前に、全員がビッグファイブ診断テストに答えます。"),
      bullet("大人向け：20問（約5分）"),
      bullet("子ども向け：15問（約3分）"),
      bullet("5つの特性それぞれにスコアが出ます（1〜5の5段階）"),
      bullet("このスコアがそのままあなたのキャラクターの「初期値」になります"),
      spacer(60),
      box([
        "例）田中さんの初期値",
        "　O（開放性）：3　　C（誠実性）：4　　E（外向性）：2　　A（協調性）：4　　N（情緒安定性）：3",
      ], "E8F5E9", "70AD47"),
      spacer(80),

      h2("STEP 2｜フィールドカードを選ぶ"),
      body("フィールドカードをシャッフルして、1枚をオープンします。"),
      body("このカードがゲームの「舞台」と「勝利条件（理想値）」になります。"),
      spacer(40),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [2200, 1320, 1320, 1320, 1320, 1320],
        rows: [
          new TableRow({ children: [hcell("フィールド例", 2200, "375623"), hcell("O 開放性", 1320, "375623"), hcell("C 誠実性", 1320, "375623"), hcell("E 外向性", 1320, "375623"), hcell("A 協調性", 1320, "375623"), hcell("N 情緒安定性", 1320, "375623")] }),
          new TableRow({ children: [cell("🔬 研究者・博士", 2200, { bold: true }), cell("5 ●●●●●", 1320, { align: AlignmentType.CENTER }), cell("5 ●●●●●", 1320, { align: AlignmentType.CENTER }), cell("2 ●●○○○", 1320, { align: AlignmentType.CENTER }), cell("3 ●●●○○", 1320, { align: AlignmentType.CENTER }), cell("2 ●●○○○", 1320, { align: AlignmentType.CENTER })] }),
          new TableRow({ children: [cell("💼 営業職", 2200, { bold: true, fill: "F9F9F9" }), cell("3 ●●●○○", 1320, { align: AlignmentType.CENTER, fill: "F9F9F9" }), cell("3 ●●●○○", 1320, { align: AlignmentType.CENTER, fill: "F9F9F9" }), cell("5 ●●●●●", 1320, { align: AlignmentType.CENTER, fill: "F9F9F9" }), cell("4 ●●●●○", 1320, { align: AlignmentType.CENTER, fill: "F9F9F9" }), cell("2 ●●○○○", 1320, { align: AlignmentType.CENTER, fill: "F9F9F9" })] }),
        ]
      }),
      spacer(80),

      h2("STEP 3｜カードを配る"),
      bullet("攻撃カード（赤）と防御カード（青）を合わせてシャッフルします"),
      bullet("全員に5枚ずつ配ります"),
      bullet("残りは山札として中央に置きます"),
      spacer(80),

      h2("STEP 4｜ゲームスタート（手番のルール）"),
      body("時計回りで手番を行います。1回の手番でやることは以下の通りです。"),
      spacer(40),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [600, 2400, 5800],
        rows: [
          new TableRow({ children: [hcell("順番", 600), hcell("行動", 2400), hcell("詳細", 5800)] }),
          ...[
            ["①", "カードを1枚引く", "山札からカードを1枚引きます"],
            ["②", "カードを1枚使う", "手札から1枚選んで使います（使わなくてもOK）"],
            ["③", "手札を調整する", "手札が5枚を超えている場合は1枚捨て札にします"],
          ].map(([n,a,d],i) => new TableRow({ children: [
            cell(n, 600, { align: AlignmentType.CENTER, fill: i%2===0?"F9F9F9":"FFFFFF", bold: true, color: "1F4E79" }),
            cell(a, 2400, { fill: i%2===0?"F9F9F9":"FFFFFF", bold: true }),
            cell(d, 5800, { fill: i%2===0?"F9F9F9":"FFFFFF" }),
          ]}))
        ]
      }),
      spacer(80),

      // ══ カードの使い方 ═══════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("カードの使い方"),

      h2("🔴 攻撃カード（赤いカード）"),
      bullet("相手1人を選んで使います"),
      bullet("カードに書かれた特性値が、相手のキャラクターシートの値に加算されます（マイナス）"),
      bullet("使い終わったカードは捨て札にします"),
      spacer(40),
      box([
        "例）「睡眠不足」カード　効果：N -3、C -2、E -1",
        "　➡ 相手を選んで宣言：「田中さんに睡眠不足カードを使います！」",
        "　➡ 田中さんのキャラクターシートを書き換える",
        "　　N：3 → 0（最低値は0）　C：4 → 2　E：2 → 1",
      ], "FFF0F0", "C00000"),
      spacer(80),

      h2("🔵 防御カード（青いカード）"),
      bullet("自分自身に使います"),
      bullet("カードに書かれた特性値が、自分のキャラクターシートの値に加算されます（プラス）"),
      bullet("最大値は5（5を超えることはありません）"),
      bullet("使い終わったカードは捨て札にします"),
      spacer(40),
      box([
        "例）「読書」カード　効果：O +3、C +1、N +1",
        "　➡ 「読書カードを自分に使います！」と宣言",
        "　➡ 自分のキャラクターシートを書き換える",
        "　　O：3 → 5（最大値は5）　C：4 → 5　N：3 → 4",
      ], "F0F7FF", "1F4E79"),
      spacer(80),

      h2("特性値の範囲"),
      box([
        "・最小値は 0　　最大値は 5",
        "・攻撃で0を下回った場合は0のまま",
        "・防御で5を超えた場合は5のまま",
      ], "FAFAFA", "CCCCCC"),
      spacer(80),

      // ══ 状態カード ══════════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("状態カード（茶色いカード）"),
      body("攻撃を受け続けてキャラクターの特性値が特定のパターンになったとき、"),
      body("「状態カード」を受け取ります。状態カードは自分の前に置きます。"),
      spacer(60),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [1800, 1200, 1200, 1200, 1200, 1200, 2200],
        rows: [
          new TableRow({ children: [hcell("状態名", 1800, "5C3317"), hcell("O", 1200, "5C3317"), hcell("C", 1200, "5C3317"), hcell("E", 1200, "5C3317"), hcell("A", 1200, "5C3317"), hcell("N", 1200, "5C3317"), hcell("意味", 2200, "5C3317")] }),
          ...[
            ["サイコパス",     "4以上", "1以下", "4以上", "1以下", "3以上", "感情なく行動する"],
            ["ナルシスト",     "3以上", "2以下", "4以上", "1以下", "3以上", "自分しか見えない"],
            ["メンヘラ",       "3以上", "3以上", "1以下", "4以上", "1以下", "不安定で依存的"],
            ["ボトムギバー",   "2以下", "1以下", "3以上", "5以上", "2以下", "尽くしすぎて疲弊"],
            ["アベレージ",     "2前後", "3前後", "4前後", "3前後", "2前後", "平均的・不安定"],
            ["ロールモデル",   "5",     "5",     "5",     "5",     "5",     "理想的・全員の目標"],
          ].map(([n,o,c,e,a,ni,m],i) => new TableRow({ children: [
            cell(n, 1800, { fill: i%2===0?"FFF4EC":"FFFFFF", bold: true }),
            cell(o, 1200, { fill: i%2===0?"FFF4EC":"FFFFFF", align: AlignmentType.CENTER, size: 19 }),
            cell(c, 1200, { fill: i%2===0?"FFF4EC":"FFFFFF", align: AlignmentType.CENTER, size: 19 }),
            cell(e, 1200, { fill: i%2===0?"FFF4EC":"FFFFFF", align: AlignmentType.CENTER, size: 19 }),
            cell(a, 1200, { fill: i%2===0?"FFF4EC":"FFFFFF", align: AlignmentType.CENTER, size: 19 }),
            cell(ni, 1200, { fill: i%2===0?"FFF4EC":"FFFFFF", align: AlignmentType.CENTER, size: 19 }),
            cell(m, 2200, { fill: i%2===0?"FFF4EC":"FFFFFF", size: 19 }),
          ]}))
        ]
      }),
      spacer(60),
      box([
        "💡 状態カードを受け取ったら：",
        "　「今こういう状態だよ」というサインです。",
        "　防御カードを使って状態から抜け出すことができます。",
        "　ゲームを通して「この状態になるのはどんな行動をしたとき？」を感じ取ってください。"
      ], "FFF4EC", "C55A11"),
      spacer(80),

      // ══ ゲーム終了・勝利判定 ════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("ゲーム終了と勝利判定"),

      h2("ゲーム終了のタイミング"),
      bullet("山札がなくなったとき（全員が同じ回数手番を行った後）"),
      bullet("または全員の合意でゲームを終了したとき"),
      spacer(80),

      h2("勝利判定の計算方法"),
      body("各プレイヤーのキャラクター値と、フィールドカードの理想値を比較します。"),
      spacer(40),
      box([
        "【計算式】　各特性の差の絶対値 を5つ合計する",
        "",
        "例）フィールド：研究者（O:5 C:5 E:2 A:3 N:2）",
        "　　自分の値　（O:4 C:5 E:2 A:4 N:3）",
        "",
        "　O：|4-5|=1　C：|5-5|=0　E：|2-2|=0　A：|4-3|=1　N：|3-2|=1",
        "　合計 = 1+0+0+1+1 = 3点",
        "",
        "　→ この合計点が最も少ない人が勝利！"
      ], "EAF4FF", "4472C4"),
      spacer(80),

      h2("同点の場合"),
      bullet("C（誠実性）の差が少ない方が優先"),
      bullet("それも同点の場合はN（情緒安定性）の差が少ない方"),
      bullet("それも同点の場合はじゃんけんで決定"),
      spacer(80),

      // ══ ゲーム後のふりかえり ═════════════════════════════════
      h1("ゲーム後のふりかえり（大切にしてほしい時間）"),
      body("このゲームで一番大切なのは、ゲームが終わった後のふりかえりです。"),
      spacer(40),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [4400, 4400],
        rows: [
          new TableRow({ children: [hcell("自分に聞いてみよう", 4400, "375623"), hcell("みんなで話し合おう", 4400, "1F4E79")] }),
          new TableRow({ children: [
            new TableCell({
              borders: borders(), shading: { fill: "F0FFF0", type: ShadingType.CLEAR }, margins: pad,
              children: [
                bullet("今の自分はどんなビッグファイブ？"),
                bullet("どのフィールドが自分に向いていた？"),
                bullet("どのカードで一番ダメージを受けた？"),
                bullet("どの防御カードが効果的だったか？"),
                bullet("現実の自分に当てはめると…？"),
              ]
            }),
            new TableCell({
              borders: borders(), shading: { fill: "F0F7FF", type: ShadingType.CLEAR }, margins: pad,
              children: [
                bullet("なぜその職業にその特性が必要なの？"),
                bullet("現実でも同じことが言えると思う？"),
                bullet("ダークな状態はどうやって回避できた？"),
                bullet("自分の特性を伸ばすには現実で何をする？"),
                bullet("どの職業が自分に向いていると思う？"),
              ]
            })
          ]}),
        ]
      }),
      spacer(80),

      // ══ バリアントルール ═════════════════════════════════════
      new Paragraph({ children: [new PageBreak()] }),
      h1("バリアントルール（応用編）"),

      h2("【協力モード】全員で戦う"),
      body("フィールドに「ボス敵」カードを置き、全員で協力してボスを倒すモード。"),
      bullet("ボスは毎ターン全員に攻撃カードを1枚ずつ使う"),
      bullet("全員の合計スコアがボスの理想値を超えたら勝利"),
      bullet("家族・クラスでワイワイ楽しめるモード"),
      spacer(60),

      h2("【子どもモード】シンプルルール"),
      bullet("攻撃カードなし（防御カードのみ使用）"),
      bullet("自分の特性値をフィールドに近づけることだけを考える"),
      bullet("競争ではなく「自己成長」に集中するモード"),
      spacer(60),

      h2("【成長記録モード】"),
      bullet("1ヶ月後にもう一度診断テストを受けて、値の変化を確認する"),
      bullet("実際の生活で防御カードの行動（読書・運動など）を実践してみる"),
      bullet("現実と連動させることで最大の教育効果が得られる"),
      spacer(80),

      // ══ Q&A ════════════════════════════════════════════════
      h1("よくある質問"),
      new Table({
        width: { size: 8800, type: WidthType.DXA },
        columnWidths: [3200, 5600],
        rows: [
          new TableRow({ children: [hcell("質問", 3200), hcell("答え", 5600)] }),
          ...[
            ["診断テストの結果が変わってもいい？", "もちろんです！環境や習慣が変わると特性も変化します。それがこのゲームで伝えたいことです。"],
            ["ダークな状態になったら負け？", "いいえ。ゲームの勝敗はフィールドとの近さで決まります。ただし、ダークな状態はフィールドでの勝利から遠ざかることが多いです。"],
            ["全部の特性が5になれば最強？", "フィールドによって違います。研究者フィールドではE（外向性）が低い方が有利です。万能な性格は存在しません。"],
            ["カードが足りなくなったら？", "捨て札をシャッフルして山札に加えてください。"],
            ["一人でも遊べる？", "診断テストだけ行って、フィールドカードで「このフィールドに近づくには何をすれば良いか」を考えるソロ学習モードとして使えます。"],
          ].map(([q,a],i) => new TableRow({ children: [
            cell(q, 3200, { fill: i%2===0?"FFF8E1":"FFFFFF", bold: true, size: 20 }),
            cell(a, 5600, { fill: i%2===0?"FFF8E1":"FFFFFF", size: 20 }),
          ]}))
        ]
      }),
      spacer(80),

      // ══ クレジット ═══════════════════════════════════════════
      spacer(200),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "このゲームはビッグファイブ理論（BFI-10・TIPI-J）をもとに設計されています。", size: 18, font: "メイリオ", color: "888888" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "無料で自由にお使いください。商業利用はご遠慮ください。", size: 18, font: "メイリオ", color: "888888" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60 },
        children: [new TextRun({ text: "ビッグファイブカードゲーム  ver.0.1  ／  2026年3月", size: 18, font: "メイリオ", color: "AAAAAA" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = "C:\\Users\\user\\Desktop\\Claude Code\\ビッグファイブカードゲーム\\ルールブック_v0.1.docx";
  fs.writeFileSync(out, buf);
  console.log("作成完了:", out);
});
