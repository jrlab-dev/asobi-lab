const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerShading = { fill: "2E4057", type: ShadingType.CLEAR };
const lightShading = { fill: "F2F4F8", type: ShadingType.CLEAR };
const darkHeaderShading = { fill: "1A1A2E", type: ShadingType.CLEAR };

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: headerShading,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18, font: "Yu Gothic" })]
    })]
  });
}

function dataCell(text, width, shade = false) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: shade ? lightShading : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, size: 18, font: "Yu Gothic" })]
    })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Yu Gothic", color: "2E4057" })]
  });
}

function subTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Yu Gothic", color: "1B6CA8" })]
  });
}

function bodyText(text) {
  return new Paragraph({
    spacing: { before: 80, after: 120 },
    children: [new TextRun({ text, size: 20, font: "Yu Gothic", color: "444444" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun("")] });
}

// ===== 第1章: 特性ルール表 =====
function makeTraitTable(traitName, traitLetter, rows) {
  const tableRows = [
    new TableRow({
      children: [
        headerCell(`${traitLetter}値`, 1500),
        headerCell("表情・仕草の特徴", 7800),
      ]
    }),
    ...rows.map((r, i) => new TableRow({
      children: [
        dataCell(r[0], 1500, i % 2 === 0),
        dataCell(r[1], 7800, i % 2 === 0),
      ]
    }))
  ];
  return [
    new Paragraph({
      spacing: { before: 180, after: 80 },
      children: [new TextRun({ text: traitName, bold: true, size: 22, font: "Yu Gothic", color: "2E4057" })]
    }),
    new Table({
      width: { size: 9300, type: WidthType.DXA },
      columnWidths: [1500, 7800],
      rows: tableRows
    }),
    spacer()
  ];
}

const traitData = [
  ["O（開放性）", "O", [
    ["O=5（最高）", "目がキラキラ輝いている。口を開けて驚き・感嘆する表情。両手を広げて「発見！」のポーズ。"],
    ["O=4（高い）", "ニコニコと好奇心あふれる笑顔。頭を少し傾けて興味深そうに見る仕草。"],
    ["O=3（中間）", "穏やかな微笑み。程よい関心を示す目線。"],
    ["O=2（低い）", "落ち着いた表情。慎重そうに確認する仕草。"],
    ["O=1（最低）", "真顔に近い。変化を好まない安定感のある佇まい。"],
  ]],
  ["C（誠実性）", "C", [
    ["C=5（最高）", "真剣な表情。背筋を伸ばした姿勢。手帳・道具・ペンなどを持つ。整然とした印象。"],
    ["C=4（高い）", "きちんとした佇まい。集中している目線。"],
    ["C=3（中間）", "自然体のリラックスした姿勢。"],
    ["C=2（低い）", "ゆったりとした姿勢。少しくだけた雰囲気。"],
    ["C=1（最低）", "だらっとしたポーズ。物事を適当に扱う雰囲気。"],
  ]],
  ["E（外向性）", "E", [
    ["E=5（最高）", "満面の笑顔で両腕を広げる。前のめりで活発な印象。周囲を巻き込む雰囲気。"],
    ["E=4（高い）", "大きな笑顔。身振り手振りが大きい。存在感がある。"],
    ["E=3（中間）", "自然な笑顔。程よいエネルギー感。"],
    ["E=2（低い）", "控えめな表情。静かにじっくり取り組む姿勢。"],
    ["E=1（最低）", "内向きのポーズ。1人で集中するような静かな雰囲気。"],
  ]],
  ["A（協調性）", "A", [
    ["A=5（最高）", "優しく温かい笑顔。手を差し伸べるポーズ。包み込むような雰囲気。"],
    ["A=4（高い）", "思いやりのある目線。相手に寄り添うような立ち方。"],
    ["A=3（中間）", "バランスのとれた表情。協力的でも自立した雰囲気。"],
    ["A=2（低い）", "鋭い眼差し。腕を組む仕草。"],
    ["A=1（最低）", "冷たい目線。距離を置くポーズ。独立した佇まい。"],
  ]],
  ["N（情緒安定性）", "N", [
    ["N=5（最高）", "どっしりとして余裕のある表情。落ち着いた目線。風格がある佇まい。"],
    ["N=4（高い）", "穏やかな余裕の笑み。プレッシャーを感じさせない自信のある立ち方。"],
    ["N=3（中間）", "普通の落ち着いた表情。"],
    ["N=2（低い）", "やや繊細そうな表情。感情が少し表に出やすい。"],
    ["N=1（最低）", "うつむき加減。感受性豊かで感情が顔に出やすい。傷つきやすそうな雰囲気。"],
  ]],
];

// ===== 第2章: フィールドカード表 =====
const fieldData = [
  ["F-001","研究者・博士","5/5/2/3/3","知的でワクワクした目。口元は静かな微笑み。","虫眼鏡や本を持つ。やや前かがみで考え込む姿勢。","好奇心と真面目さが同居する知識人。"],
  ["F-002","カウンセラー","3/3/3/5/3","温かく受け入れる優しい笑顔。目線は相手に寄り添う。","手を胸の前で組む。前のめりで相手の話を聞くポーズ。","包容力と共感力を体現する安心感のある存在。"],
  ["F-003","農家・職人","3/5/2/3/4","落ち着いた真剣な表情。誠実さがにじむ目。","道具を丁寧に扱う手。地に足のついたどっしりした立ち方。","コツコツと積み上げる誠実な職人気質。"],
  ["F-004","芸術家","5/2/3/3/4","夢見るような目。自由で個性的な表情。","筆やパレットを持つ。少し気ままなポーズ。","独創性と感性が輝く自由人。"],
  ["F-005","看護師・介護士","3/4/3/5/3","責任感ある真剣な目と温かい笑顔が共存。","手を前に出して助けるポーズ。白衣や制服姿。","使命感と思いやりを持つ頼れる存在。"],
  ["F-006","起業家・経営者","5/4/5/2/3","自信に満ちた力強い表情。鋭くも輝く目。","腕を広げる大きなジェスチャー。前のめりで攻めの姿勢。","エネルギッシュで孤独をも恐れないリーダー。"],
  ["F-007","営業職","3/3/5/4/4","人懐っこい明るい笑顔。親しみやすい目。","手を差し伸べる握手のポーズ。前向きで積極的な立ち方。","どんな相手にも明るく接する不屈の笑顔。"],
  ["F-008","俳優・芸人","5/2/5/3/4","表情豊か。大げさな驚きや喜びを体全体で表現。","両腕を大きく広げたパフォーマンスポーズ。","感情の振れ幅が最大の武器の表現者。"],
  ["F-009","インフルエンサー","4/3/5/3/4","カメラ目線で魅力的な笑顔。少し計算されたポーズ。","スマホを持つ。自分をアピールするポーズ。","自分を商品にする覚悟を持つ自己プロデューサー。"],
  ["F-010","プロスポーツ","3/5/3/3/5","精神的に強い目。プレッシャーを感じさせない余裕の表情。","筋肉質でアスリートらしい力強いポーズ。","才能より継続でのし上がったストイックな競技者。"],
  ["F-011","政治家・官僚","3/3/5/2/4","自信と計算が混じった表情。カリスマ的な目線。","演説のように手を前に出すポーズ。","勝てるが幸せとは限らない権力者。"],
  ["F-012","弁護士・検察官","3/5/4/2/3","鋭い分析の目。感情を抑えた威圧感のある表情。","書類や証拠を持つ。腕を組む姿勢。","正義を貫くために孤立を恐れない論理の人。"],
  ["F-013","外科医・救急医","3/5/3/2/5","完全に感情を遮断した冷静な目。プロとしての無表情。","手術道具や聴診器を持つ。集中した前傾姿勢。","感情を抑制することで命を救う究極のプロ。"],
  ["F-014","証券トレーダー","3/5/3/2/5","計算高い鋭い目。感情を出さないポーカーフェイス。","複数のモニターや数字を見るポーズ。腕を組む。","勝ち続けるために人間味を削るクールな勝負師。"],
  ["F-015","ネットワークビジネス","3/3/5/2/4","過剰に明るい笑顔。少し作られた感じの表情。","手を差し伸べる。近づきすぎる距離感のポーズ。","「勝てる=幸せ」ではないことを体現するキャラ。"],
];

function makeCardTable6col(data, colWidths) {
  const [w1,w2,w3,w4,w5,w6] = colWidths;
  return new Table({
    width: { size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: [
          headerCell("No", w1), headerCell("カード名", w2), headerCell("OCEAN値", w3),
          headerCell("表情", w4), headerCell("仕草・ポーズ", w5), headerCell("全体的な印象", w6),
        ]
      }),
      ...data.map((r, i) => new TableRow({
        children: r.map((cell, j) => dataCell(cell, colWidths[j], i % 2 === 0))
      }))
    ]
  });
}

// ===== 第3章: 状態カード表 (6列) =====
const darkData = [
  ["S-001","サイコパス","O4/C1/E5/A1/N4","冷たい微笑み。感情のない澄んだ目。","腕を組む。または一本指を立てて何かを指示するポーズ。","感情がなく、何を考えているかわからない恐怖感。"],
  ["S-002","ナルシスト","O3/C2/E5/A1/N3","自分を誇示する優雅な笑顔。見下した目線。","鏡を見るポーズ。または自分の髪や服を整える仕草。","自分が一番美しいと信じて疑わない自惚れ屋。"],
  ["S-003","ドラマフォーカス","O3/C1/E3/A1/N1","過剰に感情的な表情。泣き顔と怒り顔の中間。","頭を抱えるポーズ。感情的に何かを訴えるような仕草。","常に自分がドラマの主役だと思っている感情暴走キャラ。"],
];

const petitData = [
  ["S-004","メンヘラ","O4/C3/E1/A5/N1","繊細で傷つきやすそうな表情。今にも泣きそうな目。","自分を抱きしめるポーズ。うつむき加減。","依存心が強く感受性豊かな繊細キャラ。"],
  ["S-005","ボトムギバー","O2/C1/E3/A5/N2","疲れ切った優しい笑顔。与えすぎて消耗した目。","両手を差し出す。または荷物を全部持っているような姿勢。","与え続けて自分が空になっていくキャラ。"],
  ["S-006","リアルダーク","O3/C1/E5/A3/N2","ニヤリとした不敵な笑み。策士のような目。","何かを手渡す（渡すふりをする）ような仕草。","外向的だが腹黒さが滲み出るトリックスターキャラ。"],
];

const clusterData = [
  ["S-007","アベレージ","O2/C3/E4/A3/N2","特徴のない普通の笑顔。良い意味でも悪い意味でも目立たない。","自然体で立っているだけ。特別なポーズなし。","どこにでもいる普通の人。突出したものがない平均型。"],
  ["S-008","リザーブ","O1/C4/E3/A4/N4","控えめだが誠実な笑顔。静かに何かを作っているような集中顔。","道具を丁寧に扱う職人のような姿勢。内向きのポーズ。","内向的だが堅実で信頼できる職人肌キャラ。"],
  ["S-009","ロールモデル","O5/C5/E5/A5/N5","全てがそろった完璧な笑顔。オーラがある目。輝くような表情。","堂々と立って両腕を開くポーズ。光のエフェクトに包まれる。","全ての特性が最高値。誰もが憧れる理想のキャラ。"],
  ["S-010","セルフセンター","O2/C2/E5/A2/N3","明るく自信満々だが少し無神経そうな笑顔。","自分を指差す「俺が！」のポーズ。前に出る姿勢。","行動力はあるが協力が苦手な自己中心的キャラ。"],
];

// ===== 第4章: 攻撃カード (5列) =====
const attackData = [
  ["A-001","ポテトチップス","C-2, N-1","ソファでポテチを食べながらだらっとしているシーン","ぼーっとした無気力な表情。袋を抱えてリラックスしすぎたポーズ。"],
  ["A-002","ジャンクフード","C-2, E-1, N-2","ジャンクフードを山積みにして食べまくっているシーン","幸せそうだが少し後悔の混じる表情。食べ物を抱えるポーズ。"],
  ["A-003","デザート三昧","C-1, N-1","ケーキやデザートに囲まれて我慢できずに食べているシーン","甘い幸福感と誘惑に負けた表情。スプーンを口に運ぶ仕草。"],
  ["A-004","ソーシャルゲーム","O-1, C-3, N-2","暗い部屋でスマホを一心不乱にタップしているシーン","ゲームに没入した目。意識が画面に吸い込まれているポーズ。"],
  ["A-005","SNS閲覧","C-2, A-1, N-2","スマホを見て他人の投稿と比べて落ち込んでいるシーン","羨ましさと劣等感の混じった複雑な表情。スマホを覗き込む姿勢。"],
  ["A-006","睡眠不足","C-2, E-1, N-3","目の下にクマを作ってフラフラしているシーン","眠そうな半眼。立ったまま寝そうなよろめくポーズ。"],
  ["A-007","運動不足","C-1, E-1, N-2","ずっと座ったまま体が固まってしまっているシーン","だるそうな表情。体が丸まってうまく動けない様子。"],
  ["A-008","金銭欲求（強欲）","C-1, A-3, N-1","お金を積み上げて目をギラギラさせているシーン","欲望で目が輝いているが歪んだ表情。お金を掴む手。"],
  ["A-009","ギャンブル","C-3, N-2","カジノで一喜一憂しているシーン","刺激に興奮した表情とがっくり落ち込む表情が混在。"],
  ["A-010","過労・残業強要","C-1, A-1, N-3","山積みの書類に埋もれて燃え尽きているシーン","空虚な目。疲れ果てて机に倒れそうなポーズ。"],
  ["A-011","アルコール依存","C-2, A-1, N-2","お酒を飲んで現実逃避しているシーン","酔ったような幸せそうな表情だが空虚な目。グラスを掲げる仕草。"],
  ["A-012","批判・否定","O-1, C-1, A-1, N-3","指を指して誰かを批判・攻撃しているシーン","怒りと優越感の混じった表情。指差しや腕組みの攻撃的なポーズ。"],
];

// ===== 第5章: 防御カード (5列) =====
const defenseData = [
  ["D-001","瞑想","C+1, N+3","静かな場所で目を閉じて瞑想しているシーン","穏やかで安らかな表情。蓮座を組んで目を閉じるポーズ。"],
  ["D-002","マインドフルネス","O+1, A+1, N+2","自然の中で深呼吸をしているシーン","今に集中した穏やかな笑顔。両手を広げて空気を吸い込むポーズ。"],
  ["D-003","有酸素運動","C+1, E+1, N+2","楽しそうにランニングしているシーン","爽やかで清々しい笑顔。全力で走る躍動感のあるポーズ。"],
  ["D-004","ウエイトトレーニング","C+3, E+1, N+2","筋トレで汗を流しているシーン","真剣で力強い表情。バーベルを持ち上げる力強いポーズ。"],
  ["D-005","心拍トレーニング","C+2, N+2","心拍計を見ながら集中してトレーニングしているシーン","集中した静かな表情。ゾーンに入ったような目。"],
  ["D-006","プログラミング","O+2, C+3, N+1","パソコンに向かって集中してコードを書いているシーン","知的で集中した表情。キーボードを叩く姿勢。"],
  ["D-007","人助け・ボランティア","E+1, A+3, N+1","誰かを助けているシーン","温かく自然な笑顔。手を差し伸べるポーズ。"],
  ["D-008","読書","O+3, C+1, N+1","本に夢中になっているシーン","好奇心で目が輝く表情。本を両手で持ってワクワクしている姿。"],
  ["D-009","挨拶・コミュ習慣","E+2, A+1","元気に挨拶しているシーン","明るく爽やかな笑顔。手を振る・頭を下げるポーズ。"],
  ["D-010","日記・自己分析","O+2, C+1, N+1","ノートに日記を書いているシーン","内省的でじっくり考える表情。ペンを持ってノートを開くポーズ。"],
  ["D-011","感謝を伝える","A+2, N+2","誰かに感謝を伝えているシーン","心からの温かい笑顔。両手を胸に当てる感謝のポーズ。"],
  ["D-012","早起き・規則正しい生活","C+2, N+1","朝日の中で気持ちよく目覚めているシーン","清々しい笑顔。両手を上に伸ばして背伸びするポーズ。"],
  ["D-013","目標設定","O+1, C+3","目標を書き出して気合いを入れているシーン","真剣で意欲的な表情。ノートや手帳に書き込む力強い仕草。"],
];

function makeCardTable5col(data, colWidths) {
  const [w1,w2,w3,w4,w5] = colWidths;
  return new Table({
    width: { size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: [
          headerCell("No", w1), headerCell("カード名", w2), headerCell("効果値", w3),
          headerCell("シーンの表現", w4), headerCell("キャラの表情・仕草", w5),
        ]
      }),
      ...data.map((r, i) => new TableRow({
        children: r.map((cell, j) => dataCell(cell, colWidths[j], i % 2 === 0))
      }))
    ]
  });
}

// ===== ドキュメント構築 =====
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Yu Gothic", size: 20 } }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Yu Gothic", color: "2E4057" },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Yu Gothic", color: "1B6CA8" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Yu Gothic", color: "555555" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 16838, height: 11906 }, // A4横向き
        margin: { top: 1008, right: 1008, bottom: 1008, left: 1008 }
      }
    },
    children: [
      // タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 200 },
        children: [new TextRun({ text: "ビッグファイブカードゲーム", bold: true, size: 48, font: "Yu Gothic", color: "2E4057" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "表情・仕草設計書 v0.1", bold: true, size: 36, font: "Yu Gothic", color: "1B6CA8" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "2026年3月　ビッグファイブカードゲーム制作チーム", size: 20, font: "Yu Gothic", color: "888888" })]
      }),
      bodyText("【目的】各カードのO/C/E/A/N値をもとに、AIイラスト生成プロンプトに使う「表情・仕草」を定義した参照資料です。"),
      spacer(),

      // 第1章
      sectionTitle("第1章　ビッグファイブ5特性の視覚的表現ルール"),
      bodyText("各特性の値（1〜5）に対応する表情・仕草の基本ルール。プロンプト作成時の参照用。"),
      spacer(),
      ...traitData.flatMap(([name, letter, rows]) => makeTraitTable(name, letter, rows)),

      // 第2章
      sectionTitle("第2章　フィールドカード（職種）の表情・仕草一覧"),
      bodyText("キャラクターとして最もイラスト表現が重要なカード群。OCEANの値の組み合わせから、その職種らしい表情・仕草を設計する。"),
      spacer(),
      makeCardTable6col(fieldData, [900, 1600, 1000, 2400, 2200, 2200]),
      spacer(),

      // 第3章
      sectionTitle("第3章　状態カードの表情・仕草一覧"),
      subTitle("3-1　ダークトライアド系（3枚）"),
      makeCardTable6col(darkData, [900, 1600, 1000, 2400, 2200, 2200]),
      spacer(),
      subTitle("3-2　プチダーク系（3枚）"),
      makeCardTable6col(petitData, [900, 1600, 1000, 2400, 2200, 2200]),
      spacer(),
      subTitle("3-3　クラスター系（4枚）"),
      makeCardTable6col(clusterData, [900, 1600, 1000, 2400, 2200, 2200]),
      spacer(),

      // 第4章
      sectionTitle("第4章　攻撃カード（シーン・行動表現）の表情・仕草一覧"),
      bodyText("攻撃カードは「悪い習慣」を描く。キャラクターがその習慣に溺れているシーンとして表現する。やや残念・堕落した様子が伝わるデフォルメ表現。"),
      spacer(),
      makeCardTable5col(attackData, [900, 1800, 1300, 2700, 2600]),
      spacer(),

      // 第5章
      sectionTitle("第5章　防御カード（シーン・行動表現）の表情・仕草一覧"),
      bodyText("防御カードは「良い習慣」を描く。キャラクターがその習慣を実践して輝いているシーンとして表現する。前向きで清々しい表情が基本。"),
      spacer(),
      makeCardTable5col(defenseData, [900, 1800, 1300, 2700, 2600]),
      spacer(),

      // 第6章
      sectionTitle("第6章　プロンプト作成時の活用方法"),
      bodyText("【このドキュメントの使い方】"),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "1. 各カードのOCEAN値を確認する", size: 20, font: "Yu Gothic" })] }),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "2. 第1章の特性ルール表で、各値に対応する表情・仕草を確認する", size: 20, font: "Yu Gothic" })] }),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "3. 第2〜5章でそのカードの表情・仕草設計を確認する", size: 20, font: "Yu Gothic" })] }),
      new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: "4. AIイラスト生成プロンプトの「表情」「仕草」の部分に追記・修正する", size: 20, font: "Yu Gothic" })] }),
      spacer(),
      bodyText("【次のステップ】"),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "・このドキュメントのデータをもとに「画像生成プロンプト_全カード.docx」を作成する", size: 20, font: "Yu Gothic" })] }),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "・攻撃・防御・状態カードの画像生成プロンプトを新規作成する（フィールドカードは既存プロンプトを更新する）", size: 20, font: "Yu Gothic" })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('表情仕草設計書_v0.1.docx', buffer);
  console.log('SUCCESS: 表情仕草設計書_v0.1.docx を作成しました');
}).catch(err => {
  console.error('ERROR:', err);
});
