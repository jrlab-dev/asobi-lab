const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak
} = require('docx');
const fs = require('fs');

// ボーダー設定
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

// カードセクションを作成する関数
function makeCardSection(cardNum, cardName, cardType, oceanValues, expression, pose, impression, englishPrompt, japaneseMemo) {
  const children = [];

  // カードタイトル
  children.push(new Paragraph({
    children: [new TextRun({ text: `【${cardNum}】${cardName}　（${cardType}）`, bold: true, size: 26, color: "1F1F1F" })],
    spacing: { before: 320, after: 80 },
  }));

  // OCEAN値
  children.push(new Paragraph({
    children: [new TextRun({ text: `OCEAN値：${oceanValues}`, size: 22, color: "333333" })]
  }));

  // 黄色ボックス：表情・仕草
  children.push(new Paragraph({
    children: [new TextRun({ text: "▼ 表情・仕草", bold: true, size: 22, color: "7B4F00" })],
    spacing: { before: 120, after: 40 },
  }));
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      width: { size: 9026, type: WidthType.DXA },
      shading: { fill: "FFF9C4", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [
        new Paragraph({ children: [new TextRun({ text: `表情：${expression}`, size: 22 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [new TextRun({ text: `仕草：${pose}`, size: 22 })], spacing: { before: 40, after: 40 } }),
        new Paragraph({ children: [new TextRun({ text: `印象：${impression}`, size: 22, italics: true, color: "555555" })], spacing: { before: 40, after: 40 } }),
      ]
    })]})],
  }));

  // 青色ボックス：英語プロンプト
  children.push(new Paragraph({
    children: [new TextRun({ text: "▼ コピーしてAIに貼り付けてください", bold: true, size: 22, color: "0D4B8B" })],
    spacing: { before: 120, after: 40 },
  }));
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      width: { size: 9026, type: WidthType.DXA },
      shading: { fill: "D6E8FF", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: englishPrompt, size: 22, color: "0D1F3C" })], spacing: { before: 40, after: 40 } })]
    })]})],
  }));

  // 緑色ボックス：日本語メモ
  children.push(new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders,
      width: { size: 9026, type: WidthType.DXA },
      shading: { fill: "D6F5D6", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: `日本語メモ：${japaneseMemo}`, size: 20, color: "1A4A1A" })], spacing: { before: 40, after: 40 } })]
    })]})],
  }));

  return children;
}

// ========================
// 状態カード（S-001〜S-010）
// ========================
const statusCards = [
  // ダークトライアド系
  {
    num: "S-001", name: "サイコパス", type: "状態カード（ダークトライアド）",
    ocean: "O4/C1/E5/A1/N4",
    expression: "冷たい微笑み。感情のない澄んだ目。",
    pose: "腕を組む。または一本指を立てて何かを指示するポーズ。",
    impression: "感情がなく、何を考えているかわからない恐怖感。",
    eng: "a chilling character with a cold calculating smile that never reaches their completely empty emotionless clear eyes, standing with arms crossed or holding up one index finger as if giving a silent command, radiating an unsettling eeriness of someone whose thoughts are completely unreadable, dark mysterious shadowy background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "冷たい計算高い微笑みだが目だけは感情のない澄んだキャラ。腕を組むか人差し指を立てて無言で指示するポーズ。何を考えているかわからない恐怖感。謎めいた暗い背景。"
  },
  {
    num: "S-002", name: "ナルシスト", type: "状態カード（ダークトライアド）",
    ocean: "O3/C2/E5/A1/N3",
    expression: "自分を誇示する優雅な笑顔。見下した目線。",
    pose: "鏡を見るポーズ。または自分の髪や服を整える仕草。",
    impression: "自分が一番美しいと信じて疑わない自惚れ屋。",
    eng: "a self-absorbed narcissistic character looking into a hand mirror or admiring their own reflection with an elegant proud self-loving smile, looking down slightly with a condescending gaze of someone who truly believes they are the most beautiful person in the world, impeccably styled hair and clothes, glamorous sparkling background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "手鏡を見るかポーズをとる自惚れキャラ。優雅で誇らしげな自己愛の笑顔と見下した目線。完璧に整えたヘアスタイルと服装。キラキラしたグラマラスな背景。"
  },
  {
    num: "S-003", name: "ドラマフォーカス", type: "状態カード（ダークトライアド）",
    ocean: "O3/C1/E3/A1/N1",
    expression: "過剰に感情的な表情。泣き顔と怒り顔の中間。",
    pose: "頭を抱えるポーズ。感情的に何かを訴えるような仕草。",
    impression: "常に自分がドラマの主役だと思っている感情暴走キャラ。",
    eng: "a dramatically overreacting character with an exaggerated emotional expression somewhere between crying and angry, holding their head in both hands in an over-the-top melodramatic crisis pose, appearing to believe they are the tragic hero of their own soap opera, theatrical dramatic stage-like background with dramatic lighting, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "頭を両手で抱えるメロドラマ的なキャラ。泣き顔と怒り顔の中間の過剰に感情的な表情。自分がドラマの主役と思っている大げさな仕草。演劇的なライティングの背景。"
  },
  // プチダーク系
  {
    num: "S-004", name: "メンヘラ", type: "状態カード（プチダーク）",
    ocean: "O4/C3/E1/A5/N1",
    expression: "繊細で傷つきやすそうな表情。今にも泣きそうな目。",
    pose: "自分を抱きしめるポーズ。うつむき加減。",
    impression: "依存心が強く感受性豊かな繊細キャラ。",
    eng: "a delicate and emotionally sensitive character with fragile tear-filled eyes on the verge of crying, hugging themselves tightly with arms wrapped around their own body, slightly looking down with a vulnerable expression of someone who desperately needs love and reassurance, soft pastel dreamy background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "自分を抱きしめる繊細キャラ。今にも泣きそうな潤んだ目と傷つきやすそうな表情。やや俯き加減で内向きのポーズ。柔らかなパステルカラーの背景。"
  },
  {
    num: "S-005", name: "ボトムギバー", type: "状態カード（プチダーク）",
    ocean: "O2/C1/E3/A5/N2",
    expression: "疲れ切った優しい笑顔。与えすぎて消耗した目。",
    pose: "両手を差し出す。または荷物を全部持っているような姿勢。",
    impression: "与え続けて自分が空になっていくキャラ。",
    eng: "an over-giving exhausted character holding out both arms with all their energy completely depleted, wearing a tired gentle smile with hollow drained eyes from giving too much of themselves, overloaded carrying too many bags and items for others while their own belongings are empty, simple warm background showing selfless exhaustion, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "与えすぎて消耗したキャラ。疲れた優しい笑顔だが目は消耗した空虚さを宿す。他人の荷物を全部持たされているような重そうなポーズ。自己犠牲の疲弊感のある背景。"
  },
  {
    num: "S-006", name: "リアルダーク", type: "状態カード（プチダーク）",
    ocean: "O3/C1/E5/A3/N2",
    expression: "ニヤリとした不敵な笑み。策士のような目。",
    pose: "何かを手渡す（渡すふりをする）ような仕草。",
    impression: "外向的だが腹黒さが滲み出るトリックスターキャラ。",
    eng: "a sly and cunning trickster character with a sneaky mischievous grin and scheming strategist eyes, appearing outwardly friendly and social while extending a hand as if offering something with clearly hidden ulterior motives, shadows hinting at hidden darker intentions, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "ニヤリと不敵に笑うトリックスターキャラ。表向きは社交的だが腹黒さが滲む策士の目。何かを渡すふりをするような裏のある仕草。暗い意図を示す影が差す背景。"
  },
  // クラスター系
  {
    num: "S-007", name: "アベレージ", type: "状態カード（クラスター）",
    ocean: "O2/C3/E4/A3/N2",
    expression: "特徴のない普通の笑顔。良い意味でも悪い意味でも目立たない。",
    pose: "自然体で立っているだけ。特別なポーズなし。",
    impression: "どこにでもいる普通の人。突出したものがない平均型。",
    eng: "a perfectly average and unremarkable ordinary character standing naturally with a plain ordinary smile that completely blends into any crowd, no special features memorable characteristics or distinguishing marks, plain everyday casual clothes, neutral unremarkable background, the kind of character you would forget immediately after seeing them, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "どこにでもいる平均的なキャラ。良い意味でも悪い意味でも目立たない普通の笑顔。特別なポーズや特徴なし。見た瞬間に忘れてしまうような平均的な外見。無難な背景。"
  },
  {
    num: "S-008", name: "リザーブ", type: "状態カード（クラスター）",
    ocean: "O1/C4/E3/A4/N4",
    expression: "控えめだが誠実な笑顔。静かに何かを作っているような集中顔。",
    pose: "道具を丁寧に扱う職人のような姿勢。内向きのポーズ。",
    impression: "内向的だが堅実で信頼できる職人肌キャラ。",
    eng: "a quiet and reliable reserved craftsman-type character with a sincere modest smile, carefully and precisely handling tools or working on something with patient focused attention, introverted body posture turned slightly inward, radiating steady trustworthy dependable reliability, wooden workshop or craft studio background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "控えめで誠実な職人肌キャラ。道具を丁寧に扱う内向きの姿勢。静かな集中顔と誠実な微笑み。信頼と堅実さを体現する雰囲気。木工房やクラフトスタジオの背景。"
  },
  {
    num: "S-009", name: "ロールモデル", type: "状態カード（クラスター）",
    ocean: "O5/C5/E5/A5/N5",
    expression: "全てがそろった完璧な笑顔。オーラがある目。輝くような表情。",
    pose: "堂々と立って両腕を開くポーズ。光のエフェクトに包まれる。",
    impression: "全ての特性が最高値。誰もが憧れる理想のキャラ。",
    eng: "a perfect and inspiring role model character radiating a brilliant glowing golden aura from every direction, standing confidently with both arms spread wide open in a welcoming heroic pose, the most radiant perfect smile with eyes sparkling with all five combined qualities, surrounded by magical golden sparkles and light rays, stars and light effects background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "全特性が最高値の完璧なロールモデルキャラ。黄金のオーラが全身から輝く。両腕を大きく開く堂々とした英雄ポーズ。完璧な笑顔と5つの特性が宿る輝く目。魔法の光と星のエフェクト背景。"
  },
  {
    num: "S-010", name: "セルフセンター", type: "状態カード（クラスター）",
    ocean: "O2/C2/E5/A2/N3",
    expression: "明るく自信満々だが少し無神経そうな笑顔。",
    pose: "自分を指差す「俺が！」のポーズ。前に出る姿勢。",
    impression: "行動力はあるが協力が苦手な自己中心的キャラ。",
    eng: "a self-centered overconfident character with a bright self-assured smile that is slightly unaware of others around them, dramatically pointing both thumbs toward themselves in an exaggerated look at me pose with a bold forward-leaning stance radiating pure self-absorption, casual energetic outfit, simple background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "自己中心的な自信満々キャラ。両親指を自分に向ける「俺が！」の大げさなポーズ。明るいが少し無神経な笑顔と前のめりの姿勢。エネルギッシュなカジュアルな服。シンプルな背景。"
  }
];

// ========================
// ドキュメント作成
// ========================
let allChildren = [];

// タイトルページ
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "ビッグファイブカードゲーム", bold: true, size: 32 })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 480, after: 160 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "画像生成プロンプト集　状態カード（10枚）v0.1", bold: true, size: 28 })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 160 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "表情・仕草設計書 v0.1 に基づき作成　2026年3月", size: 22, color: "666666" })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 120 }
}));

// 使い方
allChildren.push(new Table({
  width: { size: 9026, type: WidthType.DXA },
  columnWidths: [9026],
  rows: [new TableRow({ children: [new TableCell({
    borders: { top: border, bottom: border, left: border, right: border },
    width: { size: 9026, type: WidthType.DXA },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    margins: { top: 120, bottom: 120, left: 200, right: 200 },
    children: [
      new Paragraph({ children: [new TextRun({ text: "【使い方】", bold: true, size: 24 })], spacing: { before: 60, after: 60 } }),
      new Paragraph({ children: [new TextRun({ text: "青色のプロンプト文をコピーして画像生成AIに貼り付けてください。", size: 22 })], spacing: { before: 40, after: 40 } }),
      new Paragraph({ children: [new TextRun({ text: "推奨サービス：Adobe Firefly / Canva AI / Bing Image Creator（無料）", size: 22 })], spacing: { before: 40, after: 40 } }),
      new Paragraph({ children: [new TextRun({ text: "黄枠：表情・仕草メモ　青枠：英語プロンプト　緑枠：日本語メモ", size: 20, color: "666666" })], spacing: { before: 40, after: 60 } }),
    ]
  })]})],
}));

// カード種別説明
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "【状態カードとは】", bold: true, size: 24 })],
  spacing: { before: 240, after: 80 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "プレイヤーのビッグファイブ値が特定の条件を満たしたときに付与される特殊カード。キャラクターのパーソナリティを視覚的に表現します。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 80 }
}));

// ダークトライアド系
allChildren.push(new Paragraph({ children: [new PageBreak()] }));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "■ ダークトライアド系（3枚）", bold: true, size: 28, color: "3A003A" })],
  spacing: { before: 200, after: 80 },
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "暗い性格特性が突出したキャラクター。ゲーム内では強力だが道徳的に問題のある存在として描く。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 200 }
}));
for (const card of statusCards.slice(0, 3)) {
  const sections = makeCardSection(card.num, card.name, card.type, card.ocean, card.expression, card.pose, card.impression, card.eng, card.jp);
  allChildren.push(...sections);
}

// プチダーク系
allChildren.push(new Paragraph({ children: [new PageBreak()] }));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "■ プチダーク系（3枚）", bold: true, size: 28, color: "4A1A4A" })],
  spacing: { before: 200, after: 80 },
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "少し問題のある特性が見られるキャラクター。完全な悪ではないが、生きにくさを抱えている存在として描く。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 200 }
}));
for (const card of statusCards.slice(3, 6)) {
  const sections = makeCardSection(card.num, card.name, card.type, card.ocean, card.expression, card.pose, card.impression, card.eng, card.jp);
  allChildren.push(...sections);
}

// クラスター系
allChildren.push(new Paragraph({ children: [new PageBreak()] }));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "■ クラスター系（4枚）", bold: true, size: 28, color: "1A4A1A" })],
  spacing: { before: 200, after: 80 },
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "特徴的な性格タイプとして分類されるキャラクター。良い面も悪い面も含む多様なタイプを表す。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 200 }
}));
for (const card of statusCards.slice(6, 10)) {
  const sections = makeCardSection(card.num, card.name, card.type, card.ocean, card.expression, card.pose, card.impression, card.eng, card.jp);
  allChildren.push(...sections);
}

// 次のステップ
allChildren.push(new Paragraph({ children: [new PageBreak()] }));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "【次のステップ】", bold: true, size: 26 })],
  spacing: { before: 320, after: 120 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "・生成した画像を「images」フォルダに保存してください。", size: 22 })],
  spacing: { before: 80, after: 80 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "・ファイル名は カード番号_カード名.png の形式を推奨します（例：S-001_サイコパス.png）", size: 22 })],
  spacing: { before: 80, after: 80 }
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "・カードHTMLの makeCard(..., imgName: null, ...) の null を画像ファイル名に変更すると自動挿入されます。", size: 22 })],
  spacing: { before: 80, after: 80 }
}));

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: allChildren
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('画像生成プロンプト_状態カード_v0.1.docx', buffer);
  console.log('状態カードのプロンプト集を作成しました！');
});
