const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  PageBreak
} = require('docx');
const fs = require('fs');

// ボーダー設定
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// カードセクションを作成する関数
function makeCardSection(cardNum, cardName, cardType, effectValues, scene, expression, pose, impression, englishPrompt, japaneseMemo) {
  const children = [];

  // カードタイトル
  children.push(new Paragraph({
    children: [new TextRun({ text: `【${cardNum}】${cardName}　（${cardType}）`, bold: true, size: 26, color: "1F1F1F" })],
    spacing: { before: 320, after: 80 },
  }));

  // 効果値
  children.push(new Paragraph({
    children: [new TextRun({ text: `効果値：${effectValues}`, size: 22, color: "333333" })]
  }));

  // シーン
  if (scene) {
    children.push(new Paragraph({
      children: [new TextRun({ text: `シーン：${scene}`, size: 22, color: "333333" })]
    }));
  }

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
// 攻撃カード（A-001〜A-012）
// ========================
const attackCards = [
  {
    num: "A-001", name: "ポテトチップス", type: "攻撃カード",
    effect: "C-2, N-1",
    scene: "ソファでポテチを食べながらだらっとしているシーン",
    expression: "ぼーっとした無気力な表情。",
    pose: "袋を抱えてリラックスしすぎたポーズ。",
    impression: "怠惰と快楽が渦巻くだらしなさの権化。",
    eng: "a lazy and unmotivated character slouched on a couch, holding a big bag of potato chips with a vacant blank expression and heavy-lidded sleepy eyes, totally limp overly relaxed body posture, crumbs everywhere, dimly lit cozy room background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "ソファで袋ポテチを抱えてだらっとしたキャラ。ぼーっとした無気力な顔と重い瞼。リラックスしすぎてだらしない姿勢。薄暗い部屋の背景。"
  },
  {
    num: "A-002", name: "ジャンクフード", type: "攻撃カード",
    effect: "C-2, E-1, N-2",
    scene: "ジャンクフードを山積みにして食べまくっているシーン",
    expression: "幸せそうだが少し後悔の混じる表情。",
    pose: "食べ物を抱えるポーズ。",
    impression: "欲求に負けた満足と罪悪感が共存する姿。",
    eng: "a character surrounded by a towering pile of junk food including burgers fries and donuts, eating with a happy satisfied but slightly guilty mixed expression, hugging all the food tightly, fast food restaurant background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "バーガー・ポテト・ドーナツの山に囲まれたキャラ。幸せそうだが少し後悔が混じる表情。食べ物を両腕で抱え込むポーズ。ファストフード背景。"
  },
  {
    num: "A-003", name: "デザート三昧", type: "攻撃カード",
    effect: "C-1, N-1",
    scene: "ケーキやデザートに囲まれて我慢できずに食べているシーン",
    expression: "甘い幸福感と誘惑に負けた表情。",
    pose: "スプーンを口に運ぶ仕草。",
    impression: "甘いものへの誘惑に完全敗北した至福のひととき。",
    eng: "a character completely surrounded by cakes desserts and sweets unable to resist, holding a spoon about to take another bite with a blissful sweet happy expression mixed with having given in to temptation, dessert shop background overflowing with cakes everywhere, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "ケーキやスイーツに囲まれたキャラ。スプーンを口に運ぶ仕草。甘い幸福感と誘惑に負けた表情。デザートショップ背景。"
  },
  {
    num: "A-004", name: "ソーシャルゲーム", type: "攻撃カード",
    effect: "O-1, C-3, N-2",
    scene: "暗い部屋でスマホを一心不乱にタップしているシーン",
    expression: "ゲームに没入した目。",
    pose: "意識が画面に吸い込まれているポーズ。",
    impression: "画面の前でリアルを忘れた廃人一歩手前。",
    eng: "a character in a dark room hunched over a glowing smartphone, tapping intensely with eyes completely absorbed and sucked into the screen, neglecting everything around them with an obsessive focused expression, blue screen glow lighting their face in the dark room, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "暗い部屋でスマホを凝視するキャラ。ゲームに没入した目と画面に吸い込まれる姿勢。青い画面の光だけが照らす薄暗い部屋の背景。"
  },
  {
    num: "A-005", name: "SNS閲覧", type: "攻撃カード",
    effect: "C-2, A-1, N-2",
    scene: "スマホを見て他人の投稿と比べて落ち込んでいるシーン",
    expression: "羨ましさと劣等感の混じった複雑な表情。",
    pose: "スマホを覗き込む姿勢。",
    impression: "比べることでじわじわ自己肯定感が削られる日常。",
    eng: "a character looking at a smartphone with a complex expression mixing envy and inferiority, scrolling through social media comparing themselves to others with a sinking deflated expression, slumped posture hunching over the phone, colorful social media feed visible on screen, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "スマホのSNSを見て落ち込むキャラ。羨ましさと劣等感の混じった複雑な顔でスマホを覗き込む姿勢。カラフルなSNSフィードが画面に映る。"
  },
  {
    num: "A-006", name: "睡眠不足", type: "攻撃カード",
    effect: "C-2, E-1, N-3",
    scene: "目の下にクマを作ってフラフラしているシーン",
    expression: "眠そうな半眼。",
    pose: "立ったまま寝そうなよろめくポーズ。",
    impression: "慢性的な疲弊が体を蝕む危うい状態。",
    eng: "a character with dark circles under their eyes staggering and wobbling, barely keeping their eyes open with heavy drooping half-closed sleepy eyes, about to fall asleep while still standing, messy room or office background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "クマが目立つ睡眠不足のキャラ。重い瞼の半眼で立ちながら寝そうによろめくポーズ。散らかった部屋や職場の背景。"
  },
  {
    num: "A-007", name: "運動不足", type: "攻撃カード",
    effect: "C-1, E-1, N-2",
    scene: "ずっと座ったまま体が固まってしまっているシーン",
    expression: "だるそうな表情。",
    pose: "体が丸まってうまく動けない様子。",
    impression: "動くことを忘れた体が固まっていく悪循環。",
    eng: "a character who has been sitting still for way too long with a sluggish heavy expression, body hunched and stiff unable to move properly, surrounded by snack wrappers and a laptop, completely un-athletic stiff posture, dim cluttered room background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "長時間座りっぱなしで体が固まったキャラ。だるそうな表情で体が丸まり動けない様子。お菓子の袋やラップトップに囲まれた散らかった部屋の背景。"
  },
  {
    num: "A-008", name: "金銭欲求（強欲）", type: "攻撃カード",
    effect: "C-1, A-3, N-1",
    scene: "お金を積み上げて目をギラギラさせているシーン",
    expression: "欲望で目が輝いているが歪んだ表情。",
    pose: "お金を掴む手。",
    impression: "金への執着が人間性をゆがめていく強欲キャラ。",
    eng: "a greedy character sitting on a pile of gold coins and money bills with intensely gleaming obsessed eyes, hands greedily clutching fistfuls of money with a distorted expression mixing desire and desperation, stacks of coins and bills all around, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "金貨と紙幣の山の上に座るキャラ。欲望でギラギラした歪んだ目でお金を鷲掴みにする手。コインや紙幣が周囲に散乱。"
  },
  {
    num: "A-009", name: "ギャンブル", type: "攻撃カード",
    effect: "C-3, N-2",
    scene: "カジノで一喜一憂しているシーン",
    expression: "刺激に興奮した表情とがっくり落ち込む表情が混在。",
    pose: "カードを持って一喜一憂するポーズ。",
    impression: "勝ちと負けの感情ジェットコースターに乗り続けるキャラ。",
    eng: "a character at a casino table with a wildly mixed expression combining excited highs and disappointed lows all at once, holding playing cards with one fist pumping in excitement while the other hand reaches out in panic, casino table with chips and cards, neon casino lights background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "カジノテーブルで一喜一憂するキャラ。興奮と落胆が混在する表情でトランプを持つ。ネオンカラーのカジノ背景。"
  },
  {
    num: "A-010", name: "過労・残業強要", type: "攻撃カード",
    effect: "C-1, A-1, N-3",
    scene: "山積みの書類に埋もれて燃え尽きているシーン",
    expression: "空虚な目。",
    pose: "疲れ果てて机に倒れそうなポーズ。",
    impression: "限界を超えても終わらない仕事に魂が抜けた状態。",
    eng: "a completely burnt out character buried under towering piles of documents and paperwork, slumped over a desk with hollow empty blank eyes showing total mental exhaustion, about to collapse onto the desk, stacks of overflowing files covering the entire workspace, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "書類の山に埋もれて燃え尽きたキャラ。空虚な抜け殻の目で机に倒れそうなポーズ。仕事が溢れかえるデスクの背景。"
  },
  {
    num: "A-011", name: "アルコール依存", type: "攻撃カード",
    effect: "C-2, A-1, N-2",
    scene: "お酒を飲んで現実逃避しているシーン",
    expression: "酔ったような幸せそうな表情だが空虚な目。",
    pose: "グラスを掲げる仕草。",
    impression: "お酒に逃げることで現実を見なくなっていくキャラ。",
    eng: "a character raising a glass of alcohol with a seemingly happy drunken smile that hides hollow empty vacant eyes underneath, escaping from reality through drinking, surrounded by empty bottles, slightly unsteady tipsy posture, dim bar or home background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "グラスを掲げる酔ったキャラ。表面は幸せそうな笑顔だが内側に空虚な目を隠す。空き瓶が周囲に並ぶ薄暗いバーや自宅の背景。"
  },
  {
    num: "A-012", name: "批判・否定", type: "攻撃カード",
    effect: "O-1, C-1, A-1, N-3",
    scene: "指を指して誰かを批判・攻撃しているシーン",
    expression: "怒りと優越感の混じった表情。",
    pose: "指差しや腕組みの攻撃的なポーズ。",
    impression: "他者を否定することで自分を保つ毒キャラ。",
    eng: "an aggressive and critical character pointing a finger accusingly with a harsh expression mixing anger and a sense of superiority, confrontational stance with one finger pointing directly or arms crossed, exaggerated dramatic attack pose, simple background with tension and conflict energy, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "指差しで批判・攻撃するキャラ。怒りと優越感の混じった表情と攻撃的な腕組みや指差しポーズ。緊張感のある対立の背景。"
  }
];

// ========================
// 防御カード（D-001〜D-013）
// ========================
const defenseCards = [
  {
    num: "D-001", name: "瞑想", type: "防御カード",
    effect: "C+1, N+3",
    scene: "静かな場所で目を閉じて瞑想しているシーン",
    expression: "穏やかで安らかな表情。",
    pose: "蓮座を組んで目を閉じるポーズ。",
    impression: "深い内なる平和と静けさを体現するキャラ。",
    eng: "a peaceful character sitting in lotus meditation pose with eyes gently closed, showing a completely calm and serene expression radiating deep inner peace and tranquility, soft glowing light surrounding them, quiet nature garden or tatami room background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "蓮座で瞑想するキャラ。目を閉じた穏やかで安らかな表情。内なる平和を体現する柔らかな光に包まれる。自然の庭や畳の部屋の背景。"
  },
  {
    num: "D-002", name: "マインドフルネス", type: "防御カード",
    effect: "O+1, A+1, N+2",
    scene: "自然の中で深呼吸をしているシーン",
    expression: "今に集中した穏やかな笑顔。",
    pose: "両手を広げて空気を吸い込むポーズ。",
    impression: "今この瞬間に完全に存在する意識の覚醒キャラ。",
    eng: "a mindful character standing in a beautiful natural setting with eyes half-closed and a gentle present-moment smile, arms spread wide open breathing in fresh air deeply, radiating peaceful calm awareness, surrounded by trees flowers and a soft breeze, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "自然の中で深呼吸するキャラ。目を半開きに今この瞬間に集中した穏やかな笑顔。両手を広げて空気を吸い込むポーズ。木々と花に囲まれた自然背景。"
  },
  {
    num: "D-003", name: "有酸素運動", type: "防御カード",
    effect: "C+1, E+1, N+2",
    scene: "楽しそうにランニングしているシーン",
    expression: "爽やかで清々しい笑顔。",
    pose: "全力で走る躍動感のあるポーズ。",
    impression: "走ることで体も心も解放される爽快感あふれるキャラ。",
    eng: "an energetic and refreshed character happily jogging or running with a bright refreshing clean smile, dynamic action running pose radiating vitality and joy, wearing sports clothes, sunny park path with trees and blue sky in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "爽やかにランニングするキャラ。清々しい笑顔と躍動感あふれる走るポーズ。スポーツウェア姿。晴れた公園の緑道の背景。"
  },
  {
    num: "D-004", name: "ウエイトトレーニング", type: "防御カード",
    effect: "C+3, E+1, N+2",
    scene: "筋トレで汗を流しているシーン",
    expression: "真剣で力強い表情。",
    pose: "バーベルを持ち上げる力強いポーズ。",
    impression: "自分を鍛えることで精神も強くなっていく意志の人。",
    eng: "a determined and powerful character doing weight training, lifting a heavy barbell with a serious intensely focused expression showing strength and willpower, sweating and straining with muscular effort, gym background with weights and equipment, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "バーベルを持ち上げるキャラ。真剣で力強い表情と筋力を振り絞るポーズ。汗をかきながらも意志の強さが滲み出る。ジムの背景。"
  },
  {
    num: "D-005", name: "心拍トレーニング", type: "防御カード",
    effect: "C+2, N+2",
    scene: "心拍計を見ながら集中してトレーニングしているシーン",
    expression: "集中した静かな表情。",
    pose: "ゾーンに入ったような目。",
    impression: "数値を管理しながら着実に強くなる自己管理の達人。",
    eng: "a focused character in the middle of heart rate training, checking a heart rate monitor on their wrist with a quiet deeply concentrated expression, appearing to be in the zone with calm determined eyes, sports facility or running track background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "心拍モニターを確認しながらトレーニングするキャラ。ゾーンに入ったような静かで集中した表情。スポーツ施設やトラックの背景。"
  },
  {
    num: "D-006", name: "プログラミング", type: "防御カード",
    effect: "O+2, C+3, N+1",
    scene: "パソコンに向かって集中してコードを書いているシーン",
    expression: "知的で集中した表情。",
    pose: "キーボードを叩く姿勢。",
    impression: "論理と創造を組み合わせて問題を解くフロー状態のキャラ。",
    eng: "an intellectual focused programmer character typing code at a laptop or computer with a deeply concentrated clever expression, eyes lit up by the screen glow in a coding flow state, surrounded by coding books and a second monitor showing code, home office or tech workspace background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "ラップトップでコードを書くキャラ。知的で集中したフロー状態の表情。画面の光に照らされた目。コーディング本やモニターに囲まれたデスクの背景。"
  },
  {
    num: "D-007", name: "人助け・ボランティア", type: "防御カード",
    effect: "E+1, A+3, N+1",
    scene: "誰かを助けているシーン",
    expression: "温かく自然な笑顔。",
    pose: "手を差し伸べるポーズ。",
    impression: "見返りを求めず人に寄り添う本物の温かさを持つキャラ。",
    eng: "a warm-hearted volunteer character helping someone in need with a genuinely warm and natural bright smile, reaching out both hands in a caring helpful gesture, radiating kindness and goodwill, community park or volunteer activity background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "ボランティアで人を助けるキャラ。手を差し伸べる温かく自然な笑顔。見返りを求めない純粋な優しさを体現。コミュニティや公園の背景。"
  },
  {
    num: "D-008", name: "読書", type: "防御カード",
    effect: "O+3, C+1, N+1",
    scene: "本に夢中になっているシーン",
    expression: "好奇心で目が輝く表情。",
    pose: "本を両手で持ってワクワクしている姿。",
    impression: "活字の世界に没入し知識と想像力を広げ続けるキャラ。",
    eng: "a curious and excited reader character deeply absorbed in a book, holding it with both hands with sparkling curious eyes wide with wonder and imagination, totally captivated and lost in the story, cozy library or bedroom background with bookshelves and warm soft lighting, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "本に夢中になったキャラ。好奇心でキラキラ輝く目で本を両手で持つポーズ。知識の世界に完全に没入。暖かみのある本棚のある図書室や寝室の背景。"
  },
  {
    num: "D-009", name: "挨拶・コミュ習慣", type: "防御カード",
    effect: "E+2, A+1",
    scene: "元気に挨拶しているシーン",
    expression: "明るく爽やかな笑顔。",
    pose: "手を振る・頭を下げるポーズ。",
    impression: "挨拶の力で人と人をつなぎ場を明るくするキャラ。",
    eng: "a friendly and energetic character giving a cheerful greeting, waving a hand enthusiastically or giving a polite bow with a big bright refreshing smile radiating warmth and social positivity, outdoors street or school entrance background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "元気よく挨拶するキャラ。手を振るか頭を下げる明るく爽やかな笑顔のポーズ。街や学校の入口の背景。"
  },
  {
    num: "D-010", name: "日記・自己分析", type: "防御カード",
    effect: "O+2, C+1, N+1",
    scene: "ノートに日記を書いているシーン",
    expression: "内省的でじっくり考える表情。",
    pose: "ペンを持ってノートを開くポーズ。",
    impression: "自分を知ることで着実に成長していく内省型キャラ。",
    eng: "a thoughtful and introspective character writing in a diary or journal, holding a pen with a pensive reflective expression, looking inward with calm focused eyes, notebook open on a desk, cozy quiet desk setup background with soft warm lighting, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "日記をつけるキャラ。ペンを持ってノートに書き込む内省的なじっくり考える表情。温かみのある落ち着いたデスクの背景。"
  },
  {
    num: "D-011", name: "感謝を伝える", type: "防御カード",
    effect: "A+2, N+2",
    scene: "誰かに感謝を伝えているシーン",
    expression: "心からの温かい笑顔。",
    pose: "両手を胸に当てる感謝のポーズ。",
    impression: "感謝の気持ちが人間関係と自分自身を豊かにするキャラ。",
    eng: "a heartfelt and grateful character expressing sincere thanks to someone with a warm genuine smile radiating from the heart, both hands placed on chest in a sincere gratitude pose, emotional and touching moment, simple bright warm background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "感謝を伝えるキャラ。両手を胸に当てる感謝のポーズと心からの温かい笑顔。感動的な場面。明るく温かみのある背景。"
  },
  {
    num: "D-012", name: "早起き・規則正しい生活", type: "防御カード",
    effect: "C+2, N+1",
    scene: "朝日の中で気持ちよく目覚めているシーン",
    expression: "清々しい笑顔。",
    pose: "両手を上に伸ばして背伸びするポーズ。",
    impression: "朝の清々しさが一日全体をポジティブにするキャラ。",
    eng: "a cheerful early riser character waking up refreshed in bright morning sunlight, arms stretched high above their head in a big energetic morning stretch with a big refreshing smile, radiating positive morning energy, bedroom window with warm golden sunrise light background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "朝日の中で目覚めるキャラ。両手を高く上げて朝の大きな背伸びをする清々しい笑顔のポーズ。黄金色の朝日が差し込む寝室の窓の背景。"
  },
  {
    num: "D-013", name: "目標設定", type: "防御カード",
    effect: "O+1, C+3",
    scene: "目標を書き出して気合いを入れているシーン",
    expression: "真剣で意欲的な表情。",
    pose: "ノートや手帳に書き込む力強い仕草。",
    impression: "目標を言語化することで夢を現実に近づけるキャラ。",
    eng: "a motivated and determined goal-setter character writing down their goals in a notebook or planner with a serious and driven expression full of ambition and willpower, gripping a pen with strong purposeful force, sticky notes and vision board visible in background, bright workspace background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    jp: "目標を書き出すキャラ。ノートにペンを力強く走らせる真剣で意欲的な表情。付箋やビジョンボードが見える明るいデスクの背景。"
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
  children: [new TextRun({ text: "画像生成プロンプト集　攻撃カード・防御カード v0.1", bold: true, size: 28 })],
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
    borders,
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

// 攻撃カードセクション
allChildren.push(new Paragraph({
  children: [new PageBreak()]
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "■ 攻撃カード（12枚）", bold: true, size: 30, color: "8B0000" })],
  spacing: { before: 200, after: 120 },
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "攻撃カードは「悪い習慣」を描く。やや残念・堕落した様子が伝わるデフォルメ表現。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 200 }
}));

for (const card of attackCards) {
  const sections = makeCardSection(card.num, card.name, card.type, card.effect, card.scene, card.expression, card.pose, card.impression, card.eng, card.jp);
  allChildren.push(...sections);
}

// 防御カードセクション
allChildren.push(new Paragraph({ children: [new PageBreak()] }));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "■ 防御カード（13枚）", bold: true, size: 30, color: "0D4B8B" })],
  spacing: { before: 200, after: 120 },
}));
allChildren.push(new Paragraph({
  children: [new TextRun({ text: "防御カードは「良い習慣」を描く。前向きで清々しい表情が基本。", size: 22, color: "444444" })],
  spacing: { before: 0, after: 200 }
}));

for (const card of defenseCards) {
  const sections = makeCardSection(card.num, card.name, card.type, card.effect, card.scene, card.expression, card.pose, card.impression, card.eng, card.jp);
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
  children: [new TextRun({ text: "・ファイル名は カード番号_カード名.png の形式を推奨します（例：A-001_ポテトチップス.png）", size: 22 })],
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
  fs.writeFileSync('画像生成プロンプト_攻撃防御カード_v0.1.docx', buffer);
  console.log('攻撃・防御カードのプロンプト集を作成しました！');
});
