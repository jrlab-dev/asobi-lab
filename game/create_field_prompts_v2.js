const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageBreak } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 60, after: opts.after || 60 },
    children: [new TextRun({ text, font: "Yu Gothic", size: opts.size || 20, ...opts })]
  });
}

function heading(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: "Yu Gothic", size: 26, bold: true, color: "1B5E20" })]
  });
}

function cardTitle(no, name, cat) {
  const catColor = cat === 'ライト' ? "1565C0" : cat === 'ミドル' ? "E65100" : "B71C1C";
  return new Paragraph({
    spacing: { before: 360, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: catColor, space: 2 } },
    children: [
      new TextRun({ text: `【${no}】${name}　`, font: "Yu Gothic", size: 28, bold: true, color: catColor }),
      new TextRun({ text: `（${cat}）`, font: "Yu Gothic", size: 22, color: "888888" }),
    ]
  });
}

function promptBox(text) {
  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    columnWidths: [9300],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9300, type: WidthType.DXA },
        shading: { fill: "E3F2FD", type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [new Paragraph({
          children: [new TextRun({ text, font: "Courier New", size: 18, color: "0D47A1" })]
        })]
      })]
    })]
  });
}

function memoBox(text) {
  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    columnWidths: [9300],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9300, type: WidthType.DXA },
        shading: { fill: "F9FBE7", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: [new Paragraph({
          children: [new TextRun({ text, font: "Yu Gothic", size: 18, color: "33691E" })]
        })]
      })]
    })]
  });
}

function expressionBox(expression, pose, impression) {
  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    columnWidths: [9300],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9300, type: WidthType.DXA },
        shading: { fill: "FFF8E1", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: [
          new Paragraph({ children: [new TextRun({ text: `表情：${expression}`, font: "Yu Gothic", size: 18, color: "5D4037" })] }),
          new Paragraph({ children: [new TextRun({ text: `仕草：${pose}`, font: "Yu Gothic", size: 18, color: "5D4037" })] }),
          new Paragraph({ children: [new TextRun({ text: `印象：${impression}`, font: "Yu Gothic", size: 18, bold: true, color: "5D4037" })] }),
        ]
      })]
    })]
  });
}

function sp() { return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] }); }

const CARDS = [
  {
    no: "F-001", name: "研究者・博士", cat: "ライト", ocean: "O5/C5/E2/A3/N3",
    expression: "知的でワクワクした目。口元は静かな微笑み。",
    pose: "虫眼鏡や本を持つ。やや前かがみで考え込む姿勢。",
    impression: "好奇心と真面目さが同居する知識人。",
    prompt: "a cheerful scientist character wearing a white lab coat and round glasses, leaning slightly forward in curiosity, holding a magnifying glass with sparkling excited eyes and a quiet gentle smile, surrounded by books and scientific equipment, warm laboratory background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "白衣・眼鏡の博士キャラ。虫眼鏡を持ち前かがみで好奇心旺盛な目と静かな微笑み。本や実験器具が周囲にある温かみのある研究室背景。"
  },
  {
    no: "F-002", name: "カウンセラー", cat: "ライト", ocean: "O3/C3/E3/A5/N3",
    expression: "温かく受け入れる優しい笑顔。目線は相手に寄り添う。",
    pose: "手を胸の前で組む。前のめりで相手の話を聞くポーズ。",
    impression: "包容力と共感力を体現する安心感のある存在。",
    prompt: "a warm and gentle counselor character with a soft welcoming smile, leaning forward attentively with hands clasped gently in front of chest, eyes full of empathy and care, sitting in a cozy office with plants and warm lighting, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "穏やかな笑顔のカウンセラーキャラ。手を胸の前で組み前のめりで話を聞くポーズ。共感と包容力のある目線。温かみのあるオフィス背景。"
  },
  {
    no: "F-003", name: "農家・職人", cat: "ライト", ocean: "O3/C5/E2/A3/N4",
    expression: "落ち着いた真剣な表情。誠実さがにじむ目。",
    pose: "道具を丁寧に扱う手。地に足のついたどっしりした立ち方。",
    impression: "コツコツと積み上げる誠実な職人気質。",
    prompt: "a sturdy and reliable farmer or craftsman character wearing overalls and a straw hat, standing firmly with both feet planted on the ground, handling tools with careful and gentle hands, serious and sincere expression with calm focused eyes, green farm fields or wooden workshop in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "つなぎ服・麦わら帽子の職人キャラ。道具を丁寧に扱う手と地に足のついた安定感のある立ち姿。真剣で誠実さがにじむ目。"
  },
  {
    no: "F-004", name: "芸術家", cat: "ライト", ocean: "O5/C2/E3/A3/N4",
    expression: "夢見るような目。自由で個性的な表情。",
    pose: "筆やパレットを持つ。少し気ままなポーズ。",
    impression: "独創性と感性が輝く自由人。",
    prompt: "a creative and free-spirited artist character wearing a paint-splattered smock and a beret, holding a palette and paintbrush with a dreamy and imaginative gaze, relaxed and whimsical carefree pose, colorful paintings and art studio in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "絵の具で汚れたスモックとベレー帽のアーティストキャラ。夢見るような目で気ままなポーズ。パレットと筆を持つ。カラフルなアトリエ背景。"
  },
  {
    no: "F-005", name: "看護師・介護士", cat: "ライト", ocean: "O3/C4/E3/A5/N3",
    expression: "責任感ある真剣な目と温かい笑顔が共存。",
    pose: "手を前に出して助けるポーズ。白衣や制服姿。",
    impression: "使命感と思いやりを持つ頼れる存在。",
    prompt: "a kind and caring nurse character wearing a white nurse uniform with a red cross symbol, reaching one hand forward in a helping gesture, showing both a warm gentle smile and determined caring eyes full of responsibility, hospital room with flowers in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "白いナース服の看護師キャラ。手を前に差し出す助けるポーズ。責任感ある目と温かい笑顔が共存した表情。花が飾られた明るい病室背景。"
  },
  {
    no: "F-006", name: "起業家・経営者", cat: "ミドル", ocean: "O5/C4/E5/A2/N3",
    expression: "自信に満ちた力強い表情。鋭くも輝く目。",
    pose: "腕を広げる大きなジェスチャー。前のめりで攻めの姿勢。",
    impression: "エネルギッシュで孤独をも恐れないリーダー。",
    prompt: "a confident and energetic entrepreneur character wearing a sharp business suit, leaning forward aggressively with both arms spread wide open in a bold dynamic gesture, sharp and shining eyes full of ambition and confidence, modern city office building and skyline in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "颯爽としたスーツの起業家キャラ。両腕を大きく広げ前のめりの攻めの姿勢。野心と自信に満ちた鋭く輝く目。都市の背景。"
  },
  {
    no: "F-007", name: "営業職", cat: "ミドル", ocean: "O3/C3/E5/A4/N4",
    expression: "人懐っこい明るい笑顔。親しみやすい目。",
    pose: "手を差し伸べる握手のポーズ。前向きで積極的な立ち方。",
    impression: "どんな相手にも明るく接する不屈の笑顔。",
    prompt: "a cheerful and enthusiastic salesperson character in business casual clothes, extending one hand forward in a warm handshake gesture with a big friendly approachable smile, energetic forward-leaning stance, holding a briefcase in the other hand, bright city street background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "ビジネスカジュアルの元気な営業職キャラ。握手するように手を差し伸べ前向きな立ち方。人懐っこい大きな笑顔と親しみやすい目。"
  },
  {
    no: "F-008", name: "俳優・芸人", cat: "ミドル", ocean: "O5/C2/E5/A3/N4",
    expression: "表情豊か。大げさな驚きや喜びを体全体で表現。",
    pose: "両腕を大きく広げたパフォーマンスポーズ。",
    impression: "感情の振れ幅が最大の武器の表現者。",
    prompt: "a dramatic and expressive actor or comedian character wearing a flashy theatrical costume, both arms spread dramatically wide open, face showing exaggerated surprise or joy with highly expressive over-the-top emotions across the whole body, stage with bright spotlights in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "派手な舞台衣装の俳優・芸人キャラ。両腕を大きく広げた大げさなパフォーマンスポーズ。体全体で驚きや喜びを表現する豊かな表情。"
  },
  {
    no: "F-009", name: "インフルエンサー", cat: "ミドル", ocean: "O4/C3/E5/A3/N4",
    expression: "カメラ目線で魅力的な笑顔。少し計算されたポーズ。",
    pose: "スマホを持つ。自分をアピールするポーズ。",
    impression: "自分を商品にする覚悟を持つ自己プロデューサー。",
    prompt: "a trendy and stylish influencer character wearing fashionable casual clothes, holding a smartphone toward the camera with a perfectly posed charming smile that feels slightly calculated and deliberate, confident self-promoting stance, colorful social media icons and hearts floating around, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "おしゃれなカジュアル服のインフルエンサーキャラ。スマホをカメラに向け計算された魅力的な笑顔。自分をアピールする自信のあるポーズ。"
  },
  {
    no: "F-010", name: "プロスポーツ", cat: "ミドル", ocean: "O3/C5/E3/A3/N5",
    expression: "精神的に強い目。プレッシャーを感じさせない余裕の表情。",
    pose: "筋肉質でアスリートらしい力強いポーズ。",
    impression: "才能より継続でのし上がったストイックな競技者。",
    prompt: "a determined and stoic athletic sports player character wearing a dynamic sports uniform, striking a powerful muscular action pose with mentally strong calm eyes showing no pressure or fear, composed and confident expression with quiet composure, stadium crowd and bright arena lights in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "スポーツユニフォームのアスリートキャラ。力強いポーズとプレッシャーを感じさせない余裕の目。精神的な強さと落ち着きが滲み出る表情。スタジアム背景。"
  },
  {
    no: "F-011", name: "政治家・官僚", cat: "グレー⚠", ocean: "O3/C3/E5/A2/N4",
    expression: "自信と計算が混じった表情。カリスマ的な目線。",
    pose: "演説のように手を前に出すポーズ。",
    impression: "勝てるが幸せとは限らない権力者。",
    prompt: "a serious and authoritative politician character wearing a dark formal suit, raising one hand forward in a speech gesture as if giving a confident address, expression mixing self-confidence and calculation with a charismatic piercing gaze, government building and national flag in background, slightly dramatic lighting, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "ダークスーツの政治家キャラ。演説のように手を前に出すカリスマポーズ。自信と計算が混じった目線（グレーカード：勝てるが幸せとは限らない）。"
  },
  {
    no: "F-012", name: "弁護士・検察官", cat: "グレー⚠", ocean: "O3/C5/E4/A2/N3",
    expression: "鋭い分析の目。感情を抑えた威圧感のある表情。",
    pose: "書類や証拠を持つ。腕を組む姿勢。",
    impression: "正義を貫くために孤立を恐れない論理の人。",
    prompt: "a sharp and analytical lawyer character wearing a black formal suit, standing with arms crossed holding legal documents, sharp piercing analytical eyes with a suppressed emotionless expression radiating quiet intimidation, courtroom with scales of justice in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "黒いスーツの弁護士キャラ。腕を組んで法廷文書を持つ威圧感のある立ち方。感情を抑えた鋭い分析の目（グレーカード：孤立を恐れない論理の人）。"
  },
  {
    no: "F-013", name: "外科医・救急医", cat: "グレー⚠", ocean: "O3/C5/E3/A2/N5",
    expression: "完全に感情を遮断した冷静な目。プロとしての無表情。",
    pose: "手術道具や聴診器を持つ。集中した前傾姿勢。",
    impression: "感情を抑制することで命を救う究極のプロ。",
    prompt: "a calm and focused surgeon character wearing green surgical scrubs with a face mask pulled down, leaning forward in a concentrated professional stance holding surgical instruments or a stethoscope, completely emotionless stone-cold professional expression with blank detached calm eyes, operating room with medical equipment in background, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "緑の手術着の外科医キャラ。聴診器や手術器具を持ち集中した前傾姿勢。完全に感情をシャットアウトした冷静な無表情（グレーカード：感情の抑制が武器）。"
  },
  {
    no: "F-014", name: "証券トレーダー", cat: "グレー⚠", ocean: "O3/C5/E3/A2/N5",
    expression: "計算高い鋭い目。感情を出さないポーカーフェイス。",
    pose: "複数のモニターや数字を見るポーズ。腕を組む。",
    impression: "勝ち続けるために人間味を削るクールな勝負師。",
    prompt: "an intense and calculating securities trader character in a business suit, standing with arms crossed staring at multiple computer screens showing stock charts with cold sharp calculating eyes and a complete poker face expression revealing no emotions, surrounded by glowing monitors with financial data, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "スーツ姿のトレーダーキャラ。腕を組んで複数のモニターを見る。計算高い冷たい鋭い目と感情ゼロのポーカーフェイス（グレーカード：人間味を削る勝負師）。"
  },
  {
    no: "F-015", name: "ネットワークビジネス", cat: "グレー⚠", ocean: "O3/C3/E5/A2/N4",
    expression: "過剰に明るい笑顔。少し作られた感じの表情。",
    pose: "手を差し伸べる。近づきすぎる距離感のポーズ。",
    impression: "「勝てる=幸せ」ではないことを体現するキャラ。",
    prompt: "a persuasive and overly enthusiastic network marketer character in bright business casual clothes, reaching both hands forward too eagerly invading personal space, wearing an excessively big forced artificial smile that feels slightly unsettling and too calculated, network diagram of people connecting around them, anime chibi illustration, clean bold outlines, flat vibrant colors, kawaii cute style, white background, full body character, digital art",
    memo: "明るい服のネットワークビジネスキャラ。近づきすぎる距離感で両手を差し伸べる。過剰に明るく少し作られた不自然な笑顔（グレーカード：勝てる=幸せではない）。"
  },
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Yu Gothic", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Yu Gothic", color: "1B5E20" },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 0 } },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1008, right: 1008, bottom: 1008, left: 1008 } }
    },
    children: [
      // タイトル
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: "ビッグファイブカードゲーム", bold: true, size: 44, font: "Yu Gothic", color: "1B5E20" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "画像生成プロンプト集　フィールドカード（15枚）v0.2", bold: true, size: 32, font: "Yu Gothic", color: "2E7D32" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "表情・仕草設計書 v0.1 に基づき更新　2026年3月", size: 18, font: "Yu Gothic", color: "888888" })] }),

      // 凡例
      new Table({
        width: { size: 9890, type: WidthType.DXA },
        columnWidths: [9890],
        rows: [new TableRow({ children: [new TableCell({
          borders: { top: border, bottom: border, left: border, right: border },
          width: { size: 9890, type: WidthType.DXA },
          shading: { fill: "F1F8E9", type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "【使い方】", bold: true, font: "Yu Gothic", size: 20, color: "1B5E20" })] }),
            new Paragraph({ children: [new TextRun({ text: "青色のプロンプト文をコピーして画像生成AIに貼り付けてください。", font: "Yu Gothic", size: 18 })] }),
            new Paragraph({ children: [new TextRun({ text: "推奨サービス：Adobe Firefly / Canva AI / Bing Image Creator（無料）", font: "Yu Gothic", size: 18 })] }),
            new Paragraph({ children: [new TextRun({ text: "【v0.2 更新内容】青枠：プロンプト　黄枠：表情・仕草メモ（更新箇所）　緑枠：日本語メモ", font: "Yu Gothic", size: 18, color: "666666" })] }),
          ]
        })] })]
      }),
      sp(),

      // 各カード
      ...CARDS.flatMap(card => [
        cardTitle(card.no, card.name, card.cat),
        p(`OCEAN値：${card.ocean}`, { color: "888888", size: 18 }),
        sp(),
        p("▼ 表情・仕草（v0.2新規追加）", { bold: true, size: 18, color: "5D4037" }),
        expressionBox(card.expression, card.pose, card.impression),
        sp(),
        p("▼ コピーしてAIに貼り付けてください（v0.2更新）", { bold: true, size: 18, color: "0D47A1" }),
        promptBox(card.prompt),
        sp(),
        p("日本語メモ：" + card.memo, { color: "33691E", size: 18 }),
        new Paragraph({ spacing: { before: 60, after: 60 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 1 } },
          children: [new TextRun("")]
        }),
      ]),

      // フッター
      sp(),
      p("【次のステップ】", { bold: true }),
      p("・生成した画像を「images」フォルダに保存してください。"),
      p("・ファイル名は カード番号_カード名.png の形式を推奨します（例：F-001_研究者博士.png）"),
      p("・カードHTMLの makeCard(..., imgName: null, ...) の null を画像ファイル名に変更すると自動挿入されます。"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('画像生成プロンプト_フィールドカード_v0.2.docx', buffer);
  console.log('SUCCESS: 画像生成プロンプト_フィールドカード_v0.2.docx を作成しました');
}).catch(err => {
  console.error('ERROR:', err);
  process.exit(1);
});
