const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        LevelFormat, VerticalAlign } = require('docx');
const fs = require('fs');

// 共通スタイル設定
const createDocStyles = () => ({
  default: { document: { run: { font: "游ゴシック", size: 22 } } },
  paragraphStyles: [
    { id: "Title", name: "Title", basedOn: "Normal",
      run: { size: 36, bold: true, color: "000000", font: "游ゴシック" },
      paragraph: { spacing: { before: 0, after: 300 }, alignment: AlignmentType.CENTER } },
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 28, bold: true, color: "2E74B5", font: "游ゴシック" },
      paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 24, bold: true, color: "2E74B5", font: "游ゴシック" },
      paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, color: "404040", font: "游ゴシック" },
      paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    { id: "Quote", name: "Quote", basedOn: "Normal",
      run: { size: 20, italics: true, color: "404040", font: "游ゴシック" },
      paragraph: { spacing: { before: 100, after: 100 }, indent: { left: 720 } } }
  ]
});

const createNumbering = () => ({
  config: [
    { reference: "bullet-list",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbered-list-1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbered-list-2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbered-list-3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbered-list-4", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "sub-bullet",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "-", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] }
  ]
});

const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const cellBorders = { top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder };
const headerShading = { fill: "2E74B5", type: ShadingType.CLEAR };

// テーブル作成ヘルパー
const createTable = (headers, rows, colWidths) => {
  return new Table({
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders: cellBorders, width: { size: colWidths[i], type: WidthType.DXA },
          shading: headerShading, verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20 })] })]
        }))
      }),
      ...rows.map(row => new TableRow({
        children: row.map((cell, i) => new TableCell({
          borders: cellBorders, width: { size: colWidths[i], type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: cell, size: 20 })] })]
        }))
      }))
    ]
  });
};

// 保護者回答分析
const createHogosha = () => {
  const children = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("【保護者】ふれあいホリデー アンケート自由記述分析（2025年度）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("概要")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "対象", bold: true }), new TextRun(": 保護者アンケート 問7「今後に向けてお気づきの点などを教えてください」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "回答数", bold: true }), new TextRun(": 418件（総回答937件中）")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "実施日", bold: true }), new TextRun(": 2025年11月21日（金）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. 肯定的意見")] }),
    new Paragraph({ children: [new TextRun("制度を支持し、有意義だったという声も一定数存在。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「あまり平日に休めないので、家族で過ごせる良い機会となった。」"],
      ["2", "「平日に休みがあると、休日だと混み合っている場所等で、一緒に出かけづらいところにも一緒に行けるので、良いと思います。」"],
      ["3", "「教員夫婦なのでお休みが取れて子どもと遊べました。平日休みだったので、普段混み合っている施設がガラガラで遊びやすかったです。」"],
      ["4", "「4連休息抜きになって嬉しそうでいつも学校毎日お疲れ様って思いました」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「特になし。有り難い休日でした。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「親子でゆっくり過ごすことができました！」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「良い取り組みだと思うので、今後も継続していただきたい」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("傾向")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("休暇を取得できた家庭からは好評")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("平日で施設が空いている点を評価")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("教員など同日休みの職場は有効活用")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. 課題・問題点の指摘（最も多い）")] }),
    new Paragraph({ children: [new TextRun("休暇取得困難や制度設計への疑問が多数寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「ほとんどの保護者の方が、ふれあいホリデーで仕事は休めないと聞きます。自分自身も職場が推奨していないので、この為に休みは取れないです。」"],
      ["2", "「子供だけ休みにしても親は休めない。この取組をするなら親も休みにしてほしい。」"],
      ["3", "「不定休の仕事だと、ふれあいホリデー当日も、1人は仕事になるため、あまり意味のない休みだと思う。ふれあいホリデーと綺麗事を言いながらも、結局は教員の為の休みだと思う。」"],
      ["4", "「本当に迷惑なのでやめていただきたい。我が家の実態は子どもたちだけでダラダラとすごし、メディア三昧。親は昼ごはんの準備の負担増。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「有給が少ない中、有給をつかって休暇を取得しないといけないのが苦しいです」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「子どもだけで家で過ごすことになり、お弁当の準備など逆に負担が増える」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「医療、介護従事者の親が、一斉に休んだらどうなるか想像して欲しい」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「行事が多い時期に実施されると会社を何度も休まないといけなくなる」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("主要な課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "休暇取得の困難", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("職場に制度が周知されていない")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("有給休暇が足りない")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("連休前は特に休みづらい")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "子どもの過ごし方", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("子どもだけで留守番")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("学童に預けざるを得ない")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("昼食準備の負担増")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "制度設計への疑問", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("誰のための休みか不明")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("教員のための休みではないか")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("ふれあいの意味がない")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. 改善提案")] }),
    new Paragraph({ children: [new TextRun("建設的な改善案も多数寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「ふれあいホリデーを市が実施するのであれば各企業などに呼びかけや休暇を取れるように働きかけをしてくれないと、ふれあいホリデーをやる意味がないです。」"],
      ["2", "「各家庭ごとに好きな日に1日お休みを取れるようにして欲しい。各企業もそれに合わせて、ふれあい休暇を作って、積極的に休暇を取れるようにして欲しい。」"],
      ["3", "「倉吉市だけやることでさまざまな歪みを生んでしまっています。鳥取県で、あるいはせめて中部地区で揃えてできないのであれば、やめるべきだと今年も感じました。」"],
      ["4", "「一斉に実施ではなく、地域別に時期をずらし、地域のイベントに参加がいいとおもう。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("企業・職場への働きかけ")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("市から企業への周知・協力要請")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("休暇取得の義務化・推奨")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("企業への助成金制度")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("市内事業所への通達")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("実施方法の変更")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("日にちを各家庭で選べるようにする")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("学校別・地区別で分散実施")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("1日ではなく期間内で選択制に")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("実施範囲の拡大")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("鳥取県全体での実施")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("中部地区での統一実施")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("高校・保育園も同日に")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("時期の変更")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("6月など祝日のない月へ")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("テスト期間を避ける")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("連休にくっつけない")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. 廃止・中止要望")] }),
    new Paragraph({ children: [new TextRun("制度自体の廃止を求める声が非常に多い。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「これからもふれあいホリデー、実施するのですか？休みが取りづらい保護者からしたら、やめて欲しいです。」"],
      ["2", "「この休暇をやめて下さい。休めない親はただただ、心理的負担です。」"],
      ["3", "「必要ないのでやめてください」"],
      ["4", "「ふれあいホリデーは必要ない。普段から子供とはちゃんとふれあっているから大丈夫です。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("廃止要望の理由")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("休暇が取れないので子どもに申し訳ない気持ちになる")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("結局子どもだけで過ごすなら意味がない")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("昼食準備など親の負担が増えるだけ")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("普段から家族との時間は取れている")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("教員のための休みにしか見えない")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. 特徴的な指摘")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("ひとり親家庭への配慮")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「ひとり親です。この制度は体験格差が生じる内容だと思います。やめていただきたいです。」")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「シングルマザー、シングルファザーなんて尚更です。特に後者に関しては大きな痛手です。」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("職種による不公平")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「休める人休めない人で不公平感が生じないか。」")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「子育て世帯が多い職場では、全員が休めるわけもなく、誰かに負担が必ずいきます。」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("教員への批判")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「結局ふれあいホリデーの目的は休みの取れていない教員に休暇を与えるための取り組みであることを耳にしました。」")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「先生たちのためにある取り組みなのかなと感じる。」")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("集計サマリー")] }),
    createTable(["カテゴリ", "件数", "割合"], [
      ["肯定的意見", "65件", "16%"],
      ["課題・問題点の指摘", "257件", "61%"],
      ["改善提案", "46件", "11%"],
      ["廃止・中止要望", "34件", "8%"],
      ["特になし", "16件", "4%"]
    ], [4000, 2500, 2500]),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("総評")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("評価できる点")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("休暇を取得できた家庭からは概ね好評")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("平日で施設が空いている点は評価されている")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("子どもの喜ぶ姿を見られたという声も")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("深刻な課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "休暇取得の構造的問題", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("企業・職場への周知が不十分")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("休暇を取りたくても取れない現状")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("同じ職場内で対象者が重なる問題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "不公平感の発生", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("休める家庭と休めない家庭の格差")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("子どもへの申し訳なさ")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("ひとり親家庭への負担")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "制度の本末転倒", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("「ふれあい」できないのに「ふれあいホリデー」")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("昼食準備など親の負担増")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("子どもだけで過ごす＝リスク増大")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "アンケート反映への疑問", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("昨年も同様の意見が出ていたはず")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("否定的意見も公表してほしいという声")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("改善が見られないことへの不満")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("今後の検討課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-3", level: 0 }, children: [new TextRun({ text: "企業への働きかけ強化", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("市から企業・事業所への正式な協力要請")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("休暇取得促進のインセンティブ")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("商工会等を通じた周知")] }),
    new Paragraph({ numbering: { reference: "numbered-list-3", level: 0 }, children: [new TextRun({ text: "実施方法の抜本的見直し", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("分散実施（学校別・地区別）の検討")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("各家庭で日程を選べる方式")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("または制度の廃止")] }),
    new Paragraph({ numbering: { reference: "numbered-list-3", level: 0 }, children: [new TextRun({ text: "実施範囲の拡大検討", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("中部地区での統一実施")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("県全体での実施")] }),
    new Paragraph({ numbering: { reference: "numbered-list-3", level: 0 }, children: [new TextRun({ text: "休めない家庭への配慮", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("学童・イベントの充実")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("昼食支援")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("子どもの見守り体制")] }),

    new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "分析日: 2026年1月20日", italics: true, size: 18 })] })
  ];
  return new Document({ styles: createDocStyles(), numbering: createNumbering(), sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }] });
};

// 児童生徒回答分析
const createJido = () => {
  const children = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("【児童生徒】ふれあいホリデー アンケート自由記述分析（2025年度）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("概要")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "対象", bold: true }), new TextRun(": 児童生徒アンケート 問4「感想や、これからどうすればいいかを教えてください」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "回答数", bold: true }), new TextRun(": 1,559件（総回答2,223件中）")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "実施日", bold: true }), new TextRun(": 2025年11月21日（金）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. 肯定的意見（楽しかった・継続希望）")] }),
    new Paragraph({ children: [new TextRun("児童生徒からは圧倒的に肯定的な意見が寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「家族が、仕事を休んで、お出かけに連れて行ってくれたので久しぶりに遠くへお出かけできました。待ち時間が短かったので、たくさん遊べて楽しかったです。」"],
      ["2", "「とても充実した休みでした、普段感じる疲れやストレスが一気に吹っ飛んでとてもリフレッシュした休みでした。」"],
      ["3", "「4連休で、お父さんといる時間が長くなって嬉しかった。」"],
      ["4", "「家族と普段少しは触れ合いたいけど触れ合えないから触れ合いホリデーがあった方がいいし普段親と学校終わりとかは話せてないから触れ合いホリデーがあることで触れ合うことができるから嬉しいです」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「楽しかったもっと増やしてほしい」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「来年も同じで良いです」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「これからも続けてほしいです」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「年に一回家族で旅行に行けるいい機会なのでありがたいです」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「倉吉市だけなので、いろいろな施設に人が少ないので、いろいろなところを満喫できました」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("傾向")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「楽しかった」という感想が最も多い")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("連休になったことへの喜びの声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("継続希望・増加希望の声が多数")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("家族と過ごせた喜びを表現する回答が多い")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. 課題・問題点の指摘")] }),
    new Paragraph({ children: [new TextRun("親が休めず、本来の目的を達成できなかったという声も一定数存在。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「親が仕事でいなかったので特に休日と変わらなかったです。」"],
      ["2", "「自分は休みになっても、家族は休みにはならないから一緒に過ごせなかった（去年も）」"],
      ["3", "「ふれあいホリデーでも家族の大人は仕事なので触れ合えることができませんでした。」"],
      ["4", "「もうちょっと時期をずらしてほしい。学校によるかもしれないけど期末テストと被ってしまっていた」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「一人で留守番していたから、大人が仕事を休めたらもっと楽しく過ごせた」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「家族が仕事でひまだったから学校のほうがよかった」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「親休めないことだってあるからあんま意味ないと思う」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「テスト期間で、勉強せざるを得ない日になってしまった」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3大課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "親が休めない", bold: true }), new TextRun(" - 子どもだけで過ごす家庭が一定数存在")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "テスト時期との重複", bold: true }), new TextRun(" - 中学校の期末テスト直前")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "一人での留守番", bold: true }), new TextRun(" - 特に共働き家庭で発生")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. 改善提案・要望")] }),
    new Paragraph({ children: [new TextRun("具体的な改善案も児童生徒から寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「大人も仕事を休みにしたら良いと思う。」"],
      ["2", "「高校生にもこの休みがほしいです」"],
      ["3", "「もっと休みを増やしてほしいです」"],
      ["4", "「テスト期間じゃないときにやったほうがいいと思いました。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("休日の拡大")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("もっと日数を増やしてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("月1回にしてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("5連休にしてほしい")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("対象の拡大")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("高校生も対象にしてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("大人（親）も休みにしてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("全県統一にしてほしい")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("時期の変更")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("テスト期間を避けてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("6月など祝日のない月に")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("連休にくっつけなくてもいい")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("イベントの充実")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("もっとイベントを増やしてほしい")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("子どもも参加できるイベントがあると良い")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("中学生も参加できるイベントを")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. 小学生と中学生の傾向の違い")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("小学生（低学年）")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("シンプルに「楽しかった」という感想が多い")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("家族と遊んだ具体的な内容を記述")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「来年もやってほしい」という素朴な継続希望")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("中学生")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("テスト期間との重複を指摘する声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("より論理的に制度の課題を指摘")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("勉強時間が確保できた点を評価する声も")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("集計サマリー")] }),
    createTable(["カテゴリ", "件数", "割合"], [
      ["肯定的（楽しかった・継続希望）", "969件", "62%"],
      ["課題・問題点の指摘", "46件", "3%"],
      ["改善提案", "87件", "6%"],
      ["特になし・短い返答", "457件", "29%"]
    ], [4000, 2500, 2500]),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("総評")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("評価できる点")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("児童生徒からは圧倒的に好評")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("家族との時間を持てたことへの喜びの声が多数")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「続けてほしい」「増やしてほしい」という要望が非常に多い")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("平日で施設が空いていて楽しめたという声も")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("今後の検討課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "保護者の休暇取得促進", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("「大人も休みにしてほしい」という声が多数")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("親が休めない場合の子どもの過ごし方")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "実施時期の再検討", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("中学校の期末テスト時期との調整")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("テスト前を避けた日程設定")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "対象範囲の拡大検討", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("高校生への拡大要望")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("県全体・中部全体での実施要望")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "イベントの充実", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("中学生も参加できるイベント")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("地域でのイベント開催")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("特徴的な意見")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「ふれあいホリデーは大人も休みにした方がいいと思います。もしならないなら、大人が家にいなくて暇になるだけなのでなくてもいいと思います。」")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「子供だけじゃなくて親も仕事が休める仕組みを市として作って欲しい。（有給休暇扱いにするなど）」")] }),
    new Paragraph({ style: "Quote", children: [new TextRun("「家族と関わるという目的をもっと伝えないとただの祝日気分になるため、なにか子供も目に触れやすい宣伝をすることで目的を理解する住民が増え、倉吉がもっと輝くと思う。」")] }),

    new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "分析日: 2026年1月20日", italics: true, size: 18 })] })
  ];
  return new Document({ styles: createDocStyles(), numbering: createNumbering(), sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }] });
};

// 教職員回答分析
const createKyoshokuin = () => {
  const children = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("【教職員】ふれあいホリデー アンケート自由記述分析（2025年度）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("概要")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "対象", bold: true }), new TextRun(": 教職員アンケート 問4「今後に向けてお気づきの点などのご意見をお書きください」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "回答数", bold: true }), new TextRun(": 164件（総回答365件中）")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "実施日", bold: true }), new TextRun(": 2025年11月21日（金）")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. 肯定的意見（継続希望・感謝）")] }),
    new Paragraph({ children: [new TextRun("制度への支持や継続を希望する声が多数寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「学期末に向けて大変忙しい時期なので、この時期の休みはありがたい。来年も是非お願いします。」"],
      ["2", "「教職員の心身のリフレッシュにもつながり、有用性が高いと感じています。」"],
      ["3", "「今回は4連休なのでとてもありがたかったです。」"],
      ["4", "「休みがあり、精神的な余裕が生まれ、有意義に過ごすことができました。仕事へのモチベーションも上がりました。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「素晴らしい取り組みであると思います。感謝します。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「普段平日にできないことができたり、まとまった休日をすごしてリフレッシュできたりして、とてもよい取組だと思います。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「天候の良い季節にまとまった休日ができ、有意義な時間を過ごすことができた。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「教職員も年休が取得しやすく、リフレッシュできる。ぜひ続けてほしい。」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("傾向")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("継続希望の声が多数")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("特に連休になった点が好評")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("心身のリフレッシュ効果を実感")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. 課題・問題点の指摘")] }),
    new Paragraph({ children: [new TextRun("制度の理念と現実のギャップを指摘する声。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「親子のふれあいが目的なのに、仕事の関係で休みが取れない保護者も多く、祖父母宅や児童クラブで過ごしている児童も多くいた。」"],
      ["2", "「生徒から『親がいないからどこにも行けなかった』『地域の活動に参加って書いてあったけど全然なかった』などの声が出ていました。」"],
      ["3", "「中学校は期末テストの直前がふれあいホリデーなので出かけにくいところがある。」"],
      ["4", "「市外在住だと我が子と休みが揃わない。やるなら全県で県民の日を休みにするなど、みんなが一斉に休みやすいようにするなどの工夫を考えたい。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("その他の声")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「教職員を休ませるためのホリデーだと保護者から不満の声も聞かれる。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「朝の8時過ぎから大勢で学校校庭でボール遊びを楽しむ子どもたちの姿を見て、少し辛い気持ちになった。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「休みを取ることのできない保護者の場合、1人で家で留守番する児童がどれだけいるのかな、と心配になります。」")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("「実施の目的と現実との乖離が大きい。」")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3大課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "保護者が休めない", bold: true }), new TextRun(" - 子どもだけで過ごす家庭が多い")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "テスト時期との重複", bold: true }), new TextRun(" - 中学校の期末テスト直前")] }),
    new Paragraph({ numbering: { reference: "numbered-list-1", level: 0 }, children: [new TextRun({ text: "実施地域の限定", bold: true }), new TextRun(" - 倉吉市のみで他地域と休みが揃わない")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. 改善提案・解決策")] }),
    new Paragraph({ children: [new TextRun("具体的な改善案が多数寄せられた。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「祝日のない6月にして欲しい。大人も子どもも新年度の疲れやGW、運動会の疲れもあると思う。」"],
      ["2", "「鳥取県全体か中部一斉にしてほしい。倉吉だけだと、郡部で働いている家族が休みづらい。」"],
      ["3", "「日直を立てなくてもよいように閉庁日にしてほしい。」"],
      ["4", "「受け入れ体制がもっと必要に思う。企業も一緒に進めていくことが必要。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("時期の変更")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("6月（祝日がない月）への移動")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("ゴールデンウィークの活用")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("テスト期間を避ける")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("実施範囲の拡大")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("全県統一実施")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("中部地区での統一実施")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("県民の日（9月12日）の活用")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("運用の改善")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("学校閉庁日化")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("地域イベントの充実")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("企業への働きかけ強化")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. 制度・運用への要望")] }),
    new Paragraph({ children: [new TextRun("勤務形態や休暇制度に関する具体的な要望。")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("代表的な意見")] }),
    createTable(["№", "意見"], [
      ["1", "「年休を取得しないといけないのが気になります。特休か何かいただけるとうれしいです。」"],
      ["2", "「管理職が勤務するというのは、少し？を感じる。学校閉庁にするべき。」"],
      ["3", "「非常勤講師は勤務を有さない日という扱いになると聞きました。年休が使えるとありがたいです。」"],
      ["4", "「ふれあいホリデーであることを全県に周知していただき、出張等を入れないように要請していただきたいです。」"]
    ], [800, 8500]),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("要望の分類")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "休暇制度", bold: true }), new TextRun(": 年休ではなく特別休暇への変更希望")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "会計年度職員", bold: true }), new TextRun(": 非常勤・時間給職員への配慮")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "管理職", bold: true }), new TextRun(": 管理職も休暇を取得できる体制")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun({ text: "出張調整", bold: true }), new TextRun(": ふれあいホリデー当日の出張回避")] }),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("集計サマリー")] }),
    createTable(["カテゴリ", "件数", "割合"], [
      ["肯定的（継続希望含む）", "53件", "32%"],
      ["課題・問題点の指摘", "10件", "6%"],
      ["改善提案", "47件", "29%"],
      ["特になし・無回答的", "54件", "33%"]
    ], [4000, 2500, 2500]),

    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("総評")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("評価できる点")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("制度自体は教職員から概ね好評")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("心身のリフレッシュ効果を多くの教職員が実感")] }),
    new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("連休化により有意義な時間を過ごせたという声が多い")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("今後の検討課題")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "保護者の休暇取得促進", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("企業・事業所への働きかけ強化")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("年度当初からの周知徹底")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "実施時期の再検討", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("6月（祝日のない月）への移動案が多数")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("中学校の期末テスト時期との調整")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "実施範囲の拡大", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("中部地区統一、または全県統一の検討")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("県民の日との連携可能性")] }),
    new Paragraph({ numbering: { reference: "numbered-list-2", level: 0 }, children: [new TextRun({ text: "制度設計の見直し", bold: true })] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("閉庁日化の検討")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("会計年度職員への配慮")] }),
    new Paragraph({ numbering: { reference: "sub-bullet", level: 0 }, children: [new TextRun("年休ではない休暇形態の検討")] }),

    new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "分析日: 2026年1月20日", italics: true, size: 18 })] })
  ];
  return new Document({ styles: createDocStyles(), numbering: createNumbering(), sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }] });
};

// メイン実行
async function main() {
  const docs = [
    { doc: createHogosha(), name: "保護者回答分析.docx" },
    { doc: createJido(), name: "児童生徒回答分析.docx" },
    { doc: createKyoshokuin(), name: "教職員回答分析.docx" }
  ];
  for (const { doc, name } of docs) {
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(`c:/DEV/ふれあいホリデー/${name}`, buffer);
    console.log(`Created: ${name}`);
  }
}
main().catch(console.error);
