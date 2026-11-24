const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, LevelFormat, BorderStyle, Footer, PageNumber } = require('docx');

const F = { ascii: 'Georgia', hAnsi: 'Georgia', eastAsia: 'Microsoft JhengHei', cs: 'Georgia' };
const GOLD = '8A6A2F', WINE = '7A2E2E', GRAY = '666666';

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 160, line: 360 }, ...opts.para, children: [new TextRun({ text, font: F, size: 23, ...opts.run })] });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: F })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: F })] });
const bullet = (t) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 100, line: 340 }, children: [new TextRun({ text: t, font: F, size: 23 })] });
const verse = (t, ref) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 160, after: 200, line: 360 },
  children: [new TextRun({ text: t, font: F, size: 23, italics: true, color: WINE }), new TextRun({ text: '　—— ' + ref + '（和合本）', font: F, size: 20, color: GRAY })],
});
const heart = (t) => new Paragraph({
  spacing: { before: 120, after: 200, line: 360 }, indent: { left: 360, right: 240 },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: GOLD, space: 12 } },
  children: [new TextRun({ text: '我們的感動　', font: F, size: 21, bold: true, color: GOLD }), new TextRun({ text: t, font: F, size: 23, color: '333333' })],
});
const img = (name, w, h, cap) => {
  const W = 430, scale = W / w, width = h > w ? 280 : W, height = Math.round(h * (width / w));
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 200, after: 60 }, children: [new ImageRun({ type: 'jpg', data: fs.readFileSync('img/' + name + '.jpg'), transformation: { width, height }, altText: { title: cap, description: cap, name } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: cap, font: F, size: 19, color: GRAY, italics: true })] }),
  ];
};
const L = [1400, 1051], P = [1051, 1400];
const bodyK = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '咖啡館裡的生日：維也納的下午與夜晚', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Café Diglas／Art Advent Karlsplatz｜2026 年 11 月 23 日', font: F, size: 21, color: GRAY })] }),

  p('從美泉宮的雪裡回到市中心，到晚上金色大廳的音樂會之間，我們有一整個下午和傍晚：在老咖啡館吃蛋糕、認識了一位正在過生日的新朋友，吃了一大盤肋排，再到卡爾廣場的聖誕市集看小羊。旅行最有趣的，常常不是景點，而是路上遇到的人。'),

  h1('一、Café Diglas：一百年的維也納咖啡館'),
  p('我們去的咖啡館是 Café Diglas，在聖史蒂芬大教堂附近的 Wollzeile 街 10 號。Diglas 家族做咖啡館的歷史可以追到 1875 年，Wollzeile 這一家則是 1923 年開的，已經傳了三代。店裡是很典型的維也納老咖啡館：深色木牆、紅色絨布長椅、大理石小圓桌和水晶壁燈，蛋糕都是自家甜點房烤的，最有名的是香蕉蛋糕和「Scheiterhaufen」（一種用麵包、蘋果做的布丁蛋糕）。據說作曲家雷哈爾（Franz Lehár）以前是這裡的常客。'),
  p('維也納的咖啡館不只是喝咖啡的地方。2011 年，「維也納咖啡館文化」被列入聯合國教科文組織的奧地利非物質文化遺產名錄，理由裡有一句很有名的話：咖啡館是一個「消費時間和空間，但帳單上只有咖啡」的地方。一杯咖啡可以坐一整個下午，看報紙、聊天，沒有人會趕你。'),
  verse('不可忘記用愛心接待客旅；因為曾有接待客旅的，不知不覺就接待了天使。', '希伯來書 13:2'),
  heart('一杯咖啡、一塊蛋糕，就能在陌生的城市裡坐下來，像在自己家的客廳一樣。這大概就是維也納人說的「咖啡館文化」。'),

  h1('二、咖啡館裡的生日朋友'),
  p('那天下午咖啡館人很多，一個空位都沒有。剛好旁邊一位印度朋友請我們幫他拍照，拍完以後，我們很不好意思地問他能不能跟他併桌，他很大方地說可以。'),
  p('坐下來聊天才知道，那天正好是他的生日。他因為到歐洲當交換學生，一個人來維也納幫自己慶生。我們問他有沒有生日願望，他很認真地說：「希望能把德文學好。」'),
  p('就這樣，三個從不同地方來的人，在一家一百年的維也納咖啡館裡，一起吃了一頓生日下午茶。'),
  ...img('cf_three', 1100, 825, '咖啡館裡的三人自拍：太太、生日的印度朋友和我；桌上是咖啡、拿鐵和蛋糕，窗外就是 Wollzeile 街'),
  verse('你們要彼此接待，像基督接待你們一樣，使榮耀歸與神。', '羅馬書 15:7'),
  heart('如果那天咖啡館有空位，我們就不會認識他了。一個沒有位子的下午，換來一個生日的故事。也祝他德文越來越好！'),

  h1('三、晚餐：一大盤肋排'),
  p('傍晚我們找了一家餐廳吃晚餐（店名待確認），木頭長椅、粉紅色的牆、圓圓的玻璃窗，很有山屋的感覺。點了一大塊烤豬肋排，配上涼拌高麗菜和兩種沾醬，還有一盤炸蘑菇（奧地利常見的「gebackene Champignons」）配塔塔醬。份量很大，兩個人吃得很飽，剛好有力氣去聽音樂會。'),
  ...img('cf_ribs', 825, 1100, '太太正在切一大塊烤肋排，旁邊是涼拌高麗菜、塔塔醬和烤肉醬'),
  ...img('cf_mushroom', 825, 1100, '一整盤炸蘑菇配塔塔醬，太太笑得很開心；背後是圓形花紋的玻璃窗'),
  verse('他們天天同心合意恆切地在殿裡，且在家中擘餅，存著歡喜誠實的心用飯。', '使徒行傳 2:46'),
  heart('旅行很累，但只要坐下來吃一頓熱騰騰的飯，就又有力氣了。'),

  h1('四、卡爾廣場的聖誕市集與小羊'),
  p('去金色大廳之前，我們先逛了卡爾廣場（Karlsplatz）的聖誕市集「Art Advent」（推測，待確認）。它在卡爾教堂前面，照片裡遠處那棟紅黃色、燈光很亮的建築，就是等一下要去的愛樂協會大樓。'),
  p('這個市集跟其他聖誕市集不太一樣：主打藝術和手工藝品，攤位要經過評審才能進來，而且它是全奧地利唯一一個餐飲全部有機認證的聖誕市集。給小朋友的活動也很多，其中一個是「稻草地」（Strohlandschaft），我們就在那裡看到了圍欄裡的綿羊和小山羊，在聖誕樹和稻草堆中間吃草。2026 年的市集從 11 月 13 日開到 12 月 23 日，每天中午 12 點到晚上 8 點，餐飲到晚上 9 點。'),
  ...img('cf_sheep', 825, 1100, '聖誕市集裡的稻草地：兩隻綿羊和一隻棕白色的小山羊在聖誕樹旁吃草，後面是市集的小木屋'),
  ...img('cf_market', 1100, 825, '晚上 7 點多的市集：一串串燈泡、賣熱紅酒和普羅賽克的小木屋，人潮後面是亮著燈的愛樂協會大樓'),
  verse('我是好牧人；我認識我的羊，我的羊也認識我。', '約翰福音 10:14'),
  heart('聖誕市集裡有一群小羊，讓人想起第一個聖誕夜，牧羊人在野地裡聽到好消息。熱鬧的城市中間，有這樣一個安靜的角落。'),

  h1('旅人小記'),
  bullet('時間：14:53–15:01 在 Café Diglas（影片 v15 太太吃蛋糕、v16 生日朋友聊天，GPS 在 Wollzeile）；傍晚吃肋排；19:22 逛卡爾廣場市集；19:56 進金色大廳。'),
  bullet('Café Diglas：Wollzeile 10，週一到週三 8:00–22:00、週四週五 8:00–22:30；下午人很多，可能要等位子或併桌。'),
  bullet('卡爾廣場：地鐵 U1、U2、U4 Karlsplatz 站，走到愛樂協會只要幾分鐘，很適合音樂會前逛。'),
  bullet('待確認：晚餐那家餐廳的名字（你傳的兩個 Google 地圖連結我打不開）；聖誕市集是不是卡爾廣場的 Art Advent；吃了哪幾種蛋糕；印度朋友願不願意出現在公開的網站上。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：wien.info（Café Diglas、Art Advent、2026 維也納聖誕市集）、visitingvienna.com、afar.com、monocle.com（Café Diglas 歷史）、coe.int（歐洲歷史咖啡館之路）、unesco.at、austria.info、wien.orf.at（維也納咖啡館文化列入非物質文化遺產）、vol.at（Art Advent 評審制度）。', font: F, size: 17, color: GRAY })] }),
];

const mk = (children) => new Document({
  styles: {
    default: { document: { run: { font: F, size: 23 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: F, size: 52, bold: true, color: '2B2B2B' }, paragraph: { spacing: { before: 0, after: 120 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: F, size: 32, bold: true, color: WINE }, paragraph: { spacing: { before: 480, after: 200 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: F, size: 27, bold: true, color: GOLD }, paragraph: { spacing: { before: 320, after: 160 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・咖啡館裡的生日　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_咖啡館裡的生日.docx', b));
