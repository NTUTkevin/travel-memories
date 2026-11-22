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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '國王湖：湖邊的步道，和新認識的朋友', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '德國・巴伐利亞・國王湖 Königssee｜2026 年 11 月 19 日', font: F, size: 21, color: GRAY })] }),

  p('國王湖旁的健行步道。這趟旅行的第二天，我們認識了來自馬來西亞的交換生宜恩，一起走了這段行程，一起享受旅行的樂趣。'),
  p('國王湖在德國，不在奧地利，但離薩爾斯堡很近，只隔著一條國界。我們這一天的路線，是從湖邊的小鎮一路走到步道上的觀景台。這一篇照例先把查到的歷史放在前面，再放上我們的照片，最後留給宜恩一個位置。'),

  h1('一、從老車站走向湖：湖濱街'),
  p('國王湖北岸的小鎮屬於舒瑙（Schönau am Königssee）。湖的北岸很窄，是整個湖唯一能走到水邊的地方，所以早早就蓋起了房子。從以前的車站（Königsseer Bahnhof）一直通到湖邊的這條路，叫做湖濱街（Seestraße），現在是行人專用的步行街，沿路有商店、餐館和老建築。街中央的廣場上立著一座方尖碑，是 1911 年為慶祝攝政王呂特波德九十歲生日所立。'),
  ...img('k_station', 1050, 1400, '一棟有小鐘樓的老房子，旁邊的旗幟寫著 Romy Schneider 的名字（建築與旗幟的地點待確認，可能是老車站一帶）'),
  p('這個小鎮還有一位名人。羅密・施奈德（Romy Schneider）1938 年生於維也納，她的母親瑪格達・施奈德（Magda Schneider）在 1935 年於舒瑙買下了「瑪麗恩格倫特」（Mariengrund）這棟房子，羅密和弟弟就在這裡長大；她後來也曾在薩爾斯堡附近的戈爾登斯坦（Goldenstein）寄宿學校念書，1955 年第一部《茜茜公主》電影上映，讓她一夜成了國際明星。舒瑙現在有一個常設展覽，叫做「一位世界巨星回家了」（Ein Weltstar kehrt heim）。路邊那面旗幟上印著一位女子的畫像和 Romy Schneider 的名字，應該就是這個展覽的宣傳旗，不過我們沒有進去看，這點先記下來。'),
  ...img('k_street', 1050, 1400, '太太站在湖濱街的石板路上，背後是「Hotel Schiffmeister」；遠處的方尖碑看起來就是 1911 年那一座（待確認）'),
  p('「Hotel Schiffmeister」的牆面特別好看：窗戶周圍畫了華麗的花邊框，窗與窗之間畫著一位穿著傳統服裝、坐在小船上划槳的女子。這種畫在牆上的裝飾，叫做「Lüftlmalerei」，是上巴伐利亞與蒂羅爾一帶的民間壁畫，源自巴洛克時期、用畫的方式「假裝」成建築上的窗框與壁龕，十八世紀開始在阿爾卑斯山腳下流行，富有的商人、農民和工匠，用漂亮的牆面展示自己的家境。名稱的由來有幾種說法，常被提到的一種，是奧伯阿默高畫家茨溫克（Franz Seraph Zwinck）的房子叫做「Zum Lüftl」。這一面牆畫的是湖上的船，正好說明了這個地方的生活。'),
  ...img('k_facade', 1050, 1400, '「Hotel Schiffmeister」牆面的壁畫：花邊窗框，中間是划著小船的女子（壁畫內容待確認）'),
  verse('凡事都有定期，天下萬務都有定時。', '傳道書 3:1'),
  heart('從車站到湖邊，一條路上有皇家的方尖碑、電影明星的童年、畫在牆上的小船。時間在這條街上疊了好幾層，而我們只是剛好在十一月的早晨走過。'),

  h1('二、一座只准電動船的湖'),
  p('國王湖長約八公里，最深的地方將近一百九十公尺，四周是貝希特斯加登阿爾卑斯山，大部分的湖屬於貝希特斯加登國家公園。湖邊沒有環湖的公路，要去對岸的景點，主要是搭船。'),
  p('自 1909 年起，湖上只准電動船行駛：那年 7 月 15 日，西門子—舒克特公司製造的電動船「Accumulator」開始載客，能載 38 人，長 12 公尺、寬 2.15 公尺，船身與船艙都用桃花心木做成。在這之前，湖上唯一的交通工具是划船。今天湖上有 19 艘電動船，電池晚上充電，白天不必再充。我們也搭上了這樣的船，一路坐到對岸的聖巴多羅買教堂，那一段另外寫成下一篇。'),
  ...img('k_wife_dock', 1400, 1050, '太太站在碼頭邊；背後的霧貼著山腰，湖水在霧下呈綠色'),
  verse('地的深處在他手中；山的高峰也屬他。海洋屬他，是他造的；旱地也是他手造成的。', '詩篇 95:4-5'),
  heart('湖邊不吵，船也不吵。為了保護一片水，人們在一百多年前就決定讓船安靜下來，這是我們很少想到的事。'),

  h1('三、步道、觀景台與那座紅頂小教堂'),
  p('離開湖邊的小鎮，我們沿著步道往上走。走到半山腰，回頭可以看見碼頭一帶的屋頂，和湖邊的小鎮；再往前，樹林中有一座木頭欄杆的觀景台，湖水像一條長長的峽灣，一直通到雪白的群山。這樣的畫面，很像是步道上的「畫家角」（Malerwinkel）：它在湖的北岸，是一個很受歡迎的觀景點，視線可以越過湖面，看到聖巴多羅買教堂和遠處的尖頂群峰。不過我們沒有查到我們站的位置的路標，所以先寫「可能是」。'),
  ...img('k_wife_house', 1400, 1050, '太太在步道邊：背後是霧中的山尖、鋪著金屬板的屋頂和湖濱的小鎮'),
  ...img('k_roofs', 1400, 788, '從步道往下看：一片灰黑色的屋頂，後面是綠色的湖水、草坡和雲霧裡的山'),
  p('湖中那座白牆紅頂的教堂，是聖巴多羅買（St. Bartholomä）朝聖教堂，在湖的半途，從岸邊走過去非常遠，主要靠船到達。它原本是貝希特斯加登修道院領主們的夏季住所，後來成為巴伐利亞國王的狩獵行宮。'),
  ...img('k_pole', 1050, 1400, '我們兩個人背對鏡頭，站在欄杆邊看湖：左邊一團雲貼在山腰，遠處的山頂有雪'),
  ...img('k_couple', 1400, 788, '木欄杆前的合照：背景是湖、峽谷，和遠處山頭的薄雪'),
  verse('諸天述說神的榮耀；穹蒼傳揚他的手段。', '詩篇 19:1'),
  heart('我們站在這裡，沒有說很多話。霧一下子遮住山、一下子又散開，我們只是看著，這就夠了。'),

  h1('四、鹽、主教與一條看不見的邊界'),
  p('國王湖屬於巴伐利亞，不過我們兩天後登上的薩爾斯堡要塞，和這裡其實有一段共同的歷史，就是鹽。'),
  p('在 1803 年以前，貝希特斯加登並不屬於巴伐利亞，而是一個叫「貝希特斯加登侯爵修道院領地」（Fürstpropstei Berchtesgaden）的小邦：它是神聖羅馬帝國直屬的諸侯國，由修道院的首席司鐸統治，領地幾乎被薩爾斯堡大主教國包圍。十二世紀末，貝希特斯加登和薩爾斯堡就曾因為鹽礦的開採起過爭執。兩邊都靠鹽致富，也因此常常為了邊界吵架。'),
  p('1803 年 2 月 11 日，最後一位首席司鐸，施羅芬貝格（Joseph Konrad von Schroffenberg）簽署了退位文件，同一天，薩爾斯堡大主教科羅雷多（Hieronymus Colloredo）也在拿破崙戰爭與教會領地世俗化的壓力下，辭去了世俗統治者的身分，兩個教會領地同時結束。貝希特斯加登先是併入薩爾斯堡選侯國，1805 年隨薩爾斯堡一起歸奧地利，1810 年再割讓給巴伐利亞；從那時起，國王湖一帶成了皇家狩獵場，聖巴多羅買也成了狩獵行宮。'),
  p('我們後來在要塞看到的「白色黃金」、牆上的那些主教，和這一天看到的湖，其實是同一片山、同一段歷史的不同角落。'),
  verse('地和其中所充滿的，世界和住在其間的，都屬耶和華。', '詩篇 24:1'),
  heart('國界、領地、主教、國王，都會換。湖還是這一個湖，山還是這一座山。這句經文，我們讀起來特別有重量。'),

  h1('五、宜恩：第二天認識的朋友'),
  p('這趟旅行的第二天，我們認識了來自馬來西亞的交換生宜恩。之後的行程，我們一起走，一起拍照、一起笑。'),
  ...img('k_pier', 1400, 1050, '在碼頭的自拍：宜恩、太太和我；後面有穿著迷彩軍服、戴著紅色貝雷帽的人走過（人物與單位待確認）'),
  ...img('k_selfie3', 1400, 1050, '觀景台上的自拍：宜恩、我和太太，後面是國王湖'),
  verse('不可忘記用愛心接待客旅；因為曾有接待客旅的，不知不覺就接待了天使。', '希伯來書 13:2'),
  heart('宜恩和我們一樣，都是離開家的人。在別人的國家裡遇見同樣來自亞洲的年輕人，一起走一段路，這是旅行送給我們的禮物。'),

  h1('旅人小記'),
  ...img('k_food', 787, 1400, '超市冷藏櫃裡一排排用保鮮膜包好的長麵包三明治，上方是蛋糕；標籤上有 SPAR 的字樣，上山前補給用'),
  bullet('國王湖離薩爾斯堡不遠；船班、票價與停靠站（如聖巴多羅買）請以官方網站為準：seenschifffahrt.de。'),
  bullet('十一月的國王湖多霧，山上已經有雪；我們戴著帽子、穿著羽絨外套，走起路來還是很暖。'),
  bullet('這一篇還有幾處待確認：觀景台是不是畫家角、旗幟與建築的地點、湖濱街壁畫的內容、碼頭上軍人的單位。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：seenschifffahrt.de（Bayerische Seenschifffahrt 的歷史與船隊）、berchtesgaden.de（Seestraße、Romy Schneider Ausstellung、Der Königssee）、komoot（Königssee 景點與 Malerwinkel）、germansights.com（Königssee）、Wikipedia〈Berchtesgaden Provostry〉〈Lüftlmalerei〉〈Electorate of Salzburg〉、Bayern.travel 與 Ammergauer Alpen（Lüftlmalerei）、zobodat.at（貝希特斯加登侯爵修道院的紋章與 1803 年世俗化）、archivportal-d.de（Fürstpropstei Berchtesgaden）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・國王湖　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('國王湖.docx', b));
