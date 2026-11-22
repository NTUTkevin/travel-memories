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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '薩爾斯堡聖誕市集：五百年的冬夜燈火', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡・Salzburger Christkindlmarkt｜2026 年 11 月 21 日（晚上）', font: F, size: 21, color: GRAY })] }),

  p('在薩爾斯堡的那一天，我們白天走過莫札特故居、要塞和美術館，傍晚到了僧侶山上。天黑以後，整座老城亮了起來：山上的要塞、大教堂的圓頂，還有山腳下廣場上一整片聖誕市集的小木屋。我們去的時候，聖誕市集才剛開幕第三天。'),

  h1('一、黃昏的僧侶山：從上面看見整座城'),
  p('傍晚五點多，我們在僧侶山（Mönchsberg）上，也就是現代美術館所在的那座山崖，往下看老城。左邊是薩爾察赫河（Salzach）和河邊的燈，中間是我們白天去過的大學教堂和大教堂的圓頂，右上方就是山頂上打著燈的霍亨薩爾斯堡要塞。薩爾斯堡的老城區在 1997 年被列為聯合國世界遺產，從上面看，可以看得很清楚：教堂、廣場、窄巷，一層一層擠在河和山之間。'),
  ...img('x_night', 1100, 825, '從僧侶山上看薩爾斯堡夜景：右上是亮著燈的霍亨薩爾斯堡要塞，中間是大學教堂的白色立面和大教堂的圓頂與雙塔，左邊是薩爾察赫河；遠處山頂上有一座亮紅燈的發射塔（可能是蓋斯山 Gaisberg，待確認）'),
  verse('在黑暗中行走的百姓看見了大光；住在死蔭之地的人有光照耀他們。', '以賽亞書 9:2'),
  heart('白天走過的每一個地方，晚上從山上看，全都變成一顆一顆的光點。走了一整天很累，可是看到這一幕，覺得很值得。'),

  h1('二、從「雜貨市集」到聖誕市集'),
  p('薩爾斯堡聖誕市集（Salzburger Christkindlmarkt）在大教堂廣場（Domplatz）和主教宮廣場（Residenzplatz）上，被認為是世界上最古老的聖誕市集之一。它的源頭可以追到 15 世紀末：當時大教堂廣場上有一個叫「Tandlmarkt」的市集，賣各式各樣的雜貨，文獻一般寫是 1491 年。到了 17 世紀，它改叫「尼古拉市集」（Nikolaimarkt），在 12 月 6 日聖尼古拉節的前後各兩週開放，賣點心、娃娃和各種雜物。'),
  p('今天這種形式的聖誕市集是 1974 年開始的，2003 年擴大到主教宮廣場，2024 年剛過 50 週年。現在兩個廣場上有 96 間綠色小木屋，一年有將近一百萬人來逛。2026 年的市集從 11 月 19 日開到 2027 年 1 月 1 日，所以我們 11 月 21 日去的時候，才開幕第三天。'),
  p('「Christkindl」是「聖嬰」的意思。這個名字也有一段和宗教改革有關的故事：以前送禮物給孩子的是聖尼古拉，在 12 月 6 日送；馬丁・路德反對敬拜聖人，改成由「聖基督」在聖誕節送禮物，後來慢慢變成今天德語區孩子等待的「Christkind」。有趣的是，到了 19 世紀，聖嬰反而在天主教地區流行起來，薩爾斯堡就是一個天主教城市。'),
  ...img('x_trees', 825, 1100, '太太戴著紅色貝雷帽，站在一整排掛滿小燈泡的聖誕樹前；後面是一座巴洛克式的白色鐘樓和粉紅色的建築（位置與鐘樓名稱待確認，可能在莫札特廣場或主教宮廣場附近）'),
  ...img('x_selfie', 825, 1100, '在聖誕樹前自拍：太太紅帽紅圍巾，我戴毛帽、圍著格子圍巾'),
  verse('不要懼怕！我報給你們大喜的信息，是關乎萬民的；因今天在大衛的城裡，為你們生了救主，就是主基督。', '路加福音 2:10–11'),
  heart('一個從五百多年前就開始的市集，到今天還在同一個廣場上。聖誕節的燈每年都會亮起來，因為大家記得那個「大喜的信息」。'),

  h1('三、一杯熱紅酒，一個杯子'),
  p('在市集裡買熱紅酒（Glühwein）或熱潘趣（Punsch），飲料會裝在市集專屬的陶瓷杯裡，要多付一筆押金（Pfand）；喝完可以把杯子還給任何一個飲料攤拿回押金，也可以直接把杯子帶回家當紀念。今年熱紅酒一杯大約 4.5 到 5.8 歐元，每個攤位自己定價；2026 年還另外推出了限量 2,026 個金色的收藏杯。'),
  p('太太手上那個深藍色的杯子，上面用金色畫著主教宮廣場（RESIDENZPLATZ）的建築，寫著「Salzburger Christkindlmarkt」。我們這杯押金多少，待確認；杯子有沒有帶回家，也請補上。'),
  ...img('x_mug', 619, 1100, '太太捧著市集的深藍色陶瓷杯，杯上是金色的主教宮廣場圖案和「Salzburger Christkindlmarkt」字樣；背後是攤位的黑板和寫著「Kassa 2」的牌子'),
  verse('你為我擺設筵席；你用油膏了我的頭，使我的福杯滿溢。', '詩篇 23:5'),
  heart('冷冷的晚上，兩隻手捧著一杯熱熱的飲料，就覺得很滿足。這個杯子，就是這一天最好的紀念品。'),

  h1('四、莫札特廣場的溜冰場'),
  p('從主教宮廣場走幾步就到了莫札特廣場（Mozartplatz），冬天這裡會變成一個溜冰場，叫「Eiszauber」（冰之魔法）。溜冰場就在莫札特雕像的腳下，旁邊是新主教宮、聖誕博物館、主教宮廣場和大教堂，由聖誕市集協會經營，一直開到 1 月底。大人門票 6.5 歐元，晚上七點以後比較便宜，現場可以租溜冰鞋；上一季有兩萬多人來溜冰。'),
  ...img('x_rink', 618, 1100, '影片截圖：莫札特廣場的溜冰場，冰面被燈光照成藍紫色，孩子們扶著欄杆溜冰；後面是巴洛克式的白色鐘樓和掛著聖誕燈的建築'),
  verse('在至高之處榮耀歸與神！在地上平安歸與他所喜悅的人！', '路加福音 2:14'),
  heart('看著小朋友在莫札特的腳下一圈一圈地溜冰，跌倒了又爬起來，整個廣場都是笑聲。'),

  h1('五、《平安夜》的故鄉'),
  p('全世界最有名的聖誕歌《平安夜》（Stille Nacht! Heilige Nacht!），作詞的約瑟夫・莫爾（Joseph Mohr）就是在薩爾斯堡出生的。他 1792 年 12 月 11 日出生，是一位編織女工的非婚生子，父親是大主教衛隊的士兵，在他出生那年就逃走了；他和媽媽、外婆住在河對岸史坦巷（Steingasse）31 號一間又冷又暗的小房子裡，從小家境貧窮。'),
  p('莫爾當上神父以後，1816 年寫下了〈平安夜〉的詩。1818 年聖誕節前，他在奧本多夫（Oberndorf）請好朋友、教師兼風琴師法蘭茲・格魯伯（Franz Xaver Gruber）為這首詩譜曲；聖誕夜彌撒之後，兩人用一把吉他伴奏，第一次唱出了這首歌。前兩天在拉姆紹教堂那一篇，我們也提過格魯伯和莫爾。在奧地利，《平安夜》被當成國寶，傳統上在聖誕夜之前不在公開場合演唱。'),
  ...img('x_alley', 618, 1100, '影片截圖：晚上回程走過的老城小巷，兩旁掛著鍛鐵招牌和金色的聖誕燈飾，巷子盡頭是一座鐘樓（巷名待確認，可能是猶太巷 Judengasse 一帶）'),
  verse('光照在黑暗裡，黑暗卻不接受光。', '約翰福音 1:5'),
  heart('寫出〈平安夜〉的人，是一個在又冷又暗的小巷裡長大的窮孩子。最溫暖的歌，有時候就是從最冷的地方唱出來的。'),

  h1('旅人小記'),
  bullet('時間：2026 年的薩爾斯堡聖誕市集從 11 月 19 日開到 2027 年 1 月 1 日；平日大約 10:00–20:30，週五到 21:00，週六 9:00–21:00，12 月 24 日只開到 15:00（以官方網站 christkindlmarkt.co.at 公告為準）。'),
  bullet('地點：大教堂廣場和主教宮廣場，旁邊的莫札特廣場有溜冰場。週末晚上人非常多，注意隨身物品。'),
  bullet('杯子押金：喝完可以把杯子還給任何一個飲料攤拿回押金；想留作紀念就直接帶走。部分攤位不一定能刷卡，最好帶一點現金。'),
  bullet('看夜景：傍晚可以先搭電梯上僧侶山，從現代美術館旁邊的觀景台看老城和要塞，再下來逛市集。'),
  bullet('這一篇的待確認：太太站的那排聖誕樹和白色鐘樓在哪裡、遠處亮紅燈的山是不是蓋斯山、杯子押金多少和有沒有帶回家、最後那條小巷的名字。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：salzburg.info（Salzburger Christkindlmarkt、Eiszauber am Mozartplatz、Joseph Mohr 專文）、salzburgerland.com（聖誕市集傳統、Joseph Mohr 1792–1848）、salzburg24.at、weekend.at（2026 年市集新聞，2026/11/17）、tips.at（溜冰場）、stillenacht.at、Wikipedia（Joseph Mohr）、evangelisch.de 與漢堡大學（Christkind 與路德）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・薩爾斯堡聖誕市集　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('薩爾斯堡聖誕市集.docx', b));
