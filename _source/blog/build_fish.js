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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '漁人堡與馬加什教堂：雪人、聖誕樹和七座白塔', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '匈牙利・布達佩斯・Halászbástya & Mátyás-templom｜2026 年 11 月 26 日', font: F, size: 21, color: GRAY })] }),

  p('11 月 26 日下午看完圖魯鳥和鏈橋，我們沿著城堡山往北走，到城堡區另一頭的三位一體廣場。這段路走得很慢：先在公園的雪地上堆了一個小雪人，再到馬加什教堂前的聖誕樹下拍照，最後爬上漁人堡，看雪後的多瑙河、國會大廈和整個佩斯。從 14:51 到 15:10，不到半小時，卻是這趟旅行最「冬天」的一段。'),

  h1('一、雪地上的小雪人'),
  p('那天的雪一直下到下午，城堡區的草地和松樹都蓋上了厚厚一層白。我們在一片松樹下的雪地上（地點待確認），用手堆了一個小雪人：兩片黃葉子當眼睛，一團乾草當鼻子，兩根小樹枝當手。我們兩個輪流蹲在它旁邊比 YA。對從台灣來的我們來說，能親手堆雪人，本身就是一件很稀奇的事。'),
  ...img('fb_bridge', 750, 1000, '往北走之前，在城堡山的紅色木欄杆前合照：背後是鏈橋、格雷沙姆宮、左邊的國會大廈和右邊的聖伊什特萬大教堂（拍攝時間與位置待確認）'),
  ...img('fb_snow3', 750, 1000, '14:51，太太戴上帽兜，蹲在剛堆好的小雪人旁邊；雪人用黃葉當眼睛、乾草當鼻子'),
  ...img('fb_snow2', 750, 1000, '14:51，換我蹲下來，兩手比 YA；後面是積雪的松樹和亮著燈的小店'),
  ...img('fb_snow1', 750, 1000, '14:52，再一張：雪人的樹枝手，和背後被雪壓彎的松枝'),
  verse('求你洗滌我，我就比雪更白。', '詩篇 51:7'),
  heart('雪人只有一個下午的生命，第二天也許就化了。可是我們會記得那雙黃葉子的眼睛，記得兩個人蹲在雪地裡笑的樣子。'),

  h1('二、馬加什教堂：皇帝和西西在這裡加冕'),
  p('三位一體廣場上最顯眼的，是有彩色菱形屋頂和細長尖塔的馬加什教堂（Mátyás-templom），正式名稱是布達聖母教堂。它的名字來自 15 世紀的馬加什國王，他的兩次婚禮都在這裡舉行。1541 年鄂圖曼人佔領布達後，這座教堂被改成城裡最主要的清真寺；1686 年圍城時，教堂的一面牆被砲火打垮，傳說牆後露出一尊聖母像，這就是布達的「聖母奇蹟」。之後它被改回天主教堂，換成了巴洛克風格。'),
  p('1867 年匈奧「折衷方案」簽定之後，6 月 8 日，法蘭茲・約瑟夫和西西就在這座教堂裡加冕為匈牙利國王和王后，李斯特為這場典禮寫的《加冕彌撒》也在這裡首演。1916 年，最後一位匈牙利國王卡爾四世和齊塔也在這裡加冕——前一篇聖伊什特萬廳打開聖冠箱子的，就是這一次。'),
  p('今天看到的樣子，是 1874 到 1896 年建築師舒勒克（Schulek Frigyes）的新哥德式大整修：他照著中世紀的樣子把教堂重新做回來，又加上自己的設計，像菱形圖案的屋瓦和滿是怪獸滴水嘴的尖塔，當時還引起不少爭議。屋頂的彩色釉瓦是茲索奈工廠做的——就是做聖伊什特萬廳大壁爐的那家。'),
  p('教堂前的廣場上立著一棵掛滿紅色和金色大球的聖誕樹，枝葉上全是雪，旁邊擺著市集的長桌長椅（三位一體廣場聖誕市集，待確認）。'),
  ...img('fb_couple', 750, 1000, '14:54，我們在積雪的聖誕樹下合照，背後是馬加什教堂的白色鐘樓、玫瑰窗和彩色屋頂'),
  ...img('fb_tree', 750, 1000, '14:53，太太站在聖誕樹旁，右後方是教堂南側的菱形彩瓦屋頂和黑色小尖塔'),
  ...img('fb_church', 750, 1000, '14:56，抬頭看馬加什教堂的鐘樓：一層層的哥德式尖窗和雕花，門廊上方的石欄杆積了一排雪'),
  verse('耶和華啊，我喜愛你所住的殿和你顯榮耀的居所。', '詩篇 26:8'),
  heart('從維也納到布達佩斯，我們一路跟著法蘭茲・約瑟夫和西西走。在美泉宮看他們的房間，在霍夫堡看西西的體操室，最後來到他們加冕的教堂門口。故事好像在這裡繞了一圈。'),

  h1('三、漁人堡：從來沒打過仗的城牆'),
  p('教堂旁邊那一排白色的塔樓和拱廊，就是漁人堡（Halászbástya）。它看起來像中世紀的城堡，其實是 1895 到 1902 年才蓋的，同樣由舒勒克設計，是新羅馬式風格，建在一段舊城牆的基礎上；它從來沒有防守過任何東西，而是一座讓人欣賞城市的觀景台。'),
  p('名字的由來有兩種說法：一說中世紀時，這段城牆由住在山下「漁人城」的漁夫公會負責防守；也有人認為，名字其實是來自山下的那個區。七座尖頂的石塔，代表 895 年帶領馬扎爾各部落進入喀爾巴阡盆地的七位首領，也呼應布達佩斯為建國千年所做的一連串建設——就像我們那天晚上坐的千禧地鐵。'),
  p('從漁人堡往東看，多瑙河對岸左邊是新哥德式的國會大廈，右邊遠處是聖伊什特萬大教堂的圓頂；下面的屋頂全是白的，河邊的紅色尖塔是山下的教堂（名稱待確認）。我們在 15:10 離開前，又抬頭拍了一張塔樓：尖尖的屋頂上壓著一層厚厚的雪，像蓋了一床棉被。'),
  ...img('fb_wife2', 750, 1000, '14:58，太太在漁人堡的欄杆邊，背後是多瑙河、國會大廈和山下被雪覆蓋的屋頂'),
  ...img('fb_wife1', 1000, 750, '14:58，太太抬頭看塔樓，右邊是漁人堡圓形孔洞的欄杆和拱窗'),
  ...img('fb_view', 1000, 750, '14:59，從漁人堡望向佩斯：左邊國會大廈、右邊遠處聖伊什特萬大教堂，右側是漁人堡的拱廊與塔樓'),
  ...img('fb_roof', 750, 1000, '15:10，漁人堡的一座塔樓：尖屋頂上積滿了雪，拱門上有一顆顆圓形的石頭裝飾'),
  verse('眾山怎樣圍繞耶路撒冷，耶和華也照樣圍繞他的百姓。', '詩篇 125:2'),
  heart('一座從來沒打過仗的城牆，蓋來只是為了讓人好好看看自己的城市。我們在雪裡站了很久，覺得被這座城市溫柔地圍繞著。'),

  h1('旅人小記'),
  bullet('時間（依檔名）：11 月 26 日 14:31 圖魯鳥與鏈橋（見〈布達城堡〉）→ 14:51–14:52 堆雪人 → 14:53–14:56 馬加什教堂與聖誕樹 → 14:58–15:10 漁人堡 → 17:16 坐千禧地鐵去歌劇院。'),
  bullet('交通：從布達城堡沿城堡區往北走約 15–20 分鐘，或搭 16 號公車到 Szentháromság tér。漁人堡下層的露台免費，上層塔樓要買票（我們有沒有上去，待確認）；馬加什教堂內部也要買票。'),
  bullet('小提醒：雪天石板路很滑；漁人堡和教堂前人潮很多，背包要顧好。'),
  bullet('待確認：堆雪人的地點（是不是城堡區的公園、旁邊亮燈的小店是什麼）；紅色木欄杆合照的時間和位置；有沒有進馬加什教堂或上漁人堡塔樓；教堂前是不是三位一體廣場聖誕市集；河邊紅色尖塔教堂的名稱。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：Wikipedia（漁人堡、馬加什教堂）、budavar.btk.mta.hu（漁人堡名稱、布達的加冕）、matyas-templom.hu（教堂歷史、1867/6/8 加冕）、visithungary.com（清真寺、茲索奈屋瓦）、structurae、History Hit、Lonely Planet、Daily News Hungary、The Catholic Travel Guide（1874–1896 舒勒克整修）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・布達佩斯・漁人堡與馬加什教堂　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('布達佩斯_漁人堡與馬加什教堂.docx', b));
