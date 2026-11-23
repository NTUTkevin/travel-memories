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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '抵達維也納：中央車站旁的家', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Wien Hauptbahnhof／Favoriten｜2026 年 11 月 22 日', font: F, size: 21, color: GRAY })] }),

  p('離開薩爾斯堡，我們來到這趟旅行的最後一站：維也納。接下來四天三夜，我們住在中央車站（Wien Hauptbahnhof）旁邊的一間 Airbnb，從車站走過去大約 9 分鐘、600 公尺。11 月 22 日中午過後到，傍晚先到超市採購，晚上就在住處自己煮。這一篇記錄的，是我們在維也納的「家」。'),

  h1('一、中央車站：從南站到「菱形屋頂」'),
  p('維也納中央車站是一座很新的車站。19 世紀以後，維也納的鐵路是由好幾個「盡頭式」車站組成的：往南的南站（Südbahnhof）、往東的東站（Ostbahnhof），兩條鐵路在這裡呈直角，火車開到底就得調頭，旅客要轉車還得換車站。第一座南站在 1841 年啟用，大概就在今天中央車站的位置；後來又改建過好幾次。'),
  p('2009 年 12 月，舊南站停止營運並拆除，原地蓋起一座可以讓火車往東南西北直接穿過的新車站。2014 年 10 月 10 日新車站開幕，同年 12 月第一批長途火車開進來，2015 年 12 月全面營運。車站最有名的是那片由菱形格子組成的屋頂（Rautendach），面積三萬一千平方公尺，用了七千噸鋼材，據說和巴黎艾菲爾鐵塔一樣重。舊南站入口那頭帶翅膀的聖馬可獅子像，從 1873 年就守在這裡，新車站也把它留了下來。'),
  ...img('w_station', 619, 1100, '在維也納的地鐵月台上，我戴著藍色毛帽、圍格子圍巾，一手拉著行李箱，一邊用手機查路；頭頂是紅色弧線的天花板（車站與路線待確認，可能是中央車站的 U1 月台）'),
  verse('你出你入，耶和華要保護你，從今時直到永遠。', '詩篇 121:8'),
  heart('拖著行李換了好幾種交通工具，終於到了最後一站。一路上平平安安，這就是最大的恩典。'),

  h1('二、住在法沃里滕：維也納的第十區'),
  p('中央車站和我們的住處都在維也納第十區「法沃里滕」（Favoriten）。這個名字來自附近一座叫「Favorita」的皇家狩獵城堡，那座城堡今天在第四區，變成了一所叫 Theresianum 的學校。法沃里滕在 1874 年成為獨立的一區，因為有很多百年以上的市營住宅，加上工廠聚集，慢慢變成典型的「工人區」。'),
  p('今天的法沃里滕是維也納人口最多的一區，大約有二十二萬人，比奧地利第三大城林茨還多，所以常被戲稱為「奧地利第三大城」。中央車站旁邊原本南站的鐵道用地，也變成了一整個新的住宅區。'),
  p('我們的 Airbnb 是頂樓的公寓：臥室有一整面落地窗，外面是一個大陽台，可以看到對面維也納老公寓的屋頂；客廳有一張大餐桌和小廚房。桌上一個紅色的螺旋花瓶插著一朵白花，牆上掛著一幅斑馬的畫。'),
  ...img('w_bed', 619, 1100, '頂樓的臥室：灰色的床頭、白色床單，床尾放著折好的毛巾，牆上一幅黑白斑馬畫；落地窗外是陽台和維也納的屋頂'),
  ...img('w_dining', 619, 1100, '餐桌靠在落地窗旁：深灰色桌面、四張磚紅色的椅子，三盞長筒吊燈從天花板垂下來，桌上是紅色螺旋花瓶'),
  ...img('w_kitchen', 619, 1100, '從餐桌看過去的小廚房：木紋吊櫃、電磁爐、烤箱、熱水壺、烤麵包機和膠囊咖啡機，什麼都有'),
  verse('我必安然躺下睡覺，因為獨有你耶和華使我安然居住。', '詩篇 4:8'),
  heart('旅行了好幾天，終於有一個可以把行李攤開、可以自己煮飯的地方。有廚房的住處，就有家的感覺。'),

  h1('三、第一個晚上：去超市採購'),
  p('放下行李之後，我們先去超市採買接下來幾天要吃的東西。11 月 22 日是星期天，奧地利的商店週日大多休息，只有火車站裡的超市照常營業，所以星期天晚上的中央車站超市總是特別熱鬧。從地圖上看，車站和住處之間就有一家 BILLA（我們是在哪一家買的，待確認）。'),
  p('BILLA 是奧地利最常見的超市之一。它的創辦人卡爾・弗拉舍克（Karl Wlaschek）年輕時是一位酒吧鋼琴師，1953 年 12 月 7 日，他用當鋼琴師存下的錢，在維也納第五區開了第一家店，是一間賣打折名牌商品的香水店。到了 1960 年已經有 45 家分店，開始賣食品；1961 年改名為 BILLA，意思是「Billiger Laden」——「便宜的商店」，還引進了當時很新的自助購物方式。1996 年，他把 BILLA 賣給了德國的 REWE 集團。'),
  verse('我們日用的飲食，今日賜給我們。', '馬太福音 6:11'),
  heart('在異國的超市裡推著推車，一起看看當地人都吃什麼、想想晚餐要煮什麼，也是旅行裡很幸福的一段時間。'),

  h1('四、老婆的料理，加上台灣的曾拌麵'),
  p('晚餐是太太下廚：用從台灣帶來的曾拌麵，加上在超市買的鮭魚和紅椒一起拌炒，再煮一鍋高麗菜蛋花湯。曾拌麵是台灣主持人曾國城和團隊推出的乾拌麵，2015 年開始在網路上販售，第一年就很受歡迎，後來還賣到美國、加拿大、香港、馬來西亞等地。在維也納吃到熟悉的台灣味，特別溫暖。'),
  ...img('w_dinner', 825, 1100, '第一晚的晚餐：兩碗曾拌麵拌鮭魚和紅椒，兩碗高麗菜蛋花湯；中間是那個紅色螺旋花瓶，旁邊放著油瓶和鍋蓋'),
  verse('你只管去歡歡喜喜吃你的飯，心中快樂喝你的酒，因為神已經悅納你的作為。', '傳道書 9:7'),
  heart('謝謝老婆的料理，讓我們在維也納的第一個晚上就有著落。一碗台灣的拌麵、一碗熱湯，比任何餐廳都好吃。'),

  h1('旅人小記'),
  bullet('交通：維也納中央車站有長途火車、區域火車、S-Bahn 和地鐵 U1，從薩爾斯堡搭火車可以直接到。我們的住處從車站走路大約 9 分鐘。'),
  bullet('住宿：這次住 Airbnb，有廚房和陽台，四天三夜（11 月 22 日到 25 日，待確認）。為了隱私，網站上不寫住處的地址。'),
  bullet('超市：奧地利的超市週日和假日大多不營業，平日晚上也比較早關；週日想買東西，可以去中央車站裡的超市（以現場公告為準）。'),
  bullet('待確認：月台照片是在哪一站、哪一條線；我們是在哪一家超市採購的；你們說的「超市市集」是不是除了超市還去了市集；住宿日期。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：wien.orf.at（中央車站開幕、菱形屋頂、南站歷史）、Wikipedia（Wien Hauptbahnhof、Wien Südbahnhof、Favoriten、Billa、Karl Wlaschek）、meinbezirk.at（車站歷史）、austria-forum.org（Favoriten、Billa、Karl Wlaschek）、ganz-wien.at、維也納市政府《Favoriten in Zahlen 2024》、REWE Group（BILLA 70 週年）、食力 foodNEXT、天下雜誌（曾拌麵）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・中央車站旁的家　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_中央車站旁的家.docx', b));
