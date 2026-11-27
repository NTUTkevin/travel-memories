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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '布達佩斯歷史博物館：聖伊什特萬廳與地底的中世紀王宮', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '匈牙利・布達佩斯・Budapesti Történeti Múzeum – Vármúzeum｜2026 年 11 月 26 日', font: F, size: 21, color: GRAY })] }),

  p('外面雪越下越大，中午我們躲進布達城堡最南邊的 E 棟：布達佩斯歷史博物館，也叫「城堡博物館」（Vármúzeum）。我們買的是「成人聯票」，含 12:30 入場、附平板導覽的聖伊什特萬廳（Szent István-terem），一張 5,900 福林。從 12:19 看到快 14 點，先上樓看金碧輝煌的聖伊什特萬廳，再一路往下走進地底的中世紀王宮和哥德式禮拜堂。'),

  h1('一、雪中的獅子庭院'),
  p('戰後重建時，皇宮的各棟分給不同的機構，E 棟分給布達佩斯歷史博物館，博物館在 1967 年搬進來。入口在皇宮中間的獅子庭院（Oroszlános udvar）。我們從博物館的窗戶往外看：庭院整片都是白的，對面是有大拱門的巴洛克式側翼，右上角露出國家美術館的圓頂，幾個人在雪地上踩出一條條腳印；另一扇窗外，城牆和屋頂都積了雪，遠方的多瑙河和佩斯只剩一片藍灰色的霧。'),
  ...img('hm_court', 750, 1000, '12:19，從博物館窗戶看雪中的庭院：大拱門側翼、右邊的綠色圓頂，地上滿是腳印（是否就是獅子庭院，待確認）'),
  ...img('hm_fog', 750, 1000, '窗外的城牆與屋頂積滿了雪，多瑙河和對岸的佩斯都隱在霧裡'),
  verse('他對雪說：要降在地上。', '約伯記 37:6'),
  heart('窗外是雪，窗內是一千年的故事。能在下雪天躲進博物館慢慢看，是旅行裡很奢侈的一段時間。'),

  h1('二、聖伊什特萬廳：從廢墟裡回來的金色房間'),
  p('聖伊什特萬廳以匈牙利的建國君主聖伊什特萬（聖史蒂芬）命名，是二十世紀初布達皇宮最重要的室內空間之一。豪斯曼（Hauszmann Alajos）找來當時最紅的工藝家一起打造：茲索奈（Zsolnay）陶瓷工廠、木工泰克・安德烈（Thék Endre）和諾伊施洛斯兄弟。整個房間先在泰克的木工坊裡組裝起來，送去 1900 年的巴黎萬國博覽會，還拿到評審團大獎，之後才裝進皇宮。這裡曾是國王的工作午宴廳、謁見前的等候室，舞會時則是牌室；1916 年 12 月卡爾四世加冕前一天，裝聖冠的箱子就是在這裡打開的。'),
  p('第二次世界大戰時，聖伊什特萬廳整個燒毀，共產時期也沒有重建，空間被改成博物館的版畫庫房。「國家豪斯曼計畫」先把炸毀的南側連接翼重新蓋起來，再依照原始設計圖、老照片，甚至當年的發票，把房間一件件做回來；有些相配的天花板飾板還是在茲索奈區找到的。2021 年 8 月 20 日（匈牙利的聖伊什特萬節）重新開放。'),
  p('房間裡最搶眼的是茲索奈做的大壁爐，據說是全國最大的一座；壁爐上方有一尊白鬍子、戴王冠的國王胸像，中間是匈牙利的紅白條紋盾徽（胸像是不是聖伊什特萬，待確認）。牆面是深色雕花木板，家具和裝飾用了 24K 金，天花板是雕刻的格子天花板，還有三盞各重 150 公斤的吊燈和龍紋拼木地板。藍底金色的牆布上，印著匈牙利聖冠和雙十字的圖案。'),
  p('牆上還有一排阿爾帕德王朝的聖王與國王：我們拍到聖拉斯洛（Szent László）、卡爾曼（Kálmán）、貝拉三世（III. Béla），以及聖伊什特萬的兒子聖伊姆雷（Szent Imre），他手裡拿著象徵純潔的百合。（這些人像是畫的，還是茲索奈的陶板，待確認。）'),
  ...img('hm_fire', 750, 1000, '12:42，茲索奈大壁爐：上方是戴王冠的白鬍子國王胸像，中間是紅白條紋盾徽，四周是一層層的拱形、柱子和金色花紋'),
  ...img('hm_fire2', 750, 1000, '12:35，壁爐的近看：拱形格窗、扭繩紋的爐口和金色的爐柵'),
  ...img('hm_kings', 750, 1000, '12:33，牆上的三位聖王與國王：聖拉斯洛、卡爾曼、貝拉三世；下面是紅絲絨扶手椅、矮凳和雕花書桌'),
  ...img('hm_imre', 750, 1000, '12:33，手拿百合的聖伊姆雷，和一張獅子扶手的紅絲絨椅子'),
  ...img('hm_wall', 750, 1000, '12:36，雕花木牆、燭台式壁燈，以及印著聖冠與雙十字的藍金色牆布和刺繡窗簾'),
  verse('那些出於你的人必修造久已荒廢之處；你要建立拆毀累代的根基。', '以賽亞書 58:12'),
  heart('一個被燒成灰的房間，一百年後被一雙雙手重新做出來。我們看到的不只是金色，而是很多人願意把失去的東西找回來的那份心。'),

  h1('三、地圖、洪水與城市的記憶'),
  p('樓上的展覽講布達佩斯這座城市的歷史。一進展區，紅色的牆上寫著豪斯曼 1900 年的一句話，意思是「願這座王宮在未來許多年，都是我們敬愛的國王真正的家」；下面掛著一幅從山下仰望城堡山的油畫，前景是一群在坡地上的平民。'),
  p('一幅 1684 年的版畫〈布達與佩斯景觀〉（Michael Wening 等人製作），從蓋勒特山（Szent Gerhardsberg）往下看：左邊山上是布達，多瑙河另一邊是佩斯，前景山坡上排滿了軍營。1541 年鄂圖曼人佔領布達，統治了 145 年；1684 年哈布斯堡軍隊第一次圍城失敗，1686 年神聖同盟才奪回布達，但那時的布達已經是一片瓦礫。展區裡還有一塊十六世紀末的石灰岩殘片，是清真寺講壇（mimber）欄杆上的星形幾何花紋，也是那段鄂圖曼歲月留下的痕跡。'),
  p('另一面綠牆上，掛著一幅描繪 1838 年佩斯大洪水的油畫：屋子倒塌，一家人擠在漂浮的木板上，有人伸手向天呼救。那年 3 月，冰塊堵住多瑙河，河水衝破臨時的堤防，成了這座城市歷史上最慘的一次水災，佩斯有兩千多棟房子倒塌。改革派貴族韋賽雷尼・米克洛什男爵划著船在淹水的街道上救人，被稱為「洪水中的船夫」；為他寫詩的，正是我們前一篇提到的詩人沃勒什馬爾蒂。'),
  p('還有一區講戰後的「工廠造的家」：二戰後布達佩斯嚴重缺房子，從 1950 年代到 1980 年代蓋了大量的預鑄國宅社區，畫裡的老平房旁邊就是吊車和一棟棟彩色的高樓。'),
  ...img('hm_hausz', 750, 1000, '紅牆上豪斯曼 1900 年的話，和一幅從山下仰望城堡的油畫（畫名與作者待確認）'),
  ...img('hm_map', 750, 1000, '13:47，1684 年〈布達與佩斯景觀〉版畫：中間的大河寫著「DANUBIUS FLUVIUS」，右下是佩斯，前景山坡上是圍城的軍營'),
  ...img('hm_mimber', 750, 1000, '十六世紀末的清真寺講壇（mimber）欄杆殘片，石灰岩上刻著六芒星和幾何花紋'),
  ...img('hm_flood', 1000, 750, '綠牆上的〈1838 年洪水〉（右）和一幅多瑙河岸的城堡花園集市（左）'),
  ...img('hm_estate', 750, 1000, '「工廠造的家：預鑄國宅社區」展區，畫裡老平房和新高樓並排，作者簽名 V. Barta Éva'),
  verse('他從高天伸手抓住我，把我從大水中拉上來。', '詩篇 18:16'),
  heart('一座城市被圍過、淹過、炸過，又一次次站起來。看著那幅洪水裡伸手的人，會想起自己生命裡那些被拉一把的時刻。'),

  h1('四、走進地底：中世紀王宮與哥德式禮拜堂'),
  p('博物館最特別的是地下兩層。二戰後皇宮被炸毀，考古學家趁重建時往下挖，1948 年起陸續挖出被埋了幾百年的中世紀王宮，1952 年後依照格羅・拉斯洛（Gerő László）的設計重建成展區。這些房間在 1686 年大圍城後被毀，之後就一直被埋在巴洛克宮殿下面。'),
  p('長長的石階一路往下，兩邊是粗糙的老石牆，牆上掛著王宮的老照片和考古圖；展間裡有一根帶著葉飾柱頭的石柱、櫥窗裡的出土石件，以及一尊油漆斑駁、露出木頭的巴洛克天使木雕（名稱待確認）。太太戴著白色毛帽，一級一級往上爬，另一對情侶手牽手走在前面。'),
  p('最深處是哥德式的王宮禮拜堂（Várkápolna）。它和旁邊的哥德大廳一樣，是盧森堡王朝的西吉斯蒙德國王在 15 世紀初蓋的，原本是上層王宮教堂下面的「下層禮拜堂」。1686 年圍城後它被埋在巴洛克的露台底下好幾個世紀，1963 年重建後成為博物館的一部分。肋拱從四面匯到頂點，三扇彩色玻璃窗透進微光，祭壇上刻著拉丁文「VINCULUM CARITATIS」，意思是「愛的聯繫」。我們在這裡拿出門票拍了一張紀念照。'),
  ...img('hm_stairs', 750, 1000, '13:50，往地底的長石階，盡頭是掛著老照片的展間'),
  ...img('hm_stairs_wife', 750, 1000, '太太一級一級往上走，兩旁是中世紀的石牆和木扶手'),
  ...img('hm_corridor', 750, 1000, '13:47，太太在地下展廳，前方的長樓梯上有一對情侶手牽手走上去'),
  ...img('hm_column', 750, 1000, '中世紀王宮的石柱，柱頭兩側雕著葉飾與人臉（待確認）；左牆是馬加什國王時期的考古說明'),
  ...img('hm_angel', 750, 1000, '13:05，油漆斑駁、露出木頭的天使木雕，背後是一幅教士肖像（名稱待確認）'),
  ...img('hm_chapel', 750, 1000, '13:57，十五世紀的哥德式王宮禮拜堂：肋拱、三扇彩色玻璃窗、金色十字架，祭壇刻著「VINCULUM CARITATIS」'),
  ...img('hm_ticket', 1000, 750, '13:58，在禮拜堂裡拿著聖伊什特萬廳的門票：11 月 26 日週四 12:30 入場、附平板、成人聯票 5,900 福林'),
  verse('在這一切之外，要存著愛心，愛心就是聯絡全德的。', '歌羅西書 3:14'),
  heart('祭壇上刻著「愛的聯繫」。王宮被毀、被埋、被遺忘，可是這一小間禮拜堂留了下來。也許一切華麗都會過去，留下來的，是把人連在一起的愛。'),

  h1('旅人小記'),
  bullet('時間（依檔名）：11 月 26 日 12:19 入館 → 12:30 聖伊什特萬廳（12:33–12:42 拍照）→ 13:05 天使木雕 → 13:47–13:50 地圖與地下石階 → 13:57–13:58 哥德式禮拜堂。出來後 14:31 到圖魯鳥與鏈橋（見〈布達城堡〉那篇）。'),
  bullet('門票：成人聯票 5,900 福林，含聖伊什特萬廳附平板導覽（指定入場時間）。地址 Szent György tér 2（E 棟），入口在獅子庭院。'),
  bullet('小提醒：地下樓層樓梯又長又陡，地面石板不平；裡面不冷，外套可以寄放（是否有寄物處，待確認）。'),
  bullet('待確認：窗外的庭院是不是獅子庭院；壁爐上的胸像是不是聖伊什特萬；聖王人像是畫還是陶板；紅牆那幅油畫與天使木雕、石柱的名稱；綠牆左邊那幅小畫的畫名；平板導覽是什麼語言；有沒有看到伯尼尼的聖伊什特萬胸像。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：welovebudapest（聖伊什特萬廳重建、2021/8/20 開放、1838 洪水）、jegy.hu（聖伊什特萬廳平板導覽）、budavar.btk.mta.hu（E 棟分配、1967 搬入）、Wikipedia（中世紀王宮、王宮禮拜堂、1541／1686 圍城、韋賽雷尼）、Hungaricana《Budapest. The History of a Capital》、Hungarian Conservative（1838 洪水）、xpatloop、History Hit（博物館概況）；展品說明牌。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・布達佩斯・歷史博物館　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('布達佩斯_歷史博物館.docx', b));
