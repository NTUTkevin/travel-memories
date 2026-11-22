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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '拉姆紹教堂：霧裡的洋蔥頂', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '德國・巴伐利亞・拉姆紹 Ramsau bei Berchtesgaden｜2026 年 11 月 19 日', font: F, size: 21, color: GRAY })] }),

  p('拉姆紹（Ramsau，也常被寫成蘭紹、藍紹）的聖塞巴斯蒂安教堂，是巴伐利亞最有名的教堂之一。那天下午，山全都藏在雲裡，只剩下一座白牆、黑色洋蔥頂的小教堂，坐在溪邊的小坡上。沒有藍天，反而更像一幅老畫。'),

  h1('一、搭公車到山谷裡的小村'),
  p('拉姆紹是貝希特斯加登旁邊的一個山谷小村，地名最早出現在 1295 年的文獻裡，現在是德國第一個獲得「登山者村」（Bergsteigerdorf）稱號的社區。官方資料說，從貝希特斯加登可以搭 845 或 846 線公車到教堂，我們這一趟就是搭公車上山谷的（路線待確認）。'),
  ...img('r_bus', 1100, 825, '公車上的自拍：我和太太，窗外是飛快掠過的樹林與紅褐色的落葉山坡'),
  ...img('r_house', 1100, 825, '進村路邊的房子：白牆、深紅木頭陽台，後面是被雲壓低的杉林（牆上似乎寫著 Haus Cäcilia，待確認）'),
  verse('我要向山舉目；我的幫助從何而來？', '詩篇 121:1'),
  heart('山今天不給我們看。可是我們還是抬頭看了，也知道它們在雲後面。有些事情看不到，不代表它不在，這好像是這一趟旅行一直在教我們的。'),

  h1('二、1512 年的教堂與洋蔥頂'),
  p('教堂從 1512 年開始蓋，是貝希特斯加登侯爵修道院的院長格雷戈爾・萊納（Gregor Rainer）任內蓋的，獻給聖塞巴斯蒂安和聖法比安。它蓋在一條老商路上：從貝希特斯加登往奧地利方向、翻過希爾施比赫爾山口（Hirschbichlpass）的路，是運鹽、運酒的馱獸隊伍經過的地方。教堂旁邊的 Oberwirt 餐館，前身就是修道院在 1500 年左右為運鹽人和車夫蓋的落腳處。鹽的故事，在薩爾斯堡的要塞和國王湖都出現過，沒想到在這個山谷小村又遇到了。'),
  p('最早的教堂是比較窄的晚期哥德式，後來一再擴建，17 世紀末改成巴洛克風格（資料有 1692 年和 1697 到 1699 年兩種說法），1700 年前後蓋了現在這座有洋蔥頂的塔，屋頂和塔都鋪著木瓦。拉姆紹在 1812 年才成為獨立的教區，這座教堂這時才升格為教區教堂。在這之前，它大約有一百五十年是由貝希特斯加登派人照顧的。'),
  p('教堂裡面有一排基督與十二使徒的彩繪木雕，年代約 1420 到 1430 年，比教堂本身還老；高壇上的早期巴洛克祭壇畫，畫的是聖塞巴斯蒂安殉道。我們這次沒有拍到教堂內部，有沒有進去看，之後再補。'),
  ...img('r_road', 1100, 825, '沿著上坡的馬路走向教堂：灰黑色的木瓦屋頂、白牆、洋蔥頂的塔，南牆上有一個圓形的日晷'),
  ...img('r_wall', 1100, 825, '馬路邊的白色長牆，應該是墓園的外牆；塔尖在霧裡只剩下深色的剪影，右邊是長滿草的山坡'),
  verse('若不是耶和華建造房屋，建造的人就枉然勞力。', '詩篇 127:1'),
  heart('五百多年前，有人在運鹽的路旁蓋了一間教堂，旁邊就是給路人休息的客棧。大概在那個時候，他們就已經知道，趕路的人也需要一個可以安靜的地方。'),

  h1('三、小溪邊的畫家角度'),
  p('教堂前的小溪是拉姆紹阿河（Ramsauer Ache），往上游走可以走到辛特湖（Hintersee）。站在溪邊看過去，溪水、木橋、教堂和遠處的山連成一幅畫，這個角度被叫做「畫家角」（Malerwinkel）：前景是拉姆紹阿河和 Ertlsteg 人行木橋，後面本來該是 Reiter Alm 山群。'),
  p('十九世紀有很多風景畫家來拉姆紹寫生，村子現在還有一條「畫家環形步道」，把畫家當年畫畫的地點串起來，教堂是其中的第 8 站；官方的小冊子上列著 Wilhelm Bendz 畫的《拉姆紹的教堂》和 Thomas Fearnley 1830 年畫的《拉姆紹》。我們那天看不到 Reiter Alm，也不確定自己站的位置是不是畫家角本身，這兩件事都待確認。'),
  ...img('r_chA', 1100, 825, '沿著小溪往教堂走：溪水撞著石頭，岸邊是紅褐色的灌木，教堂的塔和黑屋頂從樹叢後面探出來'),
  ...img('r_chB', 1100, 825, '再靠近一點：白牆、黑屋頂、洋蔥頂，前面是木欄杆的小路，旁邊的山坡被雲蓋住了一半'),
  ...img('r_chC', 824, 1100, '直拍：教堂坐在溪邊的小坡上，塔上的鐘面與窗子很清楚'),
  ...img('r_wife', 824, 1100, '太太撐著黑色的傘站在木欄杆旁，背後就是教堂；那天一直下著小雨'),
  ...img('r_brP', 824, 1100, '沿著溪往遠處看：霧從山谷往上爬，遠處有另一座小木橋'),
  ...img('r_brL', 1100, 824, '同一個地方橫拍：右邊是老樹和黃色的灌木，溪水又清又急'),
  verse('他使我躺臥在青草地上，領我在可安歇的水邊。', '詩篇 23:2'),
  heart('水一直流，雨一直下，教堂就在那裡。我們撐著傘慢慢走，想起小時候不趕時間的下午，那種感覺好久沒有了。'),

  h1('四、墓園與兩個名字'),
  p('教堂四周是老墓園，從 1658 年就開始使用，在教堂的東邊。從鐵花格門進去，一條碎石小路一直通到教堂門口，兩邊是小屋頂的木十字架、鑄鐵十字架和有照片的墓碑。門樓的山牆上寫著 1897，下面有一幅壁畫，看起來是耶穌復活的畫面（待確認）；外牆上掛著幾個花圈，繫著藍白（巴伐利亞）和黑紅金（德國）的緞帶，緞帶上的字我們沒有全部看清楚，是哪一種紀念，待確認。'),
  p('這個墓園裡，有兩個名字特別有名。一個是約翰・格里爾（Johann Grill，外號 Kederbacher，1835 到 1917 年），被稱為德國第一位山地嚮導，1881 年 6 月 6 日他帶著奧托・許克（Otto Schück）第一次爬上瓦茨曼東壁，就是前一篇在聖巴多羅買教堂後面看到的那面大岩壁，他就葬在這裡。另一個是約瑟夫・莫爾（Josef Mohr），《平安夜》的作詞人，薩爾斯堡人，1815 年 8 月 29 日來到拉姆紹，這是他當上神父後的第一個工作，後來他和奧伯恩多夫的教堂風琴手合作寫了《平安夜》。他自己後來葬在奧地利的瓦格賴因，不是這裡，但這座教堂是他的第一站。'),
  ...img('r_gate', 825, 1100, '墓園的門樓：山牆上寫著 1897，下面是一幅小壁畫，石頭拱門後面是鐵花格門和墓園；右邊的外牆掛著花圈'),
  ...img('r_path', 825, 1100, '從門裡看進去：碎石小路通向白色的教堂，兩邊都是十字架和小小的花園，山谷的霧就在後面'),
  verse('耶穌對她說：復活在我，生命也在我；信我的人，雖然死了，也必復活。', '約翰福音 11:25'),
  heart('這裡的墓一座一座都收拾得很整齊，有花、有燈、有照片。看著它們，我們想到的不是離開，而是有人一直在照顧、一直在記得。'),

  h1('五、牆上的日晷'),
  p('教堂南牆上有一個畫出來的日晷：中間是一張有臉的太陽，下面是兩圈半圓形的刻度，一圈是羅馬數字，一圈是阿拉伯數字，靠影子報時。我們到的時候沒有太陽，所以也沒有影子，只剩下牆上畫著的那一圈圈數字。日晷旁邊的窗邊還有一塊小浮雕牌，內容待確認。'),
  ...img('r_sun', 826, 1100, '南牆上的日晷：太陽臉下面是兩圈半圓形的刻度；牆下的石頭小路上，太太撐著傘抬頭看'),
  verse('求你指教我們怎樣數算自己的日子，好叫我們得著智慧的心。', '詩篇 90:12'),
  heart('沒有太陽的日子，日晷不能報時。可是我們在那裡站了好一會兒，也沒有一個人看手錶。'),

  h1('旅人小記'),
  bullet('交通：官方資料寫的是 RVO 公車 845 或 846 線，到教堂附近下車；班次請以當天的公車時刻表為準。'),
  bullet('教堂與墓園：教堂的資料寫著「隨時可以進入」，但墓園是現在還有人在祭拜的地方，請放輕聲音，拍照也避開別人家的墓碑和前來祭拜的人。'),
  bullet('附近還有昆特韋格朝聖教堂（Kunterwegkirche）和十字架之路，從教堂旁邊的路上去；這次我們沒有走，留給下一次。'),
  bullet('這一篇還有幾處待確認：搭的公車路線、門樓壁畫的內容、花圈緞帶上的字、日晷旁的小浮雕、溪邊的位置是不是畫家角、教堂內部有沒有進去。'),
  ...img('r_tree', 824, 1100, '附記：11 月 20 日拍的紅葉樹，遠處有積雪的山，左邊是白色的房子與鐵路平交道，不確定是在哪裡拍的（待確認）'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：berchtesgaden.de（Kirche Ramsau、Kunterweg-Kirche）、ramsau.de（Pfarrkirche、Maler-Rundwanderweg 小冊子、Erlebnisweg 小冊子、Silent Night 頁面）、de.wikipedia〈St. Sebastian (Ramsau)〉、waymarking.com（Sundial on St Sebastian Church）、Historical Marker Database（Pfarrkirche St. Sebastian）、Deutsche Biographie（Johann Grill）、Alpenverein 圖書館（Kederbacher）、Find a Grave（Josef Mohr）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・拉姆紹教堂　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('拉姆紹教堂.docx', b));
