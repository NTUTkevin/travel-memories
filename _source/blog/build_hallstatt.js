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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '哈修塔特鹽礦：走進七千年的鹽山', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・上奧地利邦・哈修塔特 Salzwelten Hallstatt｜2026 年 11 月 20 日', font: F, size: 21, color: GRAY })] }),

  p('前一天在國王湖和貝希特斯加登一路聽到「鹽」這個字，這一天我們終於走進鹽的源頭：哈修塔特（Hallstatt）湖邊小鎮上方的鹽山。考古學家認為，人類在這座山裡取鹽已經大約七千年，這裡被稱為世界上最古老的鹽礦。我們搭新的纜車上山，穿上白色的礦工外套，跟著導覽走進山肚子裡，最後坐礦車回到地面，看見積雪的山頭和雲霧。'),

  h1('一、搭纜車上鹽山'),
  p('鹽礦所在的「高谷」（Hochtal）在小鎮上方的山腰，要搭登山纜車（Salzbergbahn）上去。舊的纜車從 1979 年開始營運，2025 年 9 月拆除，換成一條全新的軌道纜車；建造期間鹽礦整個關閉，到 2026 年 8 月底才重新開放。也就是說，我們搭到的是才通車不到三個月的新纜車。'),
  p('根據報導，新纜車長不到 600 公尺，從海拔 515 公尺的山下站爬到 840 公尺的山上站，高低差 325 公尺，最陡的地方坡度接近 88%，全程大約 2.6 分鐘；兩個車廂各可載 60 人，車廂會自動保持水平。整個更新工程大約花了 3,400 萬歐元，由奧地利的 Doppelmayr 集團承建。山上站還新開了一間叫「Weitblick 7000」（遠眺七千）的小餐館與觀景台。'),
  ...img('h_funi', 825, 1100, '從纜車山上站往下看：右下角是新纜車的車廂和鋼架軌道，山腳下是哈修塔特的房子，對面山坡上的落葉松已經轉成金黃色'),
  verse('你們是世上的鹽。鹽若失了味，怎能叫他再鹹呢？', '馬太福音 5:13'),
  heart('兩分多鐘，就從湖邊升到半山腰。以前的礦工每天得自己走上來，我們卻坐著看風景。往下看那一片黃色的落葉松，才知道秋天原來是從山上開始的。'),

  h1('二、高谷、魯道夫塔與「哈修塔特文化」'),
  p('鹽讓這個偏僻的山谷變得富有，也讓它變成被爭奪的地方。1282 到 1284 年間，哈布斯堡家族的阿爾布雷希特一世公爵在鹽山上蓋了一座防禦塔，用父親魯道夫的名字命名為「魯道夫塔」（Rudolfsturm），用來保護鹽礦。幾年後，哈布斯堡家族和薩爾斯堡大主教為了鹽打了一場「鹽戰爭」（1291–1297），最後哈布斯堡勝出。1311 年，阿爾布雷希特的遺孀伊莉莎白王后給了哈修塔特市集的權利，這份文件也是第一次提到中世紀的鹽礦。這座塔一直到 1954 年都是礦場主管的住所，今天是一間餐廳。'),
  p('從纜車站走到礦坑入口的路上，會經過一片史前墓地。1846 年，礦場主管約翰・格奧爾格・拉姆紹爾（Johann Georg Ramsauer）發現了這片墓地，一直挖到 1863 年，挖出將近一千座墳墓，陪葬品大多送到了維也納。因為他留下非常仔細的紀錄，歐洲早期鐵器時代（大約西元前 800 到 450 年）後來就以這個小鎮命名，叫做「哈修塔特文化」——整個時代用一個小村的名字，世界上很少見。'),
  ...img('h_path', 1100, 825, '高谷裡的小路：兩旁是礦場的老房子，小路彎彎曲曲往下，遠處山坡上有一棟白色、尖屋頂的建築（是否為魯道夫塔，待確認），後面是積雪的山'),
  ...img('h_mist', 1100, 825, '從高谷往湖的方向看：山谷裡升起一團白霧，霧前面隱約看得到一座鋼構的橋（可能是天空步道一帶，待確認），對面的山頂已經下雪'),
  verse('我要向山舉目；我的幫助從何而來？我的幫助從造天地的耶和華而來。', '詩篇 121:1–2'),
  heart('一座塔、一場戰爭、一片墓地，全都是為了鹽。今天我們只花一張門票就走過了這些地方。歷史裡有很多人辛苦、流血，我們能平安地走在這裡，是一份恩典。'),

  h1('三、走進克莉絲汀娜坑道'),
  p('導覽從礦工以前住的「礦工之家」（Knappenhaus）出發，換上白色的工作外套後，從一個石頭拱門走進山裡。拱門上方的牌子寫著「Kaiserin Christina Stollen eröffnet im Jahre 1719」——「克莉絲汀娜皇后坑道，1719 年開鑿」。一般認為「克莉絲汀娜」是指當時的神聖羅馬帝國皇后伊莉莎白・克莉絲汀（瑪麗亞・特蕾莎的母親），不過坑道命名的由來我們沒有找到正式說明，待確認。'),
  p('這座山裡藏著很多考古寶物。1734 年，礦工在山裡挖到一具被鹽保存下來的史前礦工遺體，衣服、皮膚和頭髮都還看得出來，被稱為「鹽中人」（Mann im Salz），推測是在礦坑意外中喪生的。2002 年，考古人員又在山裡發現一座木頭樓梯，經過科學鑑定，大約建於西元前 1344 年，是歐洲已知最古老的樓梯。現在它被放在地下大約 400 公尺的「青銅時代電影院」裡，用投影動畫重現史前礦工的工作。'),
  p('這裡的歷史還可以往更早推：1838 年在山裡找到一把鹿角做的鋤頭，經過放射性碳定年，大約有七千年；有系統的地下採礦則從青銅時代開始。維也納自然史博物館在這裡做研究已經一百多年，夏天還能看到考古學家在高谷裡挖掘。'),
  ...img('h_christina', 618, 1100, '影片截圖：坑道入口的石拱門，上方的牌子寫著「Kaiserin Christina Stollen eröffnet im Jahre 1719」，前面的遊客穿著 Salzwelten 的白色外套魚貫走進去'),
  verse('人為黑暗定界限，查究幽暗陰翳的石頭，直到極處。', '約伯記 28:3'),
  heart('一座三千多年前的木頭樓梯，到今天還留在山裡。當年踩著它上上下下的人，不會知道有一天會有人特地從台灣來看它。平凡的勞動，原來也可以被記得這麼久。'),

  h1('四、鹽的顏色與味道'),
  p('哈修塔特的鹽，是大約兩億五千萬年前海水蒸發後留下來的。後來阿爾卑斯山隆起，岩層被擠壓、摺疊，鹽和黏土、岩石混在一起，形成一種當地叫做「哈瑟爾山岩」（Haselgebirge）的混合岩層，所以礦坑的牆壁不是純白色，而是一層一層灰、黑、橘紅交錯的紋路。'),
  p('坑道裡有一面牆展示世界各地的鹽：挪威的海鹽片、丹麥的煙燻鹽、伊朗的藍鹽、納米比亞的珍珠鹽、夏威夷的黑色熔岩鹽，還有奧地利本地的「奧塞山岩鹽」（Ausseer Bergkernsalz）。另一處把大塊的岩鹽做成燈，從裡面打光，橘紅色、白色的鹽塊像一叢發光的花。我們拍到的一段影片裡，太太在坑道裡把手湊到嘴邊，像是在嚐鹽的味道（是不是在嚐鹽，待確認）。'),
  ...img('h_salts', 825, 1100, '世界各地的鹽：六個小展示櫃，每一格一座小鹽堆，標著德文和英文名字——挪威海鹽片、丹麥煙燻鹽、伊朗藍鹽、納米比亞珍珠鹽、夏威夷黑熔岩鹽、奧地利奧塞山岩鹽'),
  ...img('h_wall', 825, 1100, '太太拿著手機拍礦坑的岩壁：燈光下可以看到灰、黑、橘紅色一圈一圈的紋路'),
  ...img('h_ceil', 1100, 825, '抬頭看坑道頂：被鑿平的岩層表面還留著一道道工具痕跡，下方是紅色的木欄杆'),
  ...img('h_lamps', 825, 1100, '岩鹽燈：一根根鐵柱撐著大塊的岩鹽，從裡面打光，發出橘紅色和白色的光（這些鹽是否產自哈修塔特，待確認）'),
  ...img('h_selfie', 1100, 825, '在鹽燈旁自拍：我、太太，後面張大嘴巴的是宜恩（待確認）'),
  verse('你們的言語要常常帶著和氣，好像用鹽調和，就可知道該怎樣回答各人。', '歌羅西書 4:6'),
  heart('原來鹽不是只有白色。灰的、黑的、橘紅的，被壓在山裡兩億多年，最後變成餐桌上的一小撮。在山肚子裡摸到、看到真正的岩鹽，才覺得「鹽」這個每天都用的東西，原來有這麼長的故事。'),

  h1('五、礦工滑梯、地下鹽湖與礦車'),
  p('鹽礦導覽最有名的是礦工滑梯：以前礦工用木頭做的長滑道，從一層坑道快速滑到下一層。這座山的坑道分成好幾層，礦工叫做「水平層」（Horizont）。導覽中還會經過地下的鹽湖。現在的鹽不是一塊一塊挖出來的，而是把山外的清水引進礦坑，讓水溶解岩鹽，雜質沉在底下，再把鹹水（鹵水，Sole）抽出來。'),
  p('這些鹵水要送到很遠的地方去煮成鹽。1595 年，皇帝魯道夫二世下令，因為附近的森林快砍完了，與其把木頭運來，不如把鹵水送到森林去。於是工人用大約一萬三千根挖空的樹幹接成管線，經過巴特戈伊瑟恩、巴特伊舍，一路通到埃本塞（Ebensee），1607 年完工，全長三十多公里。這條管線現在換成了現代的管子，還在使用，被稱為世界上最古老的工業管線。'),
  p('導覽的最後，大家坐上窄窄的礦車，一路在黑暗的坑道裡往外衝，直到看見出口的光。'),
  verse('我是世界的光。跟從我的，就不在黑暗裡走，必要得著生命的光。', '約翰福音 8:12'),
  heart('礦車在黑暗裡跑的時候，太太一直笑。坐在最後面，看著前面的人一個一個衝進光裡，好像我們也一起從很深的地方被帶了出來。'),

  h1('六、回到地面：世界遺產的視野'),
  p('出了礦坑，雲正好散開一點，對面的山頂覆著薄雪，山谷裡一團霧慢慢往上飄。哈修塔特和旁邊的達赫施泰因山、鹽湖區在 1997 年一起被列入聯合國世界遺產，名稱是「哈修塔特－達赫施泰因／薩爾茲卡默古特文化景觀」。高谷裡的天空步道「世界遺產視野」（Welterbeblick）就懸在小鎮上方大約 360 公尺，可以俯瞰整個湖和小鎮。'),
  ...img('h_couple', 825, 1100, '出礦坑後的合照：我和太太都還穿著 Salzwelten 的白外套，背後是積雪的山頭、山谷裡的白霧，和秋天紅褐色的樹林'),
  verse('耶和華啊，你所造的何其多！都是你用智慧造成的；遍地滿了你的豐富。', '詩篇 104:24'),
  heart('在山裡待了一個多小時，出來第一眼就是這片山。地底下兩億年的鹽，地面上剛下的雪，都是同一位創造主的作品。'),

  h1('旅人小記'),
  bullet('交通：從哈修塔特小鎮走到登山纜車山下站，搭纜車約 3 分鐘上到高谷；從山上站步行約 15 到 20 分鐘，經過墓地，到礦工之家集合。'),
  bullet('導覽：地下部分約 90 分鐘，整趟行程建議預留 3 小時；4 歲以下兒童不能參加。礦坑裡溫度低，要穿保暖的衣服和好走的鞋。'),
  bullet('2026 年 8 月重新開放後，從小鎮爬上高谷的「之字形步道」因為工程還沒開放，上山請搭纜車（以官方最新公告為準）。'),
  bullet('想看更多考古內容，可以報名「史前探險」（Prähistorische Expedition），會深入山裡到「鹽中人」被發現的地方。'),
  bullet('這一篇還有幾處待確認：克莉絲汀娜坑道名稱的由來、照片中的白色尖頂建築是不是魯道夫塔、白霧前的鋼構橋是不是天空步道、鹽燈的鹽產地、鹽燈自拍裡的年輕人是不是宜恩；我們有沒有走上天空步道、有沒有坐滑梯、有沒有看到地下鹽湖，也請補上。'),
  ...img('h_wall2', 825, 1100, '附記：同一面岩壁換個角度再拍一張，燈光下的紋路像木頭的年輪'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：salzwelten.at（Hallstatt Archaeology、Hallstatt 25 Years World Heritage、新聞資料、鹵水管線步道）、維也納自然史博物館 nhm-wien.ac.at（Hallstatt 研究網站：青銅時代採礦、石器時代、Ramsauer、1311 年至今）、hallstatt.net（Rudolfsturm、哈修塔特與哈布斯堡、鹵水管線）、dachstein.salzkammergut.at／salzkammergut.at（鹽礦、天空步道、鹵水管線步道）、eap-magazin.de（New Salzbergbahn, 2026/9/3）、ic-group.org（新纜車工程）、showcaves.com（Salzwelten Hallstatt 年表）、Austria-Forum（Soleleitung）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・哈修塔特鹽礦　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('哈修塔特鹽礦.docx', b));
