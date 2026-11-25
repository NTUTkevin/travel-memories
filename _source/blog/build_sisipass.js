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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Sisi Pass：一張票，跟著西西走三個地方', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・美泉宮／西西博物館／家具博物館｜2026 年 11 月 23–24 日', font: F, size: 21, color: GRAY })] }),

  p('這次在維也納，我們用一張「Sisi Pass」（西西通行票）走完了三個跟西西有關的地方：美泉宮、霍夫堡裡的西西博物館，還有家具博物館。三個地方都由同一家公司「美泉宮集團」（Schönbrunn Group）經營，所以可以合成一張票。這篇把這張票整理一下，也算是給之後想去的朋友一份小攻略。三個地方的詳細故事，請看前面各自的文章。'),

  h1('一、Sisi Pass 是什麼？'),
  p('官方的說法是「一張票、三個皇家景點」，每個地方可以進去一次，都附語音導覽：'),
  bullet('美泉宮：「Palace Ticket」，約 75 分鐘，走完整個主樓層（bel étage）的房間，包括法蘭茲・約瑟夫和西西的私人房間、正式大廳，和瑪麗亞・特蕾莎的洛可可房間。'),
  bullet('霍夫堡西西博物館：西西博物館＋皇家公寓＋銀器館。'),
  bullet('維也納家具博物館：博物館參觀，加上「電影裡的茜茜」（Sissi im Film）小路。'),
  p('票從開票日起一年內有效。只有美泉宮要預約時段，買票時就要選好參觀的日期和時間；西西博物館和家具博物館不用預約時段，一年內哪天去都可以。六歲以下兒童免費。'),
  ...img('hb_banner', 750, 1000, '霍夫堡西西博物館入口的紫色布條，印的是溫特哈爾特 1865 年的西西肖像'),
  verse('人心籌算自己的道路；惟耶和華指引他的腳步。', '箴言 16:9'),
  heart('出發前只想著要看哪些地方，沒想到一張票就把西西的一生串起來：她住過的夏宮、冬宮，還有後人拍成電影的故事。'),

  h1('二、第一站：美泉宮（11 月 23 日）'),
  p('美泉宮是哈布斯堡家族的夏宮，1854 年西西嫁給法蘭茲・約瑟夫以後，夏天常住在這裡。Sisi Pass 包含的「Palace Ticket」以前叫「Grand Tour」，走完主樓層全部的房間。我們上午先爬到凱旋門（Gloriette）看整座宮殿和花園，中午進宮殿參觀，下午在宮前的聖誕市集遇到了雪。'),
  ...img('sout_couple2', 825, 1100, '我們在美泉宮凱旋門的欄杆邊，背後就是維也納'),
  ...img('sin_billiard', 825, 1100, '太太在美泉宮的撞球室（法蘭茲・約瑟夫的房間）'),
  verse('這是耶和華所定的日子，我們在其中要高興歡喜。', '詩篇 118:24'),
  heart('第一天就遇到下雪，好像整座宮殿都在歡迎我們。'),

  h1('三、第二站：霍夫堡西西博物館（11 月 24 日上午）'),
  p('霍夫堡是皇帝的冬宮和辦公的地方。西西博物館剛好在 2026 年 11 月 19 日整修完重新開幕，我們是開幕後第五天去的。參觀路線是西西博物館、皇家公寓和銀器館連在一起走：皇帝的謁見室與辦公桌、西西的體操室、溫特哈爾特的名畫、她遇刺的故事，最後是 30 公尺長的金色米蘭桌飾。'),
  p('入口在米歇爾圓頂（Michaelerkuppel）下方，每天 9:00–17:30 開放，16:30 最後入場。'),
  ...img('hb_couple', 750, 1000, '我們在霍夫堡米歇爾宮翼前'),
  ...img('hb_winterhalter', 750, 1000, '溫特哈爾特 1865 年的西西肖像'),
  verse('我終身的事在你手中。', '詩篇 31:15'),
  heart('從年輕的新娘到黑衣的皇后，一個早上就走完她的一生。'),

  h1('四、第三站：維也納家具博物館（11 月 24 日下午）'),
  p('家具博物館是 1747 年瑪麗亞・特蕾莎設立的宮廷家具倉庫。這張票在這裡包含的「電影裡的茜茜」小路，展出 1950 年代《茜茜公主》三部曲用過的真家具、電影片段和海報。上午在霍夫堡看真正的西西，下午來這裡看電影裡的「茜茜」，剛好可以對照。'),
  ...img('mb_throne', 1100, 825, '家具博物館裡法蘭茲・約瑟夫的寶座布置'),
  ...img('mb_salon', 1100, 825, '白底描金的洛可可椅子與桌子'),
  verse('願主我們神的榮美歸於我們身上；願你堅立我們手所做的工。', '詩篇 90:17'),
  heart('看見那些手工雕刻的椅子用了兩百多年還在，就覺得用心做的東西，真的會留下來。'),

  h1('五、算一算：划算嗎？'),
  p('官網寫著「新的票價季從 2026 年 11 月 19 日開始」，以下是我們 11 月 25 日在三個官網查到的成人價格（價格會隨季節調整，出發前請再查一次）：'),
  bullet('美泉宮 Palace Ticket：€42.00（6–18 歲 €30.00）'),
  bullet('西西博物館日票：€27.00（6–18 歲 €16.00）'),
  bullet('家具博物館日票：€20.00（6–18 歲 €12.00；19–25 歲學生 €18.00）'),
  bullet('分開買合計：成人 €89.00、兒童 €58.00'),
  bullet('Sisi Pass：成人 €69.00、6–18 歲 €46.00；持維也納城市卡（Vienna City Card）€56.50'),
  p('所以成人大約省 €20（約兩成多），兒童省 €12。如果三個地方都想去，這張票很划算；只想去其中一兩個，就單買比較好。順帶一提，11 月 18 日以前的舊票價季，Sisi Pass 是 €51，新季調漲了不少。'),
  verse('你們哪一個要蓋一座樓，不先坐下算計花費，能蓋成不能呢？', '路加福音 14:28'),
  heart('旅行也要先坐下來算一算。算清楚了，玩起來就比較安心。'),

  h1('旅人小記'),
  bullet('購票：官方網路售票只有 imperialtickets.com，其他網站可能是轉賣。也可以現場買。'),
  bullet('交通：三個地方都在地鐵 U3、U4 線上。美泉宮 → U4 Schönbrunn 站；霍夫堡 → U3 Herrengasse 站；家具博物館 → U3 Zieglergasse 站（入口在 Andreasgasse 7）。霍夫堡和家具博物館都在 U3，同一天去很方便。'),
  bullet('建議順序：美泉宮要預約時段，先排好；另外兩個不用預約，可以像我們一樣上午霍夫堡、下午家具博物館。'),
  bullet('開放時間：家具博物館週一休館（週二到週日 10:00–17:00）；西西博物館每天 9:00–17:30。'),
  bullet('待確認：我們實際買的是 Sisi Pass 嗎？什麼時候買、買了多少錢（是不是在 11/19 新票價之後）？美泉宮那天是不是用這張票的 Palace Ticket 進去的？'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：schoenbrunn.at（Sisi Pass、Palace Ticket 票價與規則）、moebelmuseumwien.at（家具博物館票價與 Sisi Pass）、sisimuseum-hofburg.at（西西博物館票價與開放時間）、imperialtickets.com（官方售票）、wien.info（舊季 Sisi Pass 票價）、visitingvienna.com（地鐵 U3 Zieglergasse）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・Sisi Pass　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_SisiPass三館.docx', b));
