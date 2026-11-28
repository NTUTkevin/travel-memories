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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '蓋勒特山岩洞教堂：被封了四十年的山洞', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '匈牙利・布達佩斯・Gellérthegyi Sziklatemplom｜2026 年 11 月 27 日', font: F, size: 21, color: GRAY })] }),

  p('11 月 27 日下午，在中央市場吃完燉牛肉湯，我們（應該是走過自由橋，待確認）來到多瑙河對岸的布達。橋頭就是蓋勒特山（Gellért-hegy），山腳下的岩壁裡藏著一座教堂——匈牙利人叫它「Sziklatemplom」，岩石教堂。我們下午三點多在山坡上看了一會兒自由橋，三點半以後走進山洞。'),

  h1('一、從自由橋到蓋勒特山'),
  p('我們中午才在橋的另一頭逛中央市場；橋和市場是同一個時代的產物。自由橋 1894 到 1896 年依照費凱特哈齊・亞諾什（Feketeházy János）的設計建造，是為了千禧年博覽會蓋的；1896 年 10 月 4 日通車，皇帝法蘭茲・約瑟夫親自把最後一顆銀鉚釘敲進佩斯那一端的橋墩，橋也以他命名，叫「法蘭茲・約瑟夫橋」。全長 333.6 公尺、寬 20.1 公尺，是布達佩斯市中心最短的橋；四根橋塔頂端各站著一隻青銅的圖魯鳥（turul）——就是我們前一天在布達城堡看到的那種匈牙利神話之鳥。'),
  p('1945 年 1 月，撤退的德軍把橋炸斷。因為橋墩夠堅固，戰後修復得很快，1946 年 8 月 20 日——匈牙利最重要的國定假日聖伊什特萬日——照原樣重新通車，改名「自由橋」。據說那時灰色油漆不夠，才漆成現在的綠色。'),
  p('山的名字來自聖蓋勒特（Szent Gellért）：他是威尼斯出身的主教，幫助匈牙利第一位國王伊什特萬傳教，1046 年在異教徒的反抗中，被人從這座山上推下去殉道。'),
  ...img('cave_view', 787, 1050, '15:19，從蓋勒特山腳的高處往南看：多瑙河岸邊的大街上，兩節橘白色的電車正要開走；遠處是另一座跨河的鐵橋（推測是裴多菲橋），左岸有一棟玻璃波浪屋頂的建築，右邊是布達這一側的老建築（待確認）'),
  ...img('cave_bridge', 787, 1050, '15:19，綠色的自由橋：四根橋塔頂端站著展翅的圖魯鳥，中間是有皇冠的匈牙利國徽；橋下有遊船駛過，前景是一輛橘白色的老電車，對岸就是中央市場所在的佩斯'),
  verse('有一道河，這河的分汊使神的城歡喜；這城就是至高者居所的聖處。', '詩篇 46:4'),
  heart('中午在橋的那一頭吃飯，下午站在橋的這一頭看它。一座被炸斷又照原樣接回來的橋，綠綠地橫在灰色的天空下，好像在說：斷掉的東西，也可以重新連起來。'),

  h1('二、隱士伊凡的山洞與露德的靈感'),
  p('這個山洞是蓋勒特山裡溫泉侵蝕出來的洞穴系統的一部分，也叫「聖伊凡洞」（Szent Iván-barlang）。傳說中世紀初有一位叫伊凡的隱士住在這裡，用洞口旁湧出來的溫泉為病人治病；山腳的「泥巴浴場」，就是今天蓋勒特溫泉浴場的前身。'),
  p('1920 年代，一群匈牙利朝聖者從法國露德（Lourdes）的聖母洞回來，想在布達佩斯也蓋一座類似的山洞聖堂。1925 年開工，由布達佩斯工業大學教授盧克斯・卡爾曼（Lux Kálmán）設計，用炸藥把天然的洞穴擴大；1926 年 5 月 23 日（聖神降臨節）祝聖。一開始祭壇就在洞口，信徒是站在外面的平台上望彌撒。後來繼續往裡面擴建，1931 年聖神降臨節，下層的小聖堂開放；1934 年，山坡上新羅馬式的修道院落成，1936 年洞口上方豎起一座會發光的大木十字架。'),
  p('管理這座教堂的是保祿會（Pálos，聖保祿首位隱修士會）。它是唯一由匈牙利人創立、至今仍存在的修會：1250 年，艾斯特根的教士真福奧澤伯（Boldog Özséb）放下職位到皮利什山隱居，把附近的隱修士聯合起來。修會鼎盛時在匈牙利有一百多座會院，但 1786 年被皇帝約瑟夫二世下令解散，在匈牙利中斷了將近一百五十年，只在波蘭保留下來——所以 1934 年，白袍的保祿會士是得到波蘭琴斯托霍瓦（Częstochowa）總會的同意，才回到蓋勒特山的。2017 年，保祿會被列為匈牙利國寶（Hungarikum）。'),
  ...img('cave_gate', 787, 1050, '15:37，從主堂往下看：一道麥穗與火焰圖案的黑色鐵門後面，是一間很小的岩洞小堂，盡頭的牆上畫著希臘字母 Χ 與 Ρ 疊在一起的基督符號，兩旁是 A 和 Ω，像早期教會地下墓窟的風格（是哪一間小堂，待確認）；左邊立著說明牌'),
  verse('你們得救在乎歸回安息；你們得力在乎平靜安穩。', '以賽亞書 30:15'),
  heart('從隱士伊凡，到七百多年前在山裡隱居的奧澤伯，再到今天的保祿會士，好像都在找一個安靜的山洞，好好地跟神說話。外面是車水馬龍的大橋，一走進來，聲音就全部不見了。'),

  h1('三、混凝土牆後的四十年'),
  p('二戰期間，這個山洞曾經當作戰地醫院和避難所，也收容過波蘭難民。1945 年以後匈牙利落入共產黨統治，保祿會士還繼續服事了幾年；但在 1951 年（復活節星期一，各家記載的日期略有不同，待確認），祕密警察闖進教堂，以叛國罪逮捕了整個修會。院長韋澤爾・費倫茨（Vezér Ferenc）被判死刑，其他修士被判五到十年以上的徒刑，洞口被一道約 2.25 公尺厚的混凝土牆封死，連山上的大十字架也被拆掉。'),
  p('就這樣，教堂在牆後沉默了將近四十年。1989 年鐵幕瓦解，同年 8 月 27 日教堂重新開放，保祿會士回來了；那道混凝土牆一直到 1992 年才完全拆除。'),
  p('主堂是一個天然的岩洞，粗糙的岩壁直接當作教堂的牆和天花板，正中央掛著一個大十字架，前面是鋪著蕾絲桌巾的祭台，兩側各有一道拱門通到旁邊的小堂。我們去的時候很安靜，只有一位長者在旁邊的小堂裡點蠟燭、整理東西。'),
  ...img('cave_altar', 787, 1050, '15:46，主堂：粗糙的岩壁下，十字架上的耶穌頭頂有金色光芒，十字架上方的小牌寫著 INRI；下面是鋪著蕾絲布的祭台和綠色的圓球裝飾，兩邊各有一道拱門和黑色鐵門，右邊的小堂裡有一位長者正彎腰點蠟燭'),
  ...img('cave_wife', 787, 1050, '15:51，太太戴著黑色報童帽、圍著中午新買的白圍巾，站在一個小堂裡，看著岩壁上一尊白色聖像；左邊牆上嵌著一塊深色的紀念銘牌，旁邊是半圓形的跪凳和鐵欄杆（銘牌內容與聖像是誰，待確認）'),
  verse('我們四面受敵，卻不被困住；心裡作難，卻不至失望；遭逼迫，卻不被丟棄；打倒了，卻不至死亡。', '哥林多後書 4:8–9'),
  heart('一道兩公尺多厚的牆，可以把一個山洞封住四十年，卻封不住人心裡的信仰。站在重新打開的洞裡，想到那位被判死刑的院長，覺得能自由地走進教堂、安靜地禱告，是一件不能視為理所當然的事。'),

  h1('四、岩洞裡的小角落'),
  p('教堂由好幾個大小不一的岩洞連在一起，一間接一間走，像繞一個小圈。因為保祿會和波蘭的淵源，這裡有一幅琴斯托霍瓦「黑面聖母」的複製像和波蘭的老鷹徽章；也有一幅聖國柏（Maximilian Kolbe）的畫像——他是在奧斯威辛集中營自願代替另一名囚犯赴死的波蘭神父。我們拍到一個岩壁上的白色大理石小龕，裡面站著一尊手按胸前、露出心的耶穌像（耶穌聖心像），下面擺著兩盆綠色植物。'),
  p('另外有一間白牆的房間，擺滿了整套木雕：祭台、讀經台、燭台、圍欄柱和牆上的大浮雕，雕滿了葡萄、麥穗、花朵和聖經故事的人物。資料說，教堂裡有一些外西凡尼亞（川西凡尼亞）匈牙利民俗風格的木雕，其中包括紀念塞凱伊人（Székely，住在今天羅馬尼亞境內的匈牙利族群）的作品；這間房間是不是就是那一組、作者是誰、浮雕講的是哪些故事，待確認。'),
  ...img('cave_heart', 787, 1050, '15:52，岩壁上的耶穌聖心像：象牙白的大理石小龕，兩根凹槽圓柱撐著半圓形的頂，頂上有一個圓窗，裡面是紅色背板；下面兩盆綠色的植物'),
  ...img('cave_wood', 787, 1050, '15:56，木雕房間：整套淺色木雕的祭台和讀經台上雕滿葡萄與麥穗，後面的小祭台中央刻著 IHS，兩旁牆上是兩大片聖經人物浮雕；左邊有向日葵花瓶，右邊立著十字架和高大的復活蠟燭（這是不是塞凱伊民俗木雕的紀念室，待確認）'),
  verse('凡勞苦擔重擔的人可以到我這裡來，我就使你們得安息。', '馬太福音 11:28'),
  heart('從冰冷的岩壁走到一整間溫暖的木頭雕刻，好像看到很多人用自己的手藝，把信仰一刀一刀刻下來。這趟旅行走過好多華麗的大教堂，這個小小的山洞，反而讓人特別想多坐一下。'),

  h1('旅人小記'),
  bullet('時間（依檔名）：11 月 27 日（週五）15:19 在山坡上看自由橋與電車；15:37–15:56 在岩洞教堂裡。'),
  bullet('地址：Budapest XI., Szent Gellért rakpart 1（自由橋布達端、蓋勒特溫泉旅館對面的山坡上）。交通：地鐵 M4 或電車 47、49 號到 Szent Gellért tér；從中央市場可以直接走過自由橋。'),
  bullet('門票與開放時間（第三方旅遊網站 2026 年的資料，以現場為準）：成人約 1,400 福林，學生與長者約 1,200 福林，10 歲以下免費，含語音導覽；週一至週六 9:30–19:30，週日不開放參觀；平日彌撒 8:30、17:00、20:00，彌撒時遊客不能入內。參觀大約 30–45 分鐘。'),
  bullet('小提醒：這是仍在使用的教堂，裡面請保持安靜。'),
  bullet('附記：雲端照片資料夾另有一張主堂十字架（04，和 05 角度幾乎一樣，所以沒放進文章）。'),
  bullet('待確認：是不是從中央市場走過自由橋過來的；門票多少、有沒有用語音導覽；鐵門後那間 Χ‑Ρ 小堂與太太站的小堂是哪一間、白色聖像是誰、銘牌寫什麼；有沒有看到黑面聖母、波蘭老鷹、聖國柏畫像；木雕房間是不是塞凱伊紀念室；之後還去了哪裡。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：Wikipedia（Gellért Hill Cave、Liberty Bridge、Order of Saint Paul the First Hermit、Eusebius of Esztergom）、N. Kósa Judit／Szablyár Péter《Underground Buda》（1925 開工、1926/5/23 祝聖、1931 下層小堂、1934 修道院、1951 封閉）、Show Caves（年表、1936 十字架）、In Your Pocket Budapest（2.25 公尺混凝土牆、1989/8/27、1992）、Atlas Obscura（1951 復活節星期一、黑面聖母）、Aleteia、Xpatloop、We Love Budapest（波蘭難民）、A View On Cities、budapest.net（聖國柏）、The Better Vacation（2026 門票與開放時間、民俗木雕）、pecs.hu、hungarikum.hu、OSZK《Pauline Guide》（保祿會）、Daily News Hungary（自由橋 1896/10/4、1946/8/20、綠色油漆）、Structurae、Ancient Origins（聖蓋勒特 1046）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・布達佩斯・蓋勒特山岩洞教堂　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('布達佩斯_岩洞教堂.docx', b));
