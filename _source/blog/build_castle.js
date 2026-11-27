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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '布達城堡：下雪了，皇宮山丘上的白色早晨', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '匈牙利・布達佩斯・Budavári Palota｜2026 年 11 月 26 日', font: F, size: 21, color: GRAY })] }),

  p('來到布達佩斯的第一個早上，就下雪了。11 月 26 日上午，我們上到多瑙河西岸的城堡山，布達城堡的屋頂、矮樹和欄杆全都蓋上一層白。11:10 太太站在城牆邊，背後是霧茫茫的多瑙河和對岸的佩斯；接著在美術館的大圓頂下拍合照、看雪片一直落下來；中午進去歷史博物館（另寫一篇），下午 14:31 出來時，又在圖魯鳥雕像和鏈橋前面停了很久。'),

  h1('一、山丘上的皇宮：七百多年的起落'),
  p('布達城堡（Budavári Palota）是歷代匈牙利國王的城堡和宮殿，蓋在城堡山的南端，旁邊就是滿是中世紀、巴洛克和新古典房子的城堡區。這裡從中世紀起就是王宮所在，15 世紀盧森堡王朝的西吉斯蒙德國王、後來的馬加什一世，都把它擴建得極其華麗。'),
  p('1541 年 8 月，鄂圖曼土耳其人佔領了布達，布達成為鄂圖曼「布丁省」的首府，被統治了一百四十多年。1686 年神聖同盟圍城 78 天才奪回布達，但中世紀的王宮也在這場大圍城中毀了。之後哈布斯堡家族在廢墟上蓋了巴洛克宮殿，1849 年匈牙利革命軍攻打布達時又整座燒毀。'),
  p('今天看到的樣子，主要是十九世紀末的大擴建：1885 年由伊柏（Ybl Miklós，就是設計歌劇院的那位）開始規劃，1891 年他過世後由豪斯曼（Hauszmann Alajos）接手，把面向多瑙河的一側擴成五棟相連、長達 304 公尺的建築，也就是今天的 A 到 E 棟。第二次世界大戰的圍城又把皇宮炸成廢墟，戰後重建時內部改成博物館和圖書館。1987 年，「布達佩斯的多瑙河兩岸與布達城堡區」被列為世界遺產。從 2019 年起，政府推動「國家豪斯曼計畫」，預計到 2030 年把城堡區恢復成二十世紀初的模樣。'),
  ...img('cs_wall', 1000, 750, '11:10，太太戴著白色毛帽、紅圍巾，靠在積雪的城牆邊；背後是飄雪中的多瑙河、碼頭上的遊船和遠方的伊莉莎白橋（待確認）'),
  verse('耶和華本為大，在我們神的城中，在他的聖山上，該大受讚美。', '詩篇 48:1'),
  heart('一座山丘上的皇宮，被毀了好幾次，又一次次蓋回來。今天它安安靜靜地站在雪裡，好像什麼都沒發生過；但每一塊石頭，其實都記得。'),

  h1('二、圓頂下的雪：國家美術館與歐根親王'),
  p('皇宮最醒目的綠色大圓頂下面，是匈牙利國家美術館（Magyar Nemzeti Galéria）。戰後皇宮各棟分給不同的機構：B、C、D 棟給國家美術館，E 棟給布達佩斯歷史博物館，F 棟給國家塞切尼圖書館。圓頂正面的柱廊上掛著美術館的大幅布條，前面的矮樹被雪蓋成一顆顆白色的圓球。'),
  p('圓頂前的大露台上，有一座歐根親王（Eugene of Savoy）的騎馬銅像，由雕刻家羅納・約瑟夫（Róna József）製作，1900 年揭幕。歐根親王是率軍打敗鄂圖曼軍隊的名將，1697 年的森塔戰役就是他的勝仗；這座像原本是為森塔鎮做的，因為小鎮付不起，後來由豪斯曼買下，放到了這裡。'),
  p('影片 v22（11:22）是從圓頂下往上拍，雪片一直從柱廊前落下來；影片 v23（11:47）是太太撐著粉紅色的傘，走在歐根親王雕像底下的石牆前。'),
  ...img('cs_couple', 750, 1000, '我們在國家美術館的大圓頂前合照：我戴藍色毛帽、穿藍色羽絨外套，太太戴白色毛帽；兩旁的矮樹積滿了雪，雪還在下'),
  verse('你們的罪雖像朱紅，必變成雪白。', '以賽亞書 1:18'),
  heart('雪落在古老的皇宮上，把所有的痕跡都蓋成一樣的白。這樣的潔白，是我們自己做不出來的，是從天上來的禮物。'),

  h1('三、圖魯鳥與鏈橋'),
  p('下午從博物館出來，往北走到哈布斯堡階梯（Habsburg-lépcső）的大鐵門，門柱上蹲著一隻張開雙翼的巨大青銅鳥：圖魯鳥（Turul）。圖魯是匈牙利神話中的神鳥，傳說牠讓阿爾帕德的祖母艾梅舍懷孕，而阿爾帕德就是帶領馬扎爾各部落進入喀爾巴阡盆地的首領，所以圖魯被看作匈牙利民族的象徵。這座雕像是唐納特・久洛（Donáth Gyula）的作品，牠的爪子抓著一把劍；揭幕年份有寫 1903 年、也有寫 1905 年（待確認）。因為年久失修，2021 年它被拆下來修復，同年 12 月再用起重機放回原位。'),
  p('從門邊的觀景台往下看，就是塞切尼鏈橋（Széchenyi Lánchíd）。它由塞切尼・伊什特萬伯爵發起興建，英國工程師克拉克（William Tierney Clark）設計、亞當・克拉克（Adam Clark）負責施工，1849 年通車，是第一座連接布達和佩斯的永久橋樑。對岸正對橋頭的是新藝術風格的格雷沙姆宮（今天是飯店），更遠的綠色大圓頂是聖伊什特萬大教堂。那天雲壓得很低，整個城市只剩下灰、白和幾棵還沒掉光葉子的紅樹。'),
  ...img('cs_turul', 750, 1000, '14:31，哈布斯堡階梯旁的圖魯鳥：張開雙翼、爪子抓著劍，站在雕著匈牙利國徽的門柱上；旁邊是積雪的紅色木樓梯和匈牙利國旗'),
  ...img('cs_bridge', 750, 1000, '14:31，從城堡山俯瞰鏈橋與多瑙河：對岸是格雷沙姆宮和聖伊什特萬大教堂的圓頂，前景的樹枝上都是雪'),
  verse('但那等候耶和華的必從新得力。他們必如鷹展翅上騰。', '以賽亞書 40:31'),
  heart('神話裡的鳥張開翅膀守著這座城，真實的橋把兩岸連在一起。我們從山上往下看，覺得旅行也是這樣：一邊仰望，一邊把不同的地方連起來。'),

  h1('旅人小記'),
  bullet('時間（依檔名與影片 metadata）：11 月 26 日（週四）11:10 城牆邊 → 11:22 國家美術館圓頂（影片 v22）→ 11:47 歐根親王雕像下（影片 v23，GPS 47.4966, 19.0403）→ 約 12:19–13:58 布達佩斯歷史博物館（另一篇）→ 14:31 圖魯鳥與鏈橋 → 傍晚 17:16 搭千禧地鐵去歌劇院。'),
  bullet('交通：可以從鏈橋旁的克拉克・亞當廣場坐纜車（Sikló）上山，或搭公車上城堡區；雪天石板路和階梯很滑，要穿防滑的鞋。'),
  bullet('待確認：我們是怎麼上山的（纜車、公車還是走路）；第一張照片遠方的白色吊橋是不是伊莉莎白橋；圖魯鳥是 1903 還是 1905 年揭幕；有沒有進國家美術館；這是不是布達佩斯今年的第一場雪。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：Wikipedia（布達城堡、1541 與 1686 年布達圍城、國家豪斯曼計畫、唐納特・久洛）、budavar.btk.mta.hu（豪斯曼擴建、A–F 棟分配）、habsburger.net（伊柏與豪斯曼）、Hungaricana《Walks in the Castle District》（歐根親王像、圖魯鳥）、aviewoncities、wanderlog（歐根親王像）、welovebudapest、Daily News Hungary（圖魯鳥修復）、Lonely Planet（圖魯傳說）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・布達佩斯・布達城堡雪景　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('布達佩斯_布達城堡雪景.docx', b));
