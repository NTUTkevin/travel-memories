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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '美泉宮（室外）：凱旋門、花園與維也納的第一場雪', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Schönbrunn Gloriette／Schlosspark｜2026 年 11 月 23 日', font: F, size: 21, color: GRAY })] }),

  p('美泉宮的上午，我們先不進宮殿，而是穿過花園、爬上後面的山丘，到最高處的凱旋門（Gloriette）看整座宮殿和維也納市區。中午過後進宮殿參觀（寫在〈美泉宮（室內）〉那一篇），出來的時候，宮殿前廣場的聖誕市集下起了雪——那是我們在維也納遇到的第一場雪。'),

  h1('一、凱旋門：山丘上的「榮耀殿堂」'),
  p('在美泉宮的山丘頂上蓋一座觀景建築，早在 17 世紀末費舍爾・馮・埃爾拉赫的設計裡就有了，但一直到瑪麗亞・特蕾莎晚年整修花園時才實現。1775 年，宮廷建築師霍恩貝格（Johann Ferdinand Hetzendorf von Hohenberg）依照新古典早期的風格，在山頂蓋了這座柱廊建築，它也是整座花園最後完成的建築。蓋的時候，還拆用了維也納郊外一座文藝復興宮殿「新建築」（Neugebäude）的石材。'),
  p('瑪麗亞・特蕾莎希望它用來榮耀哈布斯堡的力量和「正義之戰」——也就是為了和平、不得不打的戰爭。她在位期間先後經歷了奧地利王位繼承戰爭（1740–1748）和七年戰爭（1756–1763）。中間的屋頂上是一隻站在地球上的帝國大鷹，四周堆著盔甲、盾牌等戰利品雕塑；正面刻著一行金色的拉丁文：'),
  p('IOSEPHO II. AVGVSTO ET MARIA THERESIA AVGVSTA IMPERANTIB. ERECT. CIƆIƆCCLXXV.', { para: { alignment: AlignmentType.CENTER }, run: { italics: true, color: WINE } }),
  p('意思是「建於約瑟夫二世皇帝與瑪麗亞・特蕾莎女皇在位時，1775 年」。最後那串看起來很奇怪的字母，是古羅馬數字 1775 的老寫法。據說後來法蘭茲・約瑟夫皇帝曾把這裡當作早餐室，今天裡面是一間咖啡館。2016 年起它做了一次全面整修。'),
  ...img('sout_inscription', 1100, 735, '凱旋門中間的頂部：站在金色地球上的帝國大鷹和盔甲戰利品雕塑，下面是金色的拉丁文題字'),
  ...img('sout_me_gloriette', 735, 1100, '我靠在凱旋門前的石欄杆上，身後是黃色牆面、白色拱門和一排石甕'),
  verse('他止息刀兵，直到地極；他折弓、斷槍，把戰車焚燒在火中。', '詩篇 46:9'),
  heart('一座紀念戰爭的建築，今天變成大家喝咖啡、看風景、拍合照的地方。真正的勝利，大概就是有一天不再需要打仗。'),

  h1('二、從山丘俯瞰宮殿與維也納'),
  p('站在凱旋門前，可以看到整條中軸線：腳下是一個大水池，往下是一大片斜坡草地，再下去是幾何形的大花壇、兩旁成排的白色石像，最遠處就是黃色的宮殿，背後是一整片維也納市區，天氣好的時候遠處還看得到維也納森林的山丘。'),
  p('這座花園大約 120 公頃，是瑪麗亞・特蕾莎在位的最後十年請霍恩貝格重新規劃的，裡面放滿了希臘羅馬神話的雕像，用來表現皇室的地位。整修後的花園不只給宮廷使用，也開放給一般民眾散步。花園裡的動物園是 1752 年法蘭茲・史蒂芬下令建造的，是世界上現存最古老的動物園。'),
  p('從照片裡可以看到，宮殿正中央的那一段外牆罩著鷹架，正在整修（整修內容待確認）。'),
  ...img('sout_view1', 1100, 825, '從山丘上看美泉宮：前面是斜坡草地和松樹，中間是大花壇和兩排白色雕像，黃色的宮殿正中央罩著鷹架，後面是整片維也納市區'),
  ...img('sout_couple1', 825, 1100, '我們在凱旋門下的石欄杆旁合照，背後兩排秋天的樹中間是一條往下的步道，遠處是維也納的房子'),
  ...img('sout_couple2', 825, 1100, '水池邊的合照：太太抬頭看著我，身後是水池、宮殿和遠處的維也納市區'),
  ...img('sout_bench_wife', 1100, 825, '太太一個人坐在水池前的綠色長椅上，看著遠方的宮殿和城市'),
  ...img('sout_bench_me', 1100, 825, '換我坐在同一張長椅上，背著藍色背包看風景'),
  verse('他使我躺臥在青草地上，領我在可安歇的水邊。', '詩篇 23:2'),
  heart('那張長椅我們輪流坐，各拍了一張背影。天空灰灰的，可是整座城市就在眼前，兩個人什麼都不用說，就很安靜、很滿足。'),

  h1('三、宮殿前的聖誕市集與第一場雪'),
  p('從宮殿出來，正門前的「榮譽庭院」（Ehrenhof）已經變成聖誕市集。2026 年美泉宮聖誕市集從 11 月 6 日開到 2027 年 1 月 6 日，免費入場，平常每天上午 10 點到晚上 9 點，聖誕節後提早到晚上 7 點結束。市集裡有一棵掛滿金色球的大聖誕樹、許多賣手工藝品和熱紅酒的小木屋，還有給小朋友玩的旋轉木馬和稻草迷宮。我們在宮殿東翼看那塊「可以摸的牆板」時，窗外的鐵花窗後面就是市集的小木屋。'),
  p('就在這個時候，天上開始飄雪。維也納的第一層積雪，平均要到 11 月底左右才會出現，這一天我們剛好遇上。太太在大聖誕樹前伸出雙手接雪花，開心得像個小孩。'),
  ...img('sout_snow', 619, 1100, '影片截圖：宮殿前的大聖誕樹下，太太戴著黑色報童帽、圍著紅圍巾，伸出戴手套的雙手接雪花'),
  verse('他降雪如羊毛，撒霜如爐灰。', '詩篇 147:16'),
  heart('在皇宮前、聖誕樹下，和最愛的人一起接住維也納的第一場雪。這是蜜月旅行裡，最像童話的一刻。'),

  h1('旅人小記'),
  bullet('交通：地鐵 U4 到 Schönbrunn 站或 Hietzing 站，從宮殿走上凱旋門大約 15 到 20 分鐘。'),
  bullet('花園：花園免費開放，冬天開放時間比較短（以官網為準）。凱旋門的屋頂觀景台要另外買票，冬天可能不開。'),
  bullet('聖誕市集：2026 年 11 月 6 日到 2027 年 1 月 6 日，在宮殿正門前的榮譽庭院，免費入場。'),
  bullet('天氣：11 月底的維也納很冷，這天還下了雪，記得帶手套、圍巾和帽子。'),
  bullet('影片：v12 太太在聖誕樹前接雪花（13:51，GPS 在美泉宮前）。'),
  bullet('待確認：凱旋門咖啡館有沒有進去；宮殿正中央的鷹架在整修什麼；你們說的「第一場雪」是維也納這一季的第一場雪，還是我們這趟旅行第一次遇到雪；市集裡有沒有買東西或喝熱紅酒。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：schoenbrunn.at（凱旋門、花園）、Wikipedia（Gloriette、Schönbrunn Palace）、habsburger.net（美泉宮花園與古典神話）、austria.info（美泉宮、動物園）、wien.orf.at（凱旋門整修、2026 維也納聖誕市集一覽）、wien.info（美泉宮聖誕市集 2026）、weekend.at（維也納第一層積雪的平均日期）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・美泉宮（室外）　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_美泉宮（室外）.docx', b));
