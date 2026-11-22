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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '聖巴多羅買教堂：湖上的紅色洋蔥頂', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '德國・巴伐利亞・國王湖 St. Bartholomä｜2026 年 11 月 19 日', font: F, size: 21, color: GRAY })] }),

  p('國王湖的紅色洋蔥頂教堂，是我們這一天最想看的風景。從湖邊的小鎮搭船過去，湖水慢慢從綠色變成灰藍色，兩邊的山壁越來越高，最後，一座白牆、紅頂的小教堂，出現在一片黃色落葉松前面。'),

  h1('一、搭船去對岸'),
  p('國王湖沒有環湖公路，要到對岸的聖巴多羅買（St. Bartholomä），一般人都是搭船：從湖邊的碼頭單程大約 35 分鐘。夏天可以再往前坐到更深處的薩列特（Salet），不過那一站只在復活節到十月中旬開，十一月就只到聖巴多羅買。船都是電動的，從 1909 年起就是這樣。'),
  p('船開了大約二十分鐘，船長通常會把船停下來，吹一段小號或圓號，讓大家聽山壁傳回來的回音。這個「國王湖的回音」很有名，據說最多可以聽到七次。我們這一趟聽到了幾次，之後再補。'),
  ...img('k2_boat_selfie', 1400, 1050, '還停在碼頭時的船上自拍：宜恩、我和太太；窗外是碼頭邊的黑色木屋頂和黃旗，船艙裡貼著救生圈與逃生錘的標示'),
  ...img('k2_flag', 1400, 1050, '船尾的藍白旗（巴伐利亞的顏色，待確認）；遠處兩山之間的湖面上，有一小片白色的建築，是我們要去的地方'),
  ...img('k2_boatpass', 1400, 1050, '另一艘電動船從對岸駛過：紅木的船緣、白色的船身，後面是黃色與橘紅色的山坡'),
  verse('耶和華在高處大有能力，勝過諸水的響聲，洋海的大浪。', '詩篇 93:4'),
  heart('船很安靜，只有水聲和旁邊人說話的聲音。這樣的湖，好像不需要誰吵著說「我很大」，山壁一直圍著它，它就是這樣大。'),

  h1('二、1134 年的教堂與狩獵行宮'),
  p('聖巴多羅買教堂建在一個叫「希爾紹」（Hirschau）的半島上，這個半島是河水多少年來衝下來的碎石堆出來的，大約有 85 公頃，只有船能到。教堂最早在 1134 年，由貝希特斯加登的侯爵修道院（前一篇提過的那個小邦）蓋在一座更小的舊建築原址上。因為這裡可以打獵，旁邊後來也蓋了修道院領主的狩獵行宮。'),
  p('現在看到的教堂，是 1697 到 1698 年重建的，屬於巴洛克式。從照片看，它有幾個圓圓的紅色屋頂，加上兩座白色的塔樓，塔頂是兩個大小不同的洋蔥頂。1803 年修道院領地結束、1810 年轉給巴伐利亞之後，這裡成了王室的狩獵行宮，教堂旁邊的老房子，現在是一間餐廳。旁邊還有一個國王湖的漁場：這裡捕魚的歷史，可以追溯到 1160 年，當時燻過的紅點鮭（Saibling）就從國王湖運到義大利北部；對修道院來說，因為教會一年有很多齋戒日，魚也是很重要的收入。'),
  ...img('k2_church_lake', 1400, 788, '從湖上遠看聖巴多羅買教堂：紅色圓頂與兩座白塔，左邊是灰色木牆的小屋，右邊是以前的狩獵行宮'),
  ...img('k2_path', 1050, 1400, '半島岸邊的小路：教堂與旁邊的老房子貼著山腳，湖水很清，岸邊堆著一圈落葉'),
  ...img('k2_close', 1050, 1400, '近看教堂：紅色的鱗片屋頂已長了青苔，白色的塔身、黃色的窗框，還有背後金黃色的落葉松'),
  verse('主啊，你世世代代作我們的居所。', '詩篇 90:1'),
  heart('同一個地方，九百年前是修道院的教堂，後來是國王的行宮，現在是我們這樣的旅人來看的風景。房子換了很多主人，但是人來這裡抬頭看山的心，大概一直沒變。'),

  h1('三、祭壇與聖巴多羅買'),
  p('教堂的名字來自使徒巴多羅買（Bartholomäus），他是阿爾卑斯山區農民與乳農的守護聖人，高壇上的畫面是他，每年 8 月 24 日前後，這裡會有朝聖。教堂裡面白色的牆上，有淡粉色的灰泥浮雕；祭壇則像是另一個世界：深紅色的豹斑紋圓柱、金色的捲葉、中間一幅暗色的大畫。'),
  ...img('k2_altar', 1050, 1400, '教堂內部的高壇：紅色豹斑紋圓柱與金色雕飾，中間的畫據說是聖巴多羅買（畫的內容待確認）；兩側的窗讓白色牆面特別亮'),
  verse('萬軍之耶和華啊，你的居所何等可愛！', '詩篇 84:1'),
  heart('從湖邊的冷風走進來，裡面很靜，牆壁是白色的，祭壇是深紅色的。這麼偏遠的地方，也有人蓋起這麼用心的房子，這是我們很少想到的事。'),

  h1('四、像牆一樣的山'),
  p('教堂背後的那面大山壁，是瓦茨曼山（Watzmann）的東壁。瓦茨曼山高 2,713 公尺，是完全在德國境內最高的山；東壁從國王湖邊一直往上到南峰，落差約 1,800 公尺，是東阿爾卑斯山最長的連續岩壁，因為登山口就在聖巴多羅買教堂後面，也被叫做「巴多羅買壁」。我們那天看到的，是被雲遮住上半截的山壁，上面灑著一層薄薄的新雪。'),
  p('當地還有一個「瓦茨曼國王」的傳說：很久以前，有個殘忍的國王叫瓦茨曼，跟他的妻子與七個孩子，因為做惡，被上帝變成了石頭，就是現在的瓦茨曼山、小瓦茨曼山和山邊的幾座「孩子山」，他們的血流到山谷，變成了兩座湖，比較大的那座叫國王湖。'),
  ...img('k2_east1', 1050, 1400, '半島上的老栗子樹與長椅，背後是被雲遮住上半截、落了薄雪的岩壁（可能是瓦茨曼東壁，待確認）'),
  ...img('k2_east2', 1050, 1400, '同一個地方換個角度：樹枝上掛著一串燈泡，遠處的岩壁一層一層，像是用刀切出來的'),
  ...img('k2_bench', 1050, 1400, '長椅上，太太靠著我的肩膀，一起看對岸的山；左邊遠處有一座尖尖的山頭（地點待確認）'),
  verse('諸山未曾生出，地與世界你未曾造成，從亙古到永遠，你是神。', '詩篇 90:2'),
  heart('傳說說這些山是人變成的。我們看著，只覺得它們比故事老得多，比我們想過的任何事都更老。我們坐在長椅上，覺得自己很小，也覺得被好好地放在這裡。'),

  h1('旅人小記'),
  bullet('船班、票價與停靠站請以官方網站為準：seenschifffahrt.de。十一月湖上的船只到聖巴多羅買，來回算上停留，至少要留 2 個小時。'),
  bullet('教堂旁的老房子（以前的狩獵行宮）現在是餐廳，旁邊還有漁場的小吃店；岸邊的栗子樹下有戶外座位。'),
  bullet('這一篇還有幾處待確認：船上有沒有聽到回音、祭壇畫的內容、長椅與岩壁的位置、船尾旗幟的種類。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：bayernportal.de 與 berchtesgaden.de（Wallfahrtskirche St. Bartholomä、Der Königssee、Watzmann Ostwand、Die Watzmann-Sage、Fischerei Königssee）、Wikipedia〈St. Bartholomew’s Church, Berchtesgaden〉〈Watzmann〉〈Königssee〉、salzburg.info（Königssee Schifffahrt）、seenschifffahrt.de（船班與歷史）、wikiwand〈Königssee〉。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・聖巴多羅買　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('聖巴多羅買教堂.docx', b));
