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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '市政廳聖誕市集：維也納最盛大的聖誕夜色', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Wiener Christkindlmarkt am Rathausplatz｜2026 年 11 月 24 日', font: F, size: 21, color: GRAY })] }),

  p('11 月 24 日下午逛完家具博物館，天黑以後我們來到市政廳廣場。這裡的「維也納聖誕市集」是全城最大、也最有名的一個：新哥德式的市政廳整棟打亮，前面是一棵三十多公尺的聖誕樹、旋轉木馬和一排排小木屋，旁邊的市政廳公園裡還有掛滿紅色愛心的大樹。那天晚上從 17:19 待到 18:38，中間還走到附近的還願教堂，看兩座塔中間發亮的「無限」符號。'),

  h1('一、從七百年前的「十二月市集」到市政廳廣場'),
  p('維也納人很喜歡說，自己的聖誕市集是世界上最早的之一。根據市集主辦單位的說法，1296 年哈布斯堡家族的阿爾布雷希特一世（Albrecht I.）准許小商人舉辦一個「十二月市集」（Dezembermarkt），讓城裡的人在冬天買得到日用品；也有資料寫成 1298 年（年份各家說法不同，待確認）。那時還沒有「聖誕市集」這個名字，只是冬天的補給市場。'),
  p('後來這個市集在維也納城裡搬來搬去，名字也變成「聖嬰市集」（Christkindlmarkt）。奧地利的傳統裡，聖誕禮物是由「聖嬰」（Christkind，也就是嬰孩耶穌）送來的，所以市集就以祂命名。1975 年，原本在展覽宮（Messepalast，今天的博物館區 MuseumsQuartier）前面的市集，因為要蓋地下停車場，臨時搬到市政廳廣場，結果一待就是五十年，成了今天的樣子。'),
  p('入口是一座大拱門，上面掛著紅色蝴蝶結和草寫的燈字「Wiener Christkindlmarkt」，從拱門下望過去，正好框住市政廳的中央鐘樓。2026 年的市集從 11 月 13 日開到 12 月 26 日，每天 10 點到 22 點（12 月 24 日到 18:30）；旁邊的「冰上夢想」溜冰場則開到 2027 年 1 月 6 日。這裡的餐飲攤位全部是有機認證，在維也納的聖誕市集裡很少見。'),
  ...img('rh_arch', 750, 1000, '17:19，我們在市集入口的大拱門下合照：紅色蝴蝶結、燈字「Wiener Christkindlmarkt」，背後是打亮的市政廳和旋轉木馬'),
  ...img('rh_sign', 1000, 750, '拱門上的草寫燈字，正好框住市政廳的鐘樓和兩側四座尖塔'),
  ...img('rh_stall', 750, 1000, '掛滿松枝和燈泡的小木屋，紅色招牌寫著「…nitzereien / …handwerk」（應是木雕與手工藝，待確認），屋頂後面露出市政廳'),
  verse('你們要看見一個嬰孩，包著布，臥在馬槽裡，那就是記號了。', '路加福音 2:12'),
  heart('七百年前，市集是為了讓人在冬天活下去；今天，市集是為了讓人在冬天心裡溫暖。名字叫「聖嬰市集」，提醒我們聖誕節的主角，是那個躺在馬槽裡的嬰孩。'),

  h1('二、市政廳：比還願教堂矮一公尺的塔'),
  p('市政廳（Wiener Rathaus）是 1872 到 1883 年蓋的，建築師是德國人弗里德里希・馮・施密特（Friedrich von Schmidt）。他在來維也納之前，曾在科隆大教堂當過建築師，所以特別擅長哥德式。當時舊市政廳已經太小，市長菲爾德（Cajetan Felder）提議在新開闢的環城大道（Ringstraße）旁蓋新的，施密特在競圖中勝出。他選擇的是新哥德式，參考了法蘭德斯地區（例如布魯塞爾）的市政廳，象徵自治城市的市民精神。今天這裡是維也納市長、市議會和邦議會辦公的地方。'),
  p('中央鐘樓有一個有名的小故事。據說皇帝法蘭茲・約瑟夫規定，市政廳的塔不能比附近的還願教堂（99 公尺）高。施密特就把塔蓋到約 98 公尺，然後在塔尖放了一尊鐵鑄的騎士像——「市政廳人」（Rathausmann），這樣加起來就比教堂還高了。市政廳人是 1882 年 10 月 21 日安上去的，由鍛鐵師傅亞歷山大・內爾（Alexander Nehr）照著皇帝馬克西米連一世的盔甲做成，身高 3.4 公尺、連旗杆 5.4 公尺，市政府的網站還很幽默地說，他的鞋子要穿 63 號。'),
  p('我們從市集走進旁邊的市政廳公園，抬頭就是紅透了的楓樹，中間露出亮著燈的鐘樓和尖塔，是那天晚上最喜歡的畫面。'),
  ...img('rh_carousel', 1000, 750, '市政廳正面：左邊是雙層旋轉木馬，中間是大聖誕樹，右邊是一排小木屋和還沒掉光葉子的大樹'),
  ...img('rh_leaves', 750, 1000, '17:27，從市政廳公園的楓葉之間看鐘樓；鐘面上的時間大約是五點半'),
  ...img('rh_leaves2', 1000, 750, '同一棵楓樹換個角度：紅葉、黑夜和整棟打亮的新哥德式立面'),
  ...img('rh_wifecap', 1000, 750, '17:32，太太在公園步道上，背後是市政廳鐘樓和掛滿燈的樹'),
  verse('我所使你們被擄到的那城，你們要為那城求平安，為那城禱告耶和華；因為那城得平安，你們也隨著得平安。', '耶利米書 29:7'),
  heart('一棟為市民蓋的房子，比皇帝蓋的宮殿還要用心。站在楓樹下看那座鐘樓，也想為這座讓我們住了幾天的城市祝福。'),

  h1('三、公園裡的聖誕世界：愛心樹與星星樹'),
  p('市政廳公園裡是給小朋友和家庭的「聖誕世界」（Weihnachtswelt）。2026 年重新設計，共有 17 個跟聖誕節有關、和真人一樣大的互動站，還有溜冰小路從樹林中間穿過去。'),
  p('最受情侶歡迎的是「愛心樹」（Herzerlbaum）：一棵大樹上掛了 200 顆紅色 LED 愛心，整棵樹變成紅色。今年它在 11 月 10 日就先點亮，比市集開幕還早三天。2026 年還有一個新項目「愛心飛行」（Herzflug），從黃昏到晚上 9 點，每半小時從一座 13 公尺高的塔上，沿著鋼索讓一顆愛心飛過廣場。照片左邊那棟白色的建築，就是在市政廳正對面的城堡劇院（Burgtheater）：它 1741 年由瑪麗亞・特蕾莎在霍夫堡旁邊創立，現在這棟環城大道上的新劇院是森佩爾（Gottfried Semper）和哈森瑙爾（Karl von Hasenauer）設計，1888 年 10 月 14 日開幕。'),
  p('公園另一頭有一棵掛滿星星的燈樹，太太在柵欄邊拍了好幾張；旁邊還有一個小木屋舞台，裡面是會動的熊樂團（應該也是聖誕世界的其中一站，待確認）：一隻戴紅帽的熊拿著麥克風唱歌，旁邊有麋鹿吹喇叭、另一隻熊彈吉他。影片 v18、v19 是在這裡拍的。'),
  ...img('rh_hearts', 1000, 750, '17:33，愛心樹：滿樹的紅色愛心、地上一顆一顆的紅燈球、會發光的愛心拱門和薑餅屋；左邊白色建築是城堡劇院'),
  ...img('rh_startree', 750, 1000, '17:37，太太在掛滿星星的燈樹前，背後是市政廳鐘樓和被藍光照亮的樹'),
  ...img('rh_bears', 562, 1000, '影片 v19 截圖：小木屋裡會動的熊樂團，戴紅帽的熊唱歌、麋鹿吹喇叭，旁邊還有一棵小聖誕樹'),
  verse('求你將我放在你心上如印記，帶在你臂上如戳記；因為愛情如死之堅強。', '雅歌 8:6'),
  heart('一整棵樹的紅色愛心，好像整座城市都在說「我愛你」。旅行到第六天，能和太太一起站在這棵樹下，就是最好的聖誕禮物。'),

  h1('四、還願教堂與「光的無限」'),
  p('離開公園往北走幾百公尺，就是還願教堂（Votivkirche）。它的由來是一次暗殺：1853 年 2 月 18 日，23 歲的皇帝法蘭茲・約瑟夫在城牆上散步時，被匈牙利民族主義者、裁縫學徒利貝尼（János Libényi）從背後用刀刺傷，幸好隨從副官歐唐諾伯爵（O’Donnell）和一位路過的市民及時把他制服。皇帝的弟弟斐迪南・馬克西米連大公（後來的墨西哥皇帝）發起全國募款，蓋一座教堂來感謝神保住皇帝的性命，「還願」（Votiv）就是這個意思。'),
  p('設計教堂的是當時才 26 歲的建築師費斯特爾（Heinrich von Ferstel），教堂 1856 年開始蓋，1879 年 4 月 24 日祝聖——那天正好是皇帝和西西的銀婚紀念日。它高 99 公尺，是維也納第二高的教堂，只輸給聖史蒂芬大教堂，也是環城大道上最早的建築。前面提到市政廳的塔「不能比它高」，說的就是這兩座塔。'),
  p('我們看到的紅色「∞」，是維也納藝術家比莉・坦納（Billi Thanner）的光雕作品「光的無限」（Unendlichkeit des Lichtes）。它是一個 28 × 6 公尺的鋁製無限符號，架在兩座塔之間，2025 年 12 月 16 日第一次點亮，全部由私人出資。原本只展到 2026 年 8 月，後來兩度延長，目前延到 2027 年 5 月底。藝術家說，這件作品是想讓人在一瞬間想起希望、喜樂，以及彼此之間的連結。'),
  ...img('rh_votiv', 1000, 750, '17:53，還願教堂：兩座新哥德式尖塔被紅色的「∞」光環連在一起，正面中間是玫瑰窗；影片 v20（18:00）也在這裡拍'),
  verse('我拿甚麼報答耶和華向我所賜的一切厚恩？', '詩篇 116:12'),
  heart('一座為了「謝謝神救命」而蓋的教堂，一百多年後掛上了「無限」的光。感謝也許就是這樣：一次被拯救，一輩子記得。'),

  h1('五、回到市政廳前：一杯熱飲和來自福拉爾貝格的大樹'),
  p('18:38 我們繞回市政廳前，在聖誕樹下喝了一杯熱飲。今年市政廳前的聖誕樹來自奧地利最西邊的福拉爾貝格邦（Vorarlberg）達米爾斯（Damüls）：一棵 31 公尺高、約 9 噸重的雲杉，10 月底就運到維也納，掛了大約 1,000 顆紅色和金色的球、2,000 盞燈，11 月 13 日開幕時由市長路德維希和福拉爾貝格邦長瓦爾納一起點亮。由各邦輪流送聖誕樹給維也納的傳統，從 1959 年就開始了。'),
  p('市集裡的熱飲一律用可以重複使用的陶瓷杯（Häferl）裝，買飲料時要另外付杯子押金，喝完可以退杯拿回押金，也可以把杯子帶回家當紀念品。每年的杯子圖案都不一樣，今年是紅底白線，畫著小木屋、聖誕樹和星星。（我們喝的是熱紅酒還是潘趣酒、押金多少、杯子有沒有帶回家，待確認。）'),
  ...img('rh_mug', 750, 1000, '18:38，手上的紅色市集杯對著市政廳：整面新哥德式拱窗、市政廳人站在鐘樓頂端，31 公尺的聖誕樹掛滿燈和紅球'),
  verse('願田野和其中所有的都歡樂！那時，林中的樹木都要在耶和華面前歡呼。', '詩篇 96:12'),
  heart('從福拉爾貝格的山上來的一棵樹，站在維也納的市政廳前，照亮了好幾百萬人的十二月。我們捧著一杯熱飲，是其中很幸福的兩個人。'),

  h1('旅人小記'),
  bullet('時間（依檔名與影片 metadata）：17:19 入口拱門 → 17:22 小木屋與旋轉木馬 → 17:27 楓葉與鐘樓 → 17:33 愛心樹 → 17:37–17:41 星星樹與熊樂團（v18、v19）→ 17:53–18:00 還願教堂（v20）→ 18:38 回到市政廳前喝熱飲。這天下午先去了家具博物館（14:29–15:19）。'),
  bullet('交通：地鐵 U2 Rathaus 站，或電車 1、2、D 號到 Rathausplatz/Burgtheater；還願教堂在 U2 Schottentor 站旁，從市集走過去約 5–10 分鐘。'),
  bullet('2026 年市集：11/13–12/26，每天 10:00–22:00，12/24 到 18:30；溜冰場到 2027/1/6；愛心飛行黃昏到 21:00 每半小時一次。晚上人非常多，背包要顧好。'),
  bullet('待確認：「十二月市集」是 1296 還是 1298 年；紅色招牌那間小木屋賣什麼；熊樂團是不是聖誕世界的一站；有沒有看到「愛心飛行」；喝的是熱紅酒還是潘趣酒、杯子押金多少、杯子有沒有帶回家；有沒有進還願教堂裡面。'),
  bullet('附記：雲端照片資料夾另有兩張沒放進文章的照片（08 楓葉與鐘樓冷色調版本、11 星星樹前太太的另一張）。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：christkindlmarkt.at（市集歷史、1975 年搬遷）、wien.gv.at（市集歷史、市政廳人）、wien.info（2026 市集時間、市政廳、還願教堂）、wien.orf.at（2026 年達米爾斯聖誕樹、光的無限）、vienna.at（愛心樹 2026）、news.at（愛心飛行、光的無限延長至 2027 年 5 月）、Wien Geschichte Wiki（聖誕樹傳統 1959）、austria-forum.org（市政廳人、內爾、城堡劇院）、habsburger.net（1853 暗殺、市政廳塔高）、Wikipedia（市政廳、還願教堂、城堡劇院）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・市政廳聖誕市集　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_市政廳聖誕市集.docx', b));
