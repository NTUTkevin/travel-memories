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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '美泉宮（室內）：走進皇帝與西西的夏宮', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Schloss Schönbrunn｜2026 年 11 月 23 日', font: F, size: 21, color: GRAY })] }),

  p('在維也納的第二天，我們去了美泉宮（Schloss Schönbrunn，也譯作「熊布朗宮」）。上午先爬上後山的凱旋門看整座宮殿和維也納市區（寫在另一篇〈美泉宮（室外）〉），中午過後拿著語音導覽機走進宮殿，一間一間看哈布斯堡家族的房間。這一篇記錄的是宮殿裡面的故事：法蘭茲・約瑟夫皇帝和西西皇后、瑪麗亞・特蕾莎女皇和她的家人，還有大廳、中國漆器廳和「百萬廳」。'),

  h1('一、從「美麗的泉水」到皇家夏宮'),
  p('「Schönbrunn」的意思是「美麗的泉水」。傳說 1612 年，皇帝馬蒂亞斯（Matthias）來這裡打獵時發現了一處泉水，脫口說出「多美的泉水啊！」（Welch schöner Brunn!），這個名字就這樣留了下來。這裡最早是哈布斯堡家族的狩獵莊園，1683 年土耳其第二次圍攻維也納時被毀。'),
  p('之後，皇帝利奧波德一世請建築師費舍爾・馮・埃爾拉赫（Johann Bernhard Fischer von Erlach）為兒子約瑟夫一世蓋一座新的行宮，1696 年到 1701 年動工。真正讓美泉宮變成今天這個模樣的，是瑪麗亞・特蕾莎女皇：她在 1743 年把這裡定為夏季的住所，請宮廷建築師帕卡西（Nicolaus Pacassi）改建成洛可可風格。每年夏天，女皇帶著一千五百多人的宮廷搬來這裡，美泉宮成了整個帝國的政治舞台。'),
  p('1918 年帝國結束以後，宮殿開放給一般民眾參觀。1996 年，美泉宮和它的花園一起被列為聯合國教科文組織的世界遺產，是歐洲最重要的巴洛克建築群之一。'),
  verse('若不是耶和華建造房屋，建造的人就枉然勞力。', '詩篇 127:1'),
  heart('三百多年來，這座宮殿換過好幾代主人，蓋了又改、改了又修。真正長久的，不是金碧輝煌的牆，而是住在裡面的人的故事。'),

  h1('二、宮廷禮拜堂：瑪利亞的婚禮'),
  p('宮殿一樓的禮拜堂，還保留著 1700 年左右費舍爾・馮・埃爾拉赫設計的位置和結構。1743 年瑪麗亞・特蕾莎下令重新裝飾，主祭壇畫由奧地利巴洛克畫家保羅・特羅格（Paul Troger）在 1743 到 1744 年間完成，畫的是「瑪利亞的婚禮」——瑪利亞和約瑟訂婚的場景；禮拜堂也在 1745 年改為奉獻給這個節日。祭壇上方是一圈金色的光芒和天使，下面是大理石做成的圓頂聖體櫃。皇帝約瑟夫二世的第二次婚禮，就是在這座禮拜堂舉行的。'),
  ...img('sin_chapel', 825, 1100, '宮廷禮拜堂的主祭壇：上方是金色光芒與天使，中間是特羅格畫的〈瑪利亞的婚禮〉，穿藍袍的瑪利亞和約瑟站在大祭司面前'),
  verse('有人攻勝孤身一人，若有二人便能抵擋他；三股合成的繩子不容易折斷。', '傳道書 4:12'),
  heart('在一座皇宮裡，最中心的祭壇畫畫的竟然是一場婚禮。我們這趟是蜜月旅行，站在這幅畫前面，好像也被祝福了一次。'),

  h1('三、法蘭茲・約瑟夫與西西'),
  p('語音導覽從西翼法蘭茲・約瑟夫皇帝的房間開始。他 1830 年在美泉宮出生，十八歲就登基，一共統治了 68 年。他每天清晨四點多起床、五點前就開始辦公，稱自己是「國家的第一個公務員」。'),
  p('撞球室是等候覲見皇帝的地方，皇帝每週接見兩次民眾。牆上中間那幅大畫，畫的是 1758 年第一次頒發「瑪麗亞・特蕾莎軍事勳章」；兩旁的畫紀念這個勳章成立一百週年，法蘭茲・約瑟夫在大廳裡舉行了盛大的宴會。我們照片裡太太身後的那幅，就是宴會廳裡擠滿穿紅色、白色軍服的軍官的場面。'),
  ...img('sin_billiard', 825, 1100, '撞球室：太太拿著語音導覽機專心聽解說；身後的大畫是軍事勳章百年紀念的宴會，右邊是皇帝的撞球桌'),
  p('接下來的胡桃木廳（Walnut Room）保留了瑪麗亞・特蕾莎時代的洛可可木雕牆板，法蘭茲・約瑟夫就在這裡接見請求覲見的人，一個早上可以見上百人，據說他記人名和面孔的本事非常好。我們在一間深色木牆、鍍金裝飾的房間裡，看到桌上一尊寫著「RADETZKY」的胸像——拉德茨基（Joseph Radetzky，1766–1858）是 1848 年替年輕皇帝在義大利打勝仗的元帥，老約翰・史特勞斯的〈拉德茨基進行曲〉就是為他寫的（這間是不是胡桃木廳，待確認）。'),
  ...img('sin_walnut', 825, 1100, '深色木牆、鍍金洛可可雕飾的房間，大理石桌上是拉德茨基元帥的胸像，鏡子裡映出它的背面；左邊角落是一盆棕櫚和一尊白色小立像（房間名稱待確認）'),
  ...img('sin_chandelier', 825, 1100, '同一間房間的鍍金枝形吊燈，從下往上拍'),
  p('1854 年，法蘭茲・約瑟夫娶了十六歲的巴伐利亞公主伊莉莎白——大家口中的「西西」。為了這場婚禮，美泉宮的皇室寢宮重新裝修過。導覽中可以看到西西的書房（樓梯間）、梳妝室和沙龍：她在書房寫信、寫日記和詩；梳妝室展示她的美容習慣，光是整理那頭長髮每天就要好幾個小時。1898 年 9 月，西西在日內瓦被一名無政府主義者刺殺。'),
  p('法蘭茲・約瑟夫晚年把美泉宮當成一年四季的住所，1916 年 11 月 21 日在這裡的鐵床上過世，享年 86 歲。我們參觀的這一天是 11 月 23 日，剛好是他過世一百一十週年的兩天後。'),
  verse('艷麗是虛假的，美容是虛浮的；惟敬畏耶和華的婦女必得稱讚。', '箴言 31:30'),
  heart('西西是全世界都知道的美麗皇后，但導覽裡聽到的，更多是她的寂寞和不自由。皇宮再大，也比不上一個能做自己的家。'),

  h1('四、瑪麗亞・特蕾莎和她的孩子們'),
  p('瑪麗亞・特蕾莎是哈布斯堡家族唯一一位女性統治者，和丈夫法蘭茲・史蒂芬（Franz I Stephan）生了十六個孩子。法蘭茲・史蒂芬在 1745 年於法蘭克福被選為神聖羅馬帝國皇帝。導覽中有好幾間房間掛著這個家族的肖像：兒童房裡是女皇的女兒們，大多很年輕就為了政治聯姻嫁到歐洲各國的宮廷——包括後來的法國王后瑪麗・安東尼。'),
  ...img('sin_fs_portrait', 825, 1100, '金色畫框裡的皇帝肖像：黑色羽毛帽、白色捲髮、金色蕾絲長袍，手扶權杖，旁邊放著皇冠（推測是法蘭茲一世・史蒂芬，所在房間待確認）'),
  ...img('sin_kids', 825, 1100, '三個穿白衣的皇家孩子坐在白色毛皮和紅色絲絨上，最大的手裡拿著花（推測是瑪麗亞・特蕾莎的子女，人物與所在房間待確認）'),
  p('對我們來說最有意思的一間是「鏡廳」（Hall of Mirrors）：1762 年，六歲的莫札特就在這裡第一次在瑪麗亞・特蕾莎面前演奏。兩天前我們才在薩爾斯堡參觀他的出生地和故居，沒想到在維也納又遇見了小莫札特的足跡。'),
  verse('兒女是耶和華所賜的產業；所懷的胎是他所給的賞賜。', '詩篇 127:3'),
  heart('畫裡的孩子們穿著華服，但他們的人生很早就被安排好了。看著這些肖像，更覺得能自己選擇和誰共度一生，是很珍貴的事。'),

  h1('五、大廳：哈布斯堡的金色舞台'),
  p('大廳（Great Gallery）是整座宮殿的中心，長 43 公尺、寬將近 10 公尺，白色牆面配上鍍金的灰泥裝飾和高大的鏡子，是洛可可宮廷大廳最經典的樣子。皇室在這裡舉行舞會、宴會和招待會，瑪麗亞・特蕾莎也喜歡帶著丈夫和孩子在這裡出現，向大家展示家族後繼有人。'),
  p('天花板上的三幅大濕壁畫，是義大利畫家古列爾米（Gregorio Guglielmi）在 1759 到 1762 年左右畫的，主題都在歌頌女皇的統治：中間那幅是「奧地利君主國的繁榮」，女皇和丈夫坐在中央，四周是象徵各種美德的人物；另一邊是和平、藝術與科學；東邊那幅則是女皇軍隊的軍事場面，還畫了她在 1752 年創立的維也納新城軍事學院。東邊這幅濕壁畫在 1945 年 4 月二戰最後的空襲中被炸毀，後來才重新畫上。'),
  p('以前大廳的吊燈每盞要點 70 根蠟燭，1901 年宮殿通電以後才換成燈泡。1961 年，美國總統甘迺迪和蘇聯領導人赫魯雪夫就是在這個大廳會面的。'),
  ...img('sin_gallery1', 1100, 735, '大廳東側的天花板：濕壁畫裡是一排排拿著長槍的士兵和軍官，下面垂著鍍金的大吊燈'),
  ...img('sin_gallery2', 1100, 735, '整個大廳：兩幅天花板濕壁畫、一整排鍍金壁燈和大吊燈，地上擠滿了聽導覽的遊客'),
  ...img('sin_gallery_me', 1100, 735, '我在大廳裡抬頭看天花板的濕壁畫'),
  ...img('sin_gallery_wife', 1100, 735, '太太戴著黑色報童帽、圍著酒紅色圍巾，在大廳裡仰望天花板'),
  verse('然而我告訴你們，就是所羅門極榮華的時候，他所穿戴的，還不如這花一朵呢！', '馬太福音 6:29'),
  heart('站在大廳中間，兩個人都忍不住一直抬頭看。人造的榮華可以做到這麼極致，但比起神親手造的一朵花，還是差得遠。'),

  h1('六、東翼：中國漆器廳、百萬廳與壁毯廳'),
  p('宮殿東翼是瑪麗亞・特蕾莎時代最珍貴的幾間房間，也看得出她對遠方「異國情調」的喜愛。'),
  h2('中國漆器廳（Vieux-Laque Room）'),
  p('這間原本是法蘭茲・史蒂芬的私人接見室。1765 年丈夫過世後，瑪麗亞・特蕾莎把它改成紀念他的房間，牆上鑲著從中國來的黑色漆器板，女皇自己則終生穿著喪服。她請人畫了三幅畫掛在這裡：中間是義大利畫家巴托尼（Pompeo Batoni）畫的法蘭茲・史蒂芬全身像（約 1770 年），旁邊是長子約瑟夫二世和弟弟利奧波德的雙人像，第三幅是利奧波德的妻子路易莎（Ludovica）和他們最初幾個孩子——穿紫色衣服的小男孩，就是後來的皇帝法蘭茲二世。拿破崙在 1805 年和 1809 年佔領維也納時，就住在東翼這幾間紀念房間裡。'),
  ...img('sin_laque_fs', 825, 1100, '中國漆器廳：黑底金色的中國漆器牆板中間，是巴托尼畫的法蘭茲一世・史蒂芬全身像，他穿紅色禮服、斜披紅白綬帶'),
  ...img('sin_laque_family', 825, 1100, '同一間房間的另一面牆：水晶吊燈旁邊是路易莎和孩子們的畫像（穿粉紫色衣服的孩子推測是後來的法蘭茲二世）'),
  h2('百萬廳（Millions Room）'),
  p('百萬廳是美泉宮最奢華的房間之一，瑪麗亞・特蕾莎用它私下接見客人。牆板用的是珍貴的熱帶玫瑰木（當時叫 Feketin），鑲著六十幅鍍金洛可可框，框裡是 16、17 世紀印度蒙兀兒宮廷的細密畫——它們被剪開、重新拼貼成新的畫面，再配上互相對照的鏡子，讓房間看起來像無限延伸。「百萬廳」這個名字，是帝國結束以後才在導覽書裡出現的。'),
  p('出口附近有一塊可以摸的「壁板樣品」，照著百萬廳的牆板重做：玫瑰木貼皮加上鍍金的木雕裝飾，高起來的地方磨得發亮，低的地方保持霧面。說明牌上寫著，原本的牆面只要摸一下都可能造成無法修復的損害，所以請大家只摸這塊樣品。'),
  ...img('sin_millions', 825, 1100, '百萬廳（語音導覽編號 032）：玫瑰木牆上鑲滿金框細密畫，鏡子裡映出白色胸像和藍綠色椅子'),
  ...img('sin_panel', 825, 1100, '「可以摸的牆板」：照百萬廳重做的玫瑰木與鍍金雕飾樣品，太太的手正輕輕摸著它；窗外的鐵花窗後面就是聖誕市集'),
  h2('壁毯廳（Gobelin Salon）'),
  p('壁毯廳的牆上掛著 18 世紀布魯塞爾織的壁毯，畫的是市集和港口的熱鬧景象。椅子的椅背和椅面也是壁毯織成的，每張代表一年中的不同月份；這些椅子是 1873 年維也納世界博覽會時才搬進這個房間的。'),
  ...img('sin_gobelin', 825, 1100, '壁毯廳：牆上的布魯塞爾壁毯織著城牆下的魚市，下面一排鍍金扶手椅，椅背上也織著田園風景'),
  verse('因為你的財寶在哪裡，你的心也在那裡。', '馬太福音 6:21'),
  heart('一位女皇失去丈夫以後，把最珍貴的房間改成紀念他的地方，一輩子穿著黑衣。百萬廳再貴，也比不上她心裡的那個人。'),

  h1('旅人小記'),
  bullet('交通：搭地鐵 U4 到 Schönbrunn 站，走幾分鐘就到宮殿正門。'),
  bullet('門票：宮殿要買時段票，附語音導覽機。只能在官方網站 imperialtickets.com 買票，官網提醒要小心假票。宮殿每天開放（開放時間以官網為準）。'),
  bullet('拍照：我們這次在宮殿裡拍了照片；各展室的拍照規定以現場公告為準。'),
  bullet('順序：我們中午 12 點左右先看了禮拜堂，再跟著語音導覽走皇帝和西西的房間、大廳，最後是東翼，大約 13 點出來，剛好接上外面廣場的聖誕市集。'),
  bullet('待確認：拉德茨基胸像那間是不是胡桃木廳；黑帽皇帝肖像是不是法蘭茲・史蒂芬、掛在哪一間；三個孩子的畫像是誰、在哪一間；你們買的是哪一種票（官網現在叫 Palace Ticket 等）。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：schoenbrunn.at（宮殿導覽各展室介紹、語音導覽文字稿、禮拜堂說明、19 世紀建築史）、austria.info（美泉宮歷史）、habsburger.net（法蘭茲・約瑟夫與美泉宮）、Google Arts & Culture（巴托尼肖像、壁毯椅）、Wikipedia（Schönbrunn Palace、Gloriette）、austria-forum.org（Paul Troger）、oeaw.ac.at（Elfriede Iby：百萬廳）、sisimuseum-hofburg.at（拉德茨基）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・美泉宮（室內）　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_美泉宮（室內）.docx', b));
