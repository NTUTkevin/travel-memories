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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '哈修塔特小鎮：湖邊兩座教堂', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・上奧地利邦・哈修塔特 Hallstatt｜2026 年 11 月 20 日', font: F, size: 21, color: GRAY })] }),

  p('從鹽礦下山之後，我們在哈修塔特小鎮裡走了一個下午。天一直陰陰的，山頂蓋著薄雪，雲霧一下散開、一下又壓下來。小鎮夾在湖和陡峭的山壁中間，幾乎找不到平地，房子一層一層往山上疊。湖邊有一座細長尖塔的新教教堂，山坡上還有一座白色鐘樓的天主教堂，兩座教堂一高一低，是這個小鎮最有名的畫面。'),

  h1('一、湖邊：「鹽的地方」'),
  p('哈修塔特這個名字，當地的說法是來自凱爾特語的「hal」，意思就是鹽，也就是上一篇鹽山的那個「鹽」。小鎮一邊是湖、一邊是山，很長一段時間對外交通非常不方便：以前要來這裡，只能划船，或是走湖邊一條窄窄的小路。1875 年才有一條從戈薩米勒（Gosaumühle）通到哈修塔特的路，1877 年鐵路通車。'),
  p('現在小鎮裡沿著湖的那條「湖濱路」（Seestraße）是 1890 年蓋的，為了開這條路，拆掉了七棟有一部分立在水中木樁上的房子，還有一座木橋。後來遊客越來越多，鎮上停車變成大問題，1966 年開通了一條繞過小鎮的隧道，隧道口就做了兩層停車場；之後小鎮中心白天不讓車子開進去，連住在鎮上的旅館客人也要把車停在外面。'),
  ...img('t_lake', 1100, 825, '湖邊看小鎮：山坡上紅褐、金黃的秋樹，下面是一排沿著湖的木造房子和船屋，遠處是新教教堂的尖塔；水面上有一隻白天鵝，前景是兩根長滿青苔的老木樁'),
  ...img('t_duck', 825, 1100, '湖邊的綠頭鴨一點都不怕人，站在步道上看著我們；背後是倒映在湖裡的小鎮，木樁上還停著一隻海鷗'),
  ...img('t_couple1', 825, 1100, '在湖邊合照：太太戴著黑色報童帽、圍紅圍巾，我戴藍色毛帽；身後就是小鎮和湖'),
  verse('你們要休息，要知道我是神！', '詩篇 46:10'),
  heart('鹽礦裡是黑的、吵的，到了湖邊忽然安靜下來。湖面像鏡子一樣，連鴨子都慢慢走。以前的人只能划船進來，我們走著走著，好像也慢慢被這個節奏帶著放慢了。'),

  h1('二、小巷、瀑布與 1750 年的大火'),
  p('小鎮的市集廣場從 14 世紀就有了，周圍大部分的房子是 16 世紀蓋的；廣場上的聖三一紀念柱，是一位鹽商在 1743 年捐建的。1750 年 9 月 20 日，一場大火從一家叫「Keferbäck」的麵包店燒起來，燒毀了鎮中心大半：四個人喪生，三十五棟房子和官府裡所有的文件都被燒掉了，連煮鹽的工坊、醫院和小教堂也沒有倖免。今天看到的很多房子，是大火之後重建的。'),
  p('小鎮後面的山壁上有一道瀑布，叫做米爾巴赫瀑布（Mühlbach，「磨坊溪」）。這裡的氣候很難種穀物，以前磨穀子的磨坊，就有一些蓋在這道瀑布上方。我們在瀑布下面的一棟房子牆上，看到畫著一個拿著東西的人形，旁邊是兩個連在一起的螺旋圖案。這個圖案很像哈修塔特文化時期常見的「眼鏡形胸針」（Brillenfibel）——一種用一整條青銅線繞成兩個螺旋的別針，是史前時代用來別住衣服的。牆上的圖是不是在畫這個，待確認。'),
  ...img('t_pink', 825, 1100, '一棟粉紅色牆面的老房子：深綠色的百葉窗半開著，屋簷是深色的木頭；牆前有一棵貼著牆長的老果樹，枝幹彎彎曲曲，冬天葉子都掉光了（樹種待確認）'),
  ...img('t_fall', 825, 1100, '小鎮後面的米爾巴赫瀑布，從紅葉滿山的岩壁上流下來；前面房子的牆上畫著一個人形和兩個連在一起的螺旋圖案（是否為史前胸針圖案，待確認）'),
  verse('你從水中經過，我必與你同在；你蹚過江河，水必不漫過你；你從火中行過，必不被燒，火焰也不著在你身上。', '以賽亞書 43:2'),
  heart('一場大火燒掉了半個小鎮，他們又一棟一棟蓋回來。今天我們拍的這些漂亮房子，背後都有人重新站起來的故事。'),

  h1('三、湖邊的尖塔：新教基督教堂'),
  p('湖邊那座細長尖塔的教堂，是新教（路德宗）的基督教堂（Evangelische Christuskirche）。16 世紀，路德的宗教改革在鹽湖區傳得很快，特別是在礦工之間。後來反宗教改革開始，1601 年當地人起來反抗，拆橋、把煮鹽的鍋倒空、擋住運木頭的船，最後被薩爾斯堡大主教的軍隊鎮壓，帶頭的人被判死刑。'),
  p('1734 年，軍隊又來了：哈修塔特、巴特伊舍和戈伊瑟恩一帶三百多位新教徒（還不算婦女和孩子），趁著夜裡被裝上運鹽的船送走，被迫移民到遙遠的外西凡尼亞（今天的羅馬尼亞）。這和前一天在薩爾斯堡要塞讀到的 1731 年驅逐新教徒，是同一段歷史的另一頁。一直到 1781 年約瑟夫二世頒布《寬容詔令》，新教徒才能有限度地聚會，那時哈修塔特有五百人承認自己是新教徒；1785 年 10 月 30 日他們蓋了一間禱告所。後來法蘭茲・約瑟夫皇帝的「新教徒特許令」讓新教和天主教完全平等，新的教堂才開始動工，1863 年 10 月、只花了五年就蓋好，就是我們看到的這座新哥德式教堂。'),
  ...img('t_church', 825, 1100, '從巷子裡抬頭看新教教堂：灰白色的石牆、紅底的鐘面、高高的木瓦尖塔，頂上是金色的十字架；左邊是有條紋遮陽篷的 Derbl 小店，右邊掛著「Grüner Baum」的招牌'),
  verse('基督釋放了我們，叫我們得以自由，所以要站立得穩，不要再被奴僕的軛挾制。', '加拉太書 5:1'),
  heart('明信片上最美的那座尖塔，背後是一群為了信仰被迫半夜上船、離開家鄉的人。能夠自由地敬拜，原來是很多人用一輩子換來的。'),

  h1('四、山坡上的天主教堂與人骨堂'),
  p('照片右邊山坡上那座有白色鐘樓的，是天主教的「山上的聖母堂」（Maria am Berg）。這裡大約在西元 1050 年就有基督徒的團體，1181 年以前應該已經有一座小教堂，那座厚實的鐘樓就是 12 世紀留下來的；後來在陡峭的岩石上擴建，1505 年完成現在的晚期哥德式教堂。教堂裡最有名的是雷翁哈德・阿斯特爾（Leonhard Astl）大約 1515 到 1520 年做的聖母翼式祭壇，也是他唯一簽了名的作品。1987 年，一對義大利夫婦從教堂的小聖母祭壇上撬走了四幅哥德式畫板，成了一樁追查了三十多年的藝術竊案。'),
  p('教堂旁邊的聖米迦勒小禮拜堂，就是有名的人骨堂（Beinhaus）。因為山坡上墓地太小，過了一段時間就要把遺骨挖出來，放進人骨堂，讓出位置給新的人。家人會在頭骨上畫花冠、寫上名字和過世的年份，就像在墳上放花一樣；這個習俗大約從 18 世紀開始（當地資料寫 1720 年）。這裡的頭骨有一千多個（各來源從約 1,200 到約 2,000 個不等），其中六百多個有彩繪。最後一個放進來的頭骨是 1995 年，一位 1983 年過世的女士，她生前的心願就是要放進人骨堂。我們這次有沒有上去參觀，待確認。'),
  verse('死啊，你得勝的權勢在哪裡？死啊，你的毒鉤在哪裡？', '哥林多前書 15:55'),
  heart('在頭骨上畫花、寫名字，聽起來有點嚇人，可是想一想，那是一家人捨不得忘記一個人的方法。'),

  h1('五、明信片的角度'),
  p('從小鎮北邊的湖濱路往回看，湖、尖塔、山坡上的天主教堂和後面的雪山連在一起，就是哈修塔特最經典的畫面。1997 年，哈修塔特和周圍的達赫施泰因山、鹽湖區一起被列為聯合國世界遺產；但也因為太有名，小鎮常被拿來當作「過度觀光」的例子。這個畫面還紅到被「複製」：中國五礦集團在廣東惠州博羅縣蓋了一整個仿造的哈修塔特住宅區，2012 年 6 月開幕，連教堂鐘樓都照著做；哈修塔特的鎮長事後還親自去參加了開幕。'),
  p('我們在這裡和同行的朋友們拍了合照，包括從新加坡來的蘇西。往回走的時候，霧又從山上壓下來，連山頂都看不見了。'),
  ...img('t_view', 1100, 825, '經典的明信片角度：湖邊的新教教堂尖塔、右邊山坡上有白色鐘樓的天主教堂，後面是頂著薄雪的山；湖上停著一艘白色的遊船'),
  ...img('t_group', 619, 1100, '在明信片角度前合照：前排是宜恩（米白色外套，待確認）和我，後排是戴藍色太陽眼鏡、來自新加坡的蘇西，和太太'),
  ...img('t_couple2', 1100, 619, '同一個地方，我們兩個人：太太抬頭看著我笑，背後一艘遊船正開過湖面'),
  ...img('t_mist', 1100, 825, '離開前，雲霧把山頂整個蓋住，只剩下山腳的木屋和秋天的樹林'),
  verse('看哪，弟兄和睦同居是何等地善，何等地美！', '詩篇 133:1'),
  heart('同一個畫面，全世界有幾百萬人拍過。可是這一張裡面有我們，有一起走了好幾天的朋友，它就只屬於我們。'),

  h1('旅人小記'),
  bullet('交通：從鹽礦纜車山下站走到小鎮中心只要幾分鐘。搭火車的話，哈修塔特火車站在湖的對岸，要再搭渡船過來。小鎮中心白天禁止車輛進入，開車要停在隧道口的停車場。'),
  bullet('明信片角度在小鎮北邊的湖濱路上，人很多，請不要擋到路，也不要闖進私人住宅拍照。這裡是有人住的地方，請放輕聲音。'),
  bullet('天主教堂和人骨堂在山坡上，要爬不少階梯；新教教堂就在市集廣場旁邊，官方資料寫每天 10 點到 18 點開放（以現場公告為準）。'),
  bullet('這一篇還有幾處待確認：牆上的人形和螺旋圖案是不是史前胸針、粉紅房子前那棵樹的樹種、我們有沒有進教堂或人骨堂、米白外套的年輕人是不是宜恩。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：hallstatt.net（主題步道：湖與鹽、湖與其他通道、第一批遊客、哈修塔特大火、市集與鹽商、瀑布磨坊；新教教堂；天主教堂；人骨堂；哥德大師）、dachstein.salzkammergut.at／salzkammergut.at（新教基督教堂、天主教堂）、林茨教區網站 dioezese-linz.at（Astl 聖母祭壇、人骨堂）、mein-oesterreich.info（隧道）、紐約大都會博物館（眼鏡形胸針）、Wikipedia（Hallstatt、Hallstatt (China)）、China Daily（2012 年報導）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・哈修塔特小鎮　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('哈修塔特小鎮.docx', b));
