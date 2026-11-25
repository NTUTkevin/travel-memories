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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '維也納家具博物館：皇帝的家具倉庫與《茜茜公主》的片場', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Möbelmuseum Wien（舊名 Hofmobiliendepot）｜2026 年 11 月 24 日', font: F, size: 21, color: GRAY })] }),

  p('上午在霍夫堡看完皇帝和西西住過的房間，下午我們來到瑪麗亞希爾夫大街旁邊的維也納家具博物館（Möbelmuseum Wien，以前叫「宮廷家具倉庫」Hofmobiliendepot）。這裡收藏的，就是那些宮殿裡的家具本身，還有一個特別的展區：1950 年代《茜茜公主》電影裡用過的真家具。'),

  h1('一、從宮廷家具倉庫到家具博物館'),
  p('1747 年，瑪麗亞・特蕾莎女皇設立「宮廷家具管理處」（Hofmobilieninspektion），學的是法國宮廷的家具庫，負責清點、維修和搬運宮裡的家具。那時候的宮殿平常幾乎是空的：皇室要搬到哪座宮殿住，先行官就先出發準備，家具也跟著一路搬過去，所以家具常常磨損，甚至搬丟。一直到十九世紀，宮殿才慢慢有固定擺放的家具。'),
  p('1901 年，法蘭茲・約瑟夫皇帝在霍夫堡和美泉宮中間的瑪麗亞希爾夫大街 88 號，蓋了一棟中央倉庫，裡面還有工坊和馬車房。帝國結束後，1919 年它歸奧地利共和國所有，1924 年開放了第一批展覽，當時的 15 間「畢德麥雅小展間」到今天還保留著。1998 年整修後重新開幕，現在館內展出約 6,500 件，整個聯邦家具管理處的收藏超過 17 萬件，是世界最大的家具收藏之一。有趣的是，這些家具到現在還在「上班」：國家的政府機關、駐外使館和國宴，還會借用倉庫裡的家具。'),
  ...img('mb_salon', 1100, 825, '白底描金的洛可可式扶手椅、椅子和桌子，配紅色錦緞椅墊，擺在大窗前（是否就是電影裡用過的那一組，待確認）'),
  verse('房屋因智慧建造，又因聰明立穩；其中因知識充滿各樣美好寶貴的財物。', '箴言 24:3–4'),
  heart('原來皇帝的家具也要「搬家」，而且一搬就是幾百年。宮殿再大，真正讓它像個家的，還是裡面的桌椅和生活。'),

  h1('二、《茜茜公主》電影裡的真家具'),
  p('這間博物館和電影很有緣。一直到 1970 年代，在維也納拍的電影幾乎都向這裡借家具，其中最有名的，就是導演恩斯特・馬里施卡（Ernst Marischka）1950 年代拍的《茜茜公主》三部曲，由羅密・施奈德（Romy Schneider）飾演西西。拍室內戲時，美泉宮和霍夫堡都不能借來拍，所以劇組大多在維也納的玫瑰丘片廠（Rosenhügel-Studios）搭景，再擺上真正的皇家家具，讓畫面看起來像真的宮殿。'),
  p('博物館沿著一條「西西小路」（Si(s)i-Pfad）展出電影片段、道具、劇照和海報，把「電影裡的茜茜」和「歷史上真正的伊莉莎白」放在一起對照。上午在西西博物館才看到大家心中的「茜茜公主」很多是電影塑造出來的，下午就在這裡看到電影是怎麼拍出來的，剛好接起來。'),
  p('展櫃裡還有幾頂鍍金、鑲滿珍珠的皇冠，下排那頂樣子很像奧地利帝國皇冠（魯道夫皇冠）。真正的帝國皇冠收藏在霍夫堡的寶物館，所以這幾頂應該是複製品，可能是電影或典禮用的道具（待確認）。'),
  ...img('mb_crowns', 825, 1100, '展櫃裡的皇冠與紅絲絨墊子；下排是類似魯道夫皇冠的那一頂（是否為電影道具，待確認）'),
  verse('原來我們不是顧念所見的，乃是顧念所不見的；因為所見的是暫時的，所不見的是永遠的。', '哥林多後書 4:18'),
  heart('電影裡的皇冠是假的，椅子卻是真的。真真假假之間，大家記住的是一個美麗的故事。'),

  h1('三、皇帝的寶座'),
  p('博物館裡有一組「法蘭茲・約瑟夫皇帝的寶座布置」，年代是十九世紀下半。紅色絲絨的帷幔和華蓋下面，中間是一張鍍金雕花的寶座，椅背頂端有一頂小皇冠，扶手下有獅身的雕刻；兩側各一張金色錦緞扶手椅，牆上掛著繡有雙頭鷹的帝國徽章，前面還有一個繡著皇冠的跪墊。'),
  p('一位寫維也納的部落客形容，親眼看到皇帝的寶座時，其實沒有想像中那麼華麗。寶座看起來威風，其實也是一件會被搬來搬去、會舊、會需要修理的家具。'),
  ...img('mb_throne', 1100, 825, '法蘭茲・約瑟夫的寶座布置：紅色帷幔、鍍金寶座、兩張扶手椅與雙頭鷹徽章'),
  verse('神啊，你的寶座是永永遠遠的；你的國權是正直的。', '詩篇 45:6'),
  heart('地上的寶座會舊，帷幔會褪色，帝國也會結束。唯有神的寶座是永遠的。'),

  h1('四、哈布斯堡家族的房間'),
  p('博物館有幾間依人物布置的展間，從瑪麗亞・特蕾莎和她的時代，到法蘭茲二世／一世，再到法蘭茲・約瑟夫和他的家人。我們拍到的一間紅牆展間，中間掛著一幅穿紅色禮服、披斗篷的皇帝全身像，旁邊有鍍金半身像、一整排圓框的家族小肖像、捲蓋書桌和鑲嵌大理石紋的桌子；左邊架上是一排表情誇張的頭像，很像雕刻家梅瑟施密特（Franz Xaver Messerschmidt）有名的「性格頭像」（這幾件與人物是誰，都待確認）。'),
  p('另一個玻璃櫃裡有一幅全家福：父親拿著花，母親坐在旁邊，四個穿著小禮服的孩子在中間玩。還有一間掛著年輕王子穿紅色金邊外套的肖像，右邊牆上是中國風的人物壁紙，前面一排白底描金的椅子，椅背都是手工刺繡的花籃和花環，旁邊一個鼓肚形五斗櫃，用不同顏色的木片拼出花紋和中國風的圖案。十八世紀的歐洲宮廷很流行這種「中國風」。'),
  ...img('mb_joseph', 825, 1100, '紅牆展間：皇帝全身像、圓框家族肖像、捲蓋書桌與鑲嵌桌，左邊是一排性格頭像（人物待確認）'),
  ...img('mb_family', 825, 1100, '玻璃櫃裡的全家福：父親、母親和四個孩子（是哪一家人，待確認）'),
  ...img('mb_chairs', 825, 1100, '刺繡椅背的白金色椅子、拼木五斗櫃，牆上是穿紅外套的少年肖像和中國風壁紙'),
  verse('求你指教我們怎樣數算自己的日子，好叫我們得著智慧的心。', '詩篇 90:12'),
  heart('畫裡的人都已經不在了，可是他們坐過的椅子、用過的書桌還在。家具好像比人更記得一個家的樣子。'),

  h1('五、皇太子魯道夫的搖籃'),
  p('博物館裡最讓人停下腳步的，是一張深色木頭雕成、鑲著鍍金玫瑰花環的大搖籃。博物館的介紹裡提到，這裡收藏著皇太子魯道夫的搖籃。魯道夫是法蘭茲・約瑟夫和西西唯一的兒子，1858 年出生；搖籃前端雕著一隻雙頭鷹，胸前的盾牌和下方的紅底白十字，就是奧地利的徽章，頂上還有一個吹號角的小天使。'),
  p('上午在西西博物館才讀到，這個孩子 1889 年在梅耶林去世，母親從此只穿黑衣。下午看到他的搖籃，心裡特別有感觸。另一件讓我們很在意的事：搖籃後面的牆上有一幅士兵舉槍射擊的畫，看起來像在畫 1867 年法蘭茲・約瑟夫的弟弟、墨西哥皇帝馬克西米連被槍決的場面（待確認）。'),
  ...img('mb_cradle', 1100, 825, '搖籃側面：絲綢內襯、木珠邊框和鍍金玫瑰花環，後面隱約是士兵開槍的畫'),
  ...img('mb_cradle2', 825, 1100, '搖籃正面：雙頭鷹、皇冠、吹號角的小天使，最下面是紅底白十字的奧地利徽章（是否魯道夫的搖籃，待確認）'),
  verse('我的肺腑是你所造的；我在母腹中，你已覆庇我。', '詩篇 139:13'),
  heart('不管是皇太子還是平凡人家的孩子，躺在搖籃裡的時候，都是父母最大的盼望。'),

  h1('六、畢德麥雅的家與虎皮地毯'),
  p('「畢德麥雅」（Biedermeier）指的大約是 1815 到 1848 年的那段時間，維也納的家具從宮廷的金碧輝煌，變成簡單、舒適、看得到木紋的市民風格。這間博物館被認為是全維也納看畢德麥雅家具最好的地方，1924 年就有的那些小展間，也是在展示這種居家生活。'),
  p('我們拍到的一間，牆上是小碎花壁紙，有花布扶手椅、拼木抽屜櫃和細腳的桌子，很溫馨；可是地上卻鋪著一整張張著嘴的老虎皮，讓人嚇一跳。這間是誰的房間、老虎皮從哪裡來，待確認。'),
  ...img('mb_tiger', 1100, 825, '碎花壁紙的房間：花布扶手椅、拼木抽屜櫃，地上是一張老虎皮地毯'),
  verse('用繩量給我的地界，坐落在佳美之處；我的產業實在美好。', '詩篇 16:6'),
  heart('看了一整天的宮殿和寶座，最想住的反而是這種小小的、有木頭香味的房間。'),

  h1('旅人小記'),
  bullet('時間：11/24 下午，照片檔名時間是 14:29–15:19，在上午霍夫堡之後。'),
  bullet('地址與交通：入口在 Andreasgasse 7（1070 第七區），地鐵 U3 Zieglergasse 站最近，就在瑪麗亞希爾夫購物大街旁邊。'),
  bullet('開放時間：週二到週日與國定假日 10:00–17:00，週一休館（官網）。有「Sisi Ticket」套票可以一起參觀美泉宮、西西博物館和家具博物館。'),
  bullet('沒放進文章的照片：寶座的另一個角度（雲端照片 05）。'),
  bullet('待確認：我們買的是哪一種票、有沒有參加「茜茜電影」導覽；白金色洛可可椅子那一組是不是電影裡用過的；皇冠是不是電影道具；紅牆展間的皇帝全身像與頭像是誰；全家福是哪一家人；穿紅外套的少年是誰；搖籃是不是魯道夫的、後面那幅畫是不是馬克西米連被槍決；老虎皮房間是誰的。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：moebelmuseumwien.at／hofmobiliendepot.at（官網、開放時間、展間介紹）、schoenbrunn-group.com（家具博物館新聞資料）、bmwet.gv.at（聯邦家具管理處歷史）、habsburger.net（皇帝的家具、魯道夫的生平）、wien.info（Sissi im Film、家具博物館、寶座布置）、meinbezirk.at（魯道夫的搖籃、西西小路）、visitingvienna.com（寶座與畢德麥雅家具）、Wikipedia（Imperial Furniture Collection）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・家具博物館　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_家具博物館.docx', b));
