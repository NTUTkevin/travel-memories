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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '匈牙利國家歌劇院：在伊柏的金色殿堂看《卡門》', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '匈牙利・布達佩斯・Magyar Állami Operaház｜2026 年 11 月 26 日', font: F, size: 21, color: GRAY })] }),

  p('坐千禧地鐵到歌劇院站，一出站就是它。11 月 26 日晚上，我們在匈牙利國家歌劇院看了比才的歌劇《卡門》。17:26 到門口，先在大廳和觀眾廳裡到處拍照，18:00 開演，散場出來已經是 21:19。三天前在維也納的金色大廳聽了莫札特，這一晚在布達佩斯看歌劇，音樂一路跟著我們走。'),

  h1('一、伊柏的新文藝復興宮殿'),
  p('布達佩斯蓋歌劇院的競圖從 1873 年開始。主辦單位邀請了六位建築師參加：伊柏・米克洛什（Ybl Miklós）、施泰因德爾・伊姆雷（Steindl Imre，後來設計國會大廈）等四位匈牙利人，還有兩位外國建築師；評審壓倒性地選了伊柏的方案。1875 年開工，九年後的 1884 年 9 月 27 日，在法蘭茲・約瑟夫皇帝出席下盛大開幕。當時叫「匈牙利皇家歌劇院」，1945 年後才改成現在的名字。'),
  p('外觀是對稱的新文藝復興風格，裝飾全都圍繞音樂。正門兩側的壁龕裡，坐著兩位匈牙利作曲家的石灰岩雕像：李斯特（Liszt Ferenc）和艾爾凱爾（Erkel Ferenc）。艾爾凱爾是匈牙利國歌的作曲者，也是這座歌劇院的第一任音樂總監。兩座像都是當年才 26 歲的雕刻家史特羅布・阿洛伊斯（Stróbl Alajos）做的，據說兩位作曲家都親自讓他寫生：李斯特的姿態比較浪漫、有動感，艾爾凱爾則顯得沉靜。最上面的主簷上，是一排歐洲作曲家的石像；側門旁還有兩隻白色卡拉拉大理石的人面獅身像，也是史特羅布的作品。'),
  p('建造時有一個條件：盡量用匈牙利的工匠和材料。只有少數例外，例如卡拉拉的大理石、奧地利的花崗岩柱子、維也納的舞台設備，以及美因茲做的大吊燈。'),
  ...img('op_ext', 750, 1000, '17:26，開演前的正門：兩層拱窗都亮著燈，主簷上站著一排作曲家石像；左邊路燈上掛著豎琴形狀的節慶燈飾（待確認），右下角是黃色計程車和準備進場的觀眾'),
  ...img('op_night', 750, 1000, '21:19，看完《卡門》出來，從門口往上看：古典路燈、拱廊、上層的格子拱頂和主簷上的石像，空中好像飄著細雪（待確認）'),
  verse('智慧建造房屋，鑿成七根柱子。', '箴言 9:1'),
  heart('一座房子，從外面的雕像就開始講音樂的故事。走進去之前先抬頭看看那些作曲家，好像他們也在門口歡迎每一個來聽音樂的人。'),

  h1('二、走進大廳'),
  p('進門後是有大理石柱和雙向大樓梯的門廳，天花板和牆上滿是壁畫與鍍金裝飾。歌劇院內外的繪畫和雕塑，大多出自當時最好的匈牙利藝術家，例如畫家洛茨・卡羅伊（Lotz Károly）、塞凱伊・貝爾塔蘭（Székely Bertalan）、坦・莫爾（Than Mór），以及雕刻家史特羅布。'),
  p('我們在 17:29 拍到一處拱頂：紅金兩色的格子天花板，中間畫著穿紅衣的女神，四周是抱在一起的小天使，正中央垂下一盞金色吊燈。（這是門廳、樓梯間還是哪一處走廊，畫的是誰、出自哪位畫家，待確認。）'),
  ...img('op_lobby', 750, 1000, '17:29，大廳裡的拱頂：紅金格子天花板上的人物壁畫、小天使，和一盞點著蠟燭形燈泡的金色吊燈'),
  verse('當稱謝進入他的門；當讚美進入他的院。', '詩篇 100:4'),
  heart('走進這樣的大廳，自然就把腳步放輕、把聲音放小。好的房子會教人怎麼走進去。'),

  h1('三、馬蹄形觀眾廳與〈音樂的神化〉'),
  p('觀眾廳是馬蹄形的，三層包廂一圈圍住舞台，最上面還有頂層座位。整個空間只有紅和金兩種顏色。官方介紹說，觀眾廳裡所有的畫都是洛茨・卡羅伊畫的；圓形天花板上是他的代表作〈音樂的神化〉（The Apotheosis of Music），周長約 45 公尺：畫面中心是阿波羅在演奏，奧林帕斯山的眾神、美惠三女神和繆思都在聽。'),
  p('天花板正中央垂下的青銅大吊燈，也是伊柏自己設計、在德國美因茲製造的。它原本點煤氣，後來改成電燈，重將近 3 噸。每年兩次，工作人員會用絞盤把它降到一樓，更換燒壞的燈泡。舞台上方的紅色布幔中間，掛著匈牙利的國徽。'),
  p('歌劇院 1980 到 1984 年整修過一次，2017 年又關閉，大規模翻新和現代化，將近五年後，在 2022 年 3 月 12 日重新開幕。我們看到的金色，是剛整修好幾年的樣子。影片 v21（17:51）從天花板一路拍到舞台和包廂：樂池裡的樂手已經在暖身，觀眾陸續入座。'),
  ...img('op_ceiling', 750, 1000, '17:41，抬頭看〈音樂的神化〉：圓形天花板上的眾神與繆思，中間金色的圓環垂下三層大吊燈'),
  ...img('op_hall', 1000, 750, '17:42，從高處看舞台和右側包廂：掛著國徽的紅色布幔、紅色大理石柱、一層層點著燭燈的包廂'),
  ...img('op_pit', 750, 1000, '開演前，天花板、大吊燈、舞台和三層包廂一次入鏡；最下面的樂池裡，樂手已經就位'),
  verse('諸天述說神的榮耀；穹蒼傳揚他的手段。', '詩篇 19:1'),
  heart('畫家把天上畫成一場音樂會，所有的神都在安靜地聽。坐在下面的我們，好像也被邀請加入這場聆聽。'),

  h1('四、在頂層看《卡門》'),
  p('那晚的演出是法國作曲家比才（Georges Bizet）的《卡門》，以法文演唱，有匈牙利文、英文、法文字幕，從 18:00 演到 21:15，中間有休息。這個版本是加泰隆尼亞導演比埃托（Calixto Bieito）的製作：他把故事從 19 世紀的塞維亞，搬到佛朗哥政權結束後的西班牙，用寫實、不走民俗風的方式演出。這個製作在世界各地巡演過，2021 年來到匈牙利。'),
  p('故事從塞維亞的廣場開始：士兵唐・荷西被吉普賽女郎卡門吸引，為她放棄了軍職和家鄉的未婚妻米凱拉，最後卻因為卡門愛上鬥牛士埃斯卡米洛，在鬥牛場外由愛生恨。〈哈巴奈拉〉（愛情是一隻不羈的鳥）和〈鬥牛士之歌〉都是大家耳熟能詳的旋律。'),
  p('我們的座位在很高的樓層，往下看可以同時看到整個舞台、樂池和一層層的包廂。開演前，太太站在金色欄杆邊拍了好幾張照片，我們也在天花板和吊燈下自拍了一張。'),
  ...img('op_wife_rail', 750, 1000, '太太戴著黑色報童帽，扶著頂層的金色欄杆；左邊是彎彎的三層包廂，前面是舞台和樂池'),
  ...img('op_wife_chand', 750, 1000, '太太站在大吊燈和〈音樂的神化〉下面，背後是掛著國徽的紅色布幔'),
  ...img('op_selfie', 750, 1000, '開演前的兩人自拍：背後是舞台、吊燈和坐滿觀眾的包廂'),
  verse('愛情，眾水不能熄滅，大水也不能淹沒。', '雅歌 8:7'),
  heart('《卡門》講的是一段燒得太猛、最後毀了兩個人的愛情。看完走出來，在冬夜裡牽著太太的手，覺得細水長流的愛比較好。'),

  h1('旅人小記'),
  bullet('時間（依檔名與影片 metadata）：11 月 26 日（週四）17:16 在沃勒什馬爾蒂廣場搭千禧地鐵 → 17:26 到歌劇院正門 → 17:29 大廳 → 17:41–17:51 觀眾廳（影片 v21）→ 18:00《卡門》開演 → 21:15 散場 → 21:19 門口夜景。'),
  bullet('交通：M1 黃線歌劇院站（Opera）就在對面；地址 Andrássy út 22。'),
  bullet('2026/27 樂季的《卡門》在 11 月 14–29 日演出；字幕有匈牙利文、英文、法文。沒有看演出的話，歌劇院也有導覽。'),
  bullet('待確認：座位在哪一層、票價多少；中間休息幾次；那晚卡門、唐・荷西是誰唱的；我們喜不喜歡這個現代版的製作；17:29 那張拱頂照片是在門廳、樓梯間還是哪裡；路燈上的節慶燈飾是不是豎琴的造型；散場時是不是真的在下雪。'),
  bullet('附記：雲端照片資料夾另有一張 17:26 的正門照片（01，和 02 幾乎同一角度，所以沒放進文章）。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：opera.hu（歌劇院興建、觀眾廳、石灰岩立面、人面獅身像台階、2022 年重新開幕、2026 年 11 月 26 日《卡門》節目頁）、Google Arts & Culture（洛茨天花板壁畫、大吊燈）、Hungaricana《Andrássy Avenue》（Gábor Eszter，觀眾廳）、Wikipedia（匈牙利國家歌劇院、匈牙利國家歌劇團、史特羅布・阿洛伊斯）、Daily News Hungary（材料來源）、welovebudapest（立面與競圖）、Bachtrack、Jegymester（演出檔期）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・布達佩斯・匈牙利國家歌劇院　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('布達佩斯_匈牙利國家歌劇院.docx', b));
