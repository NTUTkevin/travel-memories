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

const body = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '走進莫札特的家（上）', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡｜莫札特故居 Mozart-Wohnhaus｜2026 年 11 月 21 日', font: F, size: 21, color: GRAY })] }),

  p('十一月的薩爾斯堡早上，我和太太走過薩爾察赫河，來到馬卡特廣場 8 號。這裡不是莫札特出生的地方（出生地在對岸的糧食胡同 9 號），而是 1773 年全家搬來的「舞蹈教師之家」——一層有八個房間的寬敞公寓。莫札特從 17 歲住到離開薩爾斯堡，父親利奧波德則一直住到 1787 年過世。館內燈光昏黃，一件件手稿、畫作和樂器，像在低聲訴說這個家的故事。'),
  p('莫札特出生隔天，就在薩爾斯堡主座教堂受洗，洗禮名裡有一個「Theophilus」，希臘文的意思是「被神所愛的」，拉丁文寫作「Amadeus」——後來全世界都用這個名字記得他。帶著這個名字走進他的家，感覺格外不同。'),
  verse('神造萬物，各按其時成為美好，又將永生安置在世人心裡。', '傳道書 3:11'),

  h1('一、先看位置：廣場模型與聖三一教堂'),
  p('展廳裡有一座老城模型，紅色小箭頭標著「Mozarts Wohnhaus」和「Theater」。中間那座綠色圓頂的巴洛克教堂是聖三一教堂（Dreifaltigkeitskirche），1694 年由巴洛克大師費舍爾・馮・埃爾拉赫設計。這片廣場莫札特那時叫「漢尼拔廣場」，故居正對面就是今天的州立劇院和米拉貝爾花園大門。'),
  p('也就是說，莫札特家門一開，抬頭就是一座獻給「聖父、聖子、聖靈」的教堂。今天看到的故居其實是重建的：原建築在二戰後拆除，後來依照舊圖重建，1996 年重新開幕。'),
  p('這棟房子的歷史可以追溯到 1617 年，1685 年以前其實是兩棟。1711 年起，宮廷舞蹈教師羅倫茲・史佩克納在這裡教貴族子弟跳舞——當時的舞蹈學校不只教跳舞，也教上流社會的禮儀，讓年輕貴族準備進入宮廷生活，「舞蹈教師之家」的名字由此而來。巧的是，他的兒子、接任宮廷舞蹈教師的法蘭茲・史佩克納，正是 1747 年莫札特父母結婚時的證婚人。'),
  p('1773 年，史佩克納過世六年後，莫札特一家搬進二樓這間八個房間的大公寓，包括原本用來上舞蹈課的大廳。莫札特從 17 歲住到 25 歲；父親則一直住到 1787 年 5 月 28 日過世，晚年還在這裡照顧南妮爾的長子。'),
  p('1944 年 10 月 16 日，一顆空襲炸彈摧毀了建築的三分之二。莫札特基金會 1955 年買下倖存的部分，1989 年又買下戰後蓋在廢墟上的辦公大樓，拆除後依照舊圖重建，1996 年 1 月 26 日——莫札特生日前一天——重新開幕。'),
  ...img('model', ...P, '館內的馬卡特廣場模型，紅箭頭標示莫札特故居與劇院'),
  heart('看著這座模型，我們想像兩百五十年前的莫札特，每天出門就看見教堂的圓頂、聽見鐘聲。信仰不是掛在牆上的裝飾，而是生活的地景。'),

  h1('二、莫札特全家福：牆上的母親'),
  p('這幅大油畫一般認為出自薩爾斯堡畫家德拉・克羅切之手，繪於 1780 年底到 1781 年初。畫中莫札特和姊姊南妮爾四手聯彈，父親利奧波德倚著琴、手持小提琴；背後橢圓畫框裡的，是 1778 年陪兒子旅行時在巴黎病逝的母親安娜・瑪麗亞。一家人用「畫中畫」的方式，讓她依然在場。'),
  p('莫札特在母親臨終時寫信回家，信裡一再提到要把一切交託在神的旨意裡。南妮爾日後也說，這幅畫裡的弟弟是最像本人的一張。'),
  p('這幅畫的背景其實很沉重。1777–1779 年，莫札特由母親陪同，到曼海姆和巴黎找工作，結果一無所獲；母親在巴黎病逝，他只能獨自回到薩爾斯堡，繼續為他不喜歡的科洛雷多大主教工作。所以畫中這一家三口加上牆上的母親，是一個剛經歷失去、彼此依靠的家庭。'),
  ...img('family', ...L, '莫札特全家福，約 1780–1781 年'),
  verse('我們若信耶穌死而復活了，那已經在耶穌裡睡了的人，神也必將他與耶穌一同帶來。', '帖撒羅尼迦前書 4:14'),
  heart('站在這幅畫前，太太輕聲說：「媽媽雖然不在了，可是她還在這個家裡。」我們都想起了自己已經離開的家人。在基督裡，離別不是終點，這幅畫好像也在這樣盼望著。'),

  h2('姊姊南妮爾'),
  p('另一面牆上掛著一位頂著高聳髮型、胸前別著粉紅緞帶的年輕女子。這幅畫從 19 世紀起就被視為南妮爾的肖像，但館方也坦白，畫中人身分並不確定，所以稱它為「據稱」的南妮爾肖像，2017 年購入後才在故居展出。南妮爾自己也是出色的鍵盤演奏家，小時候和弟弟一起巡迴歐洲，長大後在薩爾斯堡教貴族千金彈琴。'),
  ...img('nannerl', ...P, '據稱為南妮爾・莫札特的肖像'),
  heart('南妮爾的才華不輸弟弟，卻因為時代而留在幕後。我們很喜歡她畫裡帶點俏皮的眼神——默默陪伴、成全家人的人，在神眼中一樣是寶貴的。'),

  h1('三、一封寫給父親的信（1780 年 11 月 15 日）'),
  p('展品 40 是莫札特的親筆信。1780 年秋天，巴伐利亞選帝侯委託他寫正歌劇《伊多梅紐斯，克里特之王》（K. 366），他 11 月初到慕尼黑和歌手一起排練，並寫信向父親報告：'),
  bullet('飾演國王的男高音拉夫年紀大了，想在第三幕最後唱一首好聽的詠嘆調取代四重唱。莫札特說他有道理，就算沒道理，也該「給他的白髮一點面子」。'),
  bullet('對飾演王子的閹人歌手達爾・普拉托就沒那麼客氣，諷刺地稱他「我最親愛的閹人歌手」，抱怨得把整部歌劇一句句教給他。'),
  p('《伊多梅紐斯》的故事很像舊約亞伯拉罕獻以撒：國王在海上遇難時向海神許願，平安上岸後要獻上第一個遇見的人，沒想到遇見的竟是自己的兒子。最後一個從天而來的聲音解除了誓約，兒子得以存活。這部歌劇在 1781 年 1 月 29 日於慕尼黑首演。'),
  p('這份委託來得正是時候。巴伐利亞選帝侯卡爾・特奧多爾熱愛音樂，他從曼海姆帶來的宮廷樂團被公認是當時歐洲最好的樂團，英國音樂史家伯尼形容它是「一支由將軍組成的軍隊」。1778 年他繼承巴伐利亞、把宮廷遷到慕尼黑，1780 年委託莫札特為嘉年華季寫新歌劇；劇本由薩爾斯堡大主教的宮廷神父瓦雷斯科撰寫，這份委託可能是信中提到的老男高音拉夫幫忙爭取來的。'),
  p('首演在慕尼黑王宮裡華麗的屈維利埃劇院舉行，總彩排那天正好是莫札特 25 歲生日。排練時選帝侯對他說：誰能相信這麼小的腦袋裡藏著這麼偉大的東西？有趣的是，當時唯一的報紙評論大讚布景，卻連作曲者的名字都沒提。幾個月後，莫札特就和科洛雷多大主教決裂，離開薩爾斯堡到維也納當自由音樂家。《伊多梅紐斯》因此被視為他第一部成熟的偉大歌劇，也是他告別故鄉的轉捩點。'),
  ...img('letter', ...L, '莫札特 1780 年 11 月 15 日寫給父親的親筆信（展品 40）'),
  verse('亞伯拉罕給那地方起名叫耶和華以勒。', '創世記 22:14'),
  heart('讀著這封信，我們看到的不是遙遠的天才，而是一個跟父親報告工作、對同事又尊重又抱怨的年輕人。他寫父子之間的獻祭與赦免，同時在信裡和自己的父親說心事——這份父子情，讓我們很感動。'),

  h1('四、《魔笛》的源頭：一本 1789 年的童話集'),
  p('展櫃裡攤開的是維蘭德編的童話集《金尼斯坦》第三冊，1789 年在瑞士溫特圖爾出版，翻開的正是〈露露，或魔笛〉。展牌說明，這篇故事成了席卡內德寫《魔笛》劇本的直接藍本。查資料時發現一個細節：學界普遍認為〈露露〉其實是維蘭德的女婿利伯斯金所寫，只是收錄在維蘭德的集子裡。'),
  p('《魔笛》是莫札特過世那年的作品，整部戲在講年輕人經過火與水的試煉，從黑夜走向光明。'),
  ...img('book', ...L, '維蘭德《金尼斯坦》第三冊，翻開的是〈露露，或魔笛〉'),
  verse('我是世界的光。跟從我的，就不在黑暗裡走，必要得著生命的光。', '約翰福音 8:12'),
  heart('一本泛黃的童話書，最後變成全世界傳唱的歌劇。我們聊到，神也常常這樣使用平凡的東西，成就意想不到的美好。'),

];
const body2 = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '走進莫札特的家（下）', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡｜莫札特故居 Mozart-Wohnhaus｜2026 年 11 月 21 日', font: F, size: 21, color: GRAY })] }),
  h1('五、《魔笛》展間：羽毛與星空'),
  p('順著童話書往裡走，就進到整間獻給《魔笛》的展廳。1791 年 9 月 30 日，《魔笛》在維也納郊區的維登劇院首演，莫札特親自指揮，寫劇本的席卡內德自己演捕鳥人帕帕基諾。兩個多月後，莫札特就過世了。'),
  p('莫札特和席卡內德的友誼其實就是在薩爾斯堡開始的：1780 年 9、10 月，席卡內德的劇團在薩爾斯堡演出，兩人一見如故。1789 年席卡內德接手維也納郊區的維登劇院，1791 年初邀莫札特合寫一部給一般市民看的歌唱劇——這就是《魔笛》。兩人都是共濟會成員，整部戲以共濟會的象徵寫成一則追尋智慧與光明的寓言。'),
  p('傳說首演前幾天，莫札特還有重要段落沒寫完，席卡內德乾脆把他「關」在劇院後面的一間小木屋裡趕稿。這間「魔笛小屋」後來被搬到薩爾斯堡，今天仍保存在薩爾斯堡莫札特基金會大樓旁的花園裡。首演由莫札特親自指揮，夜后由他的大姨子約瑟法・霍弗演唱，前一天他才剛完成單簧管協奏曲。《魔笛》立刻成為他最成功的歌劇——而他在兩個多月後的 12 月 5 日過世。'),
  p('牆上一排小小的版畫，畫的是早期演出裡各個角色的造型；下面的玻璃櫃裡，有一張大字寫著「Zauberflöte」的老海報，旁邊攤著各種早期劇本，其中一本封面印著義大利文「Il flauto magico」——這部用德語寫給維也納平民看的歌唱劇，很快就被翻譯、傳唱到整個歐洲。'),
  ...img('wife', ...P, '太太在《魔笛》展間，身後是早期演出的角色版畫與老劇本'),
  ...img('costumes', ...P, '早期《魔笛》角色版畫與海報、劇本'),
  p('轉角的牆上則是一系列舞台設計稿，展題叫「羽毛與星空：帕帕基諾與夜后的百年變身」。說明牆提到，莫札特基金會從 19 世紀末開館以來，光是和《魔笛》有關的舞台與服裝設計就收藏了一千多件。一身羽毛的帕帕基諾和星空下的夜后，是整部戲最經典的兩個形象；夜后那首〈復仇的火焰在我心中燃燒〉唱到高音 F，是歌劇史上最高的音之一。'),
  ...img('stage', ...L, '「羽毛與星空」展區：歷代《魔笛》舞台設計稿'),
  p('《魔笛》裡有一首很溫柔的二重唱，是公主帕米娜和帕帕基諾一起唱的，歌頌夫妻之愛。歌詞的結尾說：丈夫與妻子、妻子與丈夫，一同靠近神聖（Mann und Weib, und Weib und Mann, reichen an die Gottheit an）。兩百多年前的這句歌詞，正好說出婚姻在信仰裡的樣子。'),
  verse('有人攻勝孤身一人，若有二人便能抵擋他；三股合成的繩子不容易折斷。', '傳道書 4:12'),
  heart('在這間展廳，我幫太太拍了一張照片。她戴著紅色貝雷帽，站在兩百多年前的劇本旁邊笑著。《魔笛》講的是兩個人一起經過火和水的試煉，最後手牽手走進光裡。我們想，婚姻也是這樣：不是沒有試煉，而是有神在中間，兩個人一起走過去。'),

  h1('六、莫札特的鋼琴、義大利之旅與少年的友誼'),
  p('這間灑著自然光的房間擺著一架古鋼琴（fortepiano），琴架上還放著溫溼度計，看得出被小心保存。莫札特本人擁有的那架鋼琴出自維也納名匠安東・瓦爾特之手，約 1782 年製造，是他維也納時期作曲、演出的主力樂器，如今由莫札特基金會收藏在薩爾斯堡。牆上那幅紅衣少年在鍵盤前的肖像，是 1770 年莫札特 14 歲隨父親巡演義大利時，在維洛納畫下的。'),
  p('那幅維洛納肖像也有故事：1770 年 1 月，父子倆第一次義大利之旅經過維洛納，威尼斯共和國的財務總監皮耶特羅・盧吉亞蒂熱情接待，並請畫家在 1 月 6、7 日兩天為 13 歲的莫札特畫像。畫家究竟是維洛納名家奇尼亞羅利，還是他的外甥達拉・羅薩，至今仍有爭議。畫中琴架上的樂譜，是一首只靠這幅畫才流傳下來的〈維洛納快板〉。原作今為私人收藏，館內展出的應是複製品（請以展牌為準）。'),
  ...img('piano', ...L, '古鋼琴與 14 歲莫札特的維洛納肖像'),
  p('館內也掛著一幅手工上色的義大利古地圖，標題寫著「Statuum totius Italiae」（義大利全境諸國圖），出自紐倫堡的霍曼地圖社。那時的義大利還是一塊塊不同顏色的小國：薩伏依、威尼斯、教皇國、托斯卡尼、那不勒斯王國……莫札特父子在 1769 到 1773 年間三度南下，從維洛納、米蘭一路走到羅馬和那不勒斯。在羅馬，14 歲的莫札特聽了兩次西斯汀禮拜堂的《求主垂憐》（Miserere），回去就憑記憶把整首譜寫了下來。'),
  ...img('map', ...P, '霍曼地圖社的義大利全境圖'),

  verse('你們要用角聲讚美他，鼓瑟彈琴讚美他！', '詩篇 150:3'),
  heart('我們也是一路旅行到這裡的人。看著這張地圖，想到一個少年坐著馬車翻山越嶺，把在教堂裡聽到的聖樂帶回家——旅行真的會改變一個人，也正在改變我們。'),

  h2('寫給好朋友的一封義大利文信（1770 年 9 月 10 日，波隆那）'),
  p('就在這趟義大利之旅裡，14 歲的莫札特在佛羅倫斯遇見了同年出生的英國少年小提琴家湯瑪斯・林利（Thomas Linley Jr.）。兩個神童一見如故，幾天裡天天一起演奏，分別時還哭了。展櫃裡這封信，就是莫札特從波隆那用義大利文寫給林利的回信。'),
  p('信裡他說，很遺憾這次沒辦法再回佛羅倫斯，但會盡一切可能，好有「擁抱我親愛朋友」的喜樂；父親也很想再見到佛羅倫斯的朋友們。信末他署名「Amadeo Wolfgango Mozart」，自稱是林利「最忠實的僕人與最親愛的朋友」，請他永遠保守這份友誼。'),
  p('林利比莫札特只小四個月，12 歲就到佛羅倫斯跟名師納爾迪尼學琴。英國音樂史家伯尼 1770 年記錄說，「小托馬索」和「小莫札特」是全義大利都在談論、最有前途的兩位天才。兩人分別前一天，林利還送給莫札特一首由女詩人科莉拉寫的十四行詩——信裡問候的「科莉拉女士」就是她。據說莫札特後來說過：林利是真正的天才，要是活著，必定是音樂界最耀眼的人物之一。林利 1778 年 8 月 5 日在英國格林索普城堡的湖上划船時溺水。'),
  p('展牌最後一句讓人心酸：他們之後再也沒有見過面。1778 年，林利在一次划船意外中溺水過世，年僅 22 歲。'),
  ...img('linley', ...L, '莫札特 1770 年從波隆那寫給林利的親筆信'),
  verse('約拿單的心與大衛的心相連；約拿單愛大衛，如同愛自己的性命。', '撒母耳記上 18:1'),
  heart('這封信的字跡好工整，看得出一個少年多珍惜這份友誼。我們想起那些在人生不同階段陪過我們、後來卻再也沒見到的朋友。也許該趁現在，寫封信或打通電話給他們，把想說的話說出來。'),

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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・薩爾斯堡　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});

Packer.toBuffer(mk(body2)).then((b) => fs.writeFileSync('莫札特故居_下.docx', b));
