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

const bodyF = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '霍亨薩爾斯堡要塞與「白色黃金」', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡｜2026 年 11 月 21 日', font: F, size: 21, color: GRAY })] }),

  p('這天我們從舊城往上走，來到薩爾斯堡最醒目的地標——霍亨薩爾斯堡要塞（Festung Hohensalzburg）。它蓋在海拔約五百公尺的要塞山上，從舊城的任何一條街抬頭，都看得見那一圈白色的城牆。'),
  h1('一、一座蓋了九百多年的「教會要塞」'),
  p('要塞的起點在 1077 年。當時的薩爾斯堡大主教格布哈德（Gebhard von Helfenstein）正處於「敘任權之爭」的風暴中：他站在教宗一邊，與神聖羅馬帝國皇帝亨利四世對抗，於是下令在城市上方築一座防禦的城堡。最早只是一座木造的圍牆與塔樓；格布哈德 1085 年被迫流亡，後繼的大主教們才陸續把它改建成石造，並完成基本的規模。'),
  p('今天看到的樣貌，主要來自 1495 至 1519 年在位的大主教萊昂哈德・馮・克楚赫（Leonhard von Keutschach）。他把要塞擴建成後期哥德式的諸侯居所：華麗的「君主廳」、1501–1502 年的聖喬治教堂，還有 1502 年裝設的機械管風琴「薩爾斯堡公牛」，以及 1504 年用來往山上運送物資的纜索軌道 Reisszug，被認為是歐洲最早的纜車之一。整座要塞長約 250 公尺、寬約 150 公尺，是中歐保存最完整的中世紀城堡之一。'),
  p('這座要塞從來沒有被外敵攻下，唯一一次被圍困是 1525 年的農民戰爭：礦工、農民與市民起義，從 6 月到 8 月圍困要塞，大主教馬圖斯・朗（Matthäus Lang）和隨從被困在山上，但要塞沒有失守。平常大主教們大部分時間並不住在山上，而是住在山下舊城中心的主教宮（Residenz）。換句話說，它更像是權力的宣示：向城裡的市民、向外面的諸侯說，這片土地由教會的君王統治。'),
  ...img('f_window', ...P, '從要塞的菱格窗往下看：舊城的教堂塔樓與屋頂（右上為方濟會教堂方向，請以實景為準）'),
  verse('萬軍之耶和華與我們同在；雅各的神是我們的避難所。', '詩篇 46:7'),
  heart('隔著鐵窗看出去，整座城像被框在一幅畫裡。我們想，要塞的牆是人蓋的避難所，但真正讓人安心的，是不論在山上或山下，神都與我們同在。'),

  h1('二、站在高處看薩爾斯堡'),
  p('登上城牆平台，視野一下子打開：薩爾察赫河（Salzach）從遠處彎進舊城，河的一側是舊城（莫札特出生地所在），另一側是新城（莫札特故居所在）。照片右側最大的青綠色圓頂，是薩爾斯堡主座教堂；它前方的卡匹特廣場（Kapitelplatz）上，那顆金色的球體是現代雕塑，常被當作舊城裡的小地標（說明請以現場為準）。'),
  p('那天是陰雨的 11 月，山頭的樹葉有橘、有黃，屋頂被雨淋成深灰色。比起夏天明亮的景色，這種安靜的色調，反而更像兩百多年前莫札特看到的薩爾斯堡。'),
  ...img('f_selfie', ...L, '我和太太在要塞城牆上的合照，背景是薩爾斯堡舊城與主座教堂'),
  ...img('f_pano', ...L, '要塞上的全景：左邊是山林與小屋，右邊是舊城與薩爾察赫河'),
  heart('和太太站在這裡，雨滴還掛在外套上，兩個人笑得很開心。我們想起，今天早上從莫札特的故居出來，現在又站在他每天看見的山上俯瞰這座城，感覺好像在這個城市裡，被歷史輕輕擁抱。'),

  h1('三、這座城為什麼這麼富有？答案是鹽'),
  p('「薩爾斯堡」的名字，字面意思就是「鹽之堡」（Salz＝鹽，Burg＝城堡）。這不是巧合：整座城能蓋出主教座堂、宮殿和要塞，靠的是被稱為「白色黃金」的鹽。'),
  bullet('鹽的產地在附近的哈萊因（Hallein）與杜恩貝格（Dürrnberg）山區。早在鐵器時代，凱爾特人就已在這裡開採。12 世紀前後，薩爾斯堡的大主教們重新重視這個礦，哈萊因採鹽最早出現在 1198 年的文件中。'),
  bullet('為了搶鹽的利益，爭奪非常激烈：1196 年，薩爾斯堡大主教阿達爾貝特曾經把巴伐利亞的萊興哈爾城連同鹽場一把火燒掉。'),
  bullet('鹽被壓成圓錐形的木桶，裝上船，順著薩爾察赫河運到下游，沿河的城鎮因此繁榮。這筆穩定的收入，成了大主教們「富甲一方」的基礎，也讓他們得以把舊城建成今天的巴洛克樣貌。'),
  p('走進要塞內的展廳，木樑天花板下是一座用白色（鹽色）材質做成的薩爾斯堡立體模型：山丘上的要塞、舊城的教堂圓頂與屋頂、河邊成排的建築，一覽無遺。牆上的標語寫著：「Der Mensch kann ohne Gold leben, aber nicht ohne Salz」，英文是「Humankind can live without gold yet not without salt」——人沒有黃金還能活，沒有鹽卻不行。另一句標語說的是這裡是「水、鹽與阿爾卑斯山交會」的地方（Wo sich Wasser, Salz und die Alpen kreuzen）。'),
  ...img('f_model', ...P, '要塞展廳裡的薩爾斯堡城市模型，牆上寫著「人沒有黃金還能活，沒有鹽卻不行」'),
  p('在要塞的展區裡，我們看到一座沙色的立體地形模型，山丘上有一座城堡；前方是一列提著燈籠、背著籃子和鎬的小礦工剪影。標籤上有「SALE PREZIOSO」（義大利文「珍貴的鹽」），說的正是這段鹽的故事——原來這座要塞本身，就蓋在鹽所累積的財富與權力之上。（展區的確切名稱我沒拍清楚，之後可以補上。）'),
  ...img('f_salt', ...P, '要塞展區裡的地形模型與提燈礦工的剪影，標籤寫著「珍貴的鹽」'),
  verse('你們是世上的鹽。鹽若失了味，怎能叫它再鹹呢？', '馬太福音 5:13'),
  heart('原來，薩爾斯堡的繁華來自一粒粒不起眼的鹽。耶穌說我們是「世上的鹽」，一點點、看似平凡，卻能保存、調味，也能成就一座城。我們也想成為這樣的人——在平凡的日子裡，讓身邊的人嚐到神的味道。'),

  h1('四、要塞的另一面：牢房與囚禁'),
  p('要塞不只是宮殿與防禦工事，也是關人的地方。19 世紀它曾被當作兵營、倉庫和監獄，直到 1861 年才不再作為軍事要塞；第一次世界大戰時，這裡還關過戰俘。今天的導覽路線裡也有刑訊室與地牢，但那不是要塞的主要功能。'),
  p('最有名的囚犯，是薩爾斯堡自己的大主教——沃爾夫・迪特里希・馮・賴特瑙（Wolf Dietrich von Raitenau）。他 1587 年 28 歲就當上大主教，是個受馬基維利影響的權力人物，也是把中世紀舊城改建成巴洛克城市的人。他深愛一位平民女子莎樂美・阿爾特（Salome Alt），並想盡辦法與她結婚，兩人育有許多孩子。1611 年，他在與巴伐利亞的鹽業衝突中被捕、遭罷黜，先被關在霍亨維爾芬城堡，後來被轉到霍亨薩爾斯堡。他的堂弟兼繼任者馬庫斯・西提庫斯（Markus Sittikus）把他關在山上，直到 1617 年他死在要塞裡，再也沒有見到莎樂美和孩子。（關押年數各來源說法不一，約五到六年。）'),
  p('也就是說，那座為了保護大主教而蓋的要塞，最後囚禁了自己的大主教。權力可以建起高牆，卻也可能成為關住自己的牢籠。'),
  ...img('f_grate', ...P, '要塞內地面上的鐵柵與細鐵網（確切用途請以現場說明為準，可能與地牢或豎井有關）'),
  verse('求你領我出離被囚的地方，我好稱讚你的名；義人必環繞我，因為你是用厚恩待我。', '詩篇 142:7'),
  heart('站在這塊鐵柵旁，往下看是一片漆黑。我們想到那位曾經掌握整座城、最後卻被囚在自己城堡裡的大主教，也想到真正的自由不在權力與牆內牆外，而是在基督裡。'),

  h1('五、要塞也是兵營：薩爾斯堡的「自家軍團」'),
  p('要塞的故事不只有大主教。官方的導覽資料提醒我們：在 20 世紀變成觀光景點之前，霍亨薩爾斯堡一直是軍事設施。拿破崙戰爭之後，薩爾斯堡經過一番轉折，在 1816 年歸入奧地利，這座要塞便成了兵營，也是薩爾斯堡「自家軍團」的基地。'),
  ...img('s_corridor', ...P, '太太走在要塞內狹窄的石砌通道，頭頂是粗糙的石塊橫樑（確切位置請以現場為準）'),
  p('這個軍團的全名是「皇家暨王家步兵第 59 團『萊納大公』」（k. u. k. Infanterieregiment „Erzherzog Rainer“ Nr. 59），一般簡稱「萊納團」（Rainer-Regiment）。它的歷史可以追溯到 1682 年，名字來自萊納・斐迪南大公（Archduke Rainer Ferdinand），他是奧皇法蘭茲・約瑟夫一世的再從兄弟。1816 年薩爾斯堡併入奧地利後，上奧地利的因維特爾（Innviertel）與豪斯魯克維特爾（Hausruckviertel）被劃為這個團的徵兵區。自 1924 年起，要塞裡就設有「萊納團博物館」（Rainer-Regimentsmuseum），收藏制服、武器與戰時、平時的紀念物；展廳數目各來源寫的是八間到十間不等。'),
  h2('從白雲石山區到蒙特奇莫內'),
  p('第一次世界大戰時，萊納團跟著奧匈帝國打到義大利前線。有一份引用該團戰時日誌的研究提到：1915 年 10 月底的夜裡，這個團在南蒂羅爾白雲石山區的塞克斯滕石（Sasso di Sesto）發動了一次大規模進攻；同年冬天，日誌裡記下的是高山上的雪崩、砲擊與巡邏。另有一本以薩爾斯堡志願兵日記寫成的書，書名說的是 1916 年春夏，他隨「雪絨花師」（Edelweiß-Division）在南蒂羅爾的阿爾卑斯山一路打到蒙特奇莫內（Monte Cimone）。'),
  p('蒙特奇莫內（Monte Cimone di Tonezza，海拔 1,226 公尺）位於今天義大利的威尼托地區，是一戰山地坑道戰最有名的戰場之一：雙方都在山裡挖坑道、埋地雷。1916 年 9 月 23 日，奧匈軍引爆了埋在山頂下的地雷，在山頂炸出寬約 50 公尺、深約 22 公尺的彈坑；義軍在爆炸和接著的戰鬥中損失 1,137 人，其中約 500 人被俘。（我沒有查到這次引爆與萊納團之間的直接關係，這一段只當作這座山的背景。）'),
  ...img('s_mural', ...P, '要塞內的戰爭壁畫：士兵頂著長大衣、端著上刺刀的步槍，爬過滿是碎石與枯枝的山坡。說明牌只拍到最後的「…onte Cimone」，我推測畫的是這處戰場，也應是萊納團博物館的展品，請以現場說明為準'),
  p('這幅畫裡沒有爆炸的火光，只有一群人低著身子往山上爬，腳下是碎石與斷木，最前面那位仰頭望向上方。戰史裡「一千多人、一個彈坑」的數字，落在畫裡，變成每個人臉上的緊張。'),
  p('萊納團的戰史不只這座山。有位參觀過博物館的作者寫到，一戰展區涵蓋了這個團所有的戰役，包括 1917 年的卡波雷托（Caporetto），並用立體模型呈現義大利前線與加利西亞的駐地；館內還有一座紀念碑，紀念一戰中陣亡與失蹤的萊納團士兵。另外，義大利學者整理前線碑文時，記錄過一塊由萊納團立在戰場上的碑，德文大意是：旅人啊，如果你回到薩爾斯堡，請告訴家鄉的人，我們已經倒下，忠於皇帝與土地。（資料沒有寫明這塊碑在哪一處。）'),
  verse('他們要將刀打成犁頭，把槍打成鐮刀。這國不舉刀攻擊那國；他們也不再學習戰事。', '彌迦書 4:3'),
  heart('站在畫前，我們想到那些離開家鄉、再也沒有回到薩爾斯堡的年輕人。這座白牆下的要塞，有人曾用整個世代的生命，換來我們今天能安心旅行的和平。我們想把先知的話記在心裡：有一天，刀要打成犁頭，人也不再學習戰事。'),

  h1('六、牆上的主教們：哈拉赫與菲爾米安'),
  p('在要塞博物館的展廳裡，一面白色大螢幕前站著三尊戴主教冠的石像，上方的紅色跑馬燈依序列出大主教的名字與在位年份。我們拍到的那一段，剛好是「Harrach 1709–1727」與「Leopold Anton Freiherr von Firmian」。（館方資料說，要塞博物館有一個主題專講「薩爾斯堡的大主教們」，這應該就是那個展區；三尊石像分別是誰，我無法從照片確認。）'),
  ...img('s_bishops', ...L, '要塞博物館的大主教展示：三尊戴主教冠的石像，上方跑馬燈拍到「…Harrach 1709–1727」與「Leopold Anton Freiherr von Fir…」'),
  h2('什麼是「大主教」？'),
  p('薩爾斯堡的大主教有兩個身分：對教會，他是整個教區的屬靈領袖；對世俗，他是統治「薩爾斯堡大主教領地」的君王（諸侯），握有土地、鹽礦、軍隊與稅收，直接隸屬神聖羅馬帝國，直到 19 世紀初這個領地被世俗化為止。所以山上這座要塞，既是教會的產業，也是君王的城堡。前面說的格布哈德、克楚赫、沃爾夫・迪特里希，都是這樣的大主教。'),
  h2('哈拉赫（Franz Anton von Harrach，1709–1727 在位）'),
  p('哈拉赫 1665 年生於維也納，出身奧地利名門。他先當過維也納主教（約 1702 至 1705 年，各資料年份略有出入），之後回到薩爾斯堡，1709 年成為大主教，1727 年 7 月過世。英文資料形容他親切、受歡迎，愛排場與慶典，並大幅擴建、翻新了米拉貝爾宮（Schloss Mirabell）。這座宮殿最早是沃爾夫・迪特里希 1606 年為莎樂美・阿爾特蓋的——也就是前一節那位被關在要塞裡的大主教，和他深愛的平民女子。從要塞往河對岸新城看，那一片白色宮殿與花園的方向，就是米拉貝爾宮。DomQuartier 還保存著他 1710 年的畫像，是一組九幅家族肖像裡唯一留下來的一幅；他本人也留下一份「Notata」日記，近年有學者整理研究。'),
  h2('菲爾米安（Leopold Anton von Firmian，1727–1744 在位）'),
  p('菲爾米安 1679 年生於慕尼黑，出身蒂羅爾的古老貴族家族。他的舅舅約翰・恩斯特・馮・圖恩（Johann Ernst von Thun）在 1687 至 1709 年間擔任薩爾斯堡大主教——上一份紀錄裡那座大學教堂，就是在這段期間動工的。1727 年哈拉赫過世後，菲爾米安被選為大主教。和莫札特一家也有一點關係：維基百科記載，他是萊奧波德・莫札特早年的贊助者，而他的姪子卡爾・馮・菲爾米安伯爵，則是莫札特歌劇《米特里達特》的贊助者之一。'),
  p('但他留在歷史上最深的一筆，是 1731 年的驅逐令。薩爾斯堡的新教歷史可以追溯到 16 世紀：維基百科說，路德的思想隨著大主教馬圖斯・朗從薩克森招來的礦工傳進薩爾斯堡（就是前面 1525 年被農民與礦工圍困在要塞裡的那位）。菲爾米安上任後，想加強對偏遠山區的行政與教牧管理，引起反抗；1731 年 6 月，紐倫堡出版了一份文件，聲稱有一萬九千名受壓迫的新教徒，人數比他預料的多得多。他最後選擇用驅逐來解決：要求境內的新教徒必須離開薩爾斯堡。令的簽署日期，維基百科寫 1731 年 10 月 31 日（正好是宗教改革紀念日），百科全書則寫 11 月 11 日，各來源不一；維基百科也寫到，最短的離境期限只有八天，十二歲以下的孩子必須留下。'),
  p('被驅逐的人數，各資料說法是兩萬到兩萬三千人。他們大多走向北方，許多人最後被普魯士國王腓特烈・威廉一世接納，在 1732 年前後遷到普魯士。這場驅逐在當時的德語世界引起新教地區的強烈抗議，也成了新教與天主教之間記憶深刻的「傳說」；歌德的敘事詩《赫爾曼與竇綠苔》（Hermann und Dorothea）就提到過這件事。公平地說，菲爾米安與他的支持者認為這是恢復領地的信仰統一；被驅逐的人則認為，他們只是想照著良心敬拜。兩邊對這段歷史的記憶，到今天仍然不一樣。'),
  verse('這些人都是存著信心死的，並沒有得著所應許的；卻從遠處望見，且歡喜迎接，又承認自己在世上是客旅，是寄居的。', '希伯來書 11:13'),
  heart('站在這些主教的名字前，我們想到一件事：宗教與權力綁在一起的時候，信仰有時保護了人，有時也讓人被迫離開家園。我們很感恩今天能自由地敬拜，也想記住那兩萬多名被迫離開家鄉的人。信仰不能被強迫，最終的家鄉也不在任何一座城堡裡，而是在神為我們預備的城。'),

  h1('七、最頂樓：雷克塔的 360 度全景'),
  p('這趟參觀的最後一站，是要塞的最高處。官方的「全景導覽」（Panorama Tour）會帶遊客從鹽倉出發，經過地牢，最後登上雷克塔（Reckturm）的瞭望平台，被認為是城裡視野最好的地方，可以 360 度環視整座城市。'),
  p('11 月陰雨的天氣，雲壓得很低，遠方的阿爾卑斯山多半藏在雲後面，但近處的景色反而特別清楚：薩爾察赫河像一條銀灰色的緞帶，從舊城邊轉彎流過，兩岸是深色的屋頂和教堂尖塔。往河的另一邊看，是新城的白色建築；左邊的山坡是僧侶山（Mönchsberg）。'),
  ...img('t_view1', ...L, '從最高處往下看：薩爾察赫河穿過舊城，近處是教堂的尖塔與圓頂'),
  ...img('t_view2', ...P, '同一個方向，鏡頭拉高一點，可以看到河上的橋與遠方的城區'),
  p('轉向另一個方向，則是城外的住宅區與綠色的草地。畫面左邊那座白色塔樓，是要塞的一部分。'),
  ...img('t_view3', ...P, '往城外看：塔樓、住宅區、綠地與遠方被雲遮住的山'),
  p('平台的欄杆旁，有一塊黃銅色的導覽板，上面刻著周圍山景與地標的線條圖，讓遊客可以對照眼前看到的景色，這也是我們兩個人在照片裡背後的那塊金色板子。'),
  ...img('t_wife', ...P, '太太在要塞最高處的平台，後面是銅製導覽圖與城外景色'),
  ...img('t_me', ...P, '我在同一個平台，背後是薩爾斯堡舊城與僧侶山'),
  p('平台旁還有一扇木框的窗。太太戴著紅色貝雷帽，透過窗口望向河谷；窗台的圓形檯面上散落著遊客留下的硬幣，窗框上方掛著鏤空的小貓頭鷹裝飾。'),
  ...img('s_window', ...L, '太太從木窗望出去：薩爾察赫河、舊城的屋頂，遠處是被雲遮住的山'),
  verse('我要向山舉目；我的幫助從何而來？我的幫助從造天地的耶和華而來。', '詩篇 121:1-2'),
  heart('站在最高的地方，風很大，雲很低，我們兩個人都被吹得頭髮亂亂的，可是都笑得很開心。往下看，整座城那麼小；往上看，天好像更近。我們想起：不論走多高、走多遠，真正的幫助都從造天地的神而來。'),

  h1('旅人小記'),
  bullet('要塞可以搭纜索鐵路（Festungsbahn）上山，也可以步行；山上風大、陰雨時很冷，記得帶外套。'),
  bullet('票價與開放時間、套票內容請以官網為準：festung-hohensalzburg.at。'),
  bullet('想了解鹽的歷史，可以另外去哈萊因的鹽礦（Salzwelten）。'),
  bullet('要塞博物館（含大主教展區）與萊納團博物館都在要塞內，開放與票務細節請以官網為準。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：霍亨薩爾斯堡要塞官網（festung-hohensalzburg.at）、salzburg.info、Austria-Forum AEIOU〈Hohensalzburg〉、salzburg.info（Hohensalzburg Fortress、Salzburg from above）、Wikipedia〈Hohensalzburg Fortress〉〈Hallein〉、salzburg.info〈Von göttlichen Gnaden: Salzburgs Fürsten und Erzbischöfe〉、Burgen-Austria（Hohenwerfen）、Salzwelten〈History of the Salzburg Salt Mine〉、Salzburgerland.com〈Das Salz in Salzburg〉；萊納團與戰爭：salzburg.info〈Rainer Regimental Museum〉、salzburg-burgen.at〈Museums at the fortress〉、Written in the Landscape（unibz）、Wikipedia〈Monte Cimone di Tonezza〉、Roads to the Great War 部落格〈Salzburgs Rainer Regiment Museum〉、TracesOfWar；主教：Wikipedia〈Franz Anton von Harrach〉〈Leopold Anton von Firmian〉〈Salzburg Protestants〉、Encyclopedia.com〈Salzburg Expulsion〉、germanhistorydocs.org、Austria-Forum AEIOU〈Firmian〉、DomQuartier Salzburg。經文引自中文和合本。', font: F, size: 18, color: GRAY })] }),
];

const bodyK = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '薩爾斯堡大學教堂（Kollegienkirche）', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡｜2026 年 11 月 21 日', font: F, size: 21, color: GRAY })] }),

  p('離開要塞，下山走進舊城，在一座開闊的廣場邊，看見這座純白色的巴洛克教堂：大學教堂（Kollegienkirche）。陰天裡，它的深灰色圓頂和白牆特別乾淨，旁邊是秋天轉黃的樹。'),
  h1('一、一座為大學而蓋的教堂'),
  p('薩爾斯堡的本篤會大學創立於 17 世紀 20 年代初，是大主教帕里斯・馮・洛德龍（Paris von Lodron）的事業。起初大學沒有專屬的教堂，只在大禮堂（Aula）舉行彌撒。直到 1694 年，大主教約翰・恩斯特・馮・圖恩（Johann Ernst von Thun）決定建一座大教堂，並找來當時最重要的巴洛克建築師費舍爾・馮・埃爾拉赫（Johann Bernhard Fischer von Erlach）。教堂從 1694 年蓋到 1707 年，是他在薩爾斯堡最大、也被認為最傑出的作品。'),
  p('更巧的是，教堂的祝聖典禮從 1707 年 11 月 20 日開始，一連舉行八天。我們今天走到這裡是 2026 年 11 月 21 日，剛好是它祝聖 319 週年的前後。教堂獻給「聖母無原罪始胎」，祭壇畫出自畫家羅特邁爾（Rottmayr）之手。'),
  ...img('f_kolleg', ...L, '陰天裡的大學教堂：深色圓頂、白色外牆與鐘塔'),
  h1('二、教堂經歷過的事'),
  bullet('1800 年左右，拿破崙的軍隊佔領薩爾斯堡，教堂一度被拿來當乾草倉庫。'),
  bullet('1810 年，在巴伐利亞統治下，薩爾斯堡大學被關閉，教堂也失去了大學教堂的功能。'),
  bullet('1922 年，霍夫曼斯塔爾的《大世界劇場》在這裡首演，後來成為薩爾斯堡藝術節的重要場地之一。'),
  bullet('1960 年代大學重新開辦後，這裡又恢復了大學教堂的用途，同時也是音樂與藝術活動的場地。'),
  p('另外，莫札特少年時和這所大學也有關係。1767 年他 11 歲時創作的拉丁文音樂劇《阿波羅與雅辛托斯》（K. 38），是在這所大學的大禮堂演出的——也就是說，他曾經在這座教堂旁邊的校園裡演出自己的作品。'),
  verse('萬軍之耶和華啊，你的居所何等可愛！', '詩篇 84:1'),
  heart('一座教堂可以被拿來當倉庫、被關閉、又重新被使用，但它還是站在這裡，讓人抬頭就想到神。我們也想起自己生命中的低谷，神從不放棄，總會讓我們重新被使用。'),

  h1('旅人小記'),
  bullet('大學教堂位於大學廣場（Universitätsplatz），離莫札特出生地和糧食胡同走路不遠；廣場平日有市集。'),
  bullet('我這次只拍了外觀，沒有進入內部，所以沒有內部照片；下次進去的話，可以特別看圓頂與祭壇。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：salzburg.info〈Kollegienkirche〉、Wikipedia〈Kollegienkirche, Salzburg〉、Frommer\'s〈Kollegienkirche〉、Sacred Destinations、Waymarking。經文引自中文和合本。', font: F, size: 18, color: GRAY })] }),
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
Packer.toBuffer(mk(bodyF)).then((b) => fs.writeFileSync('薩爾斯堡_要塞.docx', b));
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('薩爾斯堡_大學教堂.docx', b));
