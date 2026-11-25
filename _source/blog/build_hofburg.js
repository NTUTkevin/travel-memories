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
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '霍夫堡與西西博物館：皇帝的辦公室，皇后的傳說', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・維也納・Hofburg／Kaiserappartements／Sisi Museum／Silberkammer｜2026 年 11 月 24 日', font: F, size: 21, color: GRAY })] }),

  p('維也納的第三天早上，我們走進霍夫堡（Hofburg）。這裡是哈布斯堡家族住了六百多年的冬宮，也是整個帝國的中心。一張票可以看三個地方：皇帝與皇后的「皇家公寓」（Kaiserappartements）、西西博物館（Sisi Museum），還有皇室的銀器館（Silberkammer）。'),
  p('大家常叫她「西西公主」，其實她嫁給法蘭茲・約瑟夫以後就是皇后：奧地利皇后伊莉莎白（1837–1898），後來也是匈牙利王后。很巧的是，西西博物館重新整修，2026 年 11 月 19 日才剛重新開幕，我們是開幕後第五天去的。'),

  h1('一、米歇爾宮翼與聖米迦勒教堂'),
  p('霍夫堡面向舊城的大門，是有綠色大圓頂的米歇爾宮翼（Michaelertrakt）。它的設計圖很早就有了：1726 年，宮廷建築師約瑟夫・埃馬努埃爾・費舍爾・馮・埃爾拉赫（Joseph Emanuel Fischer von Erlach）就規劃要用這一翼連接冬季騎術學校和帝國辦公廳。可是中間擋著一座老的宮廷劇院（舊城堡劇院），這個計畫一等就是一百五十年，直到 1889–1893 年，才由費迪南・基爾什納（Ferdinand Kirschner）照著舊圖、稍作修改蓋完。大門兩側是海克力斯功績的大理石雕像。'),
  p('宮門前的米歇爾廣場中間，有一個用矮牆圍起來的挖掘坑。1989–1991 年考古學家在這裡挖到中世紀和近代的遺跡，還有羅馬時代軍營外圍城鎮（canabae）的房子。1991 年建築師漢斯・霍萊因（Hans Hollein）把它設計成露天展示，所以照片前景那片石頭，其實是將近兩千年前的維也納。'),
  p('廣場另一邊的白色尖塔，是聖米迦勒教堂（Michaelerkirche）。根據教堂自己的介紹，它約在 1220 年興建，1288 年成為堂區教堂，是維也納最古老的教堂之一；一直到 1784 年它都是皇室的「宮廷堂區教堂」。教堂裡面還保留維也納唯一、也是最大的晚期羅馬式建築；現在看到的白色古典式正面，是 1792 年加上去的。'),
  ...img('hb_wing', 1000, 750, '太太在米歇爾宮翼前：綠色圓頂、柯林斯柱、海克力斯雕像；前景矮牆圍起來的就是羅馬遺跡'),
  ...img('hb_couple', 750, 1000, '我們兩個在米歇爾宮翼前合照，大門右邊的小牌子寫著 Sisi Museum'),
  ...img('hb_church', 750, 1000, '米歇爾廣場上的聖米迦勒教堂：古典式白色正面與高高的尖塔'),
  ...img('hb_banner', 750, 1000, '西西博物館入口的紫色布條，上面印的是溫特哈爾特 1865 年那幅頭戴星星的肖像'),
  verse('在你的院宇住一日，勝似在別處住千日。', '詩篇 84:10'),
  heart('宮殿的門口底下是羅馬人的房子，旁邊是八百年的老教堂。在維也納，一個廣場就疊了好幾層歷史。'),

  h1('二、謁見等候廳：克拉夫特的三幅巨畫'),
  p('進皇家公寓不久，會走進一間掛著三幅大畫的白色大廳，這裡是「謁見等候廳」（Audienzwartesaal）：要見皇帝的人，就在這裡排隊等候。三幅畫都是畫家約翰・彼得・克拉夫特（Johann Peter Krafft）在 1828–1832 年畫的，主角不是法蘭茲・約瑟夫，而是他的祖父：法蘭茲二世／一世。他原本是神聖羅馬帝國皇帝，拿破崙稱帝後，他在 1804 年建立奧地利帝國，改稱法蘭茲一世。'),
  p('三幅畫分別是：1809 年拿破崙軍隊撤出維也納後，皇帝從普雷斯堡（今天的布拉提斯拉瓦）回城；1814 年《巴黎和約》後皇帝凱旋入城；以及皇帝重病痊癒後，第一次和皇后坐馬車出門。這系列畫要表現的，是一位親民、被人民愛戴的「好皇帝法蘭茲」。法蘭茲・約瑟夫自己倒不太喜歡這間廳，曾說它是「可怕的大廳」，1870 年代就把它重新裝修成現在的樣子。'),
  ...img('hb_krafft1', 1000, 750, '騎白馬的皇帝向群眾舉手致意，穿白衣紅腰帶的小女孩們獻花；推測是 1814 年凱旋入城那一幅（待確認）'),
  ...img('hb_krafft2', 1000, 750, '馬車窗口裡就是皇帝和皇后，大家揮帽歡呼，背景是有柱廊的城門（推測是外城門 Burgtor）；應該是「重病痊癒後第一次出門」那一幅'),
  p('影片 v17（10:12 拍的）從大廳的一角慢慢轉過去，三幅畫和大水晶吊燈都拍到了。'),
  verse('帝王榮耀在乎民多；君王衰敗在乎民少。', '箴言 14:28'),
  heart('畫裡的每一個人都在看皇帝：小孩、婦人、軍人、老人。畫家要說的是：一個國家的榮耀，是在百姓的臉上。'),

  h1('三、皇帝的謁見室與辦公桌'),
  p('法蘭茲・約瑟夫（1830–1916）十八歲登基，在位六十八年。他每週舉行兩次「公開謁見」，帝國裡任何一位子民都可以申請來見皇帝，當面說自己的事。一個早上最多可以接見大約一百人，每個人通常不到三分鐘；男士穿禮服大衣、女士穿高領洋裝、要戴手套，鄉下來的人可以穿家鄉的傳統服裝。一生中，他在謁見室裡接見了大約二十六萬人。'),
  p('謁見室裡有一張皇帝站著用的小斜桌，上面展示一份手寫的「謁見名單」（Audienz-Liste），日期是 1910 年 1 月 3 日星期一，那年皇帝已經 79 歲。名單分成「姓名與身分」「謁見目的」「備註」三欄，名字裡可以讀到 Bienerth、Böhm-Bawerk、Apponyi、Trauttmansdorff 等；我們推測 Bienerth 是當時的首相比納特，Böhm-Bawerk 是經濟學家龐巴維克（待確認）。「目的」欄寫得最多的是「a. u. Dank für …」，大意是「最謙卑地感謝皇帝賜予的任命或勳章」（讀法待確認）。'),
  p('另一間房裡，桌上放著報紙、筆和一排菸斗，窗邊畫架上是一幅披著長髮的西西肖像。皇帝最喜歡的畫，就是溫特哈爾特畫的「散著頭髮的西西」，他把它掛在辦公桌上方，每天工作時抬頭就能看到。照片這一間是不是皇帝的書房，待確認。'),
  ...img('hb_audience', 1000, 750, '紅色錦緞牆的房間，畫架上是晚年的法蘭茲・約瑟夫；左邊是年輕時穿白色軍服的全身像（房間名稱待確認）'),
  ...img('hb_list', 1000, 750, '斜桌玻璃下的謁見名單，左上角寫著「Montag den 3. Jänner 1910」'),
  ...img('hb_study', 750, 1000, '書桌與菸斗架，畫架上是披著長髮的西西；鏡子裡還映出另一位遊客'),
  verse('所以，我們只管坦然無懼地來到施恩的寶座前，為要得憐恤，蒙恩惠，作隨時的幫助。', '希伯來書 4:16'),
  heart('一百多年前，要見皇帝得申請、排隊、穿好衣服，三分鐘就結束。想到我們禱告時可以隨時來到神面前，不用排隊，也不限三分鐘，真是很大的恩典。'),

  h1('四、西西的一天：體操室與梳洗室'),
  p('1854 年 4 月，十六歲的巴伐利亞女公爵伊莉莎白嫁給表哥法蘭茲・約瑟夫，從此住進霍夫堡，也走進嚴格的宮廷禮儀裡。她最常待的地方是「體操與梳妝室」（Turn- und Toilettezimmer）：冬天早上六點，一天就從梳頭開始。她的頭髮長到腳踝，由髮型師法妮・費法利克（Fanny Feifalik）每天花很長時間編成複雜的髮型；西西覺得梳頭很煩，就利用這段時間跟希臘文老師學語言。'),
  p('房間裡到現在還保留著她的體操器材：牆上的攀爬架（肋木）、單槓，還有門框上的吊環。她在房間裡經過時就順便做幾下，比較重的訓練另外有一間大健身房。洗手台上方掛著三幅深色絨布框的水彩畫，畫的是她在希臘科孚島的布拉伊拉別墅（Villa Braila），後來她在島上蓋了自己的阿基里斯宮。房裡還有她最喜歡的詩人海涅的肖像，她自己寫的詩就是模仿海涅。後面的廁所和浴室是 1876 年她住在這裡時重新裝修的。'),
  ...img('hb_gym', 750, 1000, '西西的體操器材：左邊是攀爬架，右邊是吊環的木架；牆上有她家人的照片和騎馬的畫'),
  ...img('hb_toilet', 1000, 750, '梳妝室的一角：有三個水壺的洗手台，上方三幅深色框的就是科孚島別墅的水彩畫；窗前是梳妝桌'),
  verse('你們不要以外面的辮頭髮，戴金飾，穿美衣為妝飾，只要以裡面存著長久溫柔、安靜的心為妝飾；這在神面前是極寶貴的。', '彼得前書 3:3–4'),
  heart('一位全歐洲最美的皇后，每天要花好幾個小時在頭髮上，卻一邊梳頭一邊學希臘文。原來她心裡最想要的，不是外面的美，而是自由和知識。'),

  h1('五、西西博物館：從少女到「歐洲最美的皇后」'),
  p('西西博物館 2004 年開幕，那年正好是西西和法蘭茲・約瑟夫結婚 150 週年。博物館設在霍夫堡的「史蒂芬公寓」，展出三百多件她的私人物品：陽傘、扇子、手套、衣服、美容配方、旅行梳妝箱和藥箱，甚至還有她的死亡證明正本。'),
  p('2026 年這次重新整修，花了一千五百萬歐元，新蓋了樓梯和電梯，讓「西西博物館 → 皇家公寓 → 銀器館」變成一條連貫的參觀路線（以前是從銀器館開始）。還第一次開放了原本由聯邦總理府使用的洛可可式「白金廳」（Weißgoldzimmer），展出她的蕾絲外套、緞面鞋、體操褲和梳妝袍等新展品。'),
  p('最有名的一幅畫，是德國畫家溫特哈爾特（Franz Xaver Winterhalter）1865 年畫的全身像：西西穿著白色薄紗禮服、轉身回眸，頭髮上別著一顆顆星星形狀的鑽石髮飾。這幅畫成了她最經典的形象，入口的紫色布條印的也是它。原作屬於維也納藝術史博物館的收藏，我們看到的這幅是原作還是複製品，待確認。'),
  ...img('hb_winterhalter', 750, 1000, '溫特哈爾特 1865 年的西西肖像：白紗禮服、星星髮飾、珍珠項鍊，手裡拿著扇子'),
  ...img('hb_dresses', 750, 1000, '展櫃裡的蕾絲外套、長手套、紫色盒子與鞋拔，右邊是一件粉色條紋長禮服（是否就是新展出的白金廳展品，待確認）'),
  verse('神造萬物，各按其時成為美好，又將永生安置在世人心裡。', '傳道書 3:11'),
  heart('畫裡的西西二十七歲，美得像童話。但看完她一生的故事，才知道畫框外面的她，其實常常很孤單。'),

  h1('六、刺殺、面具與電影：被創造出來的「西西」'),
  p('1889 年，西西唯一的兒子、皇太子魯道夫在梅耶林（Mayerling）狩獵小屋死去，之後她幾乎只穿黑色的衣服，也越來越常離開維也納到處旅行。1898 年 9 月 10 日，她在瑞士日內瓦湖邊準備搭船時，被義大利無政府主義者盧切尼（Luigi Lucheni）用一把磨尖的銼刀刺中心臟，上船後不久就過世了，享年 60 歲。'),
  p('博物館裡有一幅黑衣肖像，是匈牙利畫家霍洛維茨（Leopold Horovitz）在她死後的 1899 年畫的；畫前的玻璃上投影著「ASSASSINATION（刺殺）」的字。旁邊的展櫃裡有很多她過世後發行的紀念明信片，還有一張所謂的「西西死亡面具」照片：說明牌寫著「約 1900 年，阿爾弗雷德・諾西格（Alfred Nossig）作，照片翻拍」。維也納博物館把它稱為「理想化的『死亡面具』」，也就是說，它不是真的從遺體翻模，而是後人想像出來、美化過的樣子。'),
  p('這次新增的一區，就在講「後人眼中的西西」：牆上有 1935 年匈牙利電影《女王的驃騎兵》（A királyné huszárja，柏琪・莉莉飾演伊莉莎白王后）的海報、1968 年英法電影《梅耶林》（Mayerling，泰倫斯・楊導演，艾娃・嘉娜飾演西西、奧瑪・雪瑞夫飾演魯道夫）的劇照，還有一張粉紫色禮服的海報（應該是音樂劇《伊莉莎白》，待確認）。大家最熟悉、可愛天真的「茜茜公主」形象，其實大多是後來的電影塑造出來的，和真實的她不太一樣。'),
  ...img('hb_mourning', 750, 1000, '霍洛維茨 1899 年的黑衣肖像，玻璃上投影著「ASSASSINATION」'),
  ...img('hb_mask', 750, 1000, '紀念明信片與諾西格的「死亡面具」照片（約 1900 年）；明信片看起來都是倒著的（可能是展示設計或反光，待確認）'),
  ...img('hb_films', 750, 1000, '西西與電影：上方是音樂劇海報（待確認）與 1935 年匈牙利電影海報，下方是 1968 年《梅耶林》的六張劇照'),
  verse('其實明天如何，你們還不知道。你們的生命是甚麼呢？你們原來是一片雲霧，出現少時就不見了。', '雅各書 4:14'),
  heart('一百多年來，大家用畫、面具、明信片和電影，一直在「做」一個西西出來。真正的她是誰，也許只有她自己和神知道。'),

  h1('七、銀器館：三十公尺長的皇家餐桌'),
  p('1918 年帝國結束後，宮廷的餐具、銀器和桌飾統一歸到「宮廷銀器與餐桌用品室」，1923 年開始開放參觀。這個部門從十五世紀就有紀錄，負責幫皇帝的餐桌擺上金器和銀器。館裡最大的銀器組「鍍金銀餐具」（Vermeilservice）可以同時招待 140 位客人。'),
  p('最壯觀的是「米蘭桌飾」（Mailänder Tafelaufsatz）：1838 年法蘭茲・約瑟夫的伯父、皇帝斐迪南一世加冕為倫巴底—威尼斯王國國王時訂製的，是銀器館最大的一整組。它由鍍金青銅和鏡面底座組成，全部接起來可以長達 30 公尺；底座邊緣是跳舞的小天使和燭台輪流排列。館方把它擺成一張從前菜到甜點的宮廷宴席。'),
  p('另外還有一組有埃及風格人像、托著水晶碗的鍍金青銅桌飾。館方介紹說，宮廷收藏了很多來自巴黎和維也納的鍍金青銅桌飾；我們拍到的這一組確切的名稱和年代，待確認。牆上一整排金邊的花卉瓷盤，畫著百合、龍膽、山茶、玫瑰、天竺葵和康乃馨，每一個盤子一種花（產地與名稱待確認）。'),
  ...img('hb_milan', 750, 1000, '米蘭桌飾：一整排鍍金的燭台、小天使和鏡面底座，前面擺著金邊瓷盤與餐具'),
  ...img('hb_egypt', 750, 1000, '有埃及風格人像的鍍金青銅水晶碗桌飾，放在鏡面底座上（名稱待確認）'),
  ...img('hb_plates', 750, 1000, '金邊花卉瓷盤：百合、龍膽、山茶、玫瑰、天竺葵、康乃馨'),
  verse('草必枯乾，花必凋殘；惟有我們神的話必永遠立定。', '以賽亞書 40:8'),
  heart('三十公尺的金色餐桌，比我們家整條走廊還長。可是帝國已經沒有了，留下來的只是玻璃櫃裡的餐具。盤子上的花畫得再美，也提醒我們：只有神的話是永遠的。'),

  h1('旅人小記'),
  bullet('時間：11/24 上午。10:00 米歇爾宮翼前拍照、10:04 聖米迦勒教堂、10:12 謁見等候廳（影片 v17，GPS 在霍夫堡）。'),
  bullet('入口：米歇爾圓頂（Michaelerkuppel）下方；一張票包含西西博物館、皇家公寓和銀器館，三個地方要依順序一起參觀。官網門票只在 imperialtickets.com 販售；票價包含多語語音導覽，可以拍照。'),
  bullet('開放時間：每天 9:00–17:30（最後入場 16:30）；7、8 月到 18:00。'),
  bullet('交通：地鐵 U3 Herrengasse 站走幾分鐘就到米歇爾廣場。'),
  bullet('待確認：我們實際的參觀順序（照片看起來是先皇家公寓、再西西博物館、最後銀器館）；第一幅騎白馬的畫是不是 1814 年入城；紅牆、畫架上晚年皇帝的那間是不是謁見室；有菸斗架與長髮西西肖像的那間是不是皇帝的書房；謁見名單的讀法（右上角的「50 Parteien」等字也待確認）；溫特哈爾特肖像是原作還是複製品；蕾絲外套和條紋禮服是不是新開放的白金廳展品；明信片為什麼是倒的；粉紫色海報是不是音樂劇《伊莉莎白》；埃及風格桌飾與花卉瓷盤的名稱；買的是哪一種票。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：sisimuseum-hofburg.at（西西博物館、皇家公寓各房間、銀器館、語音導覽文字、導覽手冊）、wien.orf.at（2026 西西博物館重新開幕、2004 開幕）、wien.info（皇家公寓、銀器館、聖米迦勒教堂）、habsburger.net（法蘭茲・約瑟夫的謁見、米歇爾廣場遺跡、法蘭茲一世與克拉夫特）、vienna.at（克拉夫特三幅畫）、hofburg.com、ganz-wien.at（米歇爾宮翼）、michaelerkirche.at（教堂歷史）、sammlung.wienmuseum.at（理想化的死亡面具、霍洛維茨肖像）、Springer《Wiener Medizinische Wochenschrift》（刺殺與盧切尼）、Wikipedia（溫特哈爾特肖像、1968《梅耶林》、1935《女王的驃騎兵》）。', font: F, size: 17, color: GRAY })] }),
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
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '旅行的生活紀錄・維也納・霍夫堡與西西博物館　', font: F, size: 17, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 17, color: GRAY })] })] }) },
    children,
  }],
});
Packer.toBuffer(mk(bodyK)).then((b) => fs.writeFileSync('維也納_霍夫堡與西西博物館.docx', b));
