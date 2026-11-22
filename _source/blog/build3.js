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

const body3 = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: '旅行的生活紀錄', font: F, size: 22, color: GOLD, bold: true })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '走進莫札特的家（終）', font: F })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: '奧地利・薩爾斯堡｜莫札特故居 Mozart-Wohnhaus｜2026 年 11 月 21 日', font: F, size: 21, color: GRAY })] }),

  h1('七、維也納的十年（1781–1791）'),
  p('前面幾個展間看的，多半是莫札特在薩爾斯堡的少年歲月。這一間往後，故事移到維也納——他生命裡最輝煌、也最辛苦的十年。'),
  h2('1781：離開薩爾斯堡，成為自由音樂人'),
  p('1781 年 3 月，莫札特隨科洛雷多大主教到維也納。他在這裡的待遇像個僕人，與大主教的衝突愈演愈烈，終於在 6 月被解職。對當時的音樂家來說，離開宮廷等於放棄穩定的薪水，父親利奧波德極力反對；但莫札特選擇留下來，試著靠演奏、教學和作曲養活自己——這在 18 世紀是少見的冒險。'),
  p('他的收入主要來自三塊：演出自己的鋼琴協奏曲、訂閱音樂會和自己舉辦的「學院音樂會」（Akademie），再加上教貴族子弟彈琴與出版樂譜。維也納沙龍裡鋼琴是主角，他 25 歲以前寫了 10 首鋼琴協奏曲，維也納十年間又寫了 17 首。'),
  h2('1782：結婚'),
  p('1782 年 8 月，他在聖斯德望主教座堂與康絲坦采・韋伯結婚，父親一直到最後都不太贊成。婚後兩人搬過許多次家，生了六個孩子，只有卡爾・托馬斯和弗朗茨・克薩韋爾兩個兒子長大成人。展櫃裡有一份手稿信，旁邊是一幅小肖像和聖斯德望主教座堂的版畫；說明牌我當時沒拍清楚，只能說這個組合讓人想到他們在這座教堂結婚的故事（請以展牌為準）。'),
  ...img('v_letter', ...P, '展櫃：手稿信、小肖像與聖斯德望主教座堂版畫（說明牌待確認）'),
  h2('1786：《費加洛婚禮》與愛爾蘭朋友凱利'),
  p('1786 年 5 月 1 日，《費加洛婚禮》在維也納宮廷劇院（Burgtheater）首演，劇本是達・蓬特根據博馬舍的革命喜劇改寫，政治諷刺被刻意淡化，但音樂仍清楚地描繪了貴族與僕人之間的階級張力。首演獲得熱烈迴響。'),
  p('展櫃裡那張戴著高領的男子版畫，是愛爾蘭男高音凱利（Michael Kelly, 1762–1826）。他 1783 年到維也納，在《費加洛婚禮》首演中同時演出音樂教師巴西利奧和公證人庫爾奇奧。他後來在 1826 年出版的回憶錄中，留下莫札特指揮排練的珍貴紀錄：莫札特是個非常瘦小、蒼白的人，滿頭金髮，而且很自豪；排練時他穿著深紅色大衣、戴著鑲金邊的帽子指揮。1787 年凱利離開維也納時，還去向「不朽的莫札特和他可愛的妻子與家人」告別。不過學者提醒，凱利回憶錄中有些說法不能全信，只能當作參考。'),
  ...img('v_kelly', ...L, '凱利（Michael Kelly）版畫、莫札特 1789 年柏林王家劇院版畫與早期劇本'),
  p('同一櫃的另一張版畫標示「莫札特在柏林王家劇院，1789 年」——那是他 1789 年春天的北德之行，離開維也納幾個星期去找新的機會。右邊翻開的小冊子畫著一對男女與野獸，我推測是與《魔笛》相關的早期版本，說明牌看不清，僅供參考。'),
  h2('維也納的劇院'),
  p('旁邊牆上有兩張劇院圖：左邊是一幅大版畫，畫著擠滿觀眾的華麗大廳，舞台上有幾位穿紅袍的人物，頂上是畫著雲彩的穹頂；右邊是一幅劇院剖面圖，可以看到好幾層包廂、舞台下方的機關。法文標題寫著「Coupe et profil」，說明牌看不清楚，我無法確定是哪一座劇院，請你對照展牌。可以確定的是，18 世紀的歌劇是整個城市最重要的公共娛樂，莫札特的好幾部歌劇都是在這樣的環境裡誕生的。'),
  ...img('v_theatre', ...L, '劇院內部版畫與剖面圖（劇院名稱待確認）'),
  p('再往裡走，紅色牆面上嵌著一幅幅舞台模型。我拍到的其中一幅有骷髏、方尖碑和墓園，與《唐喬凡尼》的墓地場景很相似，但我沒有看清楚說明牌，請以展牌為準。《唐喬凡尼》1787 年 10 月在布拉格首演大獲成功，1788 年再在維也納上演。'),
  ...img('v_models', ...L, '紅牆上的舞台模型與設計稿（疑為墓園場景，請以展牌為準）'),
  h2('1787 年以後：盛名與開銷'),
  p('1787 年 5 月，父親利奧波德去世。從這一年開始，莫札特的處境有些微妙：他的收入遠高於一般音樂家，卻比不上他想進入的貴族圈；他和妻子養成了與身分不符的奢侈品味。1788 年爆發土耳其戰爭，壓縮了維也納的社交生活，加上康絲坦采生病，家裡出現資金缺口。他寫信向朋友借錢，也有人說他死時負債三萬弗洛林——這數字是死後流傳的謠言，學者至今仍在爭論。'),
  p('不過他並不是窮到沒飯吃。學者研究發現，他的年收入大致相當於商人或高階官員，真正的問題是病痛減少了他能教學、演出、作曲的收入，加上花費也大：他們住在維也納較高級的公寓，請僕人，送兒子卡爾進昂貴的寄宿學校，社交頻繁。最後一年（1791）反而是他最多產、收入也很可觀的一年：寫了《魔笛》、《提圖斯的仁慈》、單簧管協奏曲、《Ave verum corpus》，並開始《安魂曲》。'),
  p('維也納十年的另一件事是他的朋友圈：海頓是他最敬重的前輩，兩人互相致敬；斯維滕男爵帶他認識巴赫與韓德爾；達・蓬特給了他三部經典歌劇的劇本。這些友誼和他的音樂一樣，都是他在維也納留下的資產。'),
  verse('我知道怎樣處卑賤，也知道怎樣處豐富；或飽足或飢餓，或有餘或缺乏，隨事隨在，我都得了秘訣。', '腓立比書 4:12'),
  heart('原來莫札特也有帳單、有借錢、有被生活壓得喘不過氣的日子。可是他的音樂裡沒有一絲怨氣。我們想起自己人生中起伏的階段，也提醒自己：無論豐足或缺乏，都要學習知足、感恩。'),
  h1('八、病中的莫札特'),
  p('另一幅油畫畫的是晚年的莫札特：他靠著大枕頭、腿上蓋著毯子，伸手指向一旁的琴或樂譜。這類畫是 19 世紀以後人們想像莫札特最後時光的作品，帶著濃濃的浪漫色彩。1791 年 12 月，他在維也納過世，年僅 35 歲；臨終前仍在寫的是《安魂曲》，沒能完成。同一年，他還寫下了短短的聖體經文歌《Ave verum corpus》（K. 618），至今仍是教會最常唱的聖樂之一。'),
  p('早在 1787 年父親病重時，莫札特就寫信說，他已經學會把死亡看作「人真正的好朋友」，每晚睡前都想到自己也許明天就不在了，卻因此更感恩。'),
  ...img('ill', ...P, '描繪病中莫札特的油畫'),
  verse('我們活著是為主而活，死了是為主而死。所以，我們或活或死總是主的人。', '羅馬書 14:8'),
  heart('看完這幅畫，我們在房間裡站了很久。35 歲，留下六百多首作品。我們彼此說：日子不在長短，在於是不是用來讚美、用來愛人。走出故居，聖三一教堂的鐘聲正好響起。'),



  h1('九、身後的莫札特：誰在說他的故事？'),
  p('1791 年 12 月 5 日，莫札特過世。他的名聲在死後才真正擴大，而這背後有一位關鍵人物：遺孀康絲坦采。她安排《安魂曲》由弟子絮斯邁爾完成，賣掉手稿，舉辦紀念音樂會，最後成功地讓自己和孩子擺脫貧困。她後來改嫁丹麥外交官尼森（Georg Nikolaus von Nissen），兩人一起整理莫札特的文獻。'),
  p('這個展櫃正是談這一段歷史。左邊那本舊書是尼森的《莫札特傳》（Biographie W. A. Mozart\'s）。尼森 1826 年過世時只寫了一小部分，書由別人依據他的筆記完成，1828 年（實際出版在 1829 年）由萊比錫的布萊特科普夫出版社出版。他拿到南妮爾提供的約 400 封莫札特家書，是後世研究的重要來源；但他從沒見過莫札特，主要依賴康絲坦采的回憶，因此書中對莫札特和妻子的描寫偏向美化，學者讀的時候都要小心。'),
  ...img('v_nissen', ...P, '展櫃：尼森《莫札特傳》、「莫札特之耳」水彩、手稿與莫札特側面像；上方為紅衣軍裝肖像（身分待確認）'),
  p('書旁邊的那幅畫框，是兩隻耳朵的水彩，也就是醫學上有名的「莫札特耳」。莫札特的左耳缺少一般耳廓內的耳甲凹，這種罕見的先天畸形後來被叫作「莫札特耳」。圖上的「我的耳朵」和「一般的耳朵」是他的小兒子弗朗茨・克薩韋爾寫的，後來尼森把「我的」劃掉改成「莫札特的」。也有學者認為，這其實是他兒子的耳朵。後世曾有人從這個特徵推測他的健康問題，但這些推論始終只是推測。'),
  p('左下角的手稿像是家人的回憶紀錄，右下是版畫側面像；最上方是一幅穿紅色軍裝的肖像，反光太強，我認不出是誰。這些說明牌我都沒拍到清楚，請你日後補照或對照後告訴我，我再修正。'),
  p('展廳裡還有另一面牆，在古鋼琴上方掛了三幅肖像：左邊是兩個依偎的孩子，中間是一位穿白色薄紗、手拿書本的女士，右邊是穿深色大衣的男士。我不確定他們是誰，可能與莫札特家族或其後人有關，請以說明牌為準。'),
  ...img('v_portraits', ...P, '古鋼琴上方的三幅肖像（人物身分待確認）'),
  verse('義人的記念，必被稱頌。', '箴言 10:7'),
  heart('我們一路讀著這些書、信與畫，才發現「記得一個人」本身就是一件需要用心的事：誰在記、怎麼記、記得是否真實。我們也想成為這樣的人——不只活出自己的故事，也真誠地記住別人的故事。'),

  h1('旅人小記'),
  bullet('莫札特故居與出生地都由莫札特基金會（Mozarteum）經營，持薩爾斯堡卡可以兩館都參觀；故居通常人比較少，建議預留約 1.5 小時。'),
  bullet('一樓的「舞蹈教師大廳」（Tanzmeistersaal）是一家人辦家庭音樂會的地方。'),
  bullet('參觀完可以走到對面的聖三一教堂，或去薩爾斯堡主座教堂看莫札特受洗的洗禮盆。'),
  bullet('看展前先聽一段《伊多梅紐斯》、《魔笛》和《Ave verum corpus》，看到手稿時會更有感覺。'),
  new Paragraph({ spacing: { before: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 8 } }, children: [new TextRun({ text: '參考資料：Britannica〈Mozart: Vienna〉、Mediathek.at〈Wiener Jahre 1781–1791〉、Wikipedia〈Michael Kelly (tenor)〉〈Georg Nikolaus von Nissen〉〈Appearance and character of Wolfgang Amadeus Mozart〉、Yale 音樂圖書館〈Nissen Biographie〉、Braunbehrens《Mozart in Vienna 1781–1791》、莫札特基金會官網、Wikipedia〈Portrait of the Mozart Family〉〈Tanzmeisterhaus Salzburg〉〈Libretto of The Magic Flute〉、Projekt Gutenberg《莫札特書信集》、薩爾斯堡旅遊局、大都會歌劇院《伊多梅紐斯》介紹。經文引自中文和合本。', font: F, size: 18, color: GRAY })] }),

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
Packer.toBuffer(mk(body3)).then((b) => fs.writeFileSync('莫札特故居_終.docx', b));
