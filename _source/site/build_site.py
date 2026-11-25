import os, re, shutil, glob, html
from bs4 import BeautifulSoup
from PIL import Image

SRC = '/tmp/bl'
OUT = '/tmp/site'
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(OUT + '/img')

POSTS = [
  dict(id='koenigssee', trip='koe', dirs=['國王湖'], title='國王湖：湖邊的步道，和新認識的朋友', place='國王湖 Königssee', cover=7),
  dict(id='bartholomae', trip='koe', dirs=['聖巴多羅買教堂'], title='聖巴多羅買教堂：湖上的紅色洋蔥頂', place='聖巴多羅買教堂 St. Bartholomä', cover=3, videos=[('v1','從湖上看聖巴多羅買教堂','11 月 19 日。紅色洋蔥頂的教堂與湖（地點依畫面判斷）'),('v3','船上的窗景','船上看著窗外的山與湖')]),
  dict(id='ramsau', trip='koe', dirs=['拉姆紹教堂'], title='拉姆紹教堂：霧裡的洋蔥頂', place='拉姆紹教堂 Ramsau St. Sebastian', cover=5),
  dict(id='hallstatt', trip='hal', dirs=['哈修塔特鹽礦'], title='哈修塔特鹽礦：走進七千年的鹽山', place='哈修塔特鹽礦 Salzwelten Hallstatt', cover=9, videos=[('v5','走進克莉絲汀娜坑道','11 月 20 日。石拱門上寫著 1719 年開鑿'),('v6','坑道裡的太太','在岩壁前比手畫腳，好像在嚐鹽的味道'),('v7','坐礦車出礦坑','黑暗的坑道裡一路往外衝')]),
  dict(id='hallstatt-town', trip='hal', dirs=['哈修塔特小鎮'], title='哈修塔特小鎮：湖邊兩座教堂', place='哈修塔特小鎮 Hallstatt', cover=6, videos=[('v2','山間小村','11 月 20 日。雲和山谷之間的房子與遠處的雪山（確切地點待確認）')]),
  dict(id='residence', dirs=['莫札特故居_上','莫札特故居_下','莫札特故居_終'], title='走進莫札特的家', place='莫札特故居 Mozart-Wohnhaus', cover=0),
  dict(id='birthplace', dirs=['莫札特出生地'], title='糧食胡同 9 號：莫札特出生的地方', place='莫札特出生地 Mozarts Geburtshaus', cover=0),
  dict(id='fortress', dirs=['薩爾斯堡_要塞'], title='霍亨薩爾斯堡要塞與「白色黃金」', place='霍亨薩爾斯堡要塞 Festung Hohensalzburg', cover=1, videos=[('v4','要塞與遠山','11 月 21 日。要塞的塔樓、城外的山景，最後停在莫札特出生地的大門')]),
  dict(id='church', dirs=['薩爾斯堡_大學教堂'], title='薩爾斯堡大學教堂', place='大學教堂 Kollegienkirche', cover=0),
  dict(id='moderne', dirs=['薩爾斯堡_現代美術館'], title='薩爾斯堡現代美術館：山上的現代藝術', place='現代美術館 Museum der Moderne', cover=5),
  dict(id='christmas', dirs=['薩爾斯堡聖誕市集'], title='薩爾斯堡聖誕市集：五百年的冬夜燈火', place='聖誕市集 Salzburger Christkindlmarkt', cover=1, videos=[('v8','僧侶山上的夜景','11 月 21 日傍晚，從僧侶山看老城和要塞'),('v9','乾杯！','市集的熱飲杯'),('v10','莫札特廣場溜冰場','Eiszauber 溜冰場的燈光'),('v11','回程的小巷','掛著聖誕燈飾的老城巷子')]),
  dict(id='vienna-home', trip='vie', dirs=['維也納_中央車站旁的家'], title='抵達維也納：中央車站旁的家', place='維也納中央車站 Wien Hauptbahnhof', cover=4),
  dict(id='schoenbrunn-out', trip='vie', dirs=['維也納_美泉宮_室外'], title='美泉宮（室外）：凱旋門、花園與維也納的第一場雪', place='美泉宮凱旋門 Gloriette', cover=4, videos=[('v12','聖誕樹下接雪花','11 月 23 日 13:51，美泉宮前的聖誕市集下起了雪')]),
  dict(id='schoenbrunn-in', trip='vie', dirs=['維也納_美泉宮_室內'], title='美泉宮（室內）：走進皇帝與西西的夏宮', place='美泉宮 Schloss Schönbrunn', cover=7),
  dict(id='cafe-birthday', trip='vie', dirs=['維也納_咖啡館裡的生日'], title='咖啡館裡的生日：維也納的下午與夜晚', place='Café Diglas・卡爾廣場聖誕市集', cover=0, videos=[('v15','太太的蛋糕時間','11 月 23 日 14:53，Café Diglas'),('v16','生日的印度朋友','15:01，在咖啡館併桌聊天')]),
  dict(id='musikverein', trip='vie', dirs=['維也納_金色大廳莫札特音樂會'], title='金色大廳：在音樂的神殿聽莫札特', place='愛樂協會金色大廳 Musikverein', cover=3, videos=[('v13','全場起立鼓掌','11 月 23 日 22:01，演出結束後的掌聲'),('v14','穿古裝的謝幕','22:03，樂手和歌手在台上一起謝幕')]),
  dict(id='hofburg', trip='vie', dirs=['維也納_霍夫堡與西西博物館'], title='霍夫堡與西西博物館：皇帝的辦公室，皇后的傳說', place='霍夫堡 Hofburg・西西博物館 Sisi Museum', cover=1, videos=[('v17','謁見等候廳的三幅巨畫','11 月 24 日 10:12，克拉夫特畫的法蘭茲一世與大水晶吊燈')]),
  dict(id='moebel', trip='vie', dirs=['維也納_家具博物館'], title='維也納家具博物館：皇帝的家具倉庫與《茜茜公主》的片場', place='家具博物館 Möbelmuseum Wien', cover=2),
  dict(id='sisipass', trip='vie', dirs=['維也納_SisiPass三館'], title='Sisi Pass：一張票，跟著西西走三個地方', place='Sisi Pass 美泉宮・西西博物館・家具博物館', cover=1),
  dict(id='rathaus-xmas', trip='vie', dirs=['維也納_市政廳聖誕市集'], title='市政廳聖誕市集：維也納最盛大的聖誕夜色', place='市政廳聖誕市集 Wiener Christkindlmarkt', cover=0, videos=[('v18','星星樹下的太太','17:39，市政廳公園裡掛滿星星的燈樹，最後鏡頭帶到市政廳鐘樓'),('v19','熊樂團','17:41，公園小木屋裡會動的熊樂團，唱歌、吹喇叭、彈吉他'),('v20','還願教堂的「光的無限」','18:00，還願教堂兩座塔之間的紅色無限符號')]),
]

TRIPS = [('koe','德國・國王湖與拉姆紹','2026 年 11 月 19 日','2026 年 11 月 19 日 · 德國・國王湖'),('hal','奧地利・哈修塔特','2026 年 11 月 20 日','2026 年 11 月 20 日 · 哈修塔特'),('salz','奧地利・薩爾斯堡','2026 年 11 月 21 日','2026 年 11 月 21 日 · 薩爾斯堡'),('vie','奧地利・維也納','2026 年 11 月 22 日起','2026 年 11 月 22 日起 · 維也納')]
TRIPMAP = {k:(n,d,m) for k,n,d,m in TRIPS}
for _p in POSTS: _p.setdefault('trip','salz')
shutil.copytree('/home/claude/site/vid', OUT + '/vid')

def esc(s): return html.escape(s, quote=True)

# 每篇文章的地點（Google Maps）：第一個是地圖上顯示的主要地點，其餘只列連結
MAPS = {
  'koenigssee': [('國王湖碼頭 Seelände Schönau', 'Königssee Seelände, Schönau am Königssee')],
  'bartholomae': [('聖巴多羅買教堂 St. Bartholomä', 'St. Bartholomä, Königssee')],
  'ramsau': [('拉姆紹聖塞巴斯蒂安教堂', 'Pfarrkirche St. Sebastian, Ramsau bei Berchtesgaden')],
  'hallstatt': [('哈修塔特鹽礦 Salzwelten Hallstatt', 'Salzwelten Hallstatt'), ('登山纜車山下站 Salzbergbahn', 'Salzbergbahn Talstation Hallstatt')],
  'hallstatt-town': [('哈修塔特市集廣場', 'Marktplatz Hallstatt'), ('新教基督教堂', 'Evangelische Pfarrkirche Hallstatt'), ('天主教堂與人骨堂', 'Beinhaus Hallstatt'), ('明信片觀景點', 'Hallstatt Viewpoint Seestraße')],
  'residence': [('莫札特故居 Mozart-Wohnhaus', 'Mozart-Wohnhaus, Makartplatz 8, Salzburg')],
  'birthplace': [('莫札特出生地 Getreidegasse 9', 'Mozarts Geburtshaus, Getreidegasse 9, Salzburg')],
  'fortress': [('霍亨薩爾斯堡要塞', 'Festung Hohensalzburg, Salzburg')],
  'church': [('大學教堂 Kollegienkirche', 'Kollegienkirche, Universitätsplatz, Salzburg')],
  'moderne': [('現代美術館 Mönchsberg', 'Museum der Moderne Mönchsberg, Salzburg')],
  'vienna-home': [('維也納中央車站 Wien Hauptbahnhof', 'Wien Hauptbahnhof')],
  'schoenbrunn-out': [('凱旋門 Gloriette', 'Gloriette, Schönbrunn, Wien'), ('美泉宮聖誕市集（榮譽庭院）', 'Weihnachtsmarkt Schloss Schönbrunn, Wien')],
  'schoenbrunn-in': [('美泉宮 Schloss Schönbrunn', 'Schloss Schönbrunn, Wien')],
  'cafe-birthday': [('Café Diglas', 'Café Diglas, Wollzeile 10, Wien'), ('卡爾廣場聖誕市集 Art Advent', 'Art Advent Karlsplatz, Wien')],
  'musikverein': [('維也納愛樂協會 Musikverein', 'Musikverein Wien, Musikvereinsplatz 1')],
  'hofburg': [('西西博物館（霍夫堡）', 'Sisi Museum Hofburg, Michaelerkuppel, Wien'), ('米歇爾廣場羅馬遺跡', 'Ausgrabungen Michaelerplatz, Wien'), ('聖米迦勒教堂', 'Michaelerkirche, Michaelerplatz, Wien')],
  'moebel': [('維也納家具博物館 Möbelmuseum Wien', 'Möbelmuseum Wien, Andreasgasse 7, 1070 Wien')],
  'rathaus-xmas': [('維也納市政廳聖誕市集', 'Wiener Christkindlmarkt Rathausplatz, Wien'), ('愛心樹（市政廳公園）', 'Rathauspark Wien'), ('城堡劇院 Burgtheater', 'Burgtheater, Universitätsring 2, Wien'), ('還願教堂 Votivkirche', 'Votivkirche, Rooseveltplatz, Wien')],
  'sisipass': [('美泉宮 Schönbrunn', 'Schloss Schönbrunn, Wien'), ('西西博物館（霍夫堡）', 'Sisi Museum Hofburg, Michaelerkuppel, Wien'), ('維也納家具博物館', 'Möbelmuseum Wien, Andreasgasse 7, 1070 Wien')],
  'christmas': [('聖誕市集（大教堂廣場）', 'Salzburger Christkindlmarkt, Domplatz, Salzburg'), ('主教宮廣場 Residenzplatz', 'Residenzplatz, Salzburg'), ('莫札特廣場溜冰場', 'Mozartplatz, Salzburg'), ('僧侶山觀景台', 'Museum der Moderne Mönchsberg, Salzburg')],
}
from urllib.parse import quote_plus as _q
def map_html(pid):
    places = MAPS.get(pid)
    if not places: return ''
    name, q = places[0]
    emb = 'https://www.google.com/maps?q=' + _q(q) + '&z=15&hl=zh-TW&output=embed'
    links = ''.join(f'<li><a href="https://www.google.com/maps/search/?api=1&amp;query={_q(qq)}" target="_blank" rel="noopener">{esc(nn)}</a></li>' for nn, qq in places)
    return (f'<section class="mapbox"><h2>地點地圖</h2>'
            f'<div class="mapframe"><iframe data-src="{esc(emb)}" title="{esc(name)} 的 Google 地圖" loading="lazy" referrerpolicy="no-referrer-when-downgrade" allowfullscreen></iframe>'
            f'<p class="mapnote" hidden>這個預覽頁不能嵌入地圖，請點下面的連結在 Google Maps 開啟。</p></div>'
            f'<ul class="maplinks">{links}</ul><p class="mapsmall">點地點名稱會在 Google Maps 開啟（手機會打開 Google Maps App）。</p></section>')

imgcount = {}
def convert(post):
    parts = []
    first_para = None
    covers = []
    toc = []
    nimg = 0
    for di, d in enumerate(post['dirs']):
        soup = BeautifulSoup(open(f'{SRC}/{d}/body.html', encoding='utf-8').read(), 'html.parser')
        els = [e for e in soup.contents if getattr(e, 'name', None)]
        # drop title block
        if els and els[0].name == 'p' and els[0].get_text(strip=True) == '旅行的生活紀錄':
            els = els[3:]
        i = 0
        while i < len(els):
            e = els[i]; n = e.name
            if n == 'h1':
                t = e.get_text(strip=True)
                if t == '旅人小記' and post.get('videos'):
                    hid = f"{post['id']}-s{len(toc)+1}"; toc.append((hid, '影片'))
                    parts.append(f'<h2 id="{hid}">影片</h2>')
                    parts.append('<div class="vidrow">' + ''.join(f'<figure class="vid"><video controls playsinline preload="none" poster="vid/{v}.jpg" src="vid/{v}.mp4"></video><figcaption><b>{esc(a)}</b><br>{esc(b)}</figcaption></figure>' for v,a,b in post['videos']) + '</div>')
                hid = f"{post['id']}-s{len(toc)+1}"
                toc.append((hid, t))
                parts.append(f'<h2 id="{hid}">{esc(t)}</h2>')
            elif n == 'h2':
                parts.append(f'<h3>{esc(e.get_text(strip=True))}</h3>')
            elif n == 'h3':
                parts.append(f'<h4>{esc(e.get_text(strip=True))}</h4>')
            elif n == 'p':
                img = e.find('img')
                txt = e.get_text(strip=True)
                if img:
                    src = img['src']
                    fn = os.path.basename(src)
                    srcpath = f'{SRC}/{d}/media/{fn}'
                    nimg += 1
                    name = f"{post['id']}-{nimg:02d}.jpg"
                    shutil.copy(srcpath, f'{OUT}/img/{name}')
                    w, h = Image.open(srcpath).size
                    cap = ''
                    if i+1 < len(els) and els[i+1].name == 'p' and els[i+1].find('em') and not els[i+1].find('img'):
                        cap = els[i+1].get_text(strip=True); i += 1
                    cls = 'portrait' if h > w else 'landscape'
                    parts.append(f'<figure class="{cls}"><img src="img/{name}" width="{w}" height="{h}" loading="lazy" alt="{esc(cap)}"><figcaption>{esc(cap)}</figcaption></figure>')
                    covers.append(name)
                elif txt.startswith('參考資料'):
                    parts.append(f'<p class="refs">{esc(txt)}</p>')
                elif txt.startswith('我們的感動'):
                    body = txt[len('我們的感動'):].strip()
                    parts.append(f'<aside class="heart"><span>我們的感動</span><p>{esc(body)}</p></aside>')
                elif '（和合本）' in txt and e.find('em'):
                    em = e.find('em').get_text(strip=True)
                    ref = txt[len(em):].replace('（和合本）','').replace('——','').strip(' 　')
                    parts.append(f'<blockquote class="verse"><p>{esc(em)}</p><cite>{esc(ref)}（和合本）</cite></blockquote>')
                else:
                    if first_para is None and not toc: first_para = txt
                    parts.append(f'<p>{esc(txt) if not e.find(True) else e.decode_contents()}</p>')
            elif n == 'blockquote':
                txt = e.get_text(' ', strip=True)
                if txt.startswith('我們的感動'):
                    body = txt[len('我們的感動'):].strip()
                    parts.append(f'<aside class="heart"><span>我們的感動</span><p>{esc(body)}</p></aside>')
                else:
                    parts.append(f'<blockquote class="quote">{e.decode_contents()}</blockquote>')
            elif n in ('ul', 'ol'):
                items = ''.join(f'<li>{li.get_text(strip=True)}</li>' for li in e.find_all('li'))
                parts.append(f'<{n}>{items}</{n}>')
            elif n == 'table':
                parts.append(f'<div class="tablewrap">{str(e)}</div>')
            else:
                parts.append(str(e))
            i += 1
    chars = sum(len(re.sub('<[^>]+>', '', x)) for x in parts)
    return dict(body='\n'.join(parts), first=first_para or '', covers=covers, toc=toc, minutes=max(1, round(chars/450)), nimg=nimg)

res = {}
for p in POSTS:
    r = convert(p); res[p['id']] = r
    print(p['id'], 'imgs', r['nimg'], 'sections', len(r['toc']), 'min', r['minutes'], 'first:', r['first'][:30])

# summary override text (short, in the writer's voice)
SUMMARY = {
  'cafe-birthday': '1923 年的老咖啡館 Diglas 沒有位子，併桌認識了一個人來維也納慶生的印度朋友；晚餐一大盤肋排，再到卡爾廣場的聖誕市集看小羊。',
  'musikverein': '晚上走進新年音樂會的那座金色大廳：1812 年的音樂之友協會、漢森設計的「音樂神殿」、女像柱與繆思天花板，還有穿著古裝、戴假髮演出的莫札特音樂會。',
  'hofburg': '走進哈布斯堡的冬宮：克拉夫特畫的「好皇帝法蘭茲」、法蘭茲・約瑟夫 1910 年的謁見名單、西西的體操室，剛重新開幕的西西博物館，還有 30 公尺長的米蘭桌飾。',
  'moebel': '瑪麗亞・特蕾莎 1747 年設立的宮廷家具倉庫：皇帝的寶座、皇太子魯道夫的搖籃、畢德麥雅的小房間，還有《茜茜公主》電影裡用過的真家具。',
  'rathaus-xmas': '維也納最大的聖誕市集：從七百年前的「十二月市集」說起，市政廳人和比還願教堂矮一公尺的塔、愛心樹、來自福拉爾貝格的 31 公尺聖誕樹，還有還願教堂上的「光的無限」。',
  'sisipass': '一張 Sisi Pass 走完美泉宮、霍夫堡西西博物館和家具博物館：包含什麼、怎麼預約、2026 新票價季算一算划不划算，還有我們的路線。',
  'schoenbrunn-out': '爬上山丘頂的凱旋門：1775 年為「正義之戰」蓋的榮耀殿堂，從上面俯瞰黃色的宮殿和整座維也納；下午回到宮殿前的聖誕市集，遇上了維也納的第一場雪。',
  'schoenbrunn-in': '哈布斯堡的夏宮：特羅格畫的〈瑪利亞的婚禮〉、法蘭茲・約瑟夫和西西的房間、六歲莫札特演奏過的鏡廳、43 公尺長的金色大廳，還有中國漆器廳與百萬廳。',
  'vienna-home': '旅程最後一站。住在中央車站旁的頂樓公寓：從舊南站到菱形屋頂的新車站、「工人區」法沃里滕、「便宜商店」BILLA 的故事，和太太煮的曾拌麵晚餐。',
  'christmas': '從 1491 年的雜貨市集開始的聖誕市集。僧侶山夜景、一杯熱紅酒和紀念杯、莫札特廣場溜冰場，還有《平安夜》作詞者在薩爾斯堡的童年。',
  'hallstatt-town': '「鹽的地方」湖邊小鎮。1750 年的大火、為信仰半夜上船離鄉的新教徒、山坡上的人骨堂，還有明信片角度的合照。',
  'hallstatt': '世界最古老的鹽礦。新纜車上山、1719 年的克莉絲汀娜坑道、三千年前的木樓梯、各色岩鹽，最後坐礦車回到積雪的山前。',
  'residence': '莫札特一家住了二十多年的家。從全家福牆上的母親，到一封寫給父親的信，再到維也納的十年與身後的故事。',
  'birthplace': '糧食胡同 9 號的黃色大樓，莫札特 1756 年在這裡出生。小提琴、鋼琴與黑絨枕上的小物，還有他是個怎樣的人。',
  'fortress': '九百多年的教會要塞。鹽的財富、牢房、萊納團的戰史、1731 年被驅逐的新教徒，以及最頂樓的全景。',
  'church': '舊城廣場邊一座純白的巴洛克教堂，為大學而蓋，也經歷過戰爭與修復。',
  'bartholomae': '只有船能到的半島。1134 年的教堂、紅色洋蔥頂、祭壇，還有背後 1,800 公尺高的瓦茨曼東壁。',
  'ramsau': '1512 年蓋在運鹽古道旁的教堂，溪邊的畫家角度、墓園裡的登山嚮導與《平安夜》作詞人，還有牆上的日晷。',
  'koenigssee': '湖邊小鎮、只准電動船的湖、步道上的觀景台，還有第二天認識的新朋友宜恩。',
  'moderne': '僧侶山上的現代美術館。新即物主義展裡的鐵路看守人、被禁止的畫家，還有暴風雨前的博登湖。',
}

cards = []
articles = []
for idx, p in enumerate(POSTS):
    r = res[p['id']]
    cover = r['covers'][p['cover']] if r['covers'] else ''
    cards.append(f'''<a class="card" href="#{p['id']}" data-go="{p['id']}">
  <div class="cardimg"><img src="img/{cover}" alt="" loading="lazy"></div>
  <div class="cardtxt"><span class="place">{esc(p['place'])}</span><h3>{esc(p['title'])}</h3><p>{esc(SUMMARY[p['id']])}</p><span class="meta">約 {r['minutes']} 分鐘 · {r['nimg']} 張照片</span></div></a>''')
    toc = ''.join(f'<li><button type="button" data-scroll="{hid}">{esc(t)}</button></li>' for hid, t in r['toc'])
    prev = POSTS[idx-1] if idx > 0 else None
    nxt = POSTS[idx+1] if idx < len(POSTS)-1 else None
    def _cv(q):
        rr = res[q['id']]; return rr['covers'][q['cover']] if rr['covers'] else ''
    pn = '<nav class="pn" aria-label="上一篇與下一篇">'
    pn += f'<a href="#{prev["id"]}" data-go="{prev["id"]}"><img src="img/{_cv(prev)}" alt="" loading="lazy"><span><small>上一篇</small>{esc(prev["title"])}</span></a>' if prev else '<span></span>'
    pn += f'<a class="r" href="#{nxt["id"]}" data-go="{nxt["id"]}"><span><small>下一篇</small>{esc(nxt["title"])}</span><img src="img/{_cv(nxt)}" alt="" loading="lazy"></a>' if nxt else '<a class="r" href="#home" data-go="home"><span><small>回到</small>全部文章</span></a>'
    pn += '</nav>'
    tn = '<nav class="tn" aria-label="快速切換文章">'
    tn += f'<a href="#{prev["id"]}" data-go="{prev["id"]}">‹ 上一篇</a>' if prev else '<span></span>'
    tn += f'<button type="button" class="share" data-share="{p["id"]}" data-title="{esc(p["title"])}">分享這篇</button>'
    tn += f'<a href="#{nxt["id"]}" data-go="{nxt["id"]}">下一篇 ›</a>' if nxt else '<a href="#home" data-go="home">全部文章 ›</a>'
    tn += '</nav>'
    articles.append(f'''<article class="post" id="view-{p['id']}" hidden>
  <header class="posthead"><span class="place">{esc(p['place'])}</span><h1>{esc(p['title'])}</h1><p class="meta">{TRIPMAP[p['trip']][2]} · 約 {r['minutes']} 分鐘</p>{tn}</header>
  <div class="postgrid">
    <nav class="toc" aria-label="本篇章節"><b>本篇章節</b><ol>{toc}</ol></nav>
    <div class="prose">
{r['body']}
{map_html(p['id'])}
{pn}
    </div>
  </div>
</article>''')

SHORT = {'koenigssee':'國王湖','bartholomae':'聖巴多羅買教堂','ramsau':'拉姆紹教堂','hallstatt':'哈修塔特鹽礦','hallstatt-town':'哈修塔特小鎮','residence':'莫札特故居','birthplace':'莫札特出生地','fortress':'霍亨薩爾斯堡要塞','church':'大學教堂','moderne':'現代美術館','christmas':'薩爾斯堡聖誕市集','vienna-home':'中央車站旁的家','schoenbrunn-out':'美泉宮（室外）','schoenbrunn-in':'美泉宮（室內）','cafe-birthday':'咖啡館裡的生日','musikverein':'金色大廳莫札特音樂會','hofburg':'霍夫堡與西西博物館','moebel':'家具博物館','sisipass':'Sisi Pass 三館攻略','rathaus-xmas':'市政廳聖誕市集'}
trips_html = ''
quick_html = '  <nav class="quick" aria-label="快速目錄"><b class="qt">快速目錄</b>\n'
for k,n,d,m in TRIPS:
    cs = [c for c,p in zip(cards,POSTS) if p['trip']==k]
    ps = [p for p in POSTS if p['trip']==k]
    if ps: quick_html += f'    <div class="qrow"><span class="qtrip">{n}</span>' + ''.join(f'<a href="#{p["id"]}" data-go="{p["id"]}">{esc(SHORT.get(p["id"], p["title"]))}</a>' for p in ps) + '</div>\n'
    if cs: trips_html += f'  <section class="trip">\n    <div class="triphead"><h2>{n}</h2><span>{d}</span></div>\n    <div class="grid">\n' + '\n'.join(cs) + '\n    </div>\n  </section>\n'

# ===== 旅程路線地圖（首頁）：國界是 Natural Earth（world-atlas）資料，事先投影成 SVG =====
import json as _json
_RM = _json.load(open('/home/claude/site/routemap.json', encoding='utf-8'))
_LBL = {'koe': (-12, 24, 'end'), 'salz': (-12, -12, 'end'), 'hal': (12, 22, 'start'), 'vie': (0, -18, 'middle'), 'bud': (0, 28, 'middle')}
def route_html():
    W, H = _RM['W'], _RM['H']
    cs = {c['key']: c for c in _RM['cities']}
    order = [k for k, *_ in TRIPS if k in cs]
    svg = [f'<svg class="rmap" viewBox="85 95 655 210" role="img" aria-label="旅程路線地圖：' + ' → '.join(cs[k]['name'] for k in order) + '">']
    for c in _RM['countries']:
        cls = 'at' if c['name'] == 'Austria' else ('de' if c['name'] == 'Germany' else 'oth')
        svg.append(f'<path class="cty {cls}" d="{c["d"]}"/>')
    for c in _RM['countries']:
        if c['label'] and c['zh']:
            svg.append(f'<text class="ctl{" atl" if c["name"]=="Austria" else ""}" x="{c["label"][0]}" y="{c["label"][1]}">{c["zh"]}</text>')
    pts = ' '.join(f'{cs[k]["xy"][0]},{cs[k]["xy"][1]}' for k in order)
    svg.append(f'<polyline class="rline" points="{pts}"/>')
    nx = [c for c in _RM['cities'] if c.get('next')]
    for c in nx:
        a = cs[order[-1]]['xy']; svg.append(f'<line class="rnext" x1="{a[0]}" y1="{a[1]}" x2="{c["xy"][0]}" y2="{c["xy"][1]}"/>')
    for i, k in enumerate(order + [c['key'] for c in nx]):
        c = cs[k]; x, y = c['xy']; dx, dy, an = _LBL.get(k, (0, -16, 'middle'))
        num = '?' if c.get('next') else str(i + 1)
        svg.append(f'<g class="stop{" nx" if c.get("next") else ""}" data-stop="{k}" tabindex="0" role="button" aria-label="{c["name"]} {c["date"]}">'
                   f'<circle class="hit" cx="{x}" cy="{y}" r="22"/><circle class="dot" cx="{x}" cy="{y}" r="11"/><text class="num" x="{x}" y="{y+4.5}">{num}</text>'
                   f'<text class="sl" x="{x+dx}" y="{y+dy}" text-anchor="{an}">{c["name"]}<tspan class="sd" dx="6">{c["date"]}</tspan></text></g>')
    svg.append('</svg>')
    panels = ''
    for k in order + [c['key'] for c in nx]:
        c = cs[k]; ps = [q for q in POSTS if q['trip'] == k]
        if ps:
            chips = ''.join(f'<a href="#{q["id"]}" data-go="{q["id"]}">{esc(SHORT.get(q["id"], q["title"]))}</a>' for q in ps)
            panels += f'<div class="rpanel" data-panel="{k}" hidden><b>{c["name"]} {c["de"]}<span>{c["date"]} · {len(ps)} 篇</span></b><div class="rchips">{chips}</div></div>'
        else:
            panels += f'<div class="rpanel" data-panel="{k}" hidden><b>{c["name"]} {c["de"]}<span>{c["date"]}</span></b><p>下一站，照片還在路上。</p></div>'
    return ('<section class="route"><div class="triphead"><h2>旅程路線</h2><span>點地圖上的城市，看那一站的文章</span></div>'
            '<div class="rbox">' + ''.join(svg) + '</div>' + panels +
            '<p class="rhint" data-panel="none">點上面的 ①②③④，就會列出那一站的文章。</p>'
            '<p class="rsrc">地圖國界：Natural Earth（公有領域）。</p></section>\n')
TEMPLATE = open('/home/claude/site/template.html', encoding='utf-8').read()
quick_html += '  </nav>\n'
out = TEMPLATE.replace('{{QUICK}}', quick_html + route_html()).replace('{{TRIPS}}', trips_html).replace('{{ARTICLES}}', '\n'.join(articles))
# 快取更新：每次重建都給圖片與影片加上版本號，並請瀏覽器重新檢查網頁
import re as _re, time as _time
import datetime as _dt, zoneinfo as _zi
_now = _dt.datetime.now(_zi.ZoneInfo('Europe/Vienna'))
_ver = _now.strftime('%Y%m%d%H%M')
out = out.replace('__UPDATED__', f'{_now.year} 年 {_now.month} 月 {_now.day} 日 {_now:%H:%M}')
out = _re.sub(r'((?:img|vid)/[^"?\s]+\.(?:jpg|jpeg|png|mp4))(?=")', r'\1?v=' + _ver, out)
_meta = '<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate"><meta http-equiv="Pragma" content="no-cache"><meta http-equiv="Expires" content="0">'
if 'http-equiv="Cache-Control"' not in out:
    out = out.replace('<meta charset="utf-8">', '<meta charset="utf-8">' + _meta, 1)

# ===== 分享預覽：每篇一個 p/<id>.html，給 LINE／臉書讀封面、標題與簡介，打開後自動跳到文章 =====
BASE = 'https://ntutkevin.github.io/travel-memories/'
os.makedirs(f'{OUT}/p', exist_ok=True); os.makedirs(f'{OUT}/og', exist_ok=True)
def og_image(src, dst):
    im = Image.open(src).convert('RGB'); w, h = im.size
    tw = min(w, 1200); th = round(tw / 1.91)
    sc = max(tw / w, th / h); im = im.resize((round(w * sc), round(h * sc)), Image.LANCZOS)
    w, h = im.size; x = (w - tw) // 2; y = int((h - th) * 0.35)
    im.crop((x, y, x + tw, y + th)).save(dst, quality=82)
    return tw, th
def share_page(pid, title, desc, img, iw, ih, target):
    t, d = esc(title), esc(desc)
    return f"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{t}</title><meta name="description" content="{d}">
<meta property="og:type" content="article"><meta property="og:site_name" content="我們的旅行回憶">
<meta property="og:title" content="{t}"><meta property="og:description" content="{d}">
<meta property="og:url" content="{BASE}p/{pid}.html"><meta property="og:image" content="{BASE}og/{img}?v={_ver}">
<meta property="og:image:width" content="{iw}"><meta property="og:image:height" content="{ih}"><meta property="og:locale" content="zh_TW">
<meta name="twitter:card" content="summary_large_image">
<style>body{{font-family:serif;background:#eceff2;color:#1c2330;text-align:center;padding:40px 20px}}img{{max-width:100%;height:auto}}a{{color:#8a6a2f}}</style>
<script>location.replace('../{target}');</script></head>
<body><p><img src="../og/{img}" alt=""></p><h1 style="font-size:22px">{t}</h1><p><a href="../{target}">繼續閱讀 →</a></p></body></html>"""
_home_cover = None
for p in POSTS:
    r = res[p['id']]
    if not r['covers']: continue
    cv = r['covers'][p['cover']]
    iw, ih = og_image(f'{OUT}/img/{cv}', f'{OUT}/og/{p["id"]}.jpg')
    open(f'{OUT}/p/{p["id"]}.html', 'w', encoding='utf-8').write(share_page(p['id'], p['title'] + '｜我們的旅行回憶', SUMMARY[p['id']], f'{p["id"]}.jpg', iw, ih, '#' + p['id']))
    if p['id'] == 'schoenbrunn-out': _home_cover = (p['id'], iw, ih)
if not _home_cover: _home_cover = (POSTS[-1]['id'], iw, ih)
_home_desc = '我和太太一起走過的地方：德國國王湖、哈修塔特、薩爾斯堡、維也納。每個景點先查歷史，再放上照片、一節經文和那一刻的感動。'
_og_home = (f'<meta name="description" content="{esc(_home_desc)}"><meta property="og:type" content="website"><meta property="og:site_name" content="我們的旅行回憶">'
            f'<meta property="og:title" content="我們的旅行回憶"><meta property="og:description" content="{esc(_home_desc)}"><meta property="og:url" content="{BASE}">'
            f'<meta property="og:image" content="{BASE}og/{_home_cover[0]}.jpg?v={_ver}"><meta property="og:image:width" content="{_home_cover[1]}"><meta property="og:image:height" content="{_home_cover[2]}">'
            '<meta property="og:locale" content="zh_TW"><meta name="twitter:card" content="summary_large_image">')
out = out.replace('<meta name="viewport"', _og_home + '<meta name="viewport"', 1)
open(f'{OUT}/index.html', 'w', encoding='utf-8').write(out)
print(os.path.getsize(f'{OUT}/index.html'), 'bytes html')
