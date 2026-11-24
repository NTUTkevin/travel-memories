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
    pn = '<nav class="pn">'
    pn += f'<a href="#{prev["id"]}" data-go="{prev["id"]}"><small>上一篇</small>{esc(prev["title"])}</a>' if prev else '<span></span>'
    pn += f'<a class="r" href="#{nxt["id"]}" data-go="{nxt["id"]}"><small>下一篇</small>{esc(nxt["title"])}</a>' if nxt else '<a class="r" href="#home" data-go="home"><small>回到</small>全部文章</a>'
    pn += '</nav>'
    articles.append(f'''<article class="post" id="view-{p['id']}" hidden>
  <header class="posthead"><span class="place">{esc(p['place'])}</span><h1>{esc(p['title'])}</h1><p class="meta">{TRIPMAP[p['trip']][2]} · 約 {r['minutes']} 分鐘</p></header>
  <div class="postgrid">
    <nav class="toc" aria-label="本篇章節"><b>本篇章節</b><ol>{toc}</ol></nav>
    <div class="prose">
{r['body']}
{map_html(p['id'])}
{pn}
    </div>
  </div>
</article>''')

trips_html = ''
for k,n,d,m in TRIPS:
    cs = [c for c,p in zip(cards,POSTS) if p['trip']==k]
    if cs: trips_html += f'  <section class="trip">\n    <div class="triphead"><h2>{n}</h2><span>{d}</span></div>\n    <div class="grid">\n' + '\n'.join(cs) + '\n    </div>\n  </section>\n'
TEMPLATE = open('/home/claude/site/template.html', encoding='utf-8').read()
out = TEMPLATE.replace('{{TRIPS}}', trips_html).replace('{{ARTICLES}}', '\n'.join(articles))
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
open(f'{OUT}/index.html', 'w', encoding='utf-8').write(out)
print(os.path.getsize(f'{OUT}/index.html'), 'bytes html')
