# -*- coding: utf-8 -*-
"""実施割当と日程をA4のPDFに出力する（2026-10-07）

schedule.xlsx の「実施割当」「日程」シートを、印刷して書き込める形にしたもの。
日付が未定なので、日付欄は空けてある。決まったら手で書き込むか、
schedule.xlsx に入れてからこれを作り直す。

游ゴシックをTTCから取り出し、使う文字だけにサブセット化して埋め込む。
（サブセット化しないとPDFが28MBになる。th に background-color を指定すると
　Story が無関係な位置に灰色の矩形を描き残すので、罫線だけにする）
"""
import os, re, pathlib, datetime
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "実施割当と日程_2026-10-07.pdf")
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_plan_v2'
FONTDIR.mkdir(parents=True, exist_ok=True)

WEEK = "月火水木金土日"
DATES = [(10, 13), (10, 14), (10, 15), (10, 16), (10, 19), (10, 20), (10, 23),
         (11, 2), (11, 4), (11, 5), (11, 6), (11, 9), (11, 10), (11, 11), (11, 12), (11, 13)]

CSS = """
@font-face { font-family: jp; src: url(jp-regular.ttf); }
@font-face { font-family: jp; font-weight: bold; src: url(jp-bold.ttf); }
* { font-family: jp; }
body { font-size: 9pt; line-height: 1.45; color: #111; }
h1 { font-size: 14.5pt; margin: 0 0 1pt 0; }
.sub { font-size: 8.5pt; color: #555; margin: 0 0 10pt 0; }
h2 { font-size: 10.5pt; margin: 12pt 0 4pt 0; padding: 0 0 2pt 0; border-bottom: 1pt solid #444; }
p { margin: 0 0 4pt 0; }
table { width: 100%; border-collapse: collapse; margin: 3pt 0 6pt 0; font-size: 8.5pt; }
th { border: 0.6pt solid #888; border-bottom: 1.2pt solid #555; padding: 3pt 4pt; text-align: left; font-weight: bold; }
td { border: 0.6pt solid #888; padding: 3pt 4pt; vertical-align: top; }
td.c { text-align: center; }
td.w { height: 15pt; }
.note { font-size: 8.3pt; color: #444; margin: 0 0 5pt 0; }
.warn { font-size: 8.5pt; color: #a00; margin: 0 0 5pt 0; }
.kv { margin: 0 0 3pt 0; }
"""

# ---- §2 実施割当 ----------------------------------------------------------
rows = []
for team in (1, 2, 3, 4):
    gun = 1 if team % 2 == 1 else 2
    noon = ("ステッパー", "A", "B") if gun == 1 else ("テンキー", "D", "C")
    eve = ("テンキー", "D", "C") if gun == 1 else ("ステッパー", "A", "B")
    who = "p%02d〜p%02d" % (team * 3 - 2, team * 3)
    rows.append(("昼（空いている）", "組%d" % team, "群%d" % gun) + noon + (who,))
    rows.append(("夕方（混んでいる）", "組%d" % team, "群%d" % gun) + eve + (who,))

assign = []
for i, (waku, team, gun, hoshiki, p1, p2, who) in enumerate(rows, start=1):
    assign.append(
        '<tr><td class="c"><b>%d</b></td><td>%s</td><td class="c">%s</td><td class="c">%s</td>'
        '<td class="c">%s</td><td class="c"><b>%s</b></td><td class="c"><b>%s</b></td>'
        '<td class="c">%s</td><td class="w"></td></tr>'
        % (i, waku, team, gun, hoshiki, p1, p2, who))
for i in (9, 10):
    assign.append(
        '<tr><td class="c"><b>%d</b></td><td>昼＋夕方（4便）</td><td class="c">―</td><td class="c">―</td>'
        '<td class="c">―</td><td class="c">―</td><td class="c">―</td>'
        '<td class="c">計数役3名のみ</td><td class="w"></td></tr>' % i)

# ---- §3 日付ごとの記入表 ---------------------------------------------------
dates = []
for m, d in DATES:
    dt = datetime.date(2026, m, d)
    dates.append('<tr><td class="c"><b>%d/%d</b></td><td class="c">%s</td>'
                 '<td class="w"></td><td class="w"></td><td class="w"></td></tr>'
                 % (m, d, WEEK[dt.weekday()]))

HTML = """
<h1>実施割当と日程</h1>
<p class="sub">2026年10月7日　／　松本　　＊日付が決まったら §3 に書き込んでください。</p>

<h2>1. 使う便（どの日も同じ）</h2>
<table>
<tr><th>枠</th><th>集合</th><th>1本目</th><th>桃山台着</th><th>2本目</th><th>南口着・解散</th><th>拘束</th></tr>
<tr><td><b>昼</b>（空いている）</td><td class="c">11:35</td><td class="c">南口 <b>12:03</b>発</td><td class="c">12:48</td>
    <td class="c">桃山台 <b>13:48</b>発</td><td class="c">14:32</td><td class="c">約3時間</td></tr>
<tr><td><b>夕方</b>（混んでいる）</td><td class="c">15:25</td><td class="c">南口 <b>15:55</b>発</td><td class="c">16:40</td>
    <td class="c">桃山台 <b>17:18</b>発</td><td class="c">18:03</td><td class="c">約2時間40分</td></tr>
</table>
<p class="note">すべて阪急バス 吹田市内線2系統。集合・解散はどちらもＪＲ吹田駅（南口）。
4便とも始発便なので、計数役は車内0人から積み上げられます。</p>

<h2>2. 実施割当（全10日）</h2>
<p>日付以外はすべて決まっています。本実験が8日、計数方法の実験が2日です。</p>
<table>
<tr><th>回</th><th>枠</th><th>組</th><th>群</th><th>入力方式</th><th>1本目</th><th>2本目</th><th>参加者</th><th>日付</th></tr>
__ASSIGN__
</table>
<p class="note">1本目・2本目の欄はQRポスターの記号です。
<b>A＝ステッパー＋ボタンあり　B＝ステッパー＋ボタンなし　C＝テンキー＋ボタンあり　D＝テンキー＋ボタンなし</b></p>
<p class="note"><b>同じ組の2行（昼と夕方）は、同じ3名が別の日に来ます。</b>間が空いても構いません。
<b>組1・組3は昼がステッパー、組2・組4は昼がテンキー</b>です。揃えてしまうと、
入力方式の差なのか混雑の差なのか分からなくなります。
これでどの参加者も4乗車で「入力方式2通り × ボタン2通り」を1回ずつ経験します。</p>
<p class="note">回9・10（計数方法の実験）は参加者を呼びません。計数役3名だけで、
1日に昼と夕方の4便を回します。足し引き2名＋数え直し1名で、乗降計測アプリを使います。</p>

<h2>3. 日付ごとの記入表</h2>
<p>使える平日は16日です（10/12 スポーツの日、11/3 文化の日を除く）。
昼枠・夕方枠に、上の「回」の番号を書き込んでください。</p>
<table>
<tr><th>日付</th><th>曜</th><th>昼枠（11:35集合）に入れる回</th><th>夕方枠（15:25集合）に入れる回</th><th>備考</th></tr>
__DATES__
</table>
<p class="note">本実験8日＋計数方法2日＝<b>10日</b>。<b>予備が6日</b>残ります。
雨・欠席・遅延はこれで吸収できます。</p>
<p class="note"><b>土日は使えません。</b>10/24・25、11/14・15 は土日で、
桃山台17時台が土休日ダイヤでは 北10/20/北40/50 となり <b>17:18 が存在しない</b>ためです。
便を選び直すと平日のデータと混ぜられなくなるので、予備日からも外しています。</p>

<h2>4. 当日の流れ</h2>
<table>
<tr><th>時刻</th><th>昼の日</th><th>時刻</th><th>夕方の日</th></tr>
<tr><td class="c">11:35</td><td>ＪＲ吹田駅（南口）集合。説明・同意取得・カード配布・ログイン確認</td>
    <td class="c">15:25</td><td>同左（2日目の方は説明を短縮）</td></tr>
<tr><td class="c">11:50</td><td>のりば2へ移動。のりばで練習2回</td>
    <td class="c">15:42</td><td>のりば2へ移動。のりばで練習2回</td></tr>
<tr><td class="c"><b>12:03</b></td><td><b>1本目 発車</b>（松本は離脱）</td>
    <td class="c"><b>15:55</b></td><td><b>1本目 発車</b>（松本は離脱）</td></tr>
<tr><td class="c">12:48</td><td>桃山台駅着。折り返し待ち60分</td>
    <td class="c">16:40</td><td>桃山台駅着。折り返し待ち38分</td></tr>
<tr><td class="c"><b>13:48</b></td><td><b>2本目 発車</b></td>
    <td class="c"><b>17:18</b></td><td><b>2本目 発車</b></td></tr>
<tr><td class="c">14:32</td><td>ＪＲ吹田駅（南口）着・解散。アンケート</td>
    <td class="c">18:03</td><td>ＪＲ吹田駅（南口）着・解散。アンケート</td></tr>
</table>
<p class="note">アンケートは解散直後に①前半、送信を確認してから②後半を渡します。
2日目の方は②後半の末尾に追加の質問があります。</p>

<h2>5. 当日いちばん事故りやすいところ</h2>
<p class="warn"><b>南口のりば2からは、[2] 桃山台駅ゆき と [3] 桃山台駅ゆき の両方が出ます。</b>
行先表示はどちらも「桃山台駅」で、<b>行先番号の数字でしか見分けられません。</b>
使うのは [2]（12:03／15:55発）。乗ってはいけないのは [3]（毎時32分発。経路が違います）。</p>
<p class="warn"><b>桃山台では「[2] ＪＲ吹田駅（南口）ゆき」に乗り、終点の南口まで乗ります。</b>
「ＪＲ吹田駅（北口）」と表示して発車するのは [5] という別系統です。</p>
<p class="note">参加者には「実験者が指定した1本にだけ乗る。来た順に乗らない」と説明し、
<b>発車時刻を口に出して指定し、松本が見送るまでその場を離れません。</b>
9月29日に、この注意書きを書いた松本自身が桃山台で [3] に乗っています。
北口を経由すると思い込んでいたためで、行先表示の見落としではありませんでした。
「行先を確認してください」では防げないので、見送りを省略しません。</p>

<h2>6. 持っていくもの</h2>
<table>
<tr><th>配布カード</th><td>参加者番号・組・群・アンケートの版・ID・パスワード。番号順に3枚ずつ渡す</td></tr>
<tr><th>QRポスター</th><td><b>A・B・C・D の4枚。</b>その日に使うのは2枚（§2 の1本目・2本目の欄）</td></tr>
<tr><th>記録用紙</th><td>22停留所版。1日4枚（2便 × 乗車係・降車係）</td></tr>
<tr><th>対応表</th><td>accounts.csv。参加者がパスワードを無くしたとき用。人目に触れない場所に</td></tr>
<tr><th>この用紙</th><td>その日の回の行に印を付けておく</td></tr>
</table>
"""
HTML = HTML.replace("__ASSIGN__", "\n".join(assign)).replace("__DATES__", "\n".join(dates))

# --- 使う文字を集めてサブセット化 -------------------------------------------
text = re.sub(r'<[^>]+>', '', HTML)
chars = set(text) | set(CSS) | set(chr(c) for c in range(0x20, 0x7f))
chars |= set('０１２３４５６７８９①②③④⑤⑥⑦⑧⑨')
unicodes = {ord(c) for c in chars if ord(c) > 31}


def build(ttc_path, out_name):
    out = FONTDIR / out_name
    tmp = FONTDIR / ('_full_' + out_name)
    if not tmp.exists():
        TTCollection(ttc_path).fonts[0].save(str(tmp))
    opts = subset.Options()
    opts.drop_tables += ['DSIG']
    opts.notdef_outline = True
    opts.recalc_bounds = True
    font = subset.load_font(str(tmp), opts)
    ss = subset.Subsetter(options=opts)
    ss.populate(unicodes=unicodes)
    ss.subset(font)
    subset.save_font(font, str(out), opts)
    font.close()
    return out


r = build(r'C:\Windows\Fonts\YuGothM.ttc', 'jp-regular.ttf')
b = build(r'C:\Windows\Fonts\YuGothB.ttc', 'jp-bold.ttf')

arch = pymupdf.Archive()
arch.add(FONTDIR)
story = pymupdf.Story(html=HTML, user_css=CSS, archive=arch)
writer = pymupdf.DocumentWriter(OUT)
MEDIA = pymupdf.paper_rect("A4")
WHERE = MEDIA + (50, 44, -50, -44)

n, more = 0, 1
while more:
    dev = writer.begin_page(MEDIA)
    more, _ = story.place(WHERE)
    story.draw(dev)
    writer.end_page()
    n += 1
    if n > 20:
        raise RuntimeError("pagination runaway")
writer.close()

doc = pymupdf.open(OUT)
print("pages:", doc.page_count, "size:", os.path.getsize(OUT), "bytes")
print(doc[0].get_text()[:200])
