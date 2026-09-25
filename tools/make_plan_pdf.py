# -*- coding: utf-8 -*-
"""実施計画をA4のPDFに出力する。游ゴシックをTTCから取り出し、使う文字だけにサブセット化して埋め込む。"""
import os, re, pathlib
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

OUT = r'C:/Users/taichi/Desktop/研究/bus-ui-study/docs/実施計画_2026-09-25.pdf'
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_plan_v2'
FONTDIR.mkdir(parents=True, exist_ok=True)

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
th { border: 0.6pt solid #888; border-bottom: 1.2pt solid #555; padding: 3pt 5pt; text-align: left; font-weight: bold; }
td { border: 0.6pt solid #888; padding: 3pt 5pt; vertical-align: top; }
td.c { text-align: center; }
ul { margin: 0 0 5pt 0; }
li { margin-bottom: 3pt; }
.note { font-size: 8.3pt; color: #444; margin: 0 0 5pt 0; }
.stops { font-size: 8.3pt; line-height: 1.55; margin: 0 0 5pt 0; }
.kv { margin: 0 0 3pt 0; }
"""

HTML = """
<h1>バス車内人数の回答UI実験　実施計画</h1>
<p class="sub">2026年9月25日　／　松本</p>

<h2>1. 路線</h2>
<p class="kv"><b>事業者・路線</b>　阪急バス　吹田市内線 2系統</p>
<p class="kv"><b>区間</b>　ＪＲ吹田駅（北口） ⇔ 桃山台駅　／　片道18停留所・34〜35分</p>
<p class="kv"><b>集合・解散</b>　どちらの日も ＪＲ吹田駅（北口）</p>
<p class="stops"><b>停留所（往復とも対称）</b><br/>
ＪＲ吹田駅（北口）／片山小学校前／朝日が丘町名神下／上山手町／佐井寺南が丘／総合運動場前／竹谷／佐井寺／
佐井寺北／五月が丘／亥子谷／佐竹台六丁目／佐竹台五丁目／高野台中学校前／佐竹台診療所前／阪急南千里駅／
桃山台二丁目／桃山台駅</p>

<h2>2. 使う便</h2>
<p>路線・区間・停留所・所要時間はすべて同じで、日と便で混雑水準だけを変えます。
4便とも現行ダイヤ（2026年4月1日改正）に存在することを、停留所の時刻表（北口①のりば／桃山台2番のりば）で確認済みです。</p>
<table>
<tr><th>日</th><th>便</th><th>運行</th><th>起点</th><th>混雑</th></tr>
<tr><td rowspan="2"><b>昼の日</b><br/>空いている</td><td class="c">1本目</td>
    <td><b>12:49</b> ＪＲ吹田駅（北口）発 → 桃山台駅 13:24着</td><td>北口始発</td><td>空いている</td></tr>
<tr><td class="c">2本目</td>
    <td><b>13:48</b> 桃山台駅発 → ＪＲ吹田駅（北口）14:22着（折り返し24分）</td>
    <td>桃山台始発</td><td>空いている</td></tr>
<tr><td rowspan="2"><b>夕方の日</b><br/>混んでいる</td><td class="c">1本目</td>
    <td><b>16:20</b> ＪＲ吹田駅（北口）発 → 桃山台駅 16:55着</td><td>北口始発</td><td>中間<br/>10〜22人</td></tr>
<tr><td class="c">2本目</td>
    <td><b>17:18</b> 桃山台駅発 → ＪＲ吹田駅（北口）17:52着（折り返し23分）</td>
    <td>桃山台始発</td><td><b>実測<br/>8→39人</b></td></tr>
</table>
<p>4本すべてが始発便なので、計数役は車内0人から積み上げられます。北口を始発とする便は昼間の毎時49分発と、
16:20／16:48／17:02／17:26／18:14／18:38／18:50 に限られるため、この条件を満たす組み合わせは多くありません。</p>
<p>混雑の水準は3段階になります。空いている（昼）、中間（16:20発）、数えきれない（17:18発）。
「どこで数えられなくなるか」の境目を見るのに、中間の1便が効きます。</p>

<h2>3. 1日の流れ</h2>
<table>
<tr><th>時刻</th><th>昼の日</th><th>時刻</th><th>夕方の日</th></tr>
<tr><td class="c">12:00</td><td>集合・説明・同意取得・アカウント作成・URL配布</td>
    <td class="c">15:50</td><td>集合・説明・URL配布</td></tr>
<tr><td class="c">12:40</td><td>①のりばへ移動。停車中の車内で練習2回</td>
    <td class="c">16:10</td><td>①のりばへ移動。停車中の車内で練習2回</td></tr>
<tr><td class="c"><b>12:49</b></td><td><b>1本目 発車</b></td>
    <td class="c"><b>16:20</b></td><td><b>1本目 発車</b></td></tr>
<tr><td class="c">13:24</td><td>桃山台駅着（折り返し24分）</td>
    <td class="c">16:55</td><td>桃山台駅着（折り返し23分）</td></tr>
<tr><td class="c"><b>13:48</b></td><td><b>2本目 発車</b></td>
    <td class="c"><b>17:18</b></td><td><b>2本目 発車</b></td></tr>
<tr><td class="c">14:22</td><td>ＪＲ吹田駅（北口）着・解散</td>
    <td class="c">17:52</td><td>ＪＲ吹田駅（北口）着・解散</td></tr>
<tr><th>拘束</th><td><b>約2時間20分</b></td><th>拘束</th><td><b>約2時間</b></td></tr>
</table>
<p>事後アンケートは全乗車を終えたあと各自スマホで回答します。
私はどちらの日も送り出し（12:49／16:20）で離脱し、同乗しません。</p>

<h2>4. 日程</h2>
<p>使える平日は <b>16日</b> です（10/12 スポーツの日、11/3 文化の日を除く）。</p>
<table>
<tr><th>10月</th><td>13（火）・14（水）・15（木）・16（金）・19（月）・20（火）・23（金）</td></tr>
<tr><th>11月</th><td>2（月）・4（水）・5（木）・6（金）・9（月）・10（火）・11（水）・12（木）・13（金）</td></tr>
</table>
<p>実施は4日なので、<b>予備が12日</b>残ります。雨・欠席・遅延は十分に吸収できます。
どの日を昼にしてどの日を夕方にするかは、参加者の都合に合わせて決めます。
同じ方に昼と夕方の両方へ来ていただければ、混雑水準が同じ人の中で比較できます。
片方だけの方がいても成立するようにしてあります。</p>
<p class="note"><b>土日は使いません。</b>この4便が土休日ダイヤに存在しないためです
（土休日は北口16時台が 17/33/48 で 16:20 が無く、桃山台17時台が 北10/20/北40/50 で 17:18 が無い）。
便を選び直すと曜日差と条件差が分離できなくなるので、予備日からも外しています。</p>

<h2>5. 体制</h2>
<table>
<tr><th>役割</th><th>人数</th><th>やること</th></tr>
<tr><td>参加者</td><td class="c"><b>12名</b></td><td>アプリで混雑度と人数を回答。1便あたり3名（団体に見えないため）＝1日3名</td></tr>
<tr><td>計数役</td><td class="c"><b>2名</b></td><td>前扉・後扉に1名ずつ。両方の便に乗り、参加者とは別々に乗車して関わらない</td></tr>
<tr><td>実験者<br/>（松本）</td><td class="c">1名</td><td>集合場所での説明と送り出しのみ。同乗しない</td></tr>
</table>

<h2>6. 条件の割り付け</h2>
<p>入力方式（ステッパー／テンキー）はアプリが回答ごとに自動で切り替えます。
「わからない」ボタンの有無は配布URLで便ごとに固定します。
ボタンの有無と混雑水準が交絡しないよう、参加者を2群に分けて入れ替えます。</p>
<table>
<tr><th>群</th><th>昼 1本目</th><th>昼 2本目</th><th>夕 1本目</th><th>夕 2本目</th></tr>
<tr><td class="c"><b>群1</b></td><td class="c">あり</td><td class="c">なし</td><td class="c">なし</td><td class="c">あり</td></tr>
<tr><td class="c"><b>群2</b></td><td class="c">なし</td><td class="c">あり</td><td class="c">あり</td><td class="c">なし</td></tr>
</table>
<p class="note">ボタンあり＝パラメータなし、ボタンなし＝ ?unknown=off 。12名を各6名ずつに割り当てます。</p>

<h2>7. 基準値の取り方</h2>
<p>計数役2名が停留所ごとの乗車人数・降車人数を数え、始発の0人から積み上げます。
終点で「乗車合計 − 降車合計 ＝ 0」を2人で突き合わせ、0にならない便は基準値なしとして扱います
（無理に数を合わせません）。</p>
<p>夕方2本目は混雑で目視が崩れる前提です。先行実験では最大10人ずれました。
ここで得られるのは正解ではなく参照値であり、その前提で分析を組んでいます。</p>

<h2>8. 当日の運用</h2>
<ul>
<li><b>北口発（両日とも1本目）は①のりばに並びます。</b>3系統は北口を通らないので、取り違えが構造的に起きません。</li>
<li><b>桃山台発（両日とも2本目）が要注意です。</b>桃山台駅2番のりばには
[2][3][5][8][9][10][11][61][62][65][69] が集中しています。使う[2]の行先表示は
<b>「ＪＲ吹田駅（南口）」</b>で、<b>「ＪＲ吹田駅（北口）」と表示して発車するのは[5]という別系統</b>です。
「<b>[2] ＪＲ吹田駅（南口）ゆきに乗り、北口で降りる</b>」と手順に明記します。</li>
<li>昼の2本目は毎時24分発が3系統で、北口を通りません。乗せないよう注意します。</li>
<li>参加者には「実験者が指定した1本にだけ乗る。来た順に乗らない」と説明します。</li>
</ul>

<h2>9. ご判断いただきたい点</h2>
<p class="kv"><b>① 使う便の選定でよいか</b></p>
<p>4便とも現行ダイヤに存在することは確認しました。ただし車内がどれくらい混むかは、
実測があるのは桃山台17:18発の1便だけです。系統別・便別の乗車人員はどの機関も公表していないため、
とくに昼の便が本当に空いているかは、乗ってみないと分かりません。</p>
<p class="kv"><b>② 参加者の交通費の手続き</b></p>
<p>総額は1万3千〜1万5千円程度の見込みです。大学から出るものと思っていますが、
どの費目で、謝金扱いか実費精算か、証憑は何が要るか、事前申請が要るかが分かっていません。
事務へ確認する際の窓口と進め方をご教示いただけると助かります。</p>
"""

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
print('font sizes:', os.path.getsize(r), os.path.getsize(b))

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
print("--- page1 head ---")
print(doc[0].get_text()[:300])
