# -*- coding: utf-8 -*-
"""実施計画をA4のPDFに出力する。游ゴシックをTTCから取り出し、使う文字だけにサブセット化して埋め込む。

★2026-10-03 改訂。集合・解散を南口に変えたので全面的に書き直した。
  9/25版（docs/実施計画_2026-09-25.pdf）は提出済みなので消さずに残し、別ファイルで出す。

注意（作ったときにハマった点）：
  - サブセット化しないとPDFが28MBになる
  - th に background-color を指定すると Story が無関係な位置に灰色の矩形を描き残す。罫線だけにする
"""
import os, re, pathlib
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

OUT = r'C:/Users/taichi/Desktop/研究/bus-ui-study/docs/実施計画_2026-10-03.pdf'
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
.warn { font-size: 8.5pt; color: #a00; margin: 0 0 5pt 0; }
.stops { font-size: 8.3pt; line-height: 1.55; margin: 0 0 5pt 0; }
.kv { margin: 0 0 3pt 0; }
"""

HTML = """
<h1>バス車内人数の回答UI実験　実施計画</h1>
<p class="sub">2026年10月3日　／　松本　　＊9月25日版の差し替えです。集合・解散を南口に変更し、往路の2便を入れ替えました。</p>

<h2>1. 路線</h2>
<p class="kv"><b>事業者・路線</b>　阪急バス　吹田市内線 2系統</p>
<p class="kv"><b>区間</b>　ＪＲ吹田駅（南口） ⇔ 桃山台駅　／　片道22停留所・44〜45分</p>
<p class="kv"><b>集合・解散</b>　どちらの日も ＪＲ吹田駅（南口）</p>
<p class="stops"><b>停留所（往復とも対称）</b><br/>
ＪＲ吹田駅（南口）／田中町／吹田市役所前（阪急吹田駅）／アサヒビール会社前／ＪＲ吹田駅（北口）／
片山小学校前／朝日が丘町名神下／上山手町／佐井寺南が丘／総合運動場前／竹谷／佐井寺／佐井寺北／
五月が丘／亥子谷／佐竹台六丁目／佐竹台五丁目／高野台中学校前／佐竹台診療所前／阪急南千里駅／
桃山台二丁目／桃山台駅</p>
<p class="note">9月25日版は北口を起点にしていましたが、南口に変えたことで記録できる区間が18停留所から22停留所に伸びました。</p>

<h2>2. 使う便</h2>
<p>路線・区間・停留所・所要時間はすべて同じで、日と便で混雑水準だけを変えます。
4便とも2系統であることと、南口発の2便が南口始発であることを、公式の路線別時刻一覧表の行先番号欄で確認しました。</p>
<table>
<tr><th>日</th><th>便</th><th>運行</th><th>起点</th><th>混雑</th></tr>
<tr><td rowspan="2"><b>昼の日</b><br/>空いている</td><td class="c">1本目</td>
    <td><b>12:03</b> ＪＲ吹田駅（南口）発 → 桃山台駅 12:48着</td><td>南口始発</td><td>空いている</td></tr>
<tr><td class="c">2本目</td>
    <td><b>13:48</b> 桃山台駅発 → ＪＲ吹田駅（南口）14:32着（折り返し60分）</td>
    <td>桃山台始発</td><td>空いている</td></tr>
<tr><td rowspan="2"><b>夕方の日</b><br/>混んでいる</td><td class="c">1本目</td>
    <td><b>15:55</b> ＪＲ吹田駅（南口）発 → 桃山台駅 16:40着</td><td>南口始発</td><td>中間</td></tr>
<tr><td class="c">2本目</td>
    <td><b>17:18</b> 桃山台駅発 → ＪＲ吹田駅（南口）18:03着（折り返し38分）</td>
    <td>桃山台始発</td><td><b>北口まで<br/>実測8→39人</b></td></tr>
</table>
<p><b>4本すべてが始発便です。</b>計数役は車内0人から積み上げられます。
2本目は、以前は北口で降りる運用でしたが、終点の南口まで乗ります。</p>
<p class="note">昼の折り返し60分は動かせません。南口まで行く2系統は桃山台発が毎時48分、南口発は毎時3分で桃山台着が48分です。
到着と同時刻発になるため、どの時間帯を選んでも60分になります。</p>
<p>混雑の水準は3段階になります。空いている（昼）、中間（15:55発）、数えきれない（17:18発）。
「どこで数えられなくなるか」の境目を見るのに、中間の1便が効きます。</p>

<h2>3. 1日の流れ</h2>
<table>
<tr><th>時刻</th><th>昼の日</th><th>時刻</th><th>夕方の日</th></tr>
<tr><td class="c">11:35</td><td>集合・説明・同意取得・カード配布・ログイン確認</td>
    <td class="c">15:25</td><td>集合・説明・カード配布</td></tr>
<tr><td class="c">11:50</td><td>のりば2へ移動。停車中の車内で練習2回</td>
    <td class="c">15:42</td><td>のりば2へ移動。停車中の車内で練習2回</td></tr>
<tr><td class="c"><b>12:03</b></td><td><b>1本目 発車</b></td>
    <td class="c"><b>15:55</b></td><td><b>1本目 発車</b></td></tr>
<tr><td class="c">12:48</td><td>桃山台駅着（折り返し60分）</td>
    <td class="c">16:40</td><td>桃山台駅着（折り返し38分）</td></tr>
<tr><td class="c"><b>13:48</b></td><td><b>2本目 発車</b></td>
    <td class="c"><b>17:18</b></td><td><b>2本目 発車</b></td></tr>
<tr><td class="c">14:32</td><td>ＪＲ吹田駅（南口）着・解散</td>
    <td class="c">18:03</td><td>ＪＲ吹田駅（南口）着・解散</td></tr>
<tr><th>拘束</th><td><b>約3時間</b></td><th>拘束</th><td><b>約2時間40分</b></td></tr>
</table>
<p>事後アンケートは全乗車を終えたあと各自スマホで回答します。
私はどちらの日も送り出し（12:03／15:55）で離脱し、同乗しません。</p>
<p class="note">集合から発車までを9月25日版の49分から28分に詰めました。
参加者のアカウントを事前に作ってカードで渡す方式にしたので、その場での登録作業が無くなったためです。</p>

<h2>4. 日程</h2>
<p>使える平日は <b>16日</b> です（10/12 スポーツの日、11/3 文化の日を除く）。</p>
<table>
<tr><th>10月</th><td>13（火）・14（水）・15（木）・16（金）・19（月）・20（火）・23（金）</td></tr>
<tr><th>11月</th><td>2（月）・4（水）・5（木）・6（金）・9（月）・10（火）・11（水）・12（木）・13（金）</td></tr>
</table>
<p>実施は8日なので、<b>予備が8日</b>残ります。雨・欠席・遅延は吸収できます。
どの日を昼にしてどの日を夕方にするかは、参加者の都合に合わせて決めます。
同じ方に昼と夕方の両方へ来ていただくことで、混雑水準を同じ人の中で比較します。</p>
<p class="note"><b>土日は使いません。</b>桃山台17時台が土休日は 北10/20/北40/50 で <b>17:18 が無く</b>、
夕方の2本目が組めないためです。便を選び直すと曜日差と条件差が分離できなくなるので、予備日からも外しています。</p>

<h2>5. 体制と乗車回数</h2>
<table>
<tr><th>役割</th><th>人数</th><th>やること</th></tr>
<tr><td>参加者</td><td class="c"><b>12名</b></td><td>アプリで混雑度と人数を回答。1便あたり3名（団体に見えないため）＝1日3名</td></tr>
<tr><td>計数役</td><td class="c"><b>2名</b></td><td>前扉・後扉に1名ずつ。両方の便に乗り、参加者とは別々に乗車して関わらない</td></tr>
<tr><td>実験者<br/>（松本）</td><td class="c">1名</td><td>集合場所での説明と送り出しのみ。同乗しない</td></tr>
</table>
<table>
<tr><th>1人あたりの乗車</th><td class="c"><b>4回</b></td><th>実施日</th><td class="c"><b>8日</b></td><th>走らせる便</th><td class="c"><b>16便</b></td></tr>
<tr><th>延べ乗車</th><td class="c"><b>80回</b><br/>参加者48＋計数役32</td><th>回答数</th><td class="c"><b>384</b><br/>48乗車×8か所</td><th>条件ごと</th><td class="c">8セル<br/>各48回答</td></tr>
</table>
<p class="note">1便あたりの指定停留所は8か所です。昼と夕方で別の方になる場合は、1人2回・4日・8便・延べ40乗車まで下がります。</p>

<h2>6. 条件の割り付け</h2>
<p><b>入力方式（ステッパー／テンキー）は実施日ごとに固定します。</b>
1回答ごとに切り替える方式から変更しました。日ごとに固定すれば参加者が「その日ずっと使っていた方式」として思い出せるので、
事後アンケートで画面の写真を見せずに主観評価が聞けます。
「わからない」ボタンの有無は便ごとに配布URLで固定します。</p>
<table>
<tr><th>群</th><th>昼 1本目（12:03）</th><th>昼 2本目（13:48）</th><th>夕 1本目（15:55）</th><th>夕 2本目（17:18）</th></tr>
<tr><td class="c"><b>群1</b><br/>組1・組3</td><td class="c">ステッパー<br/>ボタンあり</td><td class="c">ステッパー<br/>ボタンなし</td>
    <td class="c">テンキー<br/>ボタンなし</td><td class="c">テンキー<br/>ボタンあり</td></tr>
<tr><td class="c"><b>群2</b><br/>組2・組4</td><td class="c">テンキー<br/>ボタンなし</td><td class="c">テンキー<br/>ボタンあり</td>
    <td class="c">ステッパー<br/>ボタンあり</td><td class="c">ステッパー<br/>ボタンなし</td></tr>
</table>
<p>どの参加者も4乗車で「入力方式2通り × ボタン2通り」の4組み合わせを1回ずつ経験し、
空いている便と混んでいる便の両方に乗ります。組によって昼と夕方を入れ替えてあるので、
全体としてはどちらの入力方式も空いている便・混んでいる便の両方に現れます。</p>
<p class="note">同じ便に乗る3名は必ず同じ群にします。群を混ぜると、同じバスの中で画面が違う人が並び、
条件の存在に気づかれるためです。配布URLは4通りになり、QRポスターを4枚用意して便ごとに掲げます。</p>

<h2>7. 基準値の取り方</h2>
<p>計数役2名が22停留所それぞれの乗車人数・降車人数を数え、始発の0人から積み上げます。
終点で「乗車合計 − 降車合計 ＝ 0」を2人で突き合わせ、0にならない便は基準値なしとして扱います
（無理に数を合わせません）。</p>
<p>夕方2本目は混雑で目視が崩れる前提です。先行実験（9月18日）では最大10人ずれ、累積誤差が+7程度ありました。
ここで得られるのは正解ではなく参照値であり、その前提で分析を組んでいます。</p>

<h2>8. 当日の運用</h2>
<ul>
<li><b>南口のりば2が最大の注意点です。</b>こののりばからは
<b>[2] 桃山台駅ゆき と [3] 桃山台駅ゆき の両方</b>が出ます。行先表示はどちらも「桃山台駅」で、
<b>行先番号の数字でしか見分けられません</b>。使うのは [2]（12:03／15:55発）、
乗ってはいけないのは [3]（毎時32分発。吹田高校・吹高口経由で停留所が違う）です。
9月25日版では北口発だったため「3系統は北口を通らない」ので構造的に安全でしたが、南口ではその保護が効きません。</li>
<li><b>桃山台発（両日とも2本目）も要注意です。</b>桃山台駅2番のりばには
[2][3][5][8][9][10][11][61][62][65][69] が集中しています。使う[2]の行先表示は「ＪＲ吹田駅（南口）」で、
<b>「ＪＲ吹田駅（北口）」と表示して発車するのは[5]という別系統</b>です。</li>
<li>参加者には「実験者が指定した1本にだけ乗る。来た順に乗らない」と説明し、
<b>発車時刻を口に出して指定し、私が見送るまでその場を離れません。</b></li>
</ul>
<p class="warn">9月29日に、この注意書きを書いた私自身が、桃山台で[3]に乗っています。
北口を経由すると思い込んでいたためで、行先表示の見落としではありませんでした。
「行先を確認してください」という声かけでは防げないので、見送りを省略しません。</p>

<h2>9. 費用</h2>
<p><b>謝礼金は支払いません。</b>持つのはバス運賃などの実費だけです。</p>
<table>
<tr><th>バス運賃</th><td>1乗車 220〜240円（吹田営業所管内はほぼ均一。2019年10月改定表。現行額は当日確認）</td></tr>
<tr><th>総額</th><td><b>延べ80乗車 × 220〜240円 ＝ 約1万8千〜2万円</b></td></tr>
<tr><th>大学 ⇔ ＪＲ吹田駅</th><td>徒歩・0円</td></tr>
</table>
<p>事務には謝金ではなく実費精算として通します。残るのは費目・証憑・事前申請の確認だけです。</p>

<h2>10. まだ分かっていないこと</h2>
<ul>
<li><b>昼の便が本当に空いているかは、乗ってみないと分かりません。</b>
系統別・便別の乗車人員は阪急バス・吹田市・豊中市のいずれも公表していないため、公開情報ではこれ以上詰められません。</li>
<li><b>計数の方法（乗降の足し引きか、車内の数え直しか）について、答えきれない点が残っています。</b>
空いている便で足し引きが正解として機能したことと、混んでいる便では両方式とも破綻することは、
先行実験の実測で示せます。ただし数え直し方式を系統的に試しておらず、計数したのが私1人だけなので、
「数え直したほうが正確では」という問いには答えきれません。別紙にまとめています。</li>
</ul>
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
