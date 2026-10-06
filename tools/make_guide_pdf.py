# -*- coding: utf-8 -*-
"""参加者にわたす案内をA4のPDFにする

    python tools/make_guide_pdf.py

docs/参加者向け案内.pdf が出る。
"""
import os, re, pathlib
import pymupdf
from fontTools.ttLib import TTCollection
from fontTools import subset

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "参加者向け案内.pdf")
FONTDIR = pathlib.Path(os.environ['TEMP']) / 'jpfonts_guide'
FONTDIR.mkdir(parents=True, exist_ok=True)

CSS = """
@font-face { font-family: jp; src: url(jp-regular.ttf); }
@font-face { font-family: jp; font-weight: bold; src: url(jp-bold.ttf); }
* { font-family: jp; }
body { font-size: 11pt; line-height: 1.7; color: #111; }
h1 { font-size: 20pt; margin: 0 0 2pt 0; }
.lead { font-size: 12pt; margin: 0 0 16pt 0; }
h2 { font-size: 13.5pt; margin: 16pt 0 6pt 0; color: #1a4b8c; }
p { margin: 0 0 7pt 0; }
table { width: 100%; border-collapse: collapse; margin: 4pt 0 10pt 0; font-size: 11pt; }
td { border: 0.8pt solid #999; padding: 6pt 8pt; vertical-align: top; }
td.k { width: 110pt; font-weight: bold; background-color: #eef3fa; }
.step { font-size: 11.5pt; margin: 0 0 7pt 0; }
.num { font-weight: bold; color: #1a4b8c; }
.big { font-size: 13pt; font-weight: bold; margin: 10pt 0 6pt 0; }
.box { border: 1.2pt solid #1a4b8c; padding: 9pt 11pt; margin: 10pt 0; }
"""

HTML = """
<h1>バスの人数アプリ　実験のご案内</h1>
<p class="lead">大和大学　情報学部　松本泰知　卒業研究</p>

<h2>やること</h2>
<p>バスに乗って、<b>停まるたびに車内の人数をスマホで答える。</b>それだけです。</p>

<h2>1回の実験は、往復2本です</h2>
<table>
<tr><td class="k">集合</td><td>ＪＲ吹田駅　<b>南口</b>　（北口ではありません）</td></tr>
<tr><td class="k">ゆき</td><td>ＪＲ吹田駅（南口） → 桃山台駅　　約45分</td></tr>
<tr><td class="k">待ち時間</td><td>桃山台駅で休憩</td></tr>
<tr><td class="k">かえり</td><td>桃山台駅 → ＪＲ吹田駅（南口）　　約45分</td></tr>
<tr><td class="k">さいごに</td><td>アンケート　5分ほど</td></tr>
</table>

<h2>当日の流れ</h2>
<p class="step"><span class="num">1.</span>　南口に集合。カードを受け取る（IDとパスワードが書いてあります）</p>
<p class="step"><span class="num">2.</span>　実験者が出すポスターのQRコードを、スマホのカメラで読む</p>
<p class="step"><span class="num">3.</span>　カードのIDとパスワードでログインする</p>
<p class="step"><span class="num">4.</span>　のりばで2回だけ練習する</p>
<p class="step"><span class="num">5.</span>　バスに乗る。<b>停まるたびに人数を答える</b></p>
<p class="step"><span class="num">6.</span>　桃山台駅で降りる。休憩</p>
<p class="step"><span class="num">7.</span>　かえりのバスでも同じことをする</p>
<p class="step"><span class="num">8.</span>　南口で解散。アンケートに答える</p>

<h2>人数の答え方</h2>
<p class="step"><span class="num">1.</span>　次に停まるバス停を選ぶ</p>
<p class="step"><span class="num">2.</span>　いま車内にいる人数を入れる</p>
<p class="step"><span class="num">3.</span>　送る</p>
<div class="box">
<p class="big" style="margin-top:0">数えきれないときは、だいたいの数でかまいません。</p>
<p style="margin-bottom:0">正解を当てる実験ではありません。<b>答えにくさそのものを調べています。</b>
無理に数えようとしなくて大丈夫です。</p>
</div>

<h2>お願い</h2>
<table>
<tr><td class="k">乗るバス</td><td><b>実験者が指定した1本にだけ乗ってください。</b>
同じ行先のバスが続けて来ます。実験者が見送ります</td></tr>
<tr><td class="k">持ちもの</td><td>充電したスマホ　／　ICカードか小銭</td></tr>
<tr><td class="k">運賃</td><td><b>実費をお支払いします。</b>謝礼はありません</td></tr>
<tr><td class="k">やめたいとき</td><td>いつでもやめられます。理由を言う必要はありません</td></tr>
</table>

<h2>記録されること</h2>
<p>答えた人数、答えるのにかかった時間、押したボタン、送った時刻。
<b>名前・電話番号・位置情報は記録しません。</b>データはIDだけで扱い、研究以外には使いません。</p>

<h2>連絡先</h2>
<table>
<tr><td class="k">松本　泰知</td><td>大和大学 情報学部　23610225tm@stu.yamato-u.ac.jp</td></tr>
</table>
"""

text = re.sub(r'<[^>]+>', '', HTML)
chars = set(text) | set(CSS) | set(chr(c) for c in range(0x20, 0x7f))
unicodes = {ord(c) for c in chars if ord(c) > 31}


def build(ttc_path, out_name):
    out = FONTDIR / out_name
    tmp = FONTDIR / ('_full_' + out_name)
    if not tmp.exists():
        TTCollection(ttc_path).fonts[0].save(str(tmp))
    opts = subset.Options()
    opts.drop_tables += ['DSIG']
    opts.notdef_outline = True
    font = subset.load_font(str(tmp), opts)
    ss = subset.Subsetter(options=opts)
    ss.populate(unicodes=unicodes)
    ss.subset(font)
    subset.save_font(font, str(out), opts)
    font.close()
    return out


build(r'C:\Windows\Fonts\YuGothM.ttc', 'jp-regular.ttf')
build(r'C:\Windows\Fonts\YuGothB.ttc', 'jp-bold.ttf')

arch = pymupdf.Archive()
arch.add(FONTDIR)
story = pymupdf.Story(html=HTML, user_css=CSS, archive=arch)
writer = pymupdf.DocumentWriter(OUT)
MEDIA = pymupdf.paper_rect("A4")
WHERE = MEDIA + (52, 46, -52, -46)

n, more = 0, 1
while more:
    dev = writer.begin_page(MEDIA)
    more, _ = story.place(WHERE)
    story.draw(dev)
    writer.end_page()
    n += 1
    if n > 10:
        raise RuntimeError("pagination runaway")
writer.close()

doc = pymupdf.open(OUT)
print("書き出しました: %s（%dページ・%d バイト）" % (OUT, doc.page_count, os.path.getsize(OUT)))
