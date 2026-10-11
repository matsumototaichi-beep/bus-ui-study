# -*- coding: utf-8 -*-
"""UI研究のアンケート（Google フォーム4つ）の今の中身を読み、控えと比べる

    python tools/read_forms.py              公開の回答ページを読み、いちばん新しい控えと比べる
    python tools/read_forms.py --save       読んだ中身を docs/questionnaire/フォームの中身_<今日>.json と .md に書き出す
    python tools/read_forms.py --html DIR   ページを読む代わりに、保存した <フォーム名>.html を使う
    python tools/read_forms.py --md         いちばん新しい .json から .md だけを作り直す

フォームは変えない（読むだけ）。正本は Google フォームで、.json と .md はその控え。
tools/join_survey_app.py は .json を使って、スプレッドシートの列と表の列名を対応させる。
URL は ../★UI研究/アンケートURL.xlsx から読み、ファイルには残さない（このリポジトリは public）。
"""
import argparse
import datetime
import json
import os
import re
import sys
import urllib.request

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import join_survey_app as J

UI_XLSX = os.path.join(os.path.dirname(J.ROOT), "★UI研究", "アンケートURL.xlsx")
KIND = {0: "短文", 1: "段落", 2: "単一選択", 3: "プルダウン", 4: "チェックボックス", 5: "5段階", 6: "説明",
        7: "表", 8: "ページ", 9: "日付", 10: "時刻"}
PLACE = {"桃山台": "桃山台駅", "南口": "南口"}


def parse(html):
    """回答ページの FB_PUBLIC_LOAD_DATA_ から、題名・説明・問の list を取り出す"""
    m = re.search(r"FB_PUBLIC_LOAD_DATA_ = (\[.*?\]);\s*</script>", html, re.S)
    if not m:
        sys.exit("フォームの中身が読めません（ログインが必要なページか、形が変わった）")
    d = json.loads(m.group(1))[1]
    items = []
    for it in d[1] or []:
        kind = KIND.get(it[3], "その他(%s)" % it[3])
        x = {"type": kind, "title": it[1], "help": it[2]}
        if kind not in ("ページ", "説明"):
            q = (it[4] or [[None, None, 0]])[0]
            x["required"] = bool(q[2])
            if kind == "5段階":
                opts = [o[0] for o in q[1]]
                x["range"] = [int(opts[0]), int(opts[-1])]
                x["labels"] = q[3]
            elif kind in ("単一選択", "プルダウン", "チェックボックス"):
                x["choices"] = [o[0] for o in q[1]]
        items.append(x)
    return {"title": d[8], "description": d[0], "items": items}


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept-Language": "ja"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8")


def read_all(html_dir):
    out = {}
    for name in J.FORMS:
        if html_dir:
            with open(os.path.join(html_dir, name + ".html"), encoding="utf-8") as f:
                out[name] = parse(f.read())
            continue
        import make_handouts                                   # アンケートURL.xlsx の読み方を同じにする
        url, _ = make_handouts.survey(UI_XLSX, name)
        if not url:
            sys.exit("アンケートURL.xlsx に「%s」の URL がありません" % name)
        out[name] = parse(fetch(url))
    return out


def diff(old, new):
    lines = []
    for name in J.FORMS:
        a, b = old.get(name), new.get(name)
        if a == b:
            continue
        for k in ("title", "description"):
            if (a or {}).get(k) != (b or {}).get(k):
                lines.append("%s の%s：「%s」→「%s」" % (name, "題名" if k == "title" else "説明", (a or {}).get(k), b.get(k)))
        ai, bi = (a or {}).get("items", []), b.get("items", [])
        for i in range(max(len(ai), len(bi))):
            x = ai[i] if i < len(ai) else None
            y = bi[i] if i < len(bi) else None
            if x != y:
                lines.append("%s の %d 番目：%s\n    → %s" % (name, i + 1, json.dumps(x, ensure_ascii=False),
                                                         json.dumps(y, ensure_ascii=False)))
    return lines


def answer_form(it):
    if it["type"] == "5段階":
        return "%d〜%d（%d＝%s、%d＝%s）" % (it["range"][0], it["range"][1], it["range"][0], it["labels"][0],
                                        it["range"][1], it["labels"][1])
    if it.get("choices"):
        return "%s：%s" % (it["type"], "／".join(it["choices"]))
    return it["type"]


def write_md(json_path):
    with open(json_path, encoding="utf-8") as f:
        forms = json.load(f)
    qmap = J.question_map(forms)
    day = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(json_path)).group(1)
    L = ["# UI研究のアンケートの中身（%s）" % day, "",
         "Google フォーム4つの、%s 時点の中身の控えです。" % day, "",
         "- 2026-10-10 に松本が Google フォームを手で直しました。**正本は Google フォーム**で、これはその控えです。"
         "公開の回答ページから読みました（`tools/read_forms.py`）。問の文は、フォームの文をそのまま写しています",
         "- 分析の道具 `tools/join_survey_app.py` は、同じ名前の .json を使って、スプレッドシートの列（見出しの「問N」）を"
         "下の表の「表の列名」に対応させます",
         "- フォームを直したら、`python tools/read_forms.py` で違いを確かめ、`--save` で控えを作り直します。"
         "紙の予備（`アンケート_*.docx`）は `python tools/make_questionnaire.py` で作り直します",
         "- フォームの URL はここに書きません（`../★UI研究/アンケートURL.xlsx`）", "",
         "| フォーム | 答える所 | 回 | 区間 | 問数 |", "|---|---|---|---|---|"]
    for name, (rnd, leg) in J.FORMS.items():
        L.append("| %s | %s | %d回目 | %s | %d |" % (name, PLACE[name.split("_")[1]], rnd, J.LEG_JA[leg], len(qmap[name])))
    for name in J.FORMS:
        fm = forms[name]
        key = {q["title"]: q["key"] + ("（SUS の点：答え−1）" if q["sus"] and q["sus"] % 2 else
                                       "（SUS の点：5−答え）" if q["sus"] else "") if q["key"] else None
               for q in qmap[name]}
        L += ["", "## %s" % name, "", "- 題名：%s" % fm["title"], "- 説明：%s" % (fm["description"] or "（なし）"), "",
              "| 問（フォームの文のまま） | 答え方 | 必須 | 表の列名 |", "|---|---|---|---|"]
        for it in fm["items"]:
            if it["type"] in ("ページ", "説明"):
                L.append("| 〈%s〉%s%s | | | |" % (it["type"], it["title"], "（説明：%s）" % it["help"] if it.get("help") else ""))
                continue
            L.append("| %s%s | %s | %s | %s |" % (it["title"], "（説明：%s）" % it["help"] if it.get("help") else "",
                                               answer_form(it), "必須" if it.get("required") else "任意",
                                               key.get(it["title"]) or "（その他の答え）"))
    md = os.path.splitext(json_path)[0] + ".md"
    with open(md, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(L) + "\n")
    return md


def main():
    ap = argparse.ArgumentParser(description="アンケートの今の中身を読み、控えと比べる")
    ap.add_argument("--save", action="store_true", help="読んだ中身を控えとして書き出す")
    ap.add_argument("--html", help="保存した回答ページ（<フォーム名>.html）のフォルダ")
    ap.add_argument("--md", action="store_true", help="いちばん新しい .json から .md だけを作り直す")
    a = ap.parse_args()
    if a.md:
        print("書き出しました: %s" % write_md(J.latest_form_json()))
        return
    now = read_all(a.html)
    J.question_map(now)                       # 分析の道具で対応が決まるか（決まらなければここで止まる）
    latest = J.latest_form_json()
    with open(latest, encoding="utf-8") as f:
        old = json.load(f)
    d = diff(old, now)
    print("控え: %s" % os.path.relpath(latest, J.ROOT))
    print("同じです" if not d else "違うところが %d か所あります：\n  %s" % (len(d), "\n  ".join(d)))
    print("問数: %s" % "・".join("%s %d" % (n, sum(1 for it in now[n]["items"] if it["type"] not in ("ページ", "説明")))
                               for n in J.FORMS))
    if a.save:
        path = os.path.join(J.QDIR, "フォームの中身_%s.json" % datetime.date.today().isoformat())
        with open(path, "w", encoding="utf-8") as f:
            json.dump(now, f, ensure_ascii=False, indent=1)
        print("書き出しました: %s\n書き出しました: %s" % (path, write_md(path)))
    elif d:
        sys.exit(1)


if __name__ == "__main__":
    main()
