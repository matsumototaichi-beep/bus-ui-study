# -*- coding: utf-8 -*-
"""アンケートの答えとアプリの記録をつなぐ（UI研究）

    python tools/join_survey_app.py --forms フォルダ --sessions ui_sessions.csv --responses ui_responses.csv
                                    --idmap id対応.csv --out 出力フォルダ [--ids A,B,C]

入れるもの（書き出し方は docs/分析の準備.md）
  --forms      スプレッドシート「UI研究 アンケート回答」の4シートを CSV にしたもの。
               ファイル名はシート名.csv（例 1回目_桃山台.csv）。「… - 1回目_桃山台.csv」でもよい
  --sessions   Supabase の ui_sessions（CSV か JSON）
  --responses  Supabase の ui_responses（CSV か JSON）
  --idmap      surveyor_id（ログインした人の番号）とログインID の対応（CSV。列 surveyor_id・login_id）。
               アプリが variant.login_id を記録していれば、そちらを先に使う
  --ids        参加する人の ID（例 A,B,C）。書けば、これ以外の ID に印を付ける
  --since      この日より前の行（実験の前のテストなど）は読まない（例 2026-10-13）。読まなかった数は問題の一覧に書く
  --form-json  フォームの中身の控え。省くと docs/questionnaire/フォームの中身_*.json のいちばん新しいもの

出すもの（--out。個人に関わるので exports/ の下かリポジトリの外に出す。exports/ は git に入らない）
  区間ごと.csv       参加者×回×区間ごとに1行。アンケートの答え・SUS の点・条件・記録から作った数
  問題の一覧.csv     見つけた問題。行は消さず、印と扱いの案を付ける
  アンケートの行.csv  アンケートの1行ごとに、どの区間につないだか
  つないだ表.xlsx    上の3つと「列の確かめ」「読み方」を1つにしたもの

決まり（docs/分析の準備.md・UI_STUDY_SPEC.md §8 と同じ）
  - アンケートの問は、列見出しの「問N」で見つけ、フォームの中身の控えの問の文で確かめる
  - 条件（入れ方・「わからない」ボタン）は、アプリの記録（variant）の値を使う。計画の値は別の列に並べる
  - 同じ人が同じ区間のアンケートに2回以上答えたら、時刻が合う行（区間の最後の回答より後、
    次の区間の最初の回答より前）のうち、いちばん早い行を使う。どれも合わなければ、いちばん早い行
  - 記録は、同じ人・同じ日の操作ログを時刻で往復に分けてから区間にする（60分以上あいたか、かえりのあとに別の回の
    「ゆき」が来たら次の往復）。往復の途中で回が替わった記録（前の回のタブを読み込み直した等）も、その往復の行に入れて印を付ける
  - 行は消さない。問題には印を付け、「問題の一覧」に扱いの案を書く
"""
import argparse
import glob
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from statistics import median

sys.dont_write_bytecode = True          # tools/ に .pyc を増やさない
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
QDIR = os.path.join(ROOT, "docs", "questionnaire")
TZ = "Asia/Tokyo"

# シート名 → (回, 区間)。フォーム・シート・アンケートURL.xlsx の名前と一字一句同じ
FORMS = {"1回目_桃山台": (1, "yuki"), "1回目_南口": (1, "kaeri"),
         "2回目_桃山台": (2, "yuki"), "2回目_南口": (2, "kaeri")}
LEG_JA = {"yuki": "ゆき", "kaeri": "かえり"}
INPUT_JA = {"numpad": "数字キー", "stepper": "＋−ボタン"}
# 組はアカウントごとの表（index.html の GROUP_OF と同じ。test_join_survey_app.py で照らし合わせる）
GROUP_OF = {c: "odd" for c in "ACEGIKMOQSUW"}
GROUP_OF.update({c: "even" for c in "BDFHJLNPRTVX"})
GROUP_JA = {"odd": "奇数組", "even": "偶数組"}

# しきい値
TOL_MIN = 1.0        # アンケートと記録の時刻を比べるときの許し（分）
GAP_MIN = 15.0       # 1つの区間の中でこれ以上あいたら、2本のバスがまざっている疑い（分）
SWITCH_GAP_MIN = 5.0  # 桃山台のアンケートをはさんでこれ以上あいていたら、ゆき・かえりの取り違えの疑い（分）
LATE_MIN = 120.0     # 区間の終わりからこれより後のアンケートは「遅れて答えた」（分）
SKEW_USE = 60.0      # 端末の時計のずれがこれより大きければ時刻を直す（秒）
SKEW_FLAG = 120.0    # これより大きければ印を付ける（秒）
LAT_FLIP = 0.004     # 区間の向きと逆に、緯度がこれ以上動いたら印（約450m。南口 34.763 → 桃山台駅 34.792）
LAT_TURN = 0.008     # 区間の途中で向きが変わった（約900m 進んでから約900m 戻った）
NOON_END = 14        # この時より前に始まった往復は「昼」、あとは「夕方」
TRIP_GAP_MIN = 60.0  # 同じ人の記録がこれ以上あいたら、別の往復とみる（分）
LEG_NEW_MIN = 20.0   # かえりの記録のあとに「ゆき」の記録が来て、これ以上あいていれば別の往復とみる（分）
PRAC_FAR_M = 300.0   # 練習の回答が南口のバス停からこれ以上離れていたら、バスの回答を練習のまま送った疑い（m）
MINAMI = (34.762883, 135.524382)   # stops.js の「JR吹田駅（南口）」（集合して練習する所）

# ---------------------------------------------------------------- 問の対応
# 表の列名と、問の文に含まれる言葉。上から順に当てる。列名の番号は 2026-10-10 のフォームの番号
RULES = [
    ("ID", ("カードに書いてあるID",)),
    ("問2_混み具合", ("混み具合",)),
    ("問4_だいたいで送った", ("だいたいの人数",)),
    ("問3_数えきれなかった", ("数えきれない",)),
    ("問5_1枚目の答えやすさ", ("1枚目",)),
    ("問6_2枚目の答えやすさ", ("2枚目",)),
    ("問17_使いたい方", ("どちらを利用したい",)),
    ("問18_ボタンがあった回", ("ボタンがあったのは",)),
    ("問19_送りやすかった回", ("送りやすかった",)),
    ("問20_意識した回", ("意識",)),
    ("問21_ボタンがあった方が良い", ("あった方が良い",)),
    ("自由記述", ("答えにくかったこと",)),
]
# SUS（使いやすさを測る決まった10問）。両端が下の2つの5段階を、出てくる順に SUS1〜SUS10 とする。
# 奇数番はよい向き（点−1）、偶数番はわるい向き（5−点）。向きを問の文の言葉で確かめる
SUS_LABELS = ["まったくそう思わない", "とてもそう思う"]
SUS_WORDS = ["何度も使いたい", "必要以上に複雑", "簡単に使える", "助けが必要", "うまくまとまって",
             "一貫性のない", "すぐ覚えられる", "使いにくい", "自信をもって", "覚える必要"]
SUS_COLS = ["SUS%d" % i for i in range(1, 11)]


def nk(s):
    """全角を半角にし、空白を消す（見出しや答えを比べるため）"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s or "")))


def qnum(title):
    """見出しの頭の「問N」の N。空白は消さずに読む（「問5 1枚目」を 51 と読まないように）"""
    m = re.match(r"\s*問\s*(\d+)", unicodedata.normalize("NFKC", str(title or "")))
    return int(m.group(1)) if m else None


def latest_form_json():
    paths = sorted(glob.glob(os.path.join(QDIR, "フォームの中身_*.json")))
    if not paths:
        sys.exit("フォームの中身の控え（docs/questionnaire/フォームの中身_*.json）がありません")
    return paths[-1]


def question_map(forms_json):
    """フォームごとに [{n, key, type, title, choices, sus}] を返す。対応が決まらなければ止める"""
    out = {}
    for name in FORMS:
        if name not in forms_json:
            sys.exit("フォームの中身の控えに「%s」がありません" % name)
        qs, sus = [], []
        for it in forms_json[name]["items"]:
            if it["type"] in ("ページ", "説明"):
                continue
            n = qnum(it["title"])
            if n is None:
                sys.exit("%s：問番号のない問があります：%s" % (name, it["title"]))
            q = {"n": n, "type": it["type"], "title": it["title"], "key": None, "sus": None,
                 "choices": it.get("choices") or [], "labels": it.get("labels") or []}
            t = nk(it["title"])
            if it["type"] == "5段階" and q["labels"] == SUS_LABELS:
                sus.append(q)
            else:
                for key, words in RULES:
                    if any(nk(w) in t for w in words):
                        q["key"] = key
                        break
            qs.append(q)
        if sus:
            if len(sus) != 10:
                sys.exit("%s：SUS の問が %d 問あります（10 問のはず）" % (name, len(sus)))
            for i, q in enumerate(sus):
                if nk(SUS_WORDS[i]) not in nk(q["title"]):
                    sys.exit("%s：SUS の %d 番目（問%d）の向きを確かめられません：%s"
                             % (name, i + 1, q["n"], q["title"]))
                q["key"], q["sus"] = SUS_COLS[i], i + 1
        keys = [q["key"] for q in qs if q["key"]]
        dup = [k for k, c in Counter(keys).items() if c > 1]
        if dup:
            sys.exit("%s：1つの列名に2つの問が当たりました：%s" % (name, dup))
        nums = [q["n"] for q in qs]
        if len(set(nums)) != len(nums):
            sys.exit("%s：同じ問番号が2つあります" % name)
        out[name] = qs
    return out


# ---------------------------------------------------------------- 問題の記録
SEV = {"外す": 0, "要確認": 1, "記録のみ": 2}
HOW = {
    "アンケートがない": "記録はそのまま使う。アンケートの比較では、この区間が欠ける",
    "同じフォームに2回": "決まり：時刻が合う行のうち、いちばん早い行を使う。使わなかった行もここに残す",
    "記録のないアンケート": "ID の書き間違い・回のない記録を確かめる。決められなければ外す",
    "かえりを切り替え忘れた": "SUS（使いやすさの10問）・問4・問18〜21 は、この回の答えとして使う（ボタンの有無は回で決まる）。"
                        "かえりのバスの回答は「ゆき」の行に入っている。問6 は行きと同じ入れ方の答えなので、入れ方の比較（主な比較①）には使わない。問17 は読み替えない",
    "回の記録が途中で替わった": "回答ごとの条件（入れ方・ボタン）は記録の値。この往復の行として数えた。比べるときに使うかは、まざった件数を見て決める",
    "ボタンの回が決まらない": "回のない記録の日時を見て、どの回の記録かを人が決める。決められなければ、問18〜20 は使わない",
    "IDが名簿にない": "時刻とその便のほかの参加者から本人を決められるときだけ直す。決められなければ外す",
    "IDが読めない": "時刻とその便のほかの参加者から本人を決められるときだけ直す。決められなければ外す",
    "IDが分からない": "id対応.csv を作り直す（docs/分析の準備.md）",
    "確認用アカウント": "分析から外す",
    "練習の回答": "数から外した",
    "練習の切り忘れの疑い": "本番の回答として使うかを決める",
    "本番の回答がない": "分析から外す",
    "入れ方がまざっている": "入れ方の比較（問6・記録の時間）から外すか、回答ごとの入れ方で分けて使うかを決める",
    "条件が計画と違う": "記録の条件で分析する。なぜ違うかを確かめる",
    "区間の名前と時刻が合わない": "時刻と位置で区間を分け直すかを決める。分けられなければ、その往復を入れ方の比較から外す",
    "版が違う": "その版の画面と記録が同じかを確かめて、使うかを決める",
    "アンケートが記録より前": "区間を終える前に答えている。主な分析から外し、含めた場合と結果を比べる",
    "間違ったフォームの疑い": "内容を見て、どの区間の答えとして使うかを決める",
    "遅れて答えた": "記録のみ",
    "回が記録されていない": "回ごとの条件が決まらないので主な分析から外す（件数を書き残す）",
    "操作ログのない回答": "人数と混み具合は使う。条件はその区間の値を使う",
    "回答の行がない操作ログ": "記録のみ",
    "再送で重なった行": "1行にまとめた（いちばん早く届いた行を残した）",
    "ほかの人のログインで送った": "答えた人（variant.login_id）の区間に入れた",
    "同じ回を2度": "設計どおりの組み合わせにならない。どちらを使うかを決める",
    "回の順番が逆": "設計どおりの順番ではない。回の順番の影響を見るときに気をつける",
    "端末の時計のずれ": "時刻を直してから比べた",
    "2台の端末": "カードの取り違え・端末の貸し借りがないかを確かめる。端末に保存できない端末（Cookie をすべて止めた Safari など）は、"
              "開き直すたびに番号が変わるので、1台でも2種類以上になる",
    "回の出どころが違う": "画面と記録は同じ回。前の回のタブを開いたままにしていないかを確かめる",
    "答えの矛盾": "外さずに印を付ける。外した場合でも結果が変わらないかを確かめる",
    "答えが空か選択肢にない": "スプレッドシートで列のずれがないかを確かめる",
    "列の見出し": "スプレッドシートの1行目を確かめる",
    "計数研究の記録": "UI研究の分析には使わない（記録のみ）",
    "実験の前の行": "--since より前なので読まなかった（テストなど）",
}


class Issues:
    def __init__(self):
        self.rows = []

    def add(self, kind, sev, detail, unit=None, where="", how=None, id_="", day=""):
        r = {"種類": kind, "重さ": sev, "ID": id_, "回": "", "区間": "", "日付": day,
             "どこ": where, "内容": detail, "扱いの案": how or HOW.get(kind, "")}
        if unit is not None:
            r.update({"ID": unit["ID"], "回": unit["回"], "区間": LEG_JA.get(unit["leg"], unit["leg"] or ""),
                      "日付": unit["日付"]})
            unit["codes"].append(kind)
            if sev == "外す":
                unit["drop"].append(kind)
        self.rows.append(r)


# ---------------------------------------------------------------- 読み込み
def read_csv_any(path):
    for enc in ("utf-8-sig", "cp932"):
        try:
            return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    sys.exit("CSV を読めません（文字コード）: %s" % path)


def read_table(path):
    """Supabase から書き出した表（CSV か JSON）を、dict の list にする"""
    if path.lower().endswith(".json"):
        with open(path, encoding="utf-8-sig") as f:
            obj = json.load(f)
        if isinstance(obj, dict):
            lists = [v for v in obj.values() if isinstance(v, list)]
            if len(lists) != 1:
                sys.exit("JSON の形が読めません（行の list が1つではない）: %s" % path)
            obj = lists[0]
        if obj and isinstance(obj[0], dict) and len(obj[0]) == 1 and isinstance(next(iter(obj[0].values())), list):
            obj = next(iter(obj[0].values()))          # select json_agg(...) の1行1列
        return [dict(r) for r in obj]
    rows = read_csv_any(path).to_dict("records")
    for r in rows:
        for k, v in list(r.items()):
            if isinstance(v, str):
                s = v.strip()
                if s == "" or s.upper() == "NULL":
                    r[k] = None
                elif s[:1] in "{[":
                    try:
                        r[k] = json.loads(s)
                    except ValueError:
                        pass
    return rows


def to_bool(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return None
    s = str(v).strip().lower()
    if s in ("true", "t", "1", "yes"):
        return True
    if s in ("false", "f", "0", "no"):
        return False
    return None


def to_num(v):
    if v is None or v == "":
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else f


def to_time(v):
    """UTC の pandas.Timestamp にする。読めなければ None"""
    if v is None or v == "":
        return None
    t = pd.to_datetime(v, utc=True, errors="coerce", format="mixed")
    return None if pd.isna(t) else t


def jget(d, *keys):
    for k in keys:
        if isinstance(d, dict) and d.get(k) is not None:
            return d[k]
    return None


def norm_login(v):
    if v is None:
        return None
    s = unicodedata.normalize("NFKC", str(v)).strip().split("@")[0].upper()
    return s or None


def read_idmap(path):
    if not path:
        return {}
    m = {}
    for r in read_csv_any(path).to_dict("records"):
        sid = (r.get("surveyor_id") or r.get("id") or "").strip()
        lid = r.get("login_id") or r.get("email") or ""
        if sid and lid:
            m[sid] = norm_login(lid)
    return m


def find_sheet(folder, name):
    hits = [p for p in glob.glob(os.path.join(folder, "*.csv"))
            if os.path.basename(p) == name + ".csv" or os.path.basename(p).endswith(" - " + name + ".csv")]
    if len(hits) > 1:
        sys.exit("「%s」の CSV が2つあります: %s" % (name, hits))
    return hits[0] if hits else None


# ---------------------------------------------------------------- アンケート
def load_surveys(folder, qmap, issues, ids):
    rows = []
    for name, (rnd, leg) in FORMS.items():
        path = find_sheet(folder, name)
        if not path:
            issues.add("列の見出し", "要確認", "「%s」の CSV がありません" % name, where=name,
                       how="スプレッドシートからこのシートを書き出す（docs/分析の準備.md）")
            continue
        df = read_csv_any(path)
        heads = list(df.columns)
        ts_col = next((h for h in heads if nk(h) in ("タイムスタンプ", "Timestamp")), heads[0] if heads else None)
        by_n = defaultdict(list)
        for h in heads:
            n = qnum(h)
            if n is not None:
                by_n[n].append(h)
            elif h != ts_col and nk(h) not in ("メールアドレス", "Emailaddress", "スコア", "Score"):
                issues.add("列の見出し", "記録のみ", "問番号のない列「%s」は読まない" % h, where=name)
        col_of = {}
        for q in qmap[name]:
            cands = by_n.get(q["n"], [])
            if not cands:
                issues.add("列の見出し", "要確認", "問%d の列がない（%s）" % (q["n"], q["title"]), where=name)
                continue
            exact = [h for h in cands if nk(h) == nk(q["title"])]
            use = exact[0] if exact else cands[0]
            if len(cands) > 1:
                use = exact[0] if exact else max(cands, key=lambda h: _similar(h, q["title"]))
                issues.add("列の見出し", "要確認", "問%d の列が %d 本ある。「%s」を使った（ほか：%s）"
                           % (q["n"], len(cands), use, "／".join(h for h in cands if h != use)), where=name)
            elif not exact:
                issues.add("列の見出し", "記録のみ", "問%d の見出しの文が今のフォームと違う：「%s」（今のフォーム「%s」）。問番号でつないだ"
                           % (q["n"], use, q["title"]), where=name)
            col_of[q["n"]] = use
        for n in sorted(set(by_n) - {q["n"] for q in qmap[name]}):
            issues.add("列の見出し", "要確認", "今のフォームにない問%d の列：%s" % (n, "／".join(by_n[n])), where=name)

        for i, rec in enumerate(df.to_dict("records")):
            if not any(str(v).strip() for v in rec.values()):
                continue
            ts = None
            if ts_col:
                t = pd.to_datetime(rec.get(ts_col, ""), errors="coerce", format="mixed")
                if not pd.isna(t):
                    ts = t.tz_localize(TZ) if t.tzinfo is None else t.tz_convert(TZ)
            r = {"form": name, "form_round": rnd, "form_leg": leg, "row": i + 2, "ts": ts,
                 "where": "%s %d行目" % (name, i + 2), "ans": {}, "other": [], "flags": [],
                 "ID_raw": "", "ID": None}
            for q in qmap[name]:
                raw = str(rec.get(col_of.get(q["n"]), "") if q["n"] in col_of else "").strip()
                if q["key"] == "ID":
                    r["ID_raw"] = raw
                    continue
                val = raw
                if q["type"] == "5段階":
                    v = nk(raw)
                    val = int(v) if re.fullmatch(r"[1-5]", v) else None
                    if val is None:
                        r["flags"].append(("答えが空か選択肢にない", "要確認", "問%d の答え「%s」" % (q["n"], raw)))
                elif q["type"] in ("単一選択", "プルダウン"):
                    hit = [c for c in q["choices"] if nk(c) == nk(raw)]
                    val = hit[0] if hit else (raw or None)
                    if not hit:
                        r["flags"].append(("答えが空か選択肢にない", "要確認", "問%d の答え「%s」" % (q["n"], raw)))
                else:
                    val = raw or None
                if q["key"]:
                    r["ans"][q["key"]] = val
                else:
                    r["other"].append("問%d=%s" % (q["n"], raw))
            idn = nk(r["ID_raw"]).upper()
            if re.fullmatch(r"[A-Z]", idn):
                r["ID"] = idn
                if idn not in GROUP_OF:
                    r["flags"].append(("IDが名簿にない", "要確認", "ID「%s」はアカウントの表にない" % r["ID_raw"]))
                elif ids and idn not in ids:
                    r["flags"].append(("IDが名簿にない", "要確認", "ID「%s」は --ids の名簿にない" % r["ID_raw"]))
            else:
                r["ID"] = idn or None
                r["flags"].append(("IDが読めない", "要確認", "問1 の答え「%s」" % r["ID_raw"]))
            if ts is None:
                r["flags"].append(("答えが空か選択肢にない", "要確認", "タイムスタンプが読めない"))
            rows.append(r)
    return rows


def apply_survey_fixes(rows, path, issues):
    """アンケートの入れ間違いを、分析のときだけ直す（スプレッドシートは触らない）。

    訂正の表（CSV・見出し行あり）
        もとのフォーム,ID,何番目,直す先のフォーム,理由
        2回目_桃山台,F,1,1回目_桃山台,1回目のつもりで2回目のフォームに答えたため

    「何番目」は、そのフォームの中でその ID の行を答えた時刻の早い順に数えた番号（1から）。
    区間（ゆき／かえり）が同じフォームどうしでしか付け替えない。設問がまったく同じで、
    中身は正しく、入った箱だけが違うときのための仕組み。
    直した行には必ず印を付けて、report にも残す。
    """
    if not path:
        return
    if not os.path.exists(path):
        sys.exit("訂正の表が見つかりません: %s" % path)
    df = read_csv_any(path)
    need = ["もとのフォーム", "ID", "何番目", "直す先のフォーム"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        sys.exit("訂正の表に列がありません: %s（要る列: %s）" % ("／".join(miss), "／".join(need + ["理由"])))
    for rule in df.to_dict("records"):
        src = str(rule.get("もとのフォーム", "")).strip()
        dst = str(rule.get("直す先のフォーム", "")).strip()
        rid = nk(str(rule.get("ID", ""))).upper()
        why = str(rule.get("理由", "")).strip()
        try:
            nth = int(str(rule.get("何番目", "")).strip())
        except ValueError:
            nth = 0
        if src not in FORMS or dst not in FORMS:
            sys.exit("訂正の表のフォーム名が違います: %s → %s（使える名前: %s）"
                     % (src, dst, "／".join(FORMS)))
        if FORMS[src][1] != FORMS[dst][1]:
            sys.exit("訂正の表: %s → %s は区間が違うので付け替えません（設問が同じフォームどうしだけ）" % (src, dst))
        hit = [r for r in rows if r["form"] == src and r["ID"] == rid]
        hit.sort(key=lambda r: (r["ts"] is None, r["ts"]))
        if nth < 1 or nth > len(hit):
            issues.add("アンケートの訂正", "要確認",
                       "%s の ID「%s」は %d 行しかないので、%d 番目を直せない" % (src, rid, len(hit), nth),
                       where=src, how="訂正の表の「何番目」を直す")
            continue
        r = hit[nth - 1]
        rnd, leg = FORMS[dst]
        r["form"], r["form_round"], r["form_leg"] = dst, rnd, leg
        r["fixed_from"] = src
        r["flags"].append(("アンケートの訂正", "記録のみ",
                           "%s から %s に付け替えた%s" % (src, dst, ("：" + why) if why else "")))
        issues.add("アンケートの訂正", "記録のみ",
                   "ID「%s」の %s %d番目（%s）を %s として扱った%s"
                   % (rid, src, nth, (r["ts"].strftime("%m/%d %H:%M") if r["ts"] is not None else "時刻不明"),
                      dst, ("：" + why) if why else ""),
                   where=r["where"], how="元のスプレッドシートは直していない")


def _similar(a, b):
    from difflib import SequenceMatcher
    return SequenceMatcher(None, nk(a), nk(b)).ratio()


# ---------------------------------------------------------------- アプリの記録
def load_app(sessions_path, responses_path, idmap, issues):
    S = read_table(sessions_path)
    R = read_table(responses_path)
    n_in = (len(S), len(R))

    def dedupe(rows, key, label, kind):
        rows = sorted(rows, key=lambda r: (to_time(r.get("created_at")) or pd.Timestamp.max.tz_localize("UTC"),
                                           to_num(r.get("id")) or 0))
        seen, keep, dup = set(), [], Counter()
        for r in rows:
            k = r.get(key)
            if k and k in seen:
                dup[k] += 1
                continue
            if k:
                seen.add(k)
            keep.append(r)
        for k, c in dup.items():
            issues.add("再送で重なった行", "記録のみ", "%s %s が %d 行重なっていた" % (label, k, c + 1), where=kind)
        return keep

    S = dedupe(S, "session_id", "session_id", "ui_sessions")
    R = dedupe(R, "client_id", "client_id", "ui_responses")

    sess = []
    for r in S:
        v = r.get("variant") or {}
        c = r.get("context") or {}
        sm = r.get("summary") or {}
        ev = r.get("events") or []
        sid_user = r.get("surveyor_id")
        lv = norm_login(jget(v, "login_id", "login") or jget(c, "login_id"))
        lm = idmap.get(str(sid_user)) if sid_user else None
        lm = lm or norm_login(r.get("login_id") or r.get("email"))     # 書き出しに login_id の列を足した場合
        fixed = v.get("url_fixed") or {}
        sess.append({
            "sid": r.get("session_id"), "created": to_time(r.get("created_at")),
            "t": to_time(r.get("started_at")), "sub": to_time(r.get("client_submitted_at")),
            "outcome": r.get("outcome"), "login": lv or lm, "login_v": lv, "login_m": lm,
            "study": v.get("study"), "app": v.get("app"), "round": to_num(v.get("round")), "leg": v.get("leg"),
            "group": v.get("group"), "test": to_bool(v.get("test_account")),
            "practice": bool(to_bool(v.get("practice")) or to_bool(c.get("practice"))),
            "input": v.get("input"), "unknown_ui": to_bool(v.get("unknown_ui")),
            "url_fixed": any(x is not None for x in fixed.values()) if isinstance(fixed, dict) else False,
            "version": c.get("app_version"), "exact": to_num(c.get("exact_count")),
            "count_unknown": bool(to_bool(c.get("count_unknown"))), "cls": to_num(c.get("congestion_class")),
            "cid": c.get("response_cid"), "device": jget(c, "device_id") or jget(v, "device_id"),
            "leg_at": jget(v, "leg_at"), "leg_by": jget(v, "leg_by"), "url_study": v.get("url_study"),
            "reason": c.get("reason"),
            "t_count": to_num(sm.get("t_count")), "t_total": to_num(sm.get("t_total")),
            "t_first": to_num(sm.get("t_first")), "taps_value": to_num(sm.get("taps_value")),
            "corrections": to_num(sm.get("corrections")),
            "n_unknown_tap": sum(1 for e in ev if isinstance(e, dict) and e.get("type") == "count_unknown"),
        })
    resp = []
    for r in R:
        sid_user = r.get("surveyor_id")
        lm = idmap.get(str(sid_user)) if sid_user else None
        lm = lm or norm_login(r.get("login_id") or r.get("email"))
        ns = r.get("next_stop") or {}
        resp.append({
            "cid": r.get("client_id"), "sid": r.get("session_id"), "created": to_time(r.get("created_at")),
            "t": to_time(r.get("answered_at")), "sub": to_time(r.get("client_submitted_at")),
            "login_m": lm, "cls": to_num(r.get("congestion_class")), "exact": to_num(r.get("exact_count")),
            "practice": bool(to_bool(r.get("practice"))), "version": r.get("app_version"),
            "lat": to_num(r.get("lat")) if r.get("lat") is not None else to_num(jget(ns, "lat")),
            "pos": (to_num(r.get("lat")), to_num(r.get("lng"))),       # 端末の位置そのもの（次のバス停の位置で埋めない）
            "acc": to_num(r.get("gps_accuracy")),
        })
    return sess, resp, n_in


# ---------------------------------------------------------------- 区間を作る
def plan_input(rnd, leg):
    if rnd not in (1, 2) or leg not in ("yuki", "kaeri"):
        return None
    return "numpad" if (rnd == 1) == (leg == "yuki") else "stepper"


def plan_unknown(group, rnd):
    if group not in ("odd", "even") or rnd not in (1, 2):
        return None
    return (rnd == 2) if group == "odd" else (rnd == 1)


def jst(t):
    return None if t is None else t.tz_convert(TZ)


def fmt(t):
    return "" if t is None else jst(t).strftime("%Y-%m-%d %H:%M:%S")


def mins(a, b):
    return (a - b).total_seconds() / 60.0


def med(xs):
    xs = [x for x in xs if x is not None]
    return median(xs) if xs else None


def hm(t):
    return "" if t is None else jst(t).strftime("%H:%M")


def dist_m(a, b):
    """2点（緯度, 経度）の距離（m）。近い所どうしなので平面で近似する。どちらかが無ければ None"""
    if a is None or b is None or None in a or None in b:
        return None
    dy = (a[0] - b[0]) * 111000.0
    dx = (a[1] - b[1]) * 111000.0 * 0.8217      # cos(34.8°)
    return (dx * dx + dy * dy) ** 0.5


def cond_desc(ss):
    """記録の条件を「数字キー 3・＋−ボタン 2、ボタンあり 5」のように書く"""
    ci = Counter(INPUT_JA.get(s["input"], str(s["input"])) for s in ss if s["input"])
    cu = Counter("ボタンあり" if s["unknown_ui"] else "ボタンなし" for s in ss if s["unknown_ui"] is not None)
    return "、".join(x for x in ("・".join("%s %d" % kv for kv in sorted(ci.items())),
                                 "・".join("%s %d" % kv for kv in sorted(cu.items()))) if x) or "条件の記録なし"


def build_units(sess, resp, issues, versions, ids):
    """区間（参加者×往復×ゆき／かえり）を作る。返すのは units（区間）と trips（往復）

    往復の分け方：同じ人・同じ日の操作ログを時刻の順に並べ、60分以上あいたら次の往復。
    かえりの記録のあとに「ゆき」の記録が来て、回が替わったか20分以上あいていても次の往復。練習だけのまとまりは次の往復に入れる。
    往復の回は、その往復で最初の本番（練習でない）の操作ログの回。途中で回が替わった記録（前の回のタブを読み込み直した等）も
    その往復の行に入れて「回の記録が途中で替わった」の印を付ける。回のない記録（?round のない画面）は区間に入れず、
    往復の途中にあれば check_timing で印を付ける
    """
    # 端末の時計のずれ（ログインID×日付ごと。サーバーが受け取った時刻 − 端末が送った時刻 の中央値）
    skew_s = defaultdict(list)
    for x in sess + resp:
        lg = x.get("login") or x.get("login_m")
        if x["created"] is not None and x["sub"] is not None and lg:
            skew_s[(lg, jst(x["sub"]).date())].append((x["created"] - x["sub"]).total_seconds())
    skew = {k: median(v) for k, v in skew_s.items()}

    def fix(x, lg):
        if x["t"] is None:
            return None
        s = skew.get((lg, jst(x["t"]).date()), 0.0)
        return x["t"] + pd.Timedelta(seconds=s) if abs(s) > SKEW_USE else x["t"]

    by_sid = {s["sid"]: s for s in sess if s["sid"]}
    units, trips = {}, []
    norounds = Counter()

    def unit_of(lg, leg, day, tr):
        key = (lg, tr["round"], leg, day, tr["no"])
        if key not in units:
            units[key] = {"ID": lg, "回": tr["round"], "leg": leg, "日付": str(day), "trip": tr, "sess": [], "resp": [],
                          "orphans": [], "codes": [], "drop": [], "survey": None, "survey_all": []}
            tr["units"][leg] = units[key]
        return units[key]

    n_count = 0
    seqs = defaultdict(list)
    for s in sess:
        s["tt"] = fix(s, s["login"])
        s["unit"], s["trip"] = None, None
        if s["app"] == "counter" or s["study"] == "count":
            n_count += 1
            continue
        s["has_round"] = s["study"] == "ui" and s["round"] in (1, 2)
        if s["tt"] is None:
            norounds[(s["login"] or "?", "")] += 1
            continue
        if not s["has_round"]:
            norounds[(s["login"] or "?", str(jst(s["tt"]).date()))] += 1
        seqs[(s["login"] or "?", jst(s["tt"]).date())].append(s)
    for (lg, day) in sorted(seqs):
        seq = sorted(seqs[(lg, day)], key=lambda s: s["tt"])
        groups, prev, prev_r = [], None, None
        for s in seq:
            new = not groups or mins(s["tt"], prev["tt"]) >= TRIP_GAP_MIN
            if not new and s["has_round"] and prev_r is not None and prev_r["leg"] == "kaeri" and s["leg"] == "yuki" \
                    and (s["round"] != prev_r["round"] or mins(s["tt"], prev_r["tt"]) >= LEG_NEW_MIN):
                new = True
            if new:
                groups.append([])
                prev_r = None
            groups[-1].append(s)
            if s["has_round"]:
                prev_r = s
            prev = s
        # 練習だけのまとまり（集合の前に試した等）は往復とみない。次の往復に入れる（なければ前の往復）
        merged, carry = [], []
        for g in groups:
            if all(s["practice"] for s in g):
                carry += g
                continue
            merged.append(carry + g)
            carry = []
        if carry:
            if merged:
                merged[-1] += carry
            else:
                merged.append(carry)
        groups = merged
        for g in groups:
            # 往復の回は、本番（練習でない）の最初の回答の回。練習しかなければ練習の最初の回答の回
            #   （前の回のタブで練習してから QR を読んだときに、往復の回を練習の回にしない）
            withr = [s for s in g if s["has_round"]]
            anchor = next((s for s in withr if not s["practice"]), withr[0] if withr else None)
            tr = {"ID": lg, "day": day, "no": 1 + sum(1 for t in trips if t["ID"] == lg and t["day"] == day),
                  "round": anchor["round"] if anchor else None, "sess": [], "nr_in": [], "nr_before": [], "units": {}}
            trips.append(tr)
            for s in g:
                s["trip"] = tr
                if s["has_round"]:
                    u = unit_of(lg, s["leg"], day, tr)
                    u["sess"].append(s)
                    s["unit"] = u
                    tr["sess"].append(s)
                elif anchor is not None and s["tt"] >= anchor["tt"]:
                    tr["nr_in"].append(s)
                else:
                    # 往復の本番が始まる前の回のない記録（QR を読む前に開いた等）は、往復の条件には関わらない
                    tr["nr_before"].append(s)
    if n_count:
        issues.add("計数研究の記録", "記録のみ", "計数研究（?study=count・乗降アプリ）の操作ログ %d 行を除いた" % n_count,
                   where="ui_sessions")
    for (lg, day), c in sorted(norounds.items()):
        issues.add("回が記録されていない", "外す", "回・区間のない操作ログが %d 行（?round のない画面で開いた）。入れ方が回答ごとに替わる"
                   % c, where="ui_sessions", id_=lg, day=day)

    orphans = []
    for r in resp:
        s = by_sid.get(r["sid"]) if r["sid"] else None
        if s is not None:
            r["login"] = s["login"]
            r["tt"] = fix(r, r["login"])
            if s.get("unit") is not None:
                s["unit"]["resp"].append(r)
                r["sess"] = s
            continue
        r["login"] = r["login_m"]
        r["tt"] = fix(r, r["login"])
        orphans.append(r)
    for s in sess:
        if s.get("unit") is not None and s["outcome"] in ("submit", "skip") and s["cid"]:
            if not any(r["cid"] == s["cid"] for r in s["unit"]["resp"]):
                issues.add("回答の行がない操作ログ", "記録のみ", "session_id %s（%s）に回答の行がない" % (s["sid"], s["outcome"]),
                           unit=s["unit"], where="ui_sessions")

    # 操作ログのない回答は、同じ人のいちばん近い区間（前後30分以内）に入れる
    for r in orphans:
        cands = [u for u in units.values() if u["ID"] == r["login"] and u["sess"]]
        best, bd = None, None
        for u in cands:
            for s in u["sess"]:
                d = abs(mins(r["tt"], s["tt"])) if r["tt"] is not None else None
                if d is not None and d <= 30 and (bd is None or d < bd):
                    best, bd = u, d
        if best is None:
            issues.add("操作ログのない回答", "要確認", "client_id %s（%s %s）の区間が分からない"
                       % (r["cid"], r["login"] or "ID 不明", fmt(r["tt"])), where="ui_responses")
            continue
        best["orphans"].append(r)
        best["resp"].append(r)

    # 区間ごとの数と印
    for u in units.values():
        real_s = [s for s in u["sess"] if not s["practice"]]
        real_r = [r for r in u["resp"] if not r["practice"]]
        prac = [s for s in u["sess"] if s["practice"]]
        # times＝操作の時刻（操作ログと、操作ログのない回答）。answers＝回答の時刻（件数を数えるとき）
        times = sorted([s["tt"] for s in real_s] + [r["tt"] for r in u["orphans"] if not r["practice"] and r["tt"] is not None])
        answers = sorted(r["tt"] for r in real_r if r["tt"] is not None)
        if not times and not answers:
            times = sorted([s["tt"] for s in u["sess"]])
        both = sorted(times + answers)
        u["first"], u["last"], u["times"], u["answers"] = both[0], both[-1], times or answers, answers
        u["real_s"], u["real_r"], u["n_prac"] = real_s, real_r, len({s["sid"] for s in prac})
        u["skew"] = skew.get((u["ID"], jst(u["first"]).date()))
        lg = u["ID"]
        u["test"] = (lg not in GROUP_OF) or any(s["test"] for s in u["sess"])
        u["group"] = GROUP_OF.get(lg)
        u["inputs"] = sorted({s["input"] for s in real_s if s["input"]})
        u["unknowns"] = sorted({s["unknown_ui"] for s in real_s if s["unknown_ui"] is not None})
        u["versions"] = sorted({x["version"] for x in real_s + real_r if x["version"]})
        u["devices"] = sorted({s["device"] for s in u["sess"] if s["device"]})
        u["leg_by"] = sorted({s["leg_by"] for s in real_s if s["leg_by"]})
        sw = [to_time(s["leg_at"]) for s in real_s if s["leg_by"] == "switch" and s["leg_at"]]
        sw = [t for t in sw if t is not None]
        u["switch_at"] = None
        if sw:
            k = u["skew"] if u["skew"] is not None and abs(u["skew"]) > SKEW_USE else 0
            u["switch_at"] = min(sw) + pd.Timedelta(seconds=k)
        # 開いた URL の回と、記録の回の食い違い（回答ごとに比べる）
        n_url = Counter((s["url_study"], s["round"]) for s in real_s
                        if s["url_study"] in ("r1", "r2") and s["url_study"] != "r%d" % (s["round"] or 0))
        if lg == "?":
            issues.add("IDが分からない", "外す", "surveyor_id とログインID の対応がない操作ログ %d 行" % len(u["sess"]), unit=u,
                       where="ui_sessions")
        elif u["test"]:
            issues.add("確認用アカウント", "外す", "%s は参加者のアカウントの表にない（確認用）" % lg, unit=u)
        elif ids and lg not in ids:
            issues.add("IDが名簿にない", "要確認", "%s は --ids の名簿にない" % lg, unit=u)
        if not real_s and not real_r:
            issues.add("本番の回答がない", "外す", "練習の回答 %d 件だけ" % u["n_prac"], unit=u)
        if u["n_prac"]:
            issues.add("練習の回答", "記録のみ", "練習の回答 %d 件を数から外した" % u["n_prac"], unit=u)
            late = [s for s in prac if real_s and s["tt"] > min(x["tt"] for x in real_s)]
            # 練習は南口で乗る前にする。南口のバス停から離れた所の練習は、バスの回答を練習のまま送った疑い
            #   （南口を出てすぐの回答は南口の近くなので、この方法では見つからない）
            far = []
            for r in u["resp"]:
                d = dist_m(r["pos"], MINAMI) if r["practice"] else None
                if d is not None and d > PRAC_FAR_M and not (r["acc"] is not None and r["acc"] > PRAC_FAR_M):   # 大まかな位置は見ない
                    far.append((r, d))
            if late or not real_s or far:
                why = []
                if late:
                    why.append("本番が始まったあとの練習の回答 %d 件" % len(late))
                if not real_s:
                    why.append("この区間の回答がすべて練習")
                if far:
                    why.append("南口のバス停から離れた所での練習の回答 %d 件（%s）" % (
                        len(far), "・".join("%s %.0fm" % (hm(r["tt"]), d) for r, d in sorted(far, key=lambda x: x[0]["tt"]))))
                issues.add("練習の切り忘れの疑い", "要確認", "。".join(why), unit=u)
        if len(u["inputs"]) > 1 or len(u["unknowns"]) > 1:
            c = Counter((s["input"], s["unknown_ui"]) for s in real_s)
            issues.add("入れ方がまざっている", "要確認", "、".join("%s・ボタン%s %d件" % (INPUT_JA.get(i, i), "あり" if k else "なし", n)
                                                         for (i, k), n in sorted(c.items(), key=str)), unit=u)
        pi, pk = plan_input(u["回"], u["leg"]), plan_unknown(u["group"], u["回"])
        if not u["test"] and lg != "?":
            diffs = []
            if u["inputs"] and pi and any(i != pi for i in u["inputs"]):
                diffs.append("入れ方（計画は%s）" % INPUT_JA[pi])
            if u["unknowns"] and pk is not None and any(k != pk for k in u["unknowns"]):
                diffs.append("ボタン（計画は%s）" % ("あり" if pk else "なし"))
            gs = {s["group"] for s in real_s if s["group"]}
            if gs and gs != {u["group"]}:
                diffs.append("組（記録は%s）" % "・".join(sorted(gs)))
            if any(s["url_fixed"] for s in real_s):
                diffs.append("URL の ?input・?unknown で固定していた")
            if diffs:
                issues.add("条件が計画と違う", "要確認", "、".join(diffs), unit=u)
        bad_v = [v for v in u["versions"] if versions and v not in versions]
        if bad_v or len(u["versions"]) > 1:
            issues.add("版が違う", "要確認", "記録の版：%s（使う版：%s）" % ("・".join(u["versions"]), "・".join(versions)), unit=u)
        if u["skew"] is not None and abs(u["skew"]) > SKEW_FLAG:
            issues.add("端末の時計のずれ", "記録のみ", "端末の時計が %.0f 秒%sいた" % (
                abs(u["skew"]), "進んで" if u["skew"] < 0 else "遅れて"), unit=u)
        for (us, rr), n in sorted(n_url.items(), key=str):
            issues.add("回の出どころが違う", "記録のみ", "%d 件は ?round=%s の画面で答えたが、記録の回は %d回目（別のタブで読んだ QR の回）"
                       % (n, us[1], rr), unit=u)
        moved = [s for s in real_s if s["round"] != u["回"]]      # 練習は数に入れないので見ない
        if moved:
            urls = sorted({s["url_study"] for s in moved if s["url_study"] in ("r1", "r2")})
            issues.add("回の記録が途中で替わった", "要確認",
                       "%d 件（%s〜%s）は %s回目として記録された。この往復は %d回目で始まったので、この行に入れた。"
                       "別の回のタブ（前の日の ?round=1 など）を読み込み直した疑い%s。その記録の条件：%s" % (
                           len(moved), hm(moved[0]["tt"]), hm(moved[-1]["tt"]),
                           "・".join(str(x) for x in sorted({s["round"] for s in moved})), u["回"],
                           "（開いた URL は %s）" % "・".join("?round=" + x[1] for x in urls) if urls else "",
                           cond_desc(moved)), unit=u)
        if len(u["devices"]) > 1:
            issues.add("2台の端末", "要確認", "device_id が %d 種類" % len(u["devices"]), unit=u)
        others = sorted({s["login_m"] for s in u["sess"] if s["login_v"] and s["login_m"] and s["login_m"] != s["login_v"]})
        if others:
            n = sum(1 for s in u["sess"] if s["login_v"] and s["login_m"] and s["login_m"] != s["login_v"])
            issues.add("ほかの人のログインで送った", "要確認", "%d 件が %s のログインで送られていた" % (n, "・".join(others)), unit=u)
        if u["orphans"]:
            issues.add("操作ログのない回答", "記録のみ", "操作ログのない回答 %d 件を時刻でこの区間に入れた" % len(u["orphans"]), unit=u)
        # 区間の中の長い空き・位置の向き
        ts = u["times"]
        u["gaps"] = [(ts[i], ts[i + 1]) for i in range(len(ts) - 1) if mins(ts[i + 1], ts[i]) >= GAP_MIN]
        pts = sorted((r["tt"], r["lat"]) for r in real_r if r["lat"] is not None and r["tt"] is not None)
        u["lat_note"] = None
        if len(pts) >= 3 and u["leg"] in ("yuki", "kaeri"):
            lats = [p[1] for p in pts]
            sign = 1 if u["leg"] == "yuki" else -1
            a, b = median(lats[:2]), median(lats[-2:])
            peak = max(lats) if sign > 0 else min(lats)
            if sign * (b - a) < -LAT_FLIP:
                u["lat_note"] = "位置が%sと逆に動いた（緯度 %.4f → %.4f）" % (
                    "南口→桃山台駅" if sign > 0 else "桃山台駅→南口", a, b)
            elif sign * (peak - a) > LAT_TURN and sign * (peak - b) > LAT_TURN:
                u["lat_note"] = "位置が途中で折り返した（緯度 %.4f → %.4f → %.4f）" % (a, peak, b)
    for tr in trips:
        firsts = [u["first"] for u in tr["units"].values()] or [s["tt"] for s in tr["nr_before"]]
        tr["first"] = min(firsts) if firsts else None
    return units, trips


# ---------------------------------------------------------------- つなぐ
def match_surveys(units, surveys, issues):
    by_id = defaultdict(list)
    for u in units.values():
        by_id[u["ID"]].append(u)
    for lg, us in by_id.items():
        us.sort(key=lambda u: u["first"])
        for i, u in enumerate(us):
            u["next_start"] = us[i + 1]["first"] if i + 1 < len(us) else None
    tol = pd.Timedelta(minutes=TOL_MIN)
    for r in surveys:
        r["unit"], r["rank"] = None, None
        if r["ID"] is None or r["ts"] is None:
            continue
        best = None
        for u in by_id.get(r["ID"], []):
            if u["leg"] != r["form_leg"]:
                continue
            in_win = (r["ts"] >= u["last"] - tol) and (u["next_start"] is None or r["ts"] < u["next_start"] + tol)
            rok = u["回"] == r["form_round"]
            same = jst(u["first"]).date() == r["ts"].date()
            rank = 0 if (in_win and rok) else 1 if in_win else 2 if (same and rok) else 3 if same else 4 if rok else None
            if rank is None:
                continue
            key = (rank, abs((r["ts"] - u["last"]).total_seconds()))
            if best is None or key < best[0]:
                best = (key, u)
        if best:
            r["unit"], r["rank"] = best[1], best[0][0]
            best[1]["survey_all"].append(r)
    # 1つの区間に2行以上：時刻が合う行（rank 0・1）を先に、その中でいちばん早い行を使う
    for u in units.values():
        rows = sorted(u["survey_all"], key=lambda r: (r["rank"], r["ts"]))
        if not rows:
            continue
        u["survey"] = rows[0]
        rows[0]["used"] = True
        if len(rows) > 1:
            same = all(x["ans"] == rows[0]["ans"] for x in rows[1:])
            issues.add("同じフォームに2回", "要確認", "%d 行あった。%s を使い、%s は使わなかった。答えは%s"
                       % (len(rows), rows[0]["where"], "・".join(x["where"] for x in rows[1:]),
                          "同じ" if same else "違う"), unit=u, where=rows[0]["where"])
            for x in rows[1:]:
                x["used"] = False


def gap_around(times, ts, min_gap):
    """ts をはさむ2つの回答の間が min_gap 分以上あいていれば、その2つの時刻を返す"""
    tol = pd.Timedelta(minutes=TOL_MIN)
    for a, b in zip(times, times[1:]):
        if a - tol <= ts <= b + tol and mins(b, a) >= min_gap:
            return a, b
    return None


def check_timing(units, trips, surveys, issues):
    tol = pd.Timedelta(minutes=TOL_MIN)
    explained = set()          # 切り替えの取り違えで時刻の食い違いが説明できたアンケートの行
    # 桃山台のアンケートの時刻を、ゆきとかえりの境目とみる（往復ごと）
    for tr in trips:
        legs = tr["units"]
        y, k = legs.get("yuki"), legs.get("kaeri")
        sm = y["survey"] if y else None
        if sm is not None:
            g = gap_around(y["times"], sm["ts"], SWITCH_GAP_MIN)
            sw = k["switch_at"] if k else None
            # 先にかえりに切り替えていた（切り替えの時刻が空きのあとの回答より前）なら、切り替え忘れではない
            if g and not (sw is not None and sw < g[1] - tol):
                after = [s for s in y["real_s"] if s["tt"] >= g[1]]
                n = sum(1 for t in y["answers"] if t >= g[1])
                # 空きのあとの「ゆき」の回答が、かえりの回答のすぐ前に続いている（桃山台駅で待った空きより短い間）なら、
                # かえりのバスに乗ってから切り替えに気づいたとみる。降りる前にアンケートに答えただけなら、
                # 空きのあとの回答とかえりの最初の回答の間に桃山台駅での待ちが入る（下の「記録より前」）
                near_k = bool(k is not None and after and mins(k["first"], max(s["tt"] for s in after)) < mins(g[1], g[0]))
                if n >= 3 or y["lat_note"] or mins(g[1], g[0]) >= GAP_MIN or near_k:
                    issues.add("区間の名前と時刻が合わない", "要確認",
                               "桃山台のアンケート（%s）の前後で「ゆき」の記録が %.0f 分あいている。あとの %d 件はかえりのバスの回答の疑い"
                               "（かえりに切り替え忘れた疑い%s）%s" % (hm(sm["ts"]), mins(g[1], g[0]), n,
                                                     "。かえりに切り替えたのは %s" % hm(sw) if sw is not None else "",
                                                     "。" + y["lat_note"] if y["lat_note"] else ""), unit=y)
                    y["switch_gap"] = g
                    y["forgot"] = after
                    explained.add(id(sm))
                    if k is not None:
                        last_y = max(s["tt"] for s in after) if after else g[1]
                        issues.add("区間の名前と時刻が合わない", "要確認",
                                   "かえりのバスの最初の %d 件（%s〜%s）は「ゆき」として記録された疑い（かえりに切り替え忘れた疑い%s）。"
                                   "この行の数には入っていない" % (n, hm(g[1]), hm(last_y),
                                                         "。切り替えたのは %s" % hm(sw) if sw is not None else ""), unit=k)
            if k is not None:
                g = gap_around(k["times"], sm["ts"], SWITCH_GAP_MIN)
                n = sum(1 for t in k["answers"] if t < sm["ts"] - tol)
                if g:
                    issues.add("区間の名前と時刻が合わない", "要確認",
                               "桃山台のアンケート（%s）の前後で「かえり」の記録が %.0f 分あいている。前の %d 件はゆきのバスの回答の疑い"
                               "（早く切り替えた疑い）%s" % (hm(sm["ts"]), mins(g[1], g[0]), n,
                                                    "。" + k["lat_note"] if k["lat_note"] else ""), unit=k)
                    k["switch_gap"] = g
                    explained.add(id(sm))
                elif n and sm["ts"] >= y["last"] - tol:
                    issues.add("区間の名前と時刻が合わない", "要確認",
                               "桃山台のアンケート（%s）より前に「かえり」の回答が %d 件。早く切り替えたか、"
                               "桃山台のアンケートに遅れて答えた疑い" % (hm(sm["ts"]), n), unit=k)
        sk = k["survey"] if k else None
        if sk is not None and not k.get("switch_gap"):
            g = gap_around(k["times"], sk["ts"], SWITCH_GAP_MIN)
            n = sum(1 for t in k["answers"] if g and t >= g[1])
            if g and (n >= 3 or k["lat_note"] or mins(g[1], g[0]) >= GAP_MIN):
                issues.add("区間の名前と時刻が合わない", "要確認",
                           "南口のアンケート（%s）のあとにも「かえり」の回答が %d 件。次の往復の回答がまざっている疑い"
                           "（その回の QR を読まなかった疑い）" % (hm(sk["ts"]), n), unit=k)
                k["switch_gap"] = g
                explained.add(id(sk))
    for u in units.values():
        r = u["survey"]
        if r is None:
            continue
        if r["form_round"] != u["回"]:
            msg = "%s に答えたが、アプリの記録は %d回目" % (r["form"], u["回"])
            # 同じ回の往復がこれより前にもある＝前の回のタブ（?round=1 など）のまま次の回に来た疑い
            before = sorted({str(t["day"]) for t in trips if t["ID"] == u["ID"] and t["round"] == u["回"]
                             and t is not u["trip"] and t["first"] is not None and t["first"] < u["trip"]["first"]})
            if before:
                msg += "。前（%s）にも %d回目の記録がある。%d回目なのに、前の回のタブ（?round=%d）のままアプリを使った疑い" % (
                    "・".join(before), u["回"], r["form_round"], u["回"])
            if u["leg"] == "yuki":
                how = "桃山台の2つのフォームは問いが同じなので、記録の回の答えとして使える"
            elif r["form_round"] == 2:
                msg += ("。アプリは2回とも%d回目の条件で動いていたので、問18〜21 の読み替えの列は空にした" % u["回"] if before
                        else "。問18〜21 は2回目を終える前の答えなので、読み替えの列は空にした")
                how = "問1〜17 は同じ問いなので記録の回の答えとして使える。問18〜21 は使わない"
            else:
                how = "問1〜17 は同じ問いなので記録の回の答えとして使える。問18〜21 は欠ける"
            issues.add("間違ったフォームの疑い", "要確認", msg, unit=u, where=r["where"], how=how)
        if id(r) in explained:
            continue
        if r["ts"] < u["first"] - tol:
            if u["leg"] == "kaeri":
                issues.add("間違ったフォームの疑い", "外す", "かえりの最初の回答（%s）より前に南口のアンケートに答えた。桃山台駅で南口の QR を読んだ疑い"
                           % hm(u["first"]), unit=u, where=r["where"],
                           how="乗る前の答えなので、主な分析から外す。あとで正しい行が来ていれば、それを使う")
            else:
                issues.add("アンケートが記録より前", "要確認", "ゆきの最初の回答（%s）より前に答えた"
                           % hm(u["first"]), unit=u, where=r["where"])
        elif r["ts"] < u["last"] - tol:
            n = sum(1 for t in u["answers"] if t > r["ts"] + tol)
            issues.add("アンケートが記録より前", "要確認", "区間の最後の回答（%s）より前に答えた。あとの回答 %d 件"
                       % (hm(u["last"]), n), unit=u, where=r["where"])
        elif u.get("next_start") is not None and r["ts"] >= u["next_start"] + tol:
            issues.add("間違ったフォームの疑い", "要確認", "次の区間の最初の回答（%s）より後に答えた。%s"
                       % (hm(u["next_start"]),
                          "南口で桃山台の QR を読んだか、答え忘れて遅れて答えた疑い" if u["leg"] == "yuki" else "答え忘れて遅れて答えた疑い"),
                       unit=u, where=r["where"], how="あとから思い出した答えとして印を付け、外した場合と含めた場合の両方で確かめる")
        elif mins(r["ts"], u["last"]) > LATE_MIN:
            issues.add("遅れて答えた", "記録のみ", "区間の終わりから %.0f 分後" % mins(r["ts"], u["last"]), unit=u, where=r["where"])
    # 使わなかった南口の行でも、かえりの最初の回答より前のもの（桃山台駅で南口の QR を読んだ）は、ゆきの区間に知らせる。
    # 問2〜6 は桃山台のフォームと同じ問いなので、桃山台のアンケートがないときの手がかりになる
    for u in units.values():
        if u["leg"] != "kaeri":
            continue
        for r in sorted(u["survey_all"], key=lambda x: x["ts"]):
            if r.get("used") is False and r["ts"] < u["first"] - tol:
                y = u["trip"]["units"].get("yuki")
                no_north = y is not None and y["survey"] is None
                issues.add("間違ったフォームの疑い", "要確認" if no_north else "記録のみ",
                           "%s（%s）は、かえりの最初の回答（%s）より前の答え。桃山台駅で南口の QR を読んだ疑い。この行はかえりには使っていない%s"
                           % (r["form"], hm(r["ts"]), hm(u["first"]), "。この往復の桃山台のアンケートはない" if no_north else ""),
                           unit=y if no_north else u, where=r["where"],
                           how="問2〜6 は桃山台のフォームと同じ問いなので、ゆきの答えとして使えるかを確かめる。SUS・問17 は乗る前の答えなので使わない")
    # 往復の途中から回のない記録になった（ログアウトして ?round のない画面で入り直した）。その区間の記録の条件は決まらない
    for tr in trips:
        tr["nr_parts"] = {}
        if not tr["nr_in"] or not tr["units"]:
            continue
        y, k = tr["units"].get("yuki"), tr["units"].get("kaeri")
        bound = y["survey"]["ts"] if y is not None and y["survey"] is not None else k["first"] if k is not None else None
        tr["nr_parts"] = {"yuki": [s for s in tr["nr_in"] if bound is None or s["tt"] < bound],
                          "kaeri": [s for s in tr["nr_in"] if bound is not None and s["tt"] >= bound]}
        for leg, u in tr["units"].items():
            ss = tr["nr_parts"].get(leg) or []
            if ss:
                issues.add("回が記録されていない", "外す",
                           "この区間の途中から、回・区間のない記録になった：%d 行（%s〜%s。%s）。ログアウトして ?round のない画面で"
                           "入り直した疑い。この行の数には入っていない" % (len(ss), hm(ss[0]["tt"]), hm(ss[-1]["tt"]), cond_desc(ss)),
                           unit=u)
    # 説明のつかない長い空き・位置の向き
    for u in units.values():
        if u.get("switch_gap"):
            continue
        for g in u["gaps"]:
            issues.add("区間の名前と時刻が合わない", "要確認", "%s〜%s に %.0f 分の空き。2本のバスの回答がまざっている疑い"
                       % (hm(g[0]), hm(g[1]), mins(g[1], g[0])), unit=u)
        if u["lat_note"]:
            issues.add("区間の名前と時刻が合わない", "要確認", u["lat_note"], unit=u)
    # 同じ回を2度・回の順番
    by_id = defaultdict(list)
    for u in units.values():
        if not u["test"] and u["ID"] != "?":
            by_id[u["ID"]].append(u)
    for lg, us in by_id.items():
        trs = defaultdict(dict)
        for u in us:
            if u["real_s"] or u["real_r"]:          # 練習だけの往復（集合の前に試した等）は数えない
                trs[u["回"]][id(u["trip"])] = u["trip"]
        for rnd, tt in trs.items():
            if len(tt) > 1:
                ts_ = sorted(tt.values(), key=lambda t: t["first"])
                days = [str(t["day"]) for t in ts_]
                label = "・".join(d if days.count(d) == 1 else "%s %s〜" % (d, hm(t["first"])) for d, t in zip(days, ts_))
                for u in us:
                    if u["回"] == rnd:
                        issues.add("同じ回を2度", "要確認", "%d回目の記録が %d 回の往復にある（%s）" % (rnd, len(tt), label), unit=u)
        f1 = min((u["first"] for u in us if u["回"] == 1), default=None)
        f2 = min((u["first"] for u in us if u["回"] == 2), default=None)
        if f1 is not None and f2 is not None and f2 < f1:
            for u in us:
                issues.add("回の順番が逆", "要確認", "2回目の記録（%s）が1回目（%s）より前" % (fmt(f2), fmt(f1)), unit=u)


# ---------------------------------------------------------------- 表にする
OUT_COLS = (["ID", "組", "回", "区間", "日付", "往復", "時間帯", "入れ方", "わからないボタン", "計画の入れ方", "計画のボタン", "版",
             "本番の回答数", "人数なしの数", "人数なしの割合", "わからないで送った数", "わからないを押した回数", "5の倍数の割合",
             "混み具合6段階の中央値", "人数入力の秒_中央値", "1回答の秒_中央値", "最初の操作までの秒_中央値",
             "値を作る操作数_中央値", "直した回数_合計", "スキップ数", "戻った数", "練習の回答数", "操作ログのない回答数",
             "途中で閉じた数", "最初の回答", "最後の回答", "端末の時計のずれ_秒", "区間の決め方", "かえりに切り替えた時刻", "端末の数",
             "フォーム", "アンケートの行", "アンケートの時刻", "区間の終わりからの分",
             "問2_混み具合", "問3_数えきれなかった", "問4_だいたいで送った", "問5_1枚目の答えやすさ", "問6_2枚目の答えやすさ"]
            + SUS_COLS +
            ["SUS点", "問17_使いたい方", "問17_選んだ入れ方", "問18_ボタンがあった回", "問18_正しいか",
             "問19_送りやすかった回", "問19_ボタンで読むと", "問20_意識した回", "問20_ボタンあり回の向き",
             "問21_ボタンがあった方が良い", "自由記述", "その他の答え", "問題", "主な分析から外す理由"])

COL_DOC = [
    ("ID・組・回・区間・日付", "カードの ID（ログインID）、組（奇数組／偶数組）、1回目／2回目、ゆき／かえり、日付（日本時間）。"
                         "回は、その往復で最初に記録された回（途中で回が替わった記録も、この行に入れて印を付ける）"),
    ("往復", "その日の何本目の往復か（ふつうは1）。同じ人の記録を時刻で分ける。60分以上あいたか、かえりの記録のあとに"
           "別の回の「ゆき」が来たら次の往復。アンケートだけの行は空"),
    ("時間帯", "その往復の最初の回答が 14時より前なら昼、あとなら夕方"),
    ("入れ方・わからないボタン", "アプリの記録（variant.input・variant.unknown_ui）。練習を除く。2つ以上まざっていれば「・」でつなぐ"),
    ("計画の入れ方・計画のボタン", "実験の計画（回・区間・組）から決まる条件。記録と違えば「条件が計画と違う」の印"),
    ("版", "アプリの版（context.app_version）"),
    ("本番の回答数", "練習を除いた回答（ui_responses。同じ client_id は1行にまとめた）"),
    ("人数なしの数・割合", "人数が空の回答（何も入れない・消した・「わからない」・スキップ）"),
    ("わからないで送った数", "「わからない」の状態で送った回答（context.count_unknown）"),
    ("わからないを押した回数", "「わからない」ボタンを押した回数（events の count_unknown）"),
    ("5の倍数の割合", "入れた人数（1人以上）のうち、5の倍数の割合（だいたいの数を入れた手がかり）"),
    ("混み具合6段階の中央値", "1枚目の6段階の答えの中央値"),
    ("人数入力の秒_中央値", "人数画面を出してから送るまでの秒（summary.t_count）。送信した回答だけ"),
    ("1回答の秒_中央値・最初の操作までの秒_中央値", "summary.t_total・t_first（秒）"),
    ("値を作る操作数_中央値・直した回数_合計", "summary.taps_value・corrections"),
    ("スキップ数・戻った数", "outcome が skip・abandon の操作ログ（戻った数には、途中で閉じたものを入れない）"),
    ("途中で閉じた数", "人数を入れている途中でページが閉じた回答（outcome abandon・context.reason pagehide。uiStudy-0.9.6 から）"),
    ("練習の回答数", "練習モードの回答。上の数には入れていない"),
    ("操作ログのない回答数", "ui_sessions の行がない回答。時刻で区間に入れた（条件はその区間の値）"),
    ("最初の回答・最後の回答", "区間の最初と最後の回答の時刻（日本時間。端末の時計のずれを直した値）"),
    ("端末の時計のずれ_秒", "サーバーが受け取った時刻 − 端末が送った時刻 の中央値。60秒より大きければ時刻を直した"),
    ("区間の決め方・かえりに切り替えた時刻", "variant.leg_by（pick 最初に選んだ／switch かえりに切り替えた／back ゆきに戻した）と leg_at。uiStudy-0.9.6 から"),
    ("端末の数", "context.device_id の種類の数。2以上なら2台から送っている。uiStudy-0.9.6 から"),
    ("フォーム・アンケートの行・アンケートの時刻", "使ったアンケートの行（シート名と行番号）とタイムスタンプ"),
    ("区間の終わりからの分", "アンケートの時刻 − 区間の最後の回答。マイナスなら区間を終える前に答えた"),
    ("問2〜問6", "5段階の答え（1〜5）。向きはフォームの中身の控え（docs/questionnaire/フォームの中身_*.md）"),
    ("SUS1〜SUS10", "南口のフォームの問7〜16 の答え（1＝まったくそう思わない … 5＝とてもそう思う）"),
    ("SUS点", "SUS（使いやすさを測る決まった10問）の点。0〜100。奇数番は答え−1、偶数番は5−答え、合計×2.5（Brooke 1996）"),
    ("問17_選んだ入れ方", "問17 の答え（行き／帰り）を、その往復の記録の入れ方に読み替えたもの。選んだ区間の記録がない・"
                     "入れ方がまざる・行きと帰りが同じ入れ方（切り替え忘れ）のときは空"),
    ("問18_正しいか", "「わからない」ボタンがあった回を正しく答えたか（記録でボタンがあった回と比べる）。"
                  "回のない記録があって、ボタンがあった回が記録から決まらないときは、問18〜20 の読み替えは空"),
    ("問19_ボタンで読むと", "問19 の答え（1回目／2回目）を、ボタンあり回／ボタンなし回に読み替えたもの"),
    ("問20_ボタンあり回の向き", "問20 を、1＝ボタンなし回の方が多かった … 5＝ボタンあり回の方が多かった に向きをそろえたもの"),
    ("その他の答え", "対応の決まっていない問の答え（ふつうは空）"),
    ("問題", "この行に付いた印（問題の一覧の「種類」）"),
    ("主な分析から外す理由", "重さが「外す」の印。空なら外す理由はない（「要確認」の印は人が見て決める）"),
]


def ja_input(xs):
    return "・".join(INPUT_JA.get(x, str(x)) for x in xs) or None


def ja_unknown(xs):
    return "・".join("あり" if x else "なし" for x in xs) or None


def button_round(units, trips, lg):
    """記録で「わからない」ボタンがあった回（1か2）と、決まらないときの理由。

    回ごとに、その回の往復の記録（往復の途中の回のない記録も入れる）のボタンを見る。
    1つの回がいつも「あり」、もう1つの回がいつも「なし」ならそれで決まる。決まらず、その人に回のない記録
    （?round のない画面。ボタンはいつも出る）があれば、どちらの回の記録か分からないので決めない。
    それ以外（片方の回の記録しかない等）は、記録で「あり」の回、なければ計画から
    """
    vals = {1: set(), 2: set()}
    for u in units.values():
        if u["ID"] == lg and u["回"] in vals:
            vals[u["回"]] |= {s["unknown_ui"] for s in u["real_s"] if s["unknown_ui"] is not None}
    nr = []
    for t in trips:
        if t["ID"] != lg:
            continue
        nr += [s for s in t["nr_in"] + t["nr_before"] if not s["practice"]]
        if t["round"] in vals:
            vals[t["round"]] |= {s["unknown_ui"] for s in t["nr_in"] if not s["practice"] and s["unknown_ui"] is not None}
    on = {r for r, v in vals.items() if v == {True}}
    off = {r for r, v in vals.items() if v == {False}}
    if len(on) == 1 and len(off) == 1:
        return next(iter(on)), None
    if nr:
        nr.sort(key=lambda s: s["tt"])
        days = sorted({str(jst(s["tt"]).date()) for s in nr})
        uk = {s["unknown_ui"] for s in nr if s["unknown_ui"] is not None}
        return None, "回・区間のない記録が %d 行ある（%s）。その記録ではボタンが%s。どちらの回にボタンがあったかを記録から決められない" % (
            len(nr), "・".join(days), "出ていた" if uk == {True} else "出ていなかった" if uk == {False} else "出ていた回答と出ていない回答があった")
    if len(on) == 1:
        return next(iter(on)), None
    g = GROUP_OF.get(lg)
    return (2 if g == "odd" else 1 if g == "even" else None), None


def trip_inputs(tr):
    """往復のゆき・かえりのバスで使った入れ方（記録）。切り替え忘れで「ゆき」に入った、かえりのバスの回答はかえりに入れる"""
    if tr is None:
        return {}
    y, k = tr["units"].get("yuki"), tr["units"].get("kaeri")
    forgot = {id(s) for s in ((y.get("forgot") or []) if y else [])}
    ys = y["real_s"] if y else []
    yi = {s["input"] for s in ys if s["input"] and id(s) not in forgot}
    ki = {s["input"] for s in (k["real_s"] if k else []) if s["input"]} | {s["input"] for s in ys if s["input"] and id(s) in forgot}
    return {"yuki": yi, "kaeri": ki}


def survey_cols(r, rnd, br, legs_in):
    """アンケートの答えと、読み替えた列。br＝ボタンがあった回（決まらなければ None）、legs_in＝その往復の区間ごとの入れ方"""
    out = {}
    a = r["ans"]
    for k in ["問2_混み具合", "問3_数えきれなかった", "問4_だいたいで送った", "問5_1枚目の答えやすさ", "問6_2枚目の答えやすさ"] + SUS_COLS + \
             ["問17_使いたい方", "問18_ボタンがあった回", "問19_送りやすかった回", "問20_意識した回", "問21_ボタンがあった方が良い", "自由記述"]:
        out[k] = a.get(k)
    sus = [a.get(c) for c in SUS_COLS]
    if all(isinstance(x, int) for x in sus):
        out["SUS点"] = 2.5 * sum((x - 1) if i % 2 == 0 else (5 - x) for i, x in enumerate(sus))
    v17 = a.get("問17_使いたい方")
    if v17 and legs_in:
        leg = "yuki" if v17.startswith("行き") else "kaeri" if v17.startswith("帰り") else None
        if leg:
            # 選んだ区間の記録が1つの入れ方で、もう1つの区間と違うときだけ読み替える
            #   （記録がない・まざる・行きと帰りが同じ入れ方（切り替え忘れ）のときは、どちらの入れ方を選んだか決まらない）
            mine = legs_in.get(leg) or set()
            other = legs_in.get("kaeri" if leg == "yuki" else "yuki") or set()
            if len(mine) == 1 and mine != other:
                out["問17_選んだ入れ方"] = INPUT_JA.get(next(iter(mine)))
    if not (r["form_round"] == 2 and rnd == 1) and br:     # 1回目の記録の人が2回目のフォームに答えたときは読み替えない
        v18 = a.get("問18_ボタンがあった回")
        if v18:
            out["問18_正しいか"] = "覚えていない" if v18 == "覚えていない" else ("正しい" if v18 == "%d回目" % br else "違う")
        v19 = a.get("問19_送りやすかった回")
        if v19:
            out["問19_ボタンで読むと"] = v19 if v19 == "変わらない" else ("ボタンあり回" if v19 == "%d回目" % br else "ボタンなし回")
        v20 = a.get("問20_意識した回")
        if isinstance(v20, int):
            out["問20_ボタンあり回の向き"] = v20 if br == 2 else 6 - v20
    if r["other"]:
        out["その他の答え"] = "；".join(r["other"])
    return out


def build_table(units, trips, surveys, issues):
    rows = []
    tol = pd.Timedelta(minutes=TOL_MIN)
    # アンケートの答えの矛盾は、使った行について印を付ける
    for r in surveys:
        if r.get("used") is False:
            continue
        a = r["ans"]
        sus = [a.get(c) for c in SUS_COLS]
        if all(isinstance(x, int) for x in sus) and len(set(sus)) == 1:
            r["flags"].append(("答えの矛盾", "要確認", "SUS の10問が全部「%d」（向きが交互なので矛盾する）" % sus[0]))
        q3, q4 = a.get("問3_数えきれなかった"), a.get("問4_だいたいで送った")
        if isinstance(q3, int) and isinstance(q4, int) and q4 > q3:
            r["flags"].append(("答えの矛盾", "要確認", "問4（だいたいで送った %d）が問3（数えきれなかった %d）より大きい" % (q4, q3)))

    def survey_flags(r, unit):
        for kind, sev, detail in r["flags"]:
            issues.add(kind, sev, detail, unit=unit, where=r["where"])

    def readings(r, unit, rnd):
        """問18〜20 を読み替える回（ボタンがあった回）。記録から決まらなければ印を付ける"""
        br, why = button_round(units, trips, unit["ID"])
        uses = any(r["ans"].get(k) is not None for k in ("問18_ボタンがあった回", "問19_送りやすかった回", "問20_意識した回"))
        if br is None and why and uses and not (r["form_round"] == 2 and rnd == 1):
            issues.add("ボタンの回が決まらない", "要確認", why + "。問18〜20 の読み替えの列は空にした", unit=unit, where=r["where"])
        return br

    for key in sorted(units, key=lambda k: (k[0], k[1] or 0, k[3], k[4], 0 if k[2] == "yuki" else 1)):
        u = units[key]
        r = u["survey"]
        br = None
        if r is not None:
            survey_flags(r, u)
            br = readings(r, u, u["回"])
        if r is None and not u["test"] and u["ID"] != "?" and (u["real_s"] or u["real_r"]):
            issues.add("アンケートがない", "要確認", "%s のアンケートの行がない" % next(
                (n for n, (rr, ll) in FORMS.items() if rr == u["回"] and ll == u["leg"]), "?"), unit=u)
        rows.append(unit_row(u, r, br, trip_inputs(u["trip"])))
    # どの区間にもつながらなかったアンケートの行も、1行ずつ出す
    for r in surveys:
        if r.get("unit") is not None:
            continue
        pseudo = {"ID": r["ID"] or "?", "回": r["form_round"], "leg": r["form_leg"],
                  "日付": str(r["ts"].date()) if r["ts"] is not None else "", "codes": [], "drop": []}
        # その人のその日の往復のうち、アンケートより前に始まったいちばん遅いもの
        tr = None
        if r["ID"] and r["ts"] is not None:
            cands = [t for t in trips if t["ID"] == r["ID"] and t["day"] == r["ts"].date()
                     and t["first"] is not None and t["first"] <= r["ts"] + tol]
            tr = max(cands, key=lambda t: t["first"]) if cands else None
        y = tr["units"].get("yuki") if tr else None
        nr = []
        if tr is not None:
            # 南口のアンケートなら、桃山台のアンケートのあと（かえりの分）の回のない記録を見る
            part = tr.get("nr_parts", {}).get("kaeri") if r["form_leg"] == "kaeri" else None
            nr = [s for s in (part or (tr["nr_before"] if tr["round"] is None else tr["nr_in"])) if s["tt"] <= r["ts"] + tol]
        kind, sev, hint, recs, legs_in = "記録のないアンケート", "外す", "", None, {}
        if nr:
            hint = "。同じ往復に、回・区間のない記録が %d 行ある（%s〜%s。%s）。?round のない画面で答えた疑い" % (
                len(nr), hm(nr[0]["tt"]), hm(nr[-1]["tt"]), cond_desc(nr))
        elif (r["form_leg"] == "kaeri" and y is not None and y.get("forgot") and tr["round"] == r["form_round"]
              and r["ts"] >= y["switch_gap"][1] - tol):
            # かえり全体で切り替えを忘れた。かえりのバスの回答は「ゆき」の行に入っている
            kind, sev, recs = "かえりを切り替え忘れた", "要確認", y["forgot"]
            legs_in = trip_inputs(tr)
            hint = "。かえりのバスの回答 %d 件は「ゆき」の行に入っている（%s〜%s。かえりに切り替え忘れた疑い）" % (
                len(recs), hm(recs[0]["tt"]), hm(recs[-1]["tt"]))
        else:
            same = [u for u in units.values() if u["ID"] == r["ID"] and r["ts"] is not None
                    and jst(u["first"]).date() == r["ts"].date()]
            if same:
                last = max(same, key=lambda u: u["last"])
                hint = "。同じ日の記録：%s（最後の回答 %s。%s）" % (
                    "・".join(sorted({"%d回目%s" % (u["回"], LEG_JA.get(u["leg"], "")) for u in same})), hm(last["last"]),
                    "切り替え忘れの疑い" if r["form_leg"] == "kaeri" and last["leg"] == "yuki" else "回や区間の取り違えを確かめる")
            elif r["ID"] and not any(u["ID"] == r["ID"] for u in units.values()):
                hint = "。この ID のアプリの記録が1行もない（ID の書き間違いの疑い）"
        issues.add(kind, sev, "%s の行に合う%sの記録がない%s" % (
            r["form"], "「かえり」" if kind == "かえりを切り替え忘れた" else "アプリ", hint), unit=pseudo, where=r["where"])
        survey_flags(r, pseudo)
        br = readings(r, pseudo, r["form_round"])
        rows.append(unit_row(pseudo, r, br, legs_in, no_app=True, recs=recs))
    for r in surveys:
        if r.get("used") is False:
            for kind, sev, detail in r["flags"]:
                issues.add(kind, "記録のみ", detail + "（使わなかった行）", where=r["where"])
    rows.sort(key=lambda x: (str(x["ID"]), x["回"] or 0, x["日付"] or "", x["往復"] or 1, 0 if x["区間"] == "ゆき" else 1))
    return rows


def unit_row(u, r, br, legs_in, no_app=False, recs=None):
    """表の1行。recs：記録のない行に、条件だけ並べる操作ログ（かえりを切り替え忘れた行の、かえりのバスの回答）"""
    lg = u["ID"]
    g = GROUP_OF.get(lg)
    row = {c: None for c in OUT_COLS}
    row.update({"ID": lg, "組": GROUP_JA.get(g, "確認用" if lg != "?" else None), "回": u["回"],
                "区間": LEG_JA.get(u["leg"], u["leg"]), "日付": u["日付"]})
    pi, pk = plan_input(u["回"], u["leg"]), plan_unknown(g, u["回"])
    row["計画の入れ方"] = INPUT_JA.get(pi)
    row["計画のボタン"] = None if pk is None else ("あり" if pk else "なし")
    if recs:
        real = [s for s in recs if not s["practice"]] or recs
        row.update({"入れ方": ja_input(sorted({s["input"] for s in real if s["input"]})),
                    "わからないボタン": ja_unknown(sorted({s["unknown_ui"] for s in real if s["unknown_ui"] is not None})),
                    "版": "・".join(sorted({s["version"] for s in real if s["version"]})) or None})
    if not no_app:
        rs, rr = u["real_s"], u["real_r"]
        row.update({"往復": u["trip"]["no"], "入れ方": ja_input(u["inputs"]), "わからないボタン": ja_unknown(u["unknowns"]),
                    "版": "・".join(u["versions"]) or None,
                    "本番の回答数": len(rr), "人数なしの数": sum(1 for x in rr if x["exact"] is None),
                    "わからないで送った数": sum(1 for s in rs if s["count_unknown"]),
                    "わからないを押した回数": sum(s["n_unknown_tap"] for s in rs),
                    "スキップ数": sum(1 for s in rs if s["outcome"] == "skip"),
                    "戻った数": sum(1 for s in rs if s["outcome"] == "abandon" and s["reason"] != "pagehide"),
                    "途中で閉じた数": sum(1 for s in rs if s["outcome"] == "abandon" and s["reason"] == "pagehide"),
                    "区間の決め方": "・".join(u["leg_by"]) or None, "かえりに切り替えた時刻": fmt(u["switch_at"]) or None,
                    "端末の数": len(u["devices"]) or None,
                    "練習の回答数": u["n_prac"], "操作ログのない回答数": len([x for x in u["orphans"] if not x["practice"]]),
                    "最初の回答": fmt(u["first"]), "最後の回答": fmt(u["last"]),
                    "端末の時計のずれ_秒": None if u["skew"] is None else round(u["skew"])})
        if rr:
            row["人数なしの割合"] = round(row["人数なしの数"] / len(rr), 3)
        nums = [x["exact"] for x in rr if isinstance(x["exact"], (int, float)) and x["exact"] > 0]
        if nums:
            row["5の倍数の割合"] = round(sum(1 for x in nums if x % 5 == 0) / len(nums), 3)
        row["混み具合6段階の中央値"] = med([x["cls"] for x in rr])
        sub = [s for s in rs if s["outcome"] == "submit"]
        for col, k in (("人数入力の秒_中央値", "t_count"), ("1回答の秒_中央値", "t_total"), ("最初の操作までの秒_中央値", "t_first")):
            m = med([s[k] for s in sub])
            row[col] = None if m is None else round(m / 1000.0, 2)
        row["値を作る操作数_中央値"] = med([s["taps_value"] for s in sub])
        row["直した回数_合計"] = sum(s["corrections"] or 0 for s in rs)
        row["時間帯"] = "昼" if jst(u["trip"]["first"] or u["first"]).hour < NOON_END else "夕方"
    if r is not None:
        row.update({"フォーム": r["form"], "アンケートの行": r["row"], "アンケートの時刻": fmt(r["ts"])})
        if no_app and r["ts"] is not None:
            row["時間帯"] = "昼" if r["ts"].hour < NOON_END else "夕方"     # 記録がないので、アンケートの時刻で決める
        if not no_app:
            row["区間の終わりからの分"] = round(mins(r["ts"], u["last"]), 1)
        row.update(survey_cols(r, u["回"], br, legs_in))
    row["問題"] = "、".join(dict.fromkeys(u["codes"])) or None
    row["主な分析から外す理由"] = "、".join(dict.fromkeys(u["drop"])) or None
    return row


def tidy(df):
    """整数だけの列は、空があっても整数のまま書き出す（1.0 にしない）"""
    for c in df.columns:
        vals = [v for v in df[c] if v is not None and not (isinstance(v, float) and pd.isna(v))]
        if vals and all(isinstance(v, int) and not isinstance(v, bool) for v in vals):
            df[c] = df[c].astype("Int64")
    return df


def write_out(out, table, issues, surveys):
    os.makedirs(out, exist_ok=True)
    t = tidy(pd.DataFrame(table, columns=OUT_COLS, dtype=object))
    iss = pd.DataFrame(issues.rows, columns=["種類", "重さ", "ID", "回", "区間", "日付", "どこ", "内容", "扱いの案"])
    iss = iss.sort_values(["重さ", "ID", "回", "区間", "種類"], key=lambda s: s.map(SEV) if s.name == "重さ" else s.astype(str),
                          kind="stable")
    srows = []
    for r in surveys:
        u = r.get("unit")
        srows.append({"フォーム": r["form"], "行": r["row"], "時刻": fmt(r["ts"]), "問1の答え": r["ID_raw"], "ID": r["ID"],
                      "つないだ区間": "" if u is None else "%s %d回目 %s %s" % (u["ID"], u["回"], LEG_JA.get(u["leg"], ""), u["日付"]),
                      "使ったか": "使った" if r.get("used") else ("使わなかった（同じ区間に別の行）" if r.get("used") is False
                                                           else "記録のない行として出した")})
    sv = pd.DataFrame(srows, columns=["フォーム", "行", "時刻", "問1の答え", "ID", "つないだ区間", "使ったか"])
    col = iss[iss["種類"] == "列の見出し"]
    t.to_csv(os.path.join(out, "区間ごと.csv"), index=False, encoding="utf-8-sig")
    iss.to_csv(os.path.join(out, "問題の一覧.csv"), index=False, encoding="utf-8-sig")
    sv.to_csv(os.path.join(out, "アンケートの行.csv"), index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(os.path.join(out, "つないだ表.xlsx"), engine="openpyxl") as xw:
        for name, df in (("区間ごと", t), ("問題の一覧", iss), ("アンケートの行", sv), ("列の確かめ", col),
                         ("読み方", pd.DataFrame(COL_DOC, columns=["列", "意味"]))):
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "B2" if name == "区間ごと" else "A2"
            for i, c in enumerate(df.columns, start=1):
                w = max([len(str(c))] + [len(str(v)) for v in df[c].head(200) if v is not None]) + 2
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(w, 6), 60)
    return t, iss


def default_versions():
    try:
        s = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()
        m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', s)
        return [m.group(1)] if m else []
    except OSError:
        return []


def main(argv=None):
    ap = argparse.ArgumentParser(description="アンケートの答えとアプリの記録をつなぐ（UI研究）")
    ap.add_argument("--forms", required=True, help="4シートの CSV を入れたフォルダ")
    ap.add_argument("--sessions", required=True, help="ui_sessions（CSV か JSON）")
    ap.add_argument("--responses", required=True, help="ui_responses（CSV か JSON）")
    ap.add_argument("--idmap", help="surveyor_id・login_id の CSV")
    ap.add_argument("--out", required=True, help="出力フォルダ")
    ap.add_argument("--ids", help="参加する人の ID（例 A,B,C）")
    ap.add_argument("--since", help="この日より前の行を読まない（例 2026-10-13）")
    ap.add_argument("--app-version", help="分析に使うアプリの版（, 区切り）。省くと index.html の APP_VERSION")
    ap.add_argument("--form-json", help="フォームの中身の控え（.json）")
    ap.add_argument("--survey-fixes",
                    help="アンケートの入れ間違いを直す表（CSV）。見出しは もとのフォーム,ID,何番目,直す先のフォーム,理由")
    a = ap.parse_args(argv)

    fj = a.form_json or latest_form_json()
    with open(fj, encoding="utf-8") as f:
        qmap = question_map(json.load(f))
    for name, qs in qmap.items():
        for q in qs:
            if not q["key"]:
                print("注意：%s の問%d は表の列名が決まっていない（「その他の答え」に入れる）：%s" % (name, q["n"], q["title"]))
    ids ={x.strip().upper() for x in (a.ids or "").split(",") if x.strip()}
    versions = [v.strip() for v in a.app_version.split(",")] if a.app_version else default_versions()
    issues = Issues()
    surveys = load_surveys(a.forms, qmap, issues, ids)
    apply_survey_fixes(surveys, a.survey_fixes, issues)   # 分析のときだけ付け替える（元の表は触らない）
    idmap = read_idmap(a.idmap)
    sess, resp, n_in = load_app(a.sessions, a.responses, idmap, issues)
    if a.since:
        since = pd.Timestamp(a.since).date()
        old = lambda t: t is not None and jst(t).date() < since
        n = (sum(1 for r in surveys if r["ts"] is not None and r["ts"].date() < since),
             sum(1 for x in sess if old(x["t"])), sum(1 for x in resp if old(x["t"])))
        surveys = [r for r in surveys if not (r["ts"] is not None and r["ts"].date() < since)]
        sess = [x for x in sess if not old(x["t"])]
        resp = [x for x in resp if not old(x["t"])]
        if any(n):
            issues.add("実験の前の行", "記録のみ", "%s より前の行を読まなかった：アンケート %d 行・操作ログ %d 行・回答 %d 行"
                       % (a.since, n[0], n[1], n[2]))
    units, trips = build_units(sess, resp, issues, versions, ids)
    match_surveys(units, surveys, issues)
    check_timing(units, trips, surveys, issues)
    table = build_table(units, trips, surveys, issues)
    t, iss = write_out(a.out, table, issues, surveys)

    print("フォームの中身の控え: %s" % os.path.relpath(fj, ROOT))
    print("読んだ行: アンケート %d 行／ui_sessions %d 行／ui_responses %d 行（Supabase の count(*) と同じか確かめる）"
          % (len(surveys), n_in[0], n_in[1]))
    print("使う版: %s" % ("・".join(versions) or "（指定なし）"))
    print("出した行: 区間ごと %d 行（うち外す理由あり %d 行）／問題 %d 件"
          % (len(t), t["主な分析から外す理由"].notna().sum(), len(iss)))
    for (kind, sev), n in sorted(Counter(zip(iss["種類"], iss["重さ"])).items(), key=lambda x: (SEV[x[0][1]], x[0][0])):
        print("  [%s] %s %d" % (sev, kind, n))
    print("書き出しました: %s" % os.path.abspath(a.out))
    return t, iss


if __name__ == "__main__":
    main()
