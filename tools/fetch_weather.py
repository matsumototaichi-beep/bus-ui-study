# -*- coding: utf-8 -*-
"""回答した時刻の天気を、あとから自動で取る（アプリは変えない）

    python tools/fetch_weather.py --dates 2026-10-13
        その日の天気の表を作る（日付はカンマ区切り、または 2026-10-13..2026-10-20）
    python tools/fetch_weather.py --responses exports/2026-10-20/ui_responses.csv
        Supabase から書き出した ui_responses（CSV か JSON）の1回答ごとに、そのときの天気を付ける。日付は回答から決める
    python tools/fetch_weather.py --sessions exports/2026-10-20/ui_sessions.csv
        操作ログの1行ごとに付ける（乗降アプリは1バス停＝1行なので、計数研究のかえりはこちら）

書き出すもの（--out。既定は exports/weather/。exports/ は .gitignore で GitHub に上げない）
  天気_10分_豊中.csv   気象庁アメダス「豊中」の10分値（降水量・気温・風速・風向）。実際に測った値
  天気_1時間_吹田.csv  Open-Meteo の1時間値（天気・雲量・降水量・気温・湿度）。地点は路線の中ほど
  回答と天気.csv       --responses を渡したとき。1回答に1行
  操作ログと天気.csv   --sessions を渡したとき。1行に1行

★2026-10-13: 参加者の声を受けて松本が決めた（HANDOFF §8 #51）。天気（とくに雨）はこの路線の混み具合に大きく効くとみる。
  アプリで取らない理由：実験の途中でアプリを変えると、同じ実験の中で記録の中身が変わる（10/13 の分には天気がない、など）。
  気象庁の過去の気象データ（10分値）と Open-Meteo は、あとから何日分でも取れる。回答の時刻があれば足りる。
  豊中のアメダス（伊丹空港の近く。桃山台駅から西へ約5km）がいちばん近い。湿度と日照は測っていないので、湿度は Open-Meteo を見る。

10分値の時刻は、その10分の終わり（「12:10」は 12:00〜12:10 の値）。回答の時刻を含む10分を使う。
Open-Meteo の降水量は、その時刻までの1時間の合計。気温・天気などはその時刻の値。回答の時刻にいちばん近い時刻の値を使う
（降水量だけは、回答の時刻を含む1時間）。
気象庁のページは1日に1ページずつ読み、あいだを1秒あける。
"""
import argparse
import csv
import datetime as dt
import html
import json
import math
import os
import re
import sys
import time
import urllib.parse
import urllib.request

sys.dont_write_bytecode = True
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JST = dt.timezone(dt.timedelta(hours=9))
UA = {"User-Agent": "bus-ui-study weather (graduation research; contact via GitHub matsumototaichi-beep)"}

# 気象庁 過去の気象データ（10分値）。豊中はアメダス（a）、大阪は気象台（s）でページが違う
STATIONS = {
    "豊中": {"kind": "a", "prec_no": 62, "block_no": "0602", "amedas": "62051"},
    "大阪": {"kind": "s", "prec_no": 62, "block_no": "47772", "amedas": "62078"},
}
ETRN = "https://www.data.jma.go.jp/stats/etrn/view/10min_%s1.php?prec_no=%d&block_no=%s&year=%d&month=%d&day=%d&view="
# 路線の中ほど（ＪＲ吹田駅南口と桃山台駅のあいだ）
LAT, LNG = 34.785, 135.505
OM_VARS = "weather_code,cloud_cover,precipitation,temperature_2m,relative_humidity_2m,wind_speed_10m"

# WMO の天気コード → 日本語（Open-Meteo の weather_code）
WMO = {0: "快晴", 1: "晴れ", 2: "晴れ時々曇り", 3: "曇り", 45: "霧", 48: "霧",
       51: "霧雨（弱）", 53: "霧雨", 55: "霧雨（強）", 56: "霧雨（着氷）", 57: "霧雨（着氷）",
       61: "雨（弱）", 63: "雨", 65: "雨（強）", 66: "雨（着氷）", 67: "雨（着氷）",
       71: "雪（弱）", 73: "雪", 75: "雪（強）", 77: "霧雪",
       80: "にわか雨（弱）", 81: "にわか雨", 82: "にわか雨（強）", 85: "にわか雪", 86: "にわか雪（強）",
       95: "雷雨", 96: "雷雨（ひょう）", 99: "雷雨（ひょう）"}


def get(url, tries=3):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as e:                     # 通信の失敗は少し待って読み直す
            if i == tries - 1:
                raise RuntimeError("読めません: %s（%s）" % (url, e))
            time.sleep(2 + 2 * i)


# ---------------------------------------------------------------- 気象庁 10分値
def num(s, zero_if_none=False):
    """気象庁の表の値を数にする。「///」「×」は欠測で None、「--」は現象なしで 0。
    「)」（準正常値）「]」（資料不足）などの印は外して数だけ使う"""
    s = (s or "").strip()
    if s == "--":
        return 0.0
    m = re.match(r"^-?\d+(\.\d+)?", s)
    if not m:
        return 0.0 if (zero_if_none and s == "") else None
    return float(m.group(0))


def parse_etrn(page, day, kind):
    """10分値のページ（1日分）を [{time(JST の datetime。10分の終わり), 降水量, 気温, 湿度, 風速, 風向, 最大瞬間風速}] にする"""
    rows = []
    for tr in re.findall(r'<tr class="mtx"[^>]*>(.*?)</tr>', page, re.S):
        cells = [html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)]
        if not cells or not re.match(r"^\d{1,2}:\d{2}$", cells[0]):
            continue
        hh, mm = map(int, cells[0].split(":"))
        t = dt.datetime(day.year, day.month, day.day, tzinfo=JST) + dt.timedelta(hours=hh, minutes=mm)   # 24:00 は次の日の 0:00
        if kind == "a":       # 時分・降水量・気温・相対湿度・平均風速・風向・最大瞬間風速・風向・日照時間
            c = cells + [""] * 9
            rows.append({"time": t, "降水量_mm": num(c[1]), "気温_C": num(c[2]), "湿度_%": num(c[3]),
                         "風速_m/s": num(c[4]), "風向": c[5] if c[5] not in ("///", "×") else "", "最大瞬間風速_m/s": num(c[6])})
        else:                 # 気象台：時分・現地気圧・海面気圧・降水量・気温・相対湿度・平均風速・風向・最大瞬間風速・風向・日照時間
            c = cells + [""] * 11
            rows.append({"time": t, "降水量_mm": num(c[3]), "気温_C": num(c[4]), "湿度_%": num(c[5]),
                         "風速_m/s": num(c[6]), "風向": c[7] if c[7] not in ("///", "×") else "", "最大瞬間風速_m/s": num(c[8])})
    return rows


BOSAI = "https://www.jma.go.jp/bosai/amedas/data/point/%s/%s_%02d.json"     # 速報（3時間ごとのファイル。直近10日ほど）
DIR16 = ["静穏", "北北東", "北東", "東北東", "東", "東南東", "南東", "南南東", "南",
         "南南西", "南西", "西南西", "西", "西北西", "北西", "北北西", "北"]


def jma_quick(station, d):
    """その日の確定値がまだ出ていないときの速報値（気象庁のアメダスのページと同じもの）"""
    rows = []
    for hh in range(0, 24, 3):
        try:
            data = json.loads(get(BOSAI % (STATIONS[station]["amedas"], d.strftime("%Y%m%d"), hh), tries=2))
        except (RuntimeError, ValueError):
            continue                      # まだ来ていない時間帯
        for k, v in sorted(data.items()):
            t = dt.datetime.strptime(k, "%Y%m%d%H%M%S").replace(tzinfo=JST)
            f = lambda name: (v.get(name) or [None, 1])[0] if (v.get(name) or [None, 1])[1] == 0 else None   # 品質が正常(0)のものだけ
            wd = f("windDirection")
            rows.append({"time": t, "降水量_mm": f("precipitation10m"), "気温_C": f("temp"), "湿度_%": f("humidity"),
                         "風速_m/s": f("wind"), "風向": DIR16[wd] if isinstance(wd, int) and 0 <= wd <= 16 else "",
                         "最大瞬間風速_m/s": None, "出どころ": "速報"})
        time.sleep(0.3)
    return [r for r in rows if r["time"].date() == d or (r["time"] - dt.timedelta(minutes=1)).date() == d]


def jma_10min(station, days):
    st = STATIONS[station]
    out = []
    for i, d in enumerate(days):
        if i:
            time.sleep(1)
        page = get(ETRN % (st["kind"], st["prec_no"], st["block_no"], d.year, d.month, d.day))
        rows = parse_etrn(page, d, st["kind"])
        for r in rows:
            r["出どころ"] = "確定"
        if not rows:
            rows = jma_quick(station, d)
            print("  %s の %s：確定した10分値がまだ出ていないので、速報の値を使います（%d 件）。あとで流し直すと確定値に替わります"
                  % (station, d, len(rows)))
        out += rows
    return out


# ---------------------------------------------------------------- Open-Meteo 1時間値
def open_meteo(days):
    """路線の中ほどの1時間値。直近（約3か月）は forecast、それより前は historical-forecast で読む"""
    a, b = min(days), max(days)
    q = urllib.parse.urlencode({"latitude": LAT, "longitude": LNG, "hourly": OM_VARS, "timezone": "Asia/Tokyo",
                                "start_date": a.isoformat(), "end_date": b.isoformat()})
    data = None
    for host in ("https://api.open-meteo.com/v1/forecast", "https://historical-forecast-api.open-meteo.com/v1/forecast"):
        try:
            data = json.loads(get(host + "?" + q, tries=2))
            if "hourly" in data:
                break
        except RuntimeError:
            data = None
    if not data or "hourly" not in data:
        print("  Open-Meteo が読めませんでした（気象庁の10分値だけ使います）")
        return []
    h = data["hourly"]
    rows = []
    for i, ts in enumerate(h["time"]):
        code = h["weather_code"][i]
        rows.append({"time": dt.datetime.fromisoformat(ts).replace(tzinfo=JST),
                     "天気": WMO.get(code, "" if code is None else str(code)), "天気コード": code,
                     "雲量_%": h["cloud_cover"][i], "降水量_mm": h["precipitation"][i],
                     "気温_C": h["temperature_2m"][i], "湿度_%": h["relative_humidity_2m"][i],
                     "風速_km/h": h["wind_speed_10m"][i]})
    return rows


# ---------------------------------------------------------------- 回答に付ける
def bin10(t):
    """回答の時刻 → その時刻を含む10分値の時刻（10分の終わり）。12:03 → 12:10、12:10:00 → 12:10、12:10:01 → 12:20"""
    t = t.astimezone(JST)
    base = t.replace(minute=0, second=0, microsecond=0)
    return base + dt.timedelta(minutes=10 * math.ceil((t - base).total_seconds() / 600))


def attach(t, jma, om, station):
    """1つの時刻に付ける天気の列"""
    out = {}
    k = bin10(t)
    j = jma.get(k)
    prev = [jma.get(k - dt.timedelta(minutes=10 * i)) for i in range(3)]
    p30 = [x["降水量_mm"] for x in prev if x and x["降水量_mm"] is not None]
    out["%s_10分の終わり" % station] = k.strftime("%Y-%m-%d %H:%M") if j else ""
    out["%s_降水量_10分_mm" % station] = "" if not j or j["降水量_mm"] is None else j["降水量_mm"]
    out["%s_降水量_30分_mm" % station] = round(sum(p30), 1) if len(p30) == 3 else ""
    out["%s_気温_C" % station] = "" if not j or j["気温_C"] is None else j["気温_C"]
    out["%s_風速_m/s" % station] = "" if not j or j["風速_m/s"] is None else j["風速_m/s"]
    out["雨（実測）"] = "" if len(p30) < 3 else ("あり" if sum(p30) > 0 else "なし")
    out["%s_出どころ" % station] = j.get("出どころ", "") if j else ""
    near = t.astimezone(JST).replace(minute=0, second=0, microsecond=0)
    if t.astimezone(JST).minute >= 30:
        near += dt.timedelta(hours=1)
    o = om.get(near)
    end = t.astimezone(JST).replace(minute=0, second=0, microsecond=0) + dt.timedelta(hours=1)   # 回答の時刻を含む1時間
    op = om.get(end)
    out["天気（Open-Meteo）"] = o["天気"] if o else ""
    out["雲量_%"] = o["雲量_%"] if o else ""
    out["湿度_%"] = o["湿度_%"] if o else ""
    out["降水量_1時間_mm（Open-Meteo）"] = op["降水量_mm"] if op else ""
    return out


def to_utc(v):
    if v in (None, ""):
        return None
    s = str(v).strip().replace(" ", "T", 1)
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    s = re.sub(r"([+-]\d\d)$", r"\1:00", s)                 # Supabase の CSV は「+00」で終わる
    s = re.sub(r"\.(\d{1,6})\d*", lambda m: "." + m.group(1).ljust(6, "0"), s)
    try:
        t = dt.datetime.fromisoformat(s)
    except ValueError:
        return None
    return t if t.tzinfo else t.replace(tzinfo=dt.timezone.utc)


def read_rows(path):
    if path.lower().endswith(".json"):
        with open(path, encoding="utf-8-sig") as f:
            obj = json.load(f)
        if isinstance(obj, dict):
            obj = next(v for v in obj.values() if isinstance(v, list))
        return obj
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    sys.exit("CSV を読めません（文字コード）: %s" % path)


def jload(v):
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v) if v else {}
    except (TypeError, ValueError):
        return {}


# ---------------------------------------------------------------- 実行
def parse_dates(s):
    out = set()
    for part in s.split(","):
        part = part.strip()
        if ".." in part:
            a, b = (dt.date.fromisoformat(x) for x in part.split(".."))
            while a <= b:
                out.add(a)
                a += dt.timedelta(days=1)
        elif part:
            out.add(dt.date.fromisoformat(part))
    return sorted(out)


def write_csv(path, rows, cols):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: (v.strftime("%Y-%m-%d %H:%M") if isinstance(v, dt.datetime) else v) for k, v in r.items()})


def inside(path, folder):
    path, folder = (os.path.normcase(os.path.realpath(x)) for x in (path, folder))
    try:
        return os.path.commonpath([path, folder]) == folder
    except ValueError:
        return False


def main():
    ap = argparse.ArgumentParser(description="回答した時刻の天気を、気象庁と Open-Meteo からあとで取る")
    ap.add_argument("--dates", help="日付（2026-10-13、カンマ区切り、2026-10-13..2026-10-20）")
    ap.add_argument("--responses", help="Supabase から書き出した ui_responses（CSV か JSON）")
    ap.add_argument("--sessions", help="Supabase から書き出した ui_sessions（CSV か JSON）")
    ap.add_argument("--station", default="豊中", choices=sorted(STATIONS), help="気象庁の観測所（既定 豊中）")
    ap.add_argument("--out", default=os.path.join(ROOT, "exports", "weather"), help="書き出すフォルダ")
    a = ap.parse_args()
    if not (a.dates or a.responses or a.sessions):
        ap.error("--dates・--responses・--sessions のどれかを渡してください")
    # 回答の行を書き出すので、public のリポジトリでは exports/（.gitignore）の中か、リポジトリの外にだけ書く
    if inside(a.out, ROOT) and not inside(a.out, os.path.join(ROOT, "exports")):
        sys.exit("--out はリポジトリの exports/ の中か、リポジトリの外にしてください（回答の行が入るため）")

    items = []          # (種類, 元の行, UTC の時刻)
    if a.responses:
        for r in read_rows(a.responses):
            items.append(("回答", r, to_utc(r.get("answered_at")) or to_utc(r.get("created_at"))))
    if a.sessions:
        for r in read_rows(a.sessions):
            items.append(("操作ログ", r, to_utc(r.get("started_at")) or to_utc(r.get("created_at"))))
    days = set(parse_dates(a.dates)) if a.dates else set()
    days |= {t.astimezone(JST).date() for _, _, t in items if t}
    today = dt.datetime.now(JST).date()
    if not days:
        sys.exit("日付が決まりません（回答の時刻が読めない）")
    days = sorted(days)
    print("日付: %s" % "、".join(d.isoformat() for d in days))

    os.makedirs(a.out, exist_ok=True)
    jrows = jma_10min(a.station, days)
    orows = open_meteo(days)
    jma = {r["time"]: r for r in jrows}
    om = {r["time"]: r for r in orows}
    write_csv(os.path.join(a.out, "天気_10分_%s.csv" % a.station), jrows,
              ["time", "降水量_mm", "気温_C", "湿度_%", "風速_m/s", "風向", "最大瞬間風速_m/s", "出どころ"])
    write_csv(os.path.join(a.out, "天気_1時間_吹田.csv"), orows,
              ["time", "天気", "天気コード", "雲量_%", "降水量_mm", "気温_C", "湿度_%", "風速_km/h"])

    # 便のある時間帯の雨（昼 11:30〜14:40、夕方 15:20〜18:10）
    for d in days:
        for name, (h1, m1), (h2, m2) in (("昼", (11, 30), (14, 40)), ("夕方", (15, 20), (18, 10))):
            a0 = dt.datetime(d.year, d.month, d.day, h1, m1, tzinfo=JST)
            b0 = dt.datetime(d.year, d.month, d.day, h2, m2, tzinfo=JST)
            sel = [r for r in jrows if a0 < r["time"] <= b0]
            ps = [r["降水量_mm"] for r in sel if r["降水量_mm"] is not None]
            ts = [r["気温_C"] for r in sel if r["気温_C"] is not None]
            ws = sorted({om[k]["天気"] for k in om if a0 <= k <= b0 and om[k]["天気"]})
            if sel:
                print("  %s %s：雨 %.1fmm（雨の10分 %d 回）・気温 %.1f〜%.1f℃・天気 %s%s" % (
                    d, name, sum(ps), sum(1 for p in ps if p > 0), min(ts) if ts else float("nan"),
                    max(ts) if ts else float("nan"), "／".join(ws) or "—",
                    "（途中まで）" if d == today else ""))

    for kind, fname in (("回答", "回答と天気.csv"), ("操作ログ", "操作ログと天気.csv")):
        rows = [(r, t) for k, r, t in items if k == kind]
        if not rows:
            continue
        out, cols, miss = [], None, 0
        for r, t in rows:
            v = jload(r.get("variant"))
            c = jload(r.get("context"))
            base = {"session_id": r.get("session_id") or "", "client_id": r.get("client_id") or c.get("response_cid") or "",
                    "login_id": v.get("login_id") or "", "回": v.get("round") or "", "区間": v.get("leg") or "",
                    "研究": v.get("study") or "", "outcome": r.get("outcome") or "",
                    "練習": r.get("practice") if r.get("practice") not in (None, "") else (v.get("practice") or ""),
                    "時刻_日本時間": t.astimezone(JST).strftime("%Y-%m-%d %H:%M:%S") if t else ""}
            w = attach(t, jma, om, a.station) if t else {}
            if not t or not w.get("%s_10分の終わり" % a.station):
                miss += 1
            row = {**base, **w}
            cols = cols or list(base) + list(attach(dt.datetime(2000, 1, 1, tzinfo=JST), {}, {}, a.station))
            out.append(row)
        write_csv(os.path.join(a.out, fname), out, cols)
        print("%s: %d 行に天気を付けました%s" % (fname, len(out),
              "（気象庁の10分値が無い行 %d。その日のうちは、次の日に流し直す）" % miss if miss else ""))
    print("書き出しました: %s" % a.out)


if __name__ == "__main__":
    main()
