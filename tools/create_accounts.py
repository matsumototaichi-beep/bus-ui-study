# -*- coding: utf-8 -*-
"""参加者アカウントを事前に作る（2026-09-26）

実験当日、バス停で登録作業をさせない。てこずると発車に間に合わない。
メールアドレスを使わない方式なので **パスワードを忘れても再設定できない**うえ、
参加者は別の日にもう一度ログインする。カードに刷って渡すのがいちばん確実。

  p01 〜 p24 のアカウントを Supabase に作り、
    accounts.csv        … 対応表（松本の手元用）
    accounts_cards.html … A4に8面で刷る配布カード（QR付き）
  を書き出す。**どちらも .gitignore 済み。public リポジトリなので絶対に commit しない。**

使い方（PowerShell）:
    $env:SUPABASE_SERVICE_ROLE_KEY = "<Supabase の Settings → API → service_role>"
    python tools/create_accounts.py

    既にあるアカウントのパスワードを作り直す場合のみ:
    python tools/create_accounts.py --force

service_role キーは全権限を持つ。**画面共有・スクショ・コミットに出さないこと。**
環境変数で渡すだけにして、このファイルにも .env にも書かない。
"""
import os, sys, json, csv, random, urllib.request, urllib.error, urllib.parse

SUPABASE_URL = "https://qhfegbptgkgccwdnrnas.supabase.co"   # index.html と同じ
APP_URL      = "https://matsumototaichi-beep.github.io/bus-ui-study/"
ID_DOMAIN    = "id.local"       # index.html の normLogin() と合わせる
N_ACCOUNTS   = 24               # 3名 × 8組。実際に使うのは組1〜4、残りは予備
ROOT         = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 紛らわしい字を抜いた英数字。カードを見ながらスマホで打つので i/l/1、o/0 は入れない
ALPHABET = "abcdefghjkmnpqrstuvwxyz23456789"


def gen_password(n=8):
    r = random.SystemRandom()
    return "".join(r.choice(ALPHABET) for _ in range(n))


def api(path, method="GET", body=None):
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    if not key:
        sys.exit("SUPABASE_SERVICE_ROLE_KEY が設定されていません。上のコメントの使い方を参照。")
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(SUPABASE_URL + path, data=data, method=method)
    req.add_header("apikey", key)
    req.add_header("Authorization", "Bearer " + key)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, json.loads(r.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except ValueError:
            return e.code, {"message": raw}


def find_user(email):
    """Admin API はメール完全一致の検索ができるので、それで既存を拾う。"""
    q = urllib.parse.quote(email)
    st, body = api("/auth/v1/admin/users?per_page=200&page=1&filter=" + q)
    if st != 200:
        return None
    for u in (body.get("users") or []):
        if (u.get("email") or "").lower() == email.lower():
            return u
    return None


def main():
    force = "--force" in sys.argv
    dry   = "--dry-run" in sys.argv      # Supabase に触らずカードの刷り上がりだけ確認する
    rows = []
    for i in range(1, N_ACCOUNTS + 1):
        pid    = "p%02d" % i                 # そのまま参加者番号になる
        email  = pid + "@" + ID_DOMAIN
        passwd = gen_password()
        group  = 1 if ((i - 1) // 3) % 2 == 0 else 2     # 3名ずつ1組、組ごとに群を交互に
        team   = (i - 1) // 3 + 1
        form_v = ((i - 1) % 4) + 1                       # 事後アンケートの版 v1〜v4

        if dry:
            rows.append({"番号": pid, "組": team, "群": group, "アンケート版": "v%d" % form_v,
                         "ID": pid, "パスワード": passwd, "結果": "dry-run（作成していない）"})
            continue

        st, body = api("/auth/v1/admin/users", "POST", {
            "email": email, "password": passwd,
            "email_confirm": True,                       # 確認メールは @id.local 宛なので届かない。ここで済ませる
            "user_metadata": {"display_name": pid},
        })
        note = "新規作成"
        if st not in (200, 201):
            msg = str(body.get("msg") or body.get("message") or body)
            if "already" in msg.lower() or "registered" in msg.lower() or st == 422:
                u = find_user(email)
                if u and force:
                    st2, _ = api("/auth/v1/admin/users/" + u["id"], "PUT", {"password": passwd})
                    note = "既存→パスワード再設定" if st2 == 200 else "既存・再設定に失敗(%d)" % st2
                elif u:
                    note = "既存（--force なしなので触らない。パスワードは不明）"
                    passwd = ""
                else:
                    note = "失敗: " + msg
                    passwd = ""
            else:
                note = "失敗(%d): %s" % (st, msg)
                passwd = ""
        rows.append({"番号": pid, "組": team, "群": group, "アンケート版": "v%d" % form_v,
                     "ID": pid, "パスワード": passwd, "結果": note})
        print("%s  群%d  組%d  %s" % (pid, group, team, note))

    csv_path = os.path.join(ROOT, "accounts.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    html_path = os.path.join(ROOT, "accounts_cards.html")
    with open(html_path, "w", encoding="utf-8", newline="") as f:
        f.write(build_cards(rows))

    ok = sum(1 for r in rows if r["パスワード"])
    print("\n%d/%d 件そろいました。" % (ok, len(rows)))
    print("  対応表   : " + csv_path)
    print("  配布カード: " + html_path + "（ブラウザで開いてA4に印刷）")
    print("\n★ この2つは .gitignore 済みです。public リポジトリなので commit しないこと。")


def build_cards(rows):
    cards = []
    for r in rows:
        cards.append(u"""
      <div class="card">
        <div class="hd"><span class="no">%(番号)s</span><span class="meta">組%(組)s ／ 群%(群)s ／ アンケート %(アンケート版)s</span></div>
        <div class="body">
          <table class="cred">
            <tr><th>ID</th><td class="mono">%(ID)s</td></tr>
            <tr><th>パスワード</th><td class="mono">%(パスワード)s</td></tr>
          </table>
          <div class="qr" data-url="%(url)s"></div>
        </div>
        <p class="note">このIDとパスワードは<b>別の日にもう一度使います。</b>カードを無くさないでください。<br>
        メールアドレスを使わない仕組みなので、<b>忘れると元に戻せません。</b></p>
      </div>""" % dict(r, url=APP_URL))

    return u"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>配布カード（参加者アカウント）</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<style>
  @page { size: A4 portrait; margin: 10mm; }
  * { box-sizing: border-box; }
  body { font-family: "Yu Gothic", "游ゴシック", "Meiryo", sans-serif; margin: 0; color: #111; }
  .sheet { display: grid; grid-template-columns: 1fr 1fr; gap: 4mm; }
  .card { border: 0.4mm solid #666; border-radius: 2mm; padding: 3mm 3.5mm;
          height: 60mm; overflow: hidden; page-break-inside: avoid; break-inside: avoid;
          display: flex; flex-direction: column; }
  .hd { display: flex; justify-content: space-between; align-items: baseline;
        border-bottom: 0.3mm solid #999; padding-bottom: 1mm; margin-bottom: 2mm; }
  .no { font-size: 15pt; font-weight: 700; letter-spacing: .05em; }
  .meta { font-size: 7.5pt; color: #555; }
  /* QRは flex の中では float が効かないので、左右に並べる箱を作る */
  .body { display: flex; align-items: flex-start; gap: 3mm; }
  .qr { flex: 0 0 22mm; width: 22mm; height: 22mm; }
  .qr img, .qr canvas { width: 22mm !important; height: 22mm !important; display: block; }
  .cred { flex: 1 1 auto; border-collapse: collapse; }
  .cred th { text-align: left; font-size: 8pt; color: #444; font-weight: 600;
             padding: 0 2mm 1mm 0; white-space: nowrap; vertical-align: bottom; }
  .cred td { padding: 0 0 1mm; border-bottom: 0.2mm dotted #aaa; }
  .mono { font-family: Consolas, "Courier New", monospace; font-size: 12.5pt; letter-spacing: .06em; }
  .note { font-size: 7pt; line-height: 1.45; color: #444; margin: auto 0 0; }
  @media screen { body { background: #eee; padding: 10mm; }
                  .sheet { background: #fff; padding: 10mm; width: 210mm; margin: 0 auto; } }
</style></head><body>
<div class="sheet">__CARDS__</div>
<script>
  document.querySelectorAll('.qr').forEach(function(el){
    new QRCode(el, { text: el.dataset.url, width: 160, height: 160,
                     correctLevel: QRCode.CorrectLevel.M });
  });
</script>
</body></html>""".replace("__CARDS__", "".join(cards))


if __name__ == "__main__":
    main()
