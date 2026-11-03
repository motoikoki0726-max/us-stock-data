"""
build_data.py
SEC EDGAR の一括データ（companyfacts.zip / submissions.zip / company_tickers_exchange.json）から、
米国上場企業の財務データを7軸スコアシート用に整形し、ティッカーの頭文字ごとの JSON にまとめる。

GitHub Actions から毎日実行する。Apps Script（Googleのサーバー）からSECへは直接アクセスできないため、
ここで整形したデータを GitHub の data ブランチに置き、Apps Script はそれを読むだけにする。

使い方:
  python build_data.py --out out
  （テスト用）python build_data.py --out out --tickers t.json --companyfacts cf.zip --submissions sub.zip

必要な環境変数（GitHubでは Settings → Secrets and variables → Actions の「Secret」に登録する）:
  SEC_USER_AGENT  例: "Taro Yamada taro@example.com"（SECのルールで、名前と連絡先メールが必要）
"""

import argparse
import datetime as dt
import json
import os
import sys
import time
import urllib.error
import urllib.request
import zipfile

TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
COMPANYFACTS_URL = "https://www.sec.gov/Archives/edgar/daily-index/xbrl/companyfacts.zip"
SUBMISSIONS_URL = "https://www.sec.gov/Archives/edgar/daily-index/bulkdata/submissions.zip"

HIST_YEARS = 6          # 5期分のCAGR等に加え、ROEの期首自己資本用に1期多く持つ
ANNUAL_DAYS = (350, 380)  # 通期とみなす期間の長さ（52週・53週決算を含む）
EARNINGS_8K_DAYS = 450  # 決算発表（8-K Item 2.02）を何日前まで残すか

# ---- 項目ごとの XBRL 要素名（上から順に、その期に値がある最初のものを使う） ----
REVENUE = [
    "Revenues",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "RevenueFromContractWithCustomerIncludingAssessedTax",
    "SalesRevenueNet",
    "SalesRevenueGoodsNet",
    "SalesRevenueServicesNet",
    "RevenuesNetOfInterestExpense",
    "RegulatedAndUnregulatedOperatingRevenue",
    "InterestAndDividendIncomeOperating",
]
OP_INCOME = ["OperatingIncomeLoss"]
NET_INCOME = ["NetIncomeLoss", "NetIncomeLossAvailableToCommonStockholdersBasic", "ProfitLoss"]
EPS = ["EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted", "EarningsPerShareBasic"]
DPS = ["CommonStockDividendsPerShareDeclared", "CommonStockDividendsPerShareCashPaid"]
SHARES_DILUTED = [
    "WeightedAverageNumberOfDilutedSharesOutstanding",
    "WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
    "WeightedAverageNumberOfSharesOutstandingBasic",
]
SHARES_BASIC = [
    "WeightedAverageNumberOfSharesOutstandingBasic",
    "WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
    "WeightedAverageNumberOfDilutedSharesOutstanding",
]
PRETAX = [
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
]
TAX = ["IncomeTaxExpenseBenefit"]
CFO = ["NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"]
CFI = ["NetCashProvidedByUsedInInvestingActivities", "NetCashProvidedByUsedInInvestingActivitiesContinuingOperations"]
CAPEX = [
    "PaymentsToAcquirePropertyPlantAndEquipment",
    "PaymentsToAcquireProductiveAssets",
    "PaymentsForCapitalImprovements",
    "PaymentsToAcquireOtherPropertyPlantAndEquipment",
]
BUYBACK = ["PaymentsForRepurchaseOfCommonStock", "PaymentsForRepurchaseOfEquity"]
DIV_PAID = ["PaymentsOfDividendsCommonStock", "PaymentsOfDividends", "DividendsCommonStockCash", "DividendsCommonStock"]

SH_EQUITY = ["StockholdersEquity"]
EQUITY_INCL_NCI = ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]
NCI = ["MinorityInterest"]
ASSETS = ["Assets"]
CASH = [
    "CashAndCashEquivalentsAtCarryingValue",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    "Cash",
    "CashAndDueFromBanks",
]
SHARES_OUT = ["CommonStockSharesOutstanding"]
DEBT_NONCURRENT = ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations"]
DEBT_CURRENT_TOTAL = ["DebtCurrent"]
DEBT_CURRENT_PARTS = ["LongTermDebtCurrent", "ShortTermBorrowings", "CommercialPaper"]
DEBT_CURRENT_ALT = ["LongTermDebtAndCapitalLeaseObligationsCurrent"]
LONG_TERM_DEBT_TOTAL = ["LongTermDebt"]
LEASE_TOTAL = ["OperatingLeaseLiability"]
LEASE_PARTS = ["OperatingLeaseLiabilityNoncurrent", "OperatingLeaseLiabilityCurrent"]

ANNUAL_FORMS = ("10-K", "10-K/A", "10-KT", "10-KT/A")
PERIODIC_FORMS = ANNUAL_FORMS + ("10-Q", "10-Q/A")


# ================= ダウンロード =================

def user_agent():
    ua = os.environ.get("SEC_USER_AGENT", "").strip()
    if "@" not in ua:
        sys.exit("エラー：GitHubの Settings → Secrets and variables → Actions に、SEC_USER_AGENT（「名前 メールアドレス」）を登録してください（SECのルール）。")
    return ua


def download(url, path):
    """SECの公開ファイルを保存する。混雑時の429/403は1分待って最大3回まで再試行する。"""
    req = urllib.request.Request(url, headers={"User-Agent": user_agent(), "Accept-Encoding": "identity"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=600) as res, open(path, "wb") as f:
                while True:
                    chunk = res.read(1 << 20)
                    if not chunk:
                        break
                    f.write(chunk)
                print(f"取得：{url} → {os.path.getsize(path) / 1e6:.0f} MB", flush=True)
                return res.headers.get("Last-Modified", "")
        except urllib.error.HTTPError as e:
            body = e.read(400).decode("utf-8", "replace")
            print(f"HTTP {e.code}：{url}（{attempt + 1}回目）{body[:200]}", flush=True)
            if e.code in (403, 429) and attempt < 2:
                time.sleep(60)
                continue
            raise
    return ""


# ================= 日付の補助 =================

def d(s):
    return dt.date.fromisoformat(s)


def days_between(start, end):
    return (d(end) - d(start)).days


def near(a, b, tol_days):
    return abs((d(a) - d(b)).days) <= tol_days


# ================= 1社分の XBRL を読む =================

class Facts:
    """companyfacts の us-gaap / dei を、期間の種類ごとに引きやすい形にする"""

    def __init__(self, cf):
        facts = cf.get("facts", {})
        self.gaap = facts.get("us-gaap", {})
        self.dei = facts.get("dei", {})
        self._cache = {}

    def _units(self, concept, unit, ns="gaap"):
        src = self.gaap if ns == "gaap" else self.dei
        c = src.get(concept)
        if not c:
            return []
        return c.get("units", {}).get(unit, [])

    def annual(self, concept, unit="USD"):
        """通期（約1年）の値 {期末日: 値}。同じ期が複数の書類にあれば最も新しく提出された値（修正後）"""
        key = ("A", concept, unit)
        if key not in self._cache:
            out = {}
            for x in self._units(concept, unit):
                if "start" not in x or x.get("form") not in ANNUAL_FORMS:
                    continue
                n = days_between(x["start"], x["end"])
                if not (ANNUAL_DAYS[0] <= n <= ANNUAL_DAYS[1]):
                    continue
                prev = out.get(x["end"])
                if prev is None or x["filed"] >= prev[1]:
                    out[x["end"]] = (x["val"], x["filed"])
            self._cache[key] = out
        return self._cache[key]

    def instant(self, concept, unit="USD"):
        """時点の値 {日付: 値}（貸借対照表の項目）"""
        key = ("I", concept, unit)
        if key not in self._cache:
            out = {}
            for x in self._units(concept, unit):
                if "start" in x or x.get("form") not in PERIODIC_FORMS:
                    continue
                prev = out.get(x["end"])
                if prev is None or x["filed"] >= prev[1]:
                    out[x["end"]] = (x["val"], x["filed"])
            self._cache[key] = out
        return self._cache[key]

    def durations(self, concept, unit="USD"):
        """期間の値 [(開始日, 終了日, 値, 提出日)]（四半期報告の累計期間を探すのに使う）"""
        key = ("D", concept, unit)
        if key not in self._cache:
            best = {}
            for x in self._units(concept, unit):
                if "start" not in x or x.get("form") not in PERIODIC_FORMS:
                    continue
                k = (x["start"], x["end"])
                if k not in best or x["filed"] >= best[k][3]:
                    best[k] = (x["start"], x["end"], x["val"], x["filed"])
            self._cache[key] = list(best.values())
        return self._cache[key]

    # ---- 候補リストから1つ選ぶ ----

    def pick_annual(self, cands, end, unit="USD"):
        for c in cands:
            v = self.annual(c, unit).get(end)
            if v is not None:
                return v[0], c
        return None, None

    def pick_instant(self, cands, date, unit="USD"):
        for c in cands:
            v = self.instant(c, unit).get(date)
            if v is not None:
                return v[0], c
        return None, None

    def dei_latest(self, concept, unit):
        arr = [x for x in self._units(concept, unit, ns="dei") if x.get("form") in PERIODIC_FORMS]
        if not arr:
            return None, None
        last = max(arr, key=lambda x: (x["end"], x["filed"]))
        return last["val"], last["end"]


# ================= 項目の組み立て =================

def fiscal_year_ends(f):
    """
    通期決算の期末日の一覧（古い順）。売上高・営業利益・純利益のいずれかがある期。
    決算期の変更などで300日未満の間隔の期末日が混ざった場合は、新しい方を残す
    """
    ends = set()
    for c in REVENUE + OP_INCOME + NET_INCOME:
        ends.update(f.annual(c).keys())
    kept = []
    for e in sorted(ends, reverse=True):
        if not kept or days_between(e, kept[-1]) >= 300:
            kept.append(e)
    return sorted(kept)


def debt_at(f, date):
    """有利子負債の合計（推定）。返り値 (金額, 使った要素の説明)"""
    used = []
    noncur, c = f.pick_instant(DEBT_NONCURRENT, date)
    if c:
        used.append(c)
    ltd_total, c_total = f.pick_instant(LONG_TERM_DEBT_TOTAL, date)
    ltd_cur, _ = f.pick_instant(["LongTermDebtCurrent"], date)

    cur, c = f.pick_instant(DEBT_CURRENT_TOTAL, date)
    if c:
        used.append(c)
    else:
        parts = []
        for name in DEBT_CURRENT_PARTS:
            v, cc = f.pick_instant([name], date)
            if cc:
                parts.append(v)
                used.append(cc)
        if parts:
            cur = sum(parts)
        else:
            cur, c = f.pick_instant(DEBT_CURRENT_ALT, date)
            if c:
                used.append(c)

    if noncur is None and ltd_total is not None:
        # LongTermDebt は本来「1年内返済分を含む長期借入金の総額」。1年内返済分が別にあれば差し引く
        noncur = ltd_total - (ltd_cur or 0)
        used.append(c_total + ("-LongTermDebtCurrent" if ltd_cur else ""))

    if noncur is None and cur is None:
        return None, ""
    return (noncur or 0) + (cur or 0), "+".join(used)


def lease_at(f, date):
    v, c = f.pick_instant(LEASE_TOTAL, date)
    if c:
        return v
    parts = [f.pick_instant([n], date)[0] for n in LEASE_PARTS]
    parts = [p for p in parts if p is not None]
    return sum(parts) if parts else None


def equity_incl_nci_at(f, date, sh_eq):
    v, c = f.pick_instant(EQUITY_INCL_NCI, date)
    if c:
        return v
    nci, c = f.pick_instant(NCI, date)
    if sh_eq is None:
        return None
    return sh_eq + (nci or 0)


def balance_at(f, date):
    sh_eq, _ = f.pick_instant(SH_EQUITY, date)
    debt, debt_src = debt_at(f, date)
    return {
        "shEq": sh_eq,
        "eq": equity_incl_nci_at(f, date, sh_eq),
        "ta": f.pick_instant(ASSETS, date)[0],
        "cash": f.pick_instant(CASH, date)[0],
        "debt": debt,
        "debtSrc": debt_src,
        "lease": lease_at(f, date),
    }


def dps_for_year(f, end, div_paid):
    v, c = f.pick_annual(DPS, end, "USD/shares")
    if c:
        return v, "1株配当の開示値"
    if div_paid:
        sh, c2 = f.pick_annual(SHARES_BASIC, end, "shares")
        if sh:
            return div_paid / sh, "配当支払総額÷平均株式数"
    return 0, "配当の記載なし→0"


def annual_row(f, end):
    sales, _ = f.pick_annual(REVENUE, end)
    div_paid, _ = f.pick_annual(DIV_PAID, end)
    dps, dps_src = dps_for_year(f, end, div_paid)
    buyback, c_bb = f.pick_annual(BUYBACK, end)
    row = {
        "end": end,
        "sales": sales,
        "op": f.pick_annual(OP_INCOME, end)[0],
        "np": f.pick_annual(NET_INCOME, end)[0],
        "eps": f.pick_annual(EPS, end, "USD/shares")[0],
        "dps": dps,
        "dpsSrc": dps_src,
        "shDil": f.pick_annual(SHARES_DILUTED, end, "shares")[0],
        "cfo": f.pick_annual(CFO, end)[0],
        "cfi": f.pick_annual(CFI, end)[0],
        "capex": f.pick_annual(CAPEX, end)[0],
        "buyback": buyback if c_bb else 0,
        "divPaid": div_paid if div_paid is not None else 0,
        "pretax": f.pick_annual(PRETAX, end)[0],
        "tax": f.pick_annual(TAX, end)[0],
    }
    row.update(balance_at(f, end))
    return row


def latest_balance_date(f):
    dates = set()
    for c in SH_EQUITY + ASSETS:
        dates.update(f.instant(c).keys())
    return max(dates) if dates else None


def find_duration(f, cands, start, end, unit="USD", tol=7):
    """開始日・終了日が近い期間の値を候補リストの順に探す"""
    for c in cands:
        for s, e, v, _ in f.durations(c, unit):
            if near(s, start, tol) and near(e, end, tol):
                return v
    return None


def ytd_block(f, fy_ends):
    """
    直近の四半期報告（10-Q）の累計期間と、その前年同期間。
    例：12月決算で直近が第3四半期なら、今期1〜9月と前期1〜9月。
    直近の書類が年次報告（10-K）なら None
    """
    if len(fy_ends) < 2:
        return None
    last_fy, prev_fy = fy_ends[-1], fy_ends[-2]
    q_ends = set()
    for c in REVENUE + NET_INCOME:
        for s, e, v, _ in f.durations(c):
            if e > last_fy:
                q_ends.add(e)
    if not q_ends:
        return None
    end = max(q_ends)
    start = (d(last_fy) + dt.timedelta(days=1)).isoformat()
    prev_start = (d(prev_fy) + dt.timedelta(days=1)).isoformat()
    prev_end = (d(end) - dt.timedelta(days=364)).isoformat()
    months = round(days_between(start, end) / 30.44)

    cur, prev = {}, {}
    for key, cands, unit in [("sales", REVENUE, "USD"), ("op", OP_INCOME, "USD"), ("np", NET_INCOME, "USD"),
                             ("eps", EPS, "USD/shares"), ("dps", DPS, "USD/shares")]:
        # 今期と前年同期で同じ要素を使う（途中で要素名が変わった場合のずれを防ぐ）
        for c in cands:
            a = find_duration(f, [c], start, end, unit)
            b = find_duration(f, [c], prev_start, prev_end, unit, tol=10)
            if a is not None and b is not None:
                cur[key], prev[key] = a, b
                break
    if not cur:
        return None
    return {"start": start, "end": end, "months": months, "cur": cur, "prev": prev}


def ttm_block(fy_last, ytd):
    """直近12か月 ＝ 直近通期 ＋ 今期累計 − 前年同期累計"""
    if not ytd:
        return {k: fy_last.get(k) for k in ("sales", "op", "np", "eps", "dps")} | {"basis": "直近通期"}
    out = {"basis": "直近通期＋今期累計−前年同期累計（" + str(ytd["months"]) + "か月）"}
    for k in ("sales", "op", "np", "eps", "dps"):
        a, b, base = ytd["cur"].get(k), ytd["prev"].get(k), fy_last.get(k)
        out[k] = base + a - b if None not in (a, b, base) else base
    return out


def earnings_dates(sub):
    """決算発表（8-K の Item 2.02「業績の発表」）の提出日"""
    rec = (sub or {}).get("filings", {}).get("recent", {})
    forms, dates, items = rec.get("form", []), rec.get("filingDate", []), rec.get("items", [])
    cutoff = (dt.date.today() - dt.timedelta(days=EARNINGS_8K_DAYS)).isoformat()
    out = set()
    for i, form in enumerate(forms):
        if form not in ("8-K", "8-K/A") or i >= len(dates) or i >= len(items):
            continue
        if "2.02" in str(items[i]) and dates[i] >= cutoff:
            out.add(dates[i])
    return sorted(out)


def build_company(cf, sub):
    f = Facts(cf)
    if not f.gaap:
        return {"error": "米国会計基準（us-gaap）の財務データがない（外国企業の20-F・40-F提出企業など）"}
    fy_ends = fiscal_year_ends(f)
    if not fy_ends:
        return {"error": "年次報告（10-K）の売上高・利益が見つからない"}

    years = [annual_row(f, e) for e in fy_ends[-HIST_YEARS:]]
    last = years[-1]
    bs_date = latest_balance_date(f)
    bs = balance_at(f, bs_date) if bs_date else {}
    # 四半期報告に載らない項目（リース負債など）は、直近通期の値で補う
    filled = []
    for k in ("shEq", "eq", "ta", "cash", "debt", "lease"):
        if bs.get(k) is None and last.get(k) is not None:
            bs[k] = last[k]
            filled.append(k)
    if "debt" in filled:
        bs["debtSrc"] = last.get("debtSrc", "")
    sh_out, sh_out_date = f.dei_latest("EntityCommonStockSharesOutstanding", "shares")
    if sh_out is None and bs_date:
        sh_out = f.pick_instant(SHARES_OUT, bs_date, "shares")[0]
        sh_out_date = bs_date
    bs.update({"end": bs_date, "shOut": sh_out, "shOutDate": sh_out_date, "filledFromFy": filled})

    # 出所の説明は直近期だけ残す（ファイルを小さくするため）
    last_dps_src = last.get("dpsSrc", "")
    for y in years:
        y.pop("dpsSrc", None)
        y.pop("debtSrc", None)

    ytd = ytd_block(f, fy_ends)
    out = {
        "dpsSrc": last_dps_src,
        "fy": years,
        "bs": bs,
        "ytd": ytd,
        "ttm": ttm_block(years[-1], ytd),
        "er": earnings_dates(sub),
    }
    latest_filed = max((x[1] for c in REVENUE + NET_INCOME for x in f.annual(c).values()), default="")
    out["lastAnnualFiled"] = latest_filed
    return out


# ================= 全社の処理 =================

def load_tickers(path):
    j = json.load(open(path, encoding="utf-8"))
    fields = j["fields"]
    rows = [dict(zip(fields, r)) for r in j["data"]]
    by_cik = {}
    for r in rows:
        if not r.get("ticker"):
            continue
        by_cik.setdefault(int(r["cik"]), []).append(r)
    return by_cik


def read_member(zf, cik):
    name = f"CIK{cik:010d}.json"
    try:
        with zf.open(name) as fp:
            return json.load(fp)
    except KeyError:
        return None


def round_floats(o):
    """JSONを小さくするため、浮動小数点は有効数字を落とさない範囲で丸める"""
    if isinstance(o, float):
        return float(f"{o:.8g}")
    if isinstance(o, dict):
        return {k: round_floats(v) for k, v in o.items()}
    if isinstance(o, list):
        return [round_floats(v) for v in o]
    return o


def shard_key(ticker):
    c = ticker[:1].upper()
    return c if "A" <= c <= "Z" else "_"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="out")
    ap.add_argument("--tickers")
    ap.add_argument("--companyfacts")
    ap.add_argument("--submissions")
    ap.add_argument("--work", default="work")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    os.makedirs(args.work, exist_ok=True)
    source = {}

    tickers_path = args.tickers or os.path.join(args.work, "tickers.json")
    cf_path = args.companyfacts or os.path.join(args.work, "companyfacts.zip")
    sub_path = args.submissions or os.path.join(args.work, "submissions.zip")
    if not args.tickers:
        source["tickers"] = download(TICKERS_URL, tickers_path)
    if not args.submissions:
        source["submissions"] = download(SUBMISSIONS_URL, sub_path)
    if not args.companyfacts:
        source["companyfacts"] = download(COMPANYFACTS_URL, cf_path)

    by_cik = load_tickers(tickers_path)
    print(f"ティッカーのある企業：{len(by_cik)}社", flush=True)

    shards = {}
    ok = err = 0
    started = time.time()
    with zipfile.ZipFile(cf_path) as cfz, zipfile.ZipFile(sub_path) as subz:
        for i, (cik, rows) in enumerate(sorted(by_cik.items())):
            sub = read_member(subz, cik)
            meta = {
                "cik": cik,
                "name": rows[0].get("name") or (sub or {}).get("name", ""),
                "sic": (sub or {}).get("sic", ""),
                "sicDesc": (sub or {}).get("sicDescription", ""),
                "fye": (sub or {}).get("fiscalYearEnd", ""),
            }
            try:
                cf = read_member(cfz, cik)
                body = build_company(cf, sub) if cf else {"error": "SECの財務データ（companyfacts）がない"}
            except Exception as e:  # 1社の失敗で全体を止めない
                body = {"error": f"整形中のエラー：{type(e).__name__}: {e}"}
            if "error" in body:
                err += 1
                body["er"] = earnings_dates(sub)
            else:
                ok += 1
            for r in rows:
                rec = {"t": r["ticker"], "ex": r.get("exchange") or ""}
                rec.update(meta)
                rec.update(body)
                shards.setdefault(shard_key(r["ticker"]), {})[r["ticker"].upper()] = round_floats(rec)
            if (i + 1) % 1000 == 0:
                print(f"  {i + 1}社 処理（{time.time() - started:.0f}秒）", flush=True)

    total = 0
    for k, recs in sorted(shards.items()):
        with open(os.path.join(args.out, k + ".json"), "w", encoding="utf-8") as fp:
            json.dump(recs, fp, ensure_ascii=False, separators=(",", ":"))
        total += len(recs)
    meta = {
        "built": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "companiesOk": ok,
        "companiesError": err,
        "tickers": total,
        "shards": sorted(shards.keys()),
        "source": source,
    }
    with open(os.path.join(args.out, "_meta.json"), "w", encoding="utf-8") as fp:
        json.dump(meta, fp, ensure_ascii=False, indent=1)
    print(f"完了：{ok}社 OK / {err}社 データなし・エラー / {total}ティッカー（{time.time() - started:.0f}秒）", flush=True)


if __name__ == "__main__":
    main()
