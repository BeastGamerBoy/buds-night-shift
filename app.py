"""BUDS Round 2.5 - Night Shift | Bengaluru Analytics & Operations Console (v2)
Run: pip install -r requirements.txt  then  streamlit run app.py"""
import re, warnings, numpy as np, pandas as pd, streamlit as st
import plotly.graph_objects as go
warnings.filterwarnings("ignore")

# ============ THEME ============
BG, CARD, BORDER, CYAN, OK, WARN, BAD = "#0B0F17", "#151D2A", "#1E293B", "#06B6D4", "#10B981", "#F59E0B", "#F43F5E"
st.set_page_config("BUDS Night Shift", layout="wide")
st.markdown(f"""<style>.stApp{{background:{BG};color:#E2E8F0}} section[data-testid=stSidebar]{{background:{CARD}}}
.card{{background:{CARD};border:1px solid {BORDER};border-radius:8px;padding:12px 14px}}
.card h4{{margin:0;font-size:11px;color:#94A3B8;text-transform:uppercase}} .card p{{margin:4px 0 0;font-size:24px;font-weight:600}}
.stTabs [aria-selected=true]{{color:{CYAN}!important}} h1,h2,h3{{color:#F1F5F9}}</style>""", unsafe_allow_html=True)
card = lambda t, v, c="#F1F5F9", f="": f'<div class="card"><h4>{t} {f}</h4><p style="color:{c}">{v}</p></div>'
inr = lambda v: f"₹{v:,.0f}"
def style(fig, h=320):
    fig.update_layout(template="plotly_dark", paper_bgcolor=CARD, plot_bgcolor=CARD, height=h, margin=dict(l=10, r=10, t=40, b=10), colorway=[CYAN, OK, WARN, BAD])
    return fig

# ============ DEMO DATA ============
def demo_data():
    r = np.random.default_rng(7); hubs = ["BLR_HUB_01", "BLR-02", "blr hub 3", "BLR_HUB_04", "BLR-01", "Hub 3"]; n = 300
    od = pd.Timestamp("2026-09-01") + pd.to_timedelta(r.integers(0, 30, n), "D"); prom = od + pd.to_timedelta(2, "D")
    ret = pd.DataFrame({"Order_ID": range(n), "Hub_ID": r.choice(hubs, n), "Order_Date": od, "Promised_Date": prom,
        "Delivery_Date": prom + pd.to_timedelta(r.integers(-1, 3, n), "D"), "Status": r.choice(["Delivered", "delivered", "In Transit", "Returned"], n, p=[.6, .2, .1, .1]),
        "Returned": r.choice([0, 1], n, p=[.93, .07]), "Revenue": r.integers(400, 4000, n)})
    ret = pd.concat([ret, ret.iloc[:6]]); ret["Order_Date"] = ret["Order_Date"].astype(object); ret.loc[ret.index[5], "Order_Date"] = "31-31-2026"; ret.loc[ret.index[9], "Revenue"] = np.nan
    ret.iloc[3, ret.columns.get_loc("Delivery_Date")] = pd.NaT
    return {"Inventory": pd.DataFrame({"SKU": [f"S{i}" for i in range(40)], "Hub_ID": r.choice(hubs, 40), "Qty": r.integers(50, 900, 40), "Unit_Cost": r.integers(80, 1500, 40)}),
            "Fleet": pd.DataFrame({"Vehicle_ID": [f"V{i}" for i in range(25)], "Hub_ID": r.choice(hubs, 25), "Capacity_Hrs": 10, "Used_Hrs": r.integers(5, 10, 25), "Km": r.integers(80, 220, 25) * 26, "Cost_Per_Km": r.integers(14, 26, 25)}),
            "Returns": ret, "Suppliers": pd.DataFrame({"Supplier_ID": [f"SUP{i}" for i in range(8)], "OnTime_Pct": r.integers(70, 98, 8), "Defect_Pct": r.integers(1, 9, 8)})}

SKIP = r"dashboard|summary|finding|assumption|analysis|action_plan|kpis|data.?quality|clean.?data|readme"
def tidy(d0):
    """Real workbooks have title rows above the table: find the true header row, drop empty rows/columns."""
    t = d0.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if t.empty: return t
    nn = t.head(10).notna().sum(axis=1); h = next((i for i, v in nn.items() if v >= max(3, .6 * nn.max())), 0)
    t.columns = [str(x).strip() if pd.notna(x) else f"col_{i}" for i, x in enumerate(t.iloc[h])]
    return t.iloc[h + 1:].reset_index(drop=True).infer_objects()

def load_files(files):
    out = {}
    for f in files:
        if f.name.lower().endswith(("xlsx", "xls")):
            for s_, d in pd.read_excel(f, sheet_name=None, header=None).items():
                if re.search(SKIP, s_, re.I): continue
                t = tidy(d)
                if len(t) < 2 or any(t.shape == o.shape and t.iloc[:, 0].astype(str).equals(o.iloc[:, 0].astype(str)) for o in out.values()): continue
                out[s_] = t
        else:
            try: out[f.name.rsplit(".", 1)[0]] = pd.read_csv(f)
            except UnicodeDecodeError: f.seek(0); out[f.name.rsplit(".", 1)[0]] = pd.read_csv(f, encoding="latin-1")
    return out

# ============ CLEANING (raw never changed) ============
HUBRX = "hub|warehouse|depot"
DATERX = r"(^|[_\s])date|date(time)?([_\s]|$)|_time$"
NOIMP = r"minute|cycle|promis|rating|distance|physical"          # never invent values for these: leave blank, exclude from KPIs
_nrm = lambda x: re.sub(r"[\s_\-]+", "-", str(x).strip().upper())
_num = lambda x: (lambda m: int(m.group(1)) if m else None)(re.search(r"(\d+)", str(x)))
def std_hub(v): return v
def hub_canon(s):
    """Map every spelling of a hub to one canonical form (h-3 / Hub 3 / BLR_HUB_03 -> same; HUB-E -> HUB-EAST)."""
    vc = s.dropna().astype(str).str.strip().value_counts(); canon = {}
    nums = [(x, n) for x, n in vc.items() if _num(x) is not None]; names = [(x, n) for x, n in vc.items() if _num(x) is None]
    if nums:
        pc = {}
        for x, n in nums: pc.setdefault(re.sub(r"\d.*$", "", _nrm(x)), []).append((n, x))
        raw = max(max(pc.values(), key=lambda l: sum(n for n, _ in l)))[1]; pref, w = re.sub(r"\d.*$", "", raw), len(re.search(r"\d+", raw).group())
        for x, _ in nums: canon[x] = f"{pref}{_num(x):0{w}d}"
    groups = {}
    for x, n in names: groups.setdefault(_nrm(x), []).append((n, x))
    full = [max(g)[1] for g in groups.values()]
    for k, g in groups.items():
        ext = [y for y in full if _nrm(y).startswith(k) and _nrm(y) != k]; best = ext[0] if len(ext) == 1 else max(g)[1]
        for _, x in g: canon[x] = best
    return canon

def clean(raw):
    out, log = {}, []
    L = lambda s, step, det, n: log.append({"Sheet": s, "Transformation": step, "Detail": det, "Rows affected": int(n)})
    for name, d0 in raw.items():
        df = d0.copy(deep=True); df.columns = [str(c).strip().replace(" ", "_") for c in df.columns]
        n = df.duplicated().sum()
        if n: df = df.drop_duplicates(); L(name, "Duplicate removal", "Exact duplicate rows dropped (cleaned copy only)", n)
        idc = df.columns[0]
        if re.search(r"(id|key)$", idc, re.I) and df[idc].duplicated().any():
            n = df[idc].duplicated().sum(); df = df.drop_duplicates(subset=[idc]); L(name, "Duplicate ID removal", f"{idc}: repeated IDs (e.g. manual adjustments) - first record kept", n)
        for c in df.columns:
            if not pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_datetime64_any_dtype(df[c]) and not re.search(DATERX + r"|(id|key|code|sku)$", c, re.I):
                t = pd.to_numeric(df[c].astype(str).str.replace(r"[₹,%\s]|INR|Rs\.?", "", regex=True), errors="coerce"); nz = df[c].notna().sum()
                if nz and t.notna().sum() >= .8 * nz: df[c] = t; L(name, "Text -> number", f"{c}: symbols removed / non-numeric entries blanked", t.notna().sum())
        hubc = [c for c in df.columns if re.search(HUBRX, c, re.I) and not re.search(r"name|region|zone|capac|cost|downtime|util|process|avg|daily", c, re.I) and not pd.api.types.is_numeric_dtype(df[c])]
        for c in hubc:
            mp = hub_canon(df[c]); old = df[c].copy(); df[c] = df[c].map(lambda v: mp.get(str(v).strip(), v) if pd.notna(v) else v); ch = ((df[c] != old) & old.notna()).sum()
            L(name, "Hub ID standardisation", f"{c}: spelling variants unified (e.g. {next((f'{k} -> {v}' for k, v in mp.items() if k != v), 'already consistent')})", ch)
            miss = df[c].isna()
            for o in [x for x in df.columns if miss.any() and x != c and re.search("hub", x, re.I) and not pd.api.types.is_numeric_dtype(df[x])]:
                ref = df.dropna(subset=[c, o]).groupby(o)[c].agg(lambda z: z.mode().iat[0])
                if len(ref) and ref.nunique() == len(ref):
                    f_ = miss & df[o].map(ref).notna(); df.loc[f_, c] = df.loc[f_, o].map(ref)
                    if f_.sum(): L(name, "Missing hub restored", f"{c} filled from {o} (one-to-one mapping)", f_.sum())
                    miss = df[c].isna()
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]) and re.search(r"minute|cycle|promis|distance|_km$|rating", c, re.I):
                bad = (df[c] <= 0) | (df[c] > (5 if "rating" in c.lower() else np.inf))
                if bad.sum(): df.loc[bad, c] = np.nan; L(name, "Invalid value -> blank", f"{c}: zero/negative/out-of-range values blanked (excluded from KPIs)", bad.sum())
        for c in [c for c in df.columns if re.search(DATERX, c, re.I)]:
            if pd.api.types.is_datetime64_any_dtype(df[c]): continue
            p = pd.to_datetime(df[c], errors="coerce").fillna(pd.to_datetime(df[c], errors="coerce", dayfirst=True)); bad = (p.isna() & df[c].notna()).sum(); df[c] = p
            if bad: L(name, "Invalid date -> blank", f"{c}: unreadable dates nulled", bad)
        for c in df.columns:
            m = df[c].isna().sum()
            if not m or pd.api.types.is_datetime64_any_dtype(df[c]): continue
            if m == len(df): L(name, "Empty column ignored", f"{c}: no values (formula placeholder)", 0); continue
            if re.search(NOIMP, c, re.I): L(name, "Left blank (not imputed)", f"{c}: excluded from KPI denominators instead of guessing", m); continue
            if pd.api.types.is_numeric_dtype(df[c]): df[c] = df[c].fillna(df[c].median()); L(name, "Missing number -> median", c, m)
            else: df[c] = df[c].fillna("UNKNOWN"); L(name, "Missing text -> UNKNOWN", c, m)
        for s_ in [c for c in df.columns if "status" in c.lower() and not pd.api.types.is_numeric_dtype(df[c]) and df[c].notna().any()]:
            k = df[s_].astype(str).str.strip().str.lower(); canon = df.groupby(k)[s_].agg(lambda z: z.value_counts().index[0]); new = k.map(canon).where(df[s_].notna())
            ch = ((new != df[s_]) & df[s_].notna()).sum(); df[s_] = new
            if ch: L(name, "Status normalisation", f"{s_}: case/space variants unified", ch)
        s_ = next((c for c in df.columns if "status" in c.lower()), None); d_ = next((c for c in df.columns if "deliver" in c.lower() and pd.api.types.is_datetime64_any_dtype(df[c])), None)
        if s_ and d_:
            cf = df[s_].astype(str).str.lower().eq("delivered") & df[d_].isna(); df.loc[cf, s_] = "In Transit"
            if cf.sum(): L(name, "Status conflict resolved", "'Delivered' but no delivery date -> 'In Transit'", cf.sum())
        a_, r_ = next((c for c in df.columns if re.search(r"availab.*status", c, re.I)), None), next((c for c in df.columns if re.search(r"runs|trips", c, re.I)), None)
        if a_ and r_:
            cf = df[a_].astype(str).str.lower().isin(["maintenance", "unavailable"]) & (df[r_] > 0)
            if cf.sum(): L(name, "Status conflict flagged", "Vehicle Maintenance/Unavailable but has assigned runs (kept, flagged)", cf.sum())
        k_, q_ = next((c for c in df.columns if re.search(r"stock_status", c, re.I)), None), next((c for c in df.columns if re.search(r"physical|system", c, re.I) and pd.api.types.is_numeric_dtype(df[c])), None)
        if k_ and q_:
            cf = df[k_].astype(str).str.lower().str.contains("out") & (df[q_] > 0)
            if cf.sum(): L(name, "Status conflict flagged", "'Out of Stock' but units > 0 (kept, flagged)", cf.sum())
        out[name] = df
    return out, log

def profile(raw):
    rows = []
    for s, df in raw.items():
        for c in df.columns: rows.append({"Sheet": s, "Column": c, "Type": str(df[c].dtype), "Missing": int(df[c].isna().sum()), "Unique": int(df[c].nunique()), "Sample": str(df[c].dropna().iloc[0]) if df[c].notna().any() else ""})
    return pd.DataFrame(rows)

# ============ AUTO COLUMN MAPPING (works on unfamiliar names) ============
OR, FL, IV, SP = "order|raw_data|r1|sales|shipment|deliver", "fleet|vehicle|truck", "invent|stock|sku", "suppl|vendor"
ROLES = {"status": (OR, "order_status|status", "a"), "promised_min": (OR, "promised.*min|promised_delivery|sla_min|promis", "n"), "actual_min": (OR, "actual.*cycle|actual.*min|cycle_time|delivery_min", "n"),
         "promised": (OR, "promis|due|sla|expect|eta", "d"), "delivered": (OR, "deliver|actual|arriv", "d"), "returned": (OR, "returned|refund_flag|return_flag|is_return", "a"),
         "revenue": (OR, "delivery_fee|revenue|sales|amount|value|price", "n"), "var_cost": (OR, "variable_cost|var_cost", "n"), "fixed": (OR, "fixed_cost", "n"),
         "order_date": (OR, "order_date|date", "d"), "aov": (OR, "order_value|basket", "n"), "ret_id": ("return", "return_id|return_key", "a"),
         "capacity": (FL, "cap", "n"), "used": (FL, "used|utili|hours|active", "n"), "cpk": (FL, "cost", "n"), "km": (FL, "km|dist", "n"),
         "fleet_status": (FL, "availab.*status|status", "a"), "runs": (FL, "assigned_runs|runs|trips", "n"), "route_cost": ("route", "base_cost|cost", "n"),
         "unit_cost": (IV, "unit_cost|cost|price", "n"), "qty": (IV, "system_stock|qty|quant|stock_units|units", "n"), "phys": (IV, "physical", "n"),
         "approval": (SP, "approved_status|approval|status", "a"), "ontime": (SP, "on_?time|punctual|deliver", "n"), "defect": (SP, "defect|reject|quality", "n")}
GROUPS = {"Orders": ["status", "promised_min", "actual_min", "promised", "delivered", "revenue", "var_cost", "fixed", "order_date", "aov"], "Returns": ["returned", "ret_id"],
          "Fleet": ["capacity", "used", "cpk", "km", "fleet_status", "runs"], "Inventory": ["unit_cost", "qty", "phys"], "Suppliers": ["approval", "ontime", "defect"]}
_ok = lambda s, k: (k == "a") or (k == "d" and pd.api.types.is_datetime64_any_dtype(s)) or (k == "n" and pd.api.types.is_numeric_dtype(s))

def detect(c):
    m, used = {}, set()
    for r, (sx, cx, k) in ROLES.items():
        m[r] = None; sh = [s_ for s_ in c if re.search(sx, s_, re.I)] or list(c)
        for s_ in sh:
            hit = None
            for alt in cx.split("|"):
                hit = next((x for x in c[s_].columns if (s_, x) not in used and re.search(alt, x, re.I) and _ok(c[s_][x], k)), None)
                if hit: break
            if hit: m[r] = (s_, hit); used.add((s_, hit)); break
    return m

def get(c, m, r):
    v = m.get(r); return c[v[0]][v[1]] if v and v[0] in c and v[1] in c[v[0]] else None

# ============ KPI ENGINE (named surrogates, never crashes) ============
def hubcol(df):
    return next((h for h in df.columns if re.search(HUBRX, h, re.I) and re.search(r"id|code|key", h, re.I)), None) or next((h for h in df.columns if re.search(HUBRX, h, re.I)), None)

def otd_frame(c, m):
    for a, b in (("promised_min", "actual_min"), ("promised", "delivered")):
        if m.get(a) and m.get(b) and m[a][0] == m[b][0] and m[a][0] in c:
            df = c[m[a][0]]; x = df[[m[a][1], m[b][1]]].copy(); x.columns = ["p", "d"]
            if m.get("status") and m["status"][0] == m[a][0]: x = x[df[m["status"][1]].astype(str).str.lower().eq("delivered")]
            x = x.dropna()
            if a == "promised_min": x = x[(x.p > 0) & (x.d > 0)]
            hc = hubcol(df); return x.assign(ok=(x.d <= x.p), hub=df.loc[x.index, hc] if hc else "ALL")
    return None

def kpis(c, m, drop=()):
    m = {r: (None if any(r in GROUPS[g] for g in drop) else v) for r, v in m.items()}; sur = {}
    miss = [g for g, rs in GROUPS.items() if all(get(c, m, r) is None for r in rs)]
    x = otd_frame(c, m); otd = x.ok.mean() * 100 if x is not None and len(x) else None
    od = next((c[m[r][0]] for r in ("promised_min", "actual_min", "revenue", "var_cost") if m.get(r) and m[r][0] in c), None); n_ord = len(od) if od is not None else None
    rt, rid, ret = get(c, m, "returned"), get(c, m, "ret_id"), None
    if rt is not None: ret = ((rt > 0) if pd.api.types.is_numeric_dtype(rt) else rt.astype(str).str.lower().isin(["1", "yes", "y", "true", "returned"])).mean() * 100
    elif rid is not None and n_ord: ret = rid.nunique() / n_ord * 100
    rv, vc = get(c, m, "revenue"), get(c, m, "var_cost"); rev = rv.sum() if rv is not None else None
    fixed, fm = 0.0, m.get("fixed")
    if fm and fm[0] in c:
        df = c[fm[0]]; hc, dm = hubcol(df), m.get("order_date"); days = df[dm[1]].dt.normalize().nunique() if dm and dm[0] == fm[0] and dm[1] in df else 1
        if hc: df = df[df[hc] != "UNKNOWN"]
        fixed = (df.groupby(hc)[fm[1]].median().sum() if hc else df[fm[1]].median()) * days
    q, uc, ph = get(c, m, "qty"), get(c, m, "unit_cost"), get(c, m, "phys"); same = lambda a, b: m.get(a) and m.get(b) and m[a][0] == m[b][0]
    hold = (q * uc).sum() * .24 / 12 if same("qty", "unit_cost") else None
    inv_acc = ((q == ph)[q.notna() & ph.notna()].mean() * 100) if same("qty", "phys") and (q.notna() & ph.notna()).any() else None
    if otd is None: otd, sur["OTD"] = 85.0, "Benchmark 85% (no promised vs actual time columns found)"
    if ret is None: ret, sur["Return"] = 6.0, "Industry surrogate 6%"
    if hold is None and q is not None and get(c, m, "aov") is not None:
        hold = get(c, m, "aov").sum() * .02; sur["Holding"] = "Inventory feed has no unit cost: estimated as 2% of order value (GMV) per month"
    if rev is None: rev, sur["Revenue"] = (hold * 6 if hold else 5e6), "Surrogate revenue (no order revenue/fee column)"
    if hold is None: hold, sur["Holding"] = rev * .02, "Surrogate: 2% of revenue"
    u, cap, km, cpk, fs, runs = (get(c, m, r) for r in ("used", "capacity", "km", "cpk", "fleet_status", "runs")); util = avail = fc = None
    if u is not None and cap is not None and cap.sum(): util = u.sum() / cap.sum() * 100
    if fs is not None:
        s_ = fs.astype(str).str.lower(); op = s_.isin(["assigned", "available", "standby"]); avail = op.mean() * 100
        if util is None and op.sum(): util = s_.eq("assigned").sum() / op.sum() * 100
    if vc is not None: fc = vc.sum()
    elif km is not None and cpk is not None: fc = (km * cpk).sum()
    elif runs is not None and get(c, m, "route_cost") is not None:
        fc = (runs[op] if fs is not None else runs).sum() * get(c, m, "route_cost").mean(); sur["Fleet cost"] = "Estimated: runs by operable vehicles × average route base cost"
    if util is None: util, sur["Util"] = 70.0, "Demand-based surrogate 70%"
    if fc is None: fc, sur["Fleet cost"] = rev * .35, "Surrogate: 35% of revenue"
    ot, df_, ap = get(c, m, "ontime"), get(c, m, "defect"), get(c, m, "approval"); sup = sup_ok = None
    if ot is not None and df_ is not None: sup = (ot * .6 + (100 - df_) * .4).mean()
    if ap is not None:
        a_ = ap.astype(str).str.strip().str.lower(); sup_ok = a_.eq("approved").mean() * 100
        if sup is None: sup = a_.map(lambda v: 100 if v == "approved" else 60 if v == "conditional" else 30).mean()
    if sup is None: sup, sur["Supplier"] = 80.0, "Neutral surrogate 80/100"
    return dict(otd=otd, util=util, ret=ret, hold=hold, sup=sup, rev=rev, fleet_cost=fc, fixed=fixed, inv_acc=inv_acc, fleet_avail=avail, sup_ok=sup_ok, n_orders=n_ord, missing=miss, surrogate=sur)

def assessment(b):
    if not b["missing"] and not b["surrogate"]: return "All feeds and KPI columns found - full-confidence assessment."
    s = "**Operational Assessment under Partial Data.** "
    if b["missing"]: s += f"Missing feeds: {', '.join(b['missing'])}. "
    return s + "Surrogates in use: " + "; ".join(f"{k}: {v}" for k, v in b["surrogate"].items()) + ". Treat affected KPIs as indicative (about ±15%) until data is restored."

# ============ SIMULATION ============
NEUTRAL = dict(demand=0, fuel=1.0, drivers=100, delay=0)
def simulate(b, p):
    dm, cap, util = 1 + p["demand"] / 100, max(p["drivers"] / 100, .05), b["util"] / 100
    load = dm * util / cap; served = min(1, 1 / load) if load else 1; rev = b["rev"] * dm * served
    cost = b["fleet_cost"] * dm * served * p["fuel"] + b.get("fixed", 0) + b["hold"] * (1 + .2 * (1 - served) + p["delay"] / 600) + b["rev"] * b["ret"] / 100 * .3 * dm * served
    risk = min(100, max(0, (100 - b["otd"]) + p["delay"] * .35 + max(0, load - 1) * 60))
    return {"Revenue": rev, "Op. Cost": cost, "Deficit (cost-rev)": cost - rev, "Margin %": (rev - cost) / rev * 100 if rev else 0, "Delay Risk %": risk, "Fleet Load %": min(150, load * 100), "Demand Served %": served * 100}

# ============ RECOMMENDATIONS (each with evidence) ============
def hub_otd(c, m):
    x = otd_frame(c, m)
    return None if x is None or not len(x) else (x.groupby("hub").ok.mean() * 100).sort_values()

def recommend(b, c, m):
    r, sim0 = [], simulate(b, NEUTRAL); h = hub_otd(c, m)
    if b["otd"] < 85: r.append(f"**Raise on-time delivery** - OTD is {b['otd']:.1f}% vs 85% target." + (f" Worst hub: {h.index[0]} at {h.iloc[0]:.1f}% - add drivers or buffer time there first." if h is not None else ""))
    if b["util"] > 90: r.append(f"**Fleet is stretched** ({b['util']:.1f}% utilised): no headroom for demand spikes - pre-arrange 3PL backup.")
    elif b["util"] < 65: r.append(f"**Fleet under-used** ({b['util']:.1f}%): consolidate routes or reduce vehicles to cut cost.")
    if b["ret"] > 8: r.append(f"**Return rate high** ({b['ret']:.1f}%): audit top-returning hubs/SKUs; each point of returns costs about {inr(b['rev'] * .01 * .3)}.")
    if b["sup"] < 80: r.append(f"**Supplier score weak** ({b['sup']:.1f}/100): renegotiate SLAs with the lowest on-time suppliers.")
    if sim0["Deficit (cost-rev)"] > 0: r.append(f"**Operating deficit of {inr(sim0['Deficit (cost-rev)'])}**: fleet cost ({inr(b['fleet_cost'])}) is the biggest lever.")
    r.append(f"Holding cost is about {inr(b['hold'])}/month - clear slow-moving stock before adding inventory.")
    if b.get("inv_acc") is not None and b["inv_acc"] < 95: r.append(f"**Inventory records unreliable**: only {b['inv_acc']:.1f}% of rows have system stock = physical count - schedule cycle counts before trusting availability.")
    if b.get("fleet_avail") is not None and b["fleet_avail"] < 70: r.append(f"**Fleet availability low** ({b['fleet_avail']:.1f}% of vehicle records operable) - fix maintenance backlog.")
    if b.get("sup_ok") is not None and b["sup_ok"] < 70: r.append(f"**Only {b['sup_ok']:.0f}% of suppliers fully approved** - move volume to approved suppliers and close conditional reviews.")
    if b["surrogate"]: r.append("⚠ Some KPIs use surrogates (see banner) - confirm with real data before acting.")
    return r

# ============ ASK YOUR DATA (no API needed) ============
def ask(q, c):
    toks = {w[:5] for w in re.findall(r"[a-z]+", q.lower()) if len(w) > 2}; ql = q.lower()
    ov = lambda x: sum(t[:5] in toks for t in re.split(r"[_\W]+", x.lower()) if len(t) > 2)
    best = max(c, key=lambda s_: sum(ov(x) for x in c[s_].columns) + 2 * ov(s_)); df = c[best]
    num = sorted([x for x in df.columns if pd.api.types.is_numeric_dtype(df[x]) and ov(x)], key=ov, reverse=True)
    cat = sorted([x for x in df.columns if not pd.api.types.is_numeric_dtype(df[x]) and not pd.api.types.is_datetime64_any_dtype(df[x]) and ov(x) and not re.search(r"(id|key)$", x, re.I)], key=ov, reverse=True)
    if "hub" in ql and not cat: cat = [x for x in df.columns if re.search(HUBRX, x, re.I)][:1]
    agg = "mean" if re.search(r"averag|mean|typical", ql) else "count" if re.search(r"how many|count|number of", ql) else "sum"
    asc = bool(re.search(r"lowest|least|minimum|worst|smallest", ql))
    if num and cat: t = df.groupby(cat[0])[num[0]].agg(agg).sort_values(ascending=asc).reset_index(); code = f"{best}.groupby('{cat[0]}')['{num[0]}'].{agg}()"
    elif cat: t = df[cat[0]].value_counts().reset_index(); t.columns = [cat[0], "count"]; code = f"{best}['{cat[0]}'].value_counts()"
    elif num: t = df[num[0]].agg(["sum", "mean", "min", "max"]).to_frame().T; code = f"{best}['{num[0]}'].agg(sum/mean/min/max)"
    else: return None, f"Couldn't match your words to a column in '{best}'. Try using column names, e.g. 'total revenue by hub'.", None
    top = t.iloc[0]; return t, f"From sheet **{best}**: top result is **{top.iloc[0]}** ({top.iloc[1]:,.2f})." if len(t) > 1 and not isinstance(top.iloc[0], (int, float)) or len(t) > 1 else f"From sheet **{best}**.", code

def llm(q, ctx, key):
    import anthropic
    r = anthropic.Anthropic(api_key=key).messages.create(model="claude-sonnet-4-6", max_tokens=800, messages=[{"role": "user", "content": f"You are an operations analyst. Answer only from this data context, cite numbers, state limits.\n\n{ctx}\n\nQuestion: {q}"}])
    return r.content[0].text

# ============ POLICY RAG (C11) ============
DEMO_DOCS = {"SOP_Operations.txt": "1. Driver Shift Policy\nDrivers shall not exceed 9 hours per shift. Standard Operating Procedure.\n2. Route Speed\nVehicles must not exceed 40 km/h in zone A routes. Standard Operating Procedure.",
             "Executive_Override.txt": "1. Night Shift Directive\nExecutive Override: drivers may extend to 11 hours per shift during hub outages.\n2. Route Speed\nExecutive Override: speed limit in zone A is 30 km/h during monsoon.",
             "Provisional_Notes.txt": "1. Draft Speed Note\nProvisional note: zone A speed limit may rise to 50 km/h. Not yet approved."}
TIERS = [("Executive", 4, r"executive|override|directive|ceo"), ("Unverified", 1, r"unverified"), ("Provisional", 2, r"provisional|draft"), ("SOP", 3, r"standard operating|sop|policy|confirmed")]
TOPICS = {"speed": r"speed|km/h", "shift": r"shift|hours", "return": r"return|refund", "hub": r"hub|outage", "route": r"route|vehicle|restrict", "ontime": r"on-?time|sla|promis", "supplier": r"supplier|approv|lead time"}
def ref_docs(c):
    """Reference sheets (routes, service rules) become policy passages; 'Confirmed' > 'Provisional' > 'Unverified'."""
    return {f"{s_}.sheet": "\n".join(" | ".join(f"{k}: {v}" for k, v in zip(df.columns, r) if v) for r in df.fillna("").astype(str).itertuples(index=False)) for s_, df in c.items() if re.search(r"route|_ref|reference", s_, re.I)}
def read_doc(f):
    if f.name.lower().endswith(".pdf"):
        from pypdf import PdfReader
        return "\n".join(p.extract_text() or "" for p in PdfReader(f).pages)
    return f.read().decode("utf-8", "ignore")
def parse_docs(docs):
    secs, title = [], "Preamble"
    for dn, txt in docs.items():
        tier = next(((n, w) for n, w, rx in TIERS if re.search(rx, dn, re.I)), ("SOP", 2))
        for ln in txt.splitlines():
            if re.match(r"^(\d+[.)]|#+|section)\s*\S", ln, re.I) and len(ln) < 70: title = ln.strip(); continue
            if ln.strip(): t2 = next(((n, w) for n, w, rx in TIERS if re.search(rx, ln, re.I)), tier); secs.append(dict(doc=dn, section=title, text=ln.strip(), tier=t2[0], w=t2[1]))
    return pd.DataFrame(secs, columns=["doc", "section", "text", "tier", "w"])
def find_conflicts(secs):
    rows = []
    for tp, rx in TOPICS.items():
        g = secs[secs.text.str.contains(rx, case=False, regex=True)]
        g = g.assign(val=g.text.str.extract(r"(\d+(?:\.\d+)?)\s*(?:km/h|hours|hrs|min(?:ute)?s?|%|days|units|km)", flags=re.I)[0]).dropna(subset=["val"])
        if g.val.nunique() > 1:
            win = g.sort_values("w", ascending=False, kind="stable").iloc[0]
            rows.append(dict(Topic=tp, Competing="; ".join(f"{r.val} ({r.tier}, {r.doc})" for r in g.itertuples()), Winner=f"{win.val} - {win.tier} ({win.doc} / {win.section})", Rule="Executive > SOP > Provisional"))
    return pd.DataFrame(rows, columns=["Topic", "Competing", "Winner", "Rule"])
def answer(q, secs, c, conf):
    stem = lambda t: {w[:5] for w in re.findall(r"[a-z0-9/]+", t.lower()) if w not in {"the", "a", "is", "what", "of", "in", "for", "and", "to", "at", "which", "have", "has", "or"}}; tok = stem(q)
    if secs.empty: return "No policy documents loaded.", None
    secs = secs.assign(score=secs.text.map(lambda t: len(tok & stem(t))) + secs.w * .1); top = secs[secs.score >= 1.1].sort_values("score", ascending=False).head(3)
    if top.empty: return "No supporting passage found in the loaded policy documents.", None
    out = "**Evidence-backed answer**\n\n" + "\n".join(f"- {r.text}  \n  _Source: {r.doc} › {r.section} [{r.tier}]_" for r in top.itertuples())
    for r in conf.itertuples():
        if re.search(TOPICS[r.Topic], q, re.I): out += f"\n\n⚠️ **Conflict on '{r.Topic}'** - resolved by hierarchy: **{r.Winner}**."
    rows = [d[d[next(x for x in d.columns if re.search(HUBRX, x, re.I))] == h].head(3).assign(_sheet=s) for h in {std_hub(m) for m in re.findall(r"hub[\s_\-]*\d+", q, re.I)} for s, d in c.items() if any(re.search(HUBRX, x, re.I) for x in d.columns)]
    return out, (pd.concat(rows) if rows else None)

# ============ CONSTRAINT CARDS (presets + free text) ============
PRESETS = {"Hub Outage - South Bengaluru": "hub outage south bengaluru", "Fleet Strike": "fleet strike", "R1/R2 Data Ingestion Disruption": "returns and fleet data missing",
           "Fuel price shock": "fuel +40%", "Festival demand spike": "demand +50%, delay 20 min"}
def parse_card(t):
    p, tl, notes, drop = dict(NEUTRAL), t.lower(), [], []
    num = lambda rx: (lambda m: int(m.group(1)) if m else None)(re.search(rx, tl))
    if "outage" in tl: p.update(demand=12, drivers=85, delay=35, fuel=1.08); notes.append("Hub outage: +12% spillover demand, 15% drivers unreachable, +8% fuel, +35 min delay (assumed defaults)")
    if "strike" in tl: p.update(drivers=55, fuel=1.05, delay=20); notes.append("Strike: 45% driver availability lost, +5% spot-hire cost, +20 min delay (assumed defaults)")
    if (v := num(r"demand\D{0,20}?([+-]?\d+)")) is not None: p["demand"] = v; notes.append(f"Demand {v:+d}% (from your text)")
    if (v := num(r"(?:fuel|cost)\D{0,20}?([+-]?\d+)")) is not None: p["fuel"] = 1 + v / 100; notes.append(f"Fuel/fleet cost {v:+d}% (from your text)")
    if (v := num(r"(?:driver|staff|availab)\w*\D{0,20}?([+-]?\d+)")) is not None: p["drivers"] = 100 - abs(v) if (v < 0 or re.search(r"loss|lost|drop|reduc|short|down", tl)) else v; notes.append(f"Driver availability {p['drivers']}% (from your text)")
    if (v := num(r"(?:delay|bottleneck)\D{0,20}?(\d+)")) is not None: p["delay"] = v; notes.append(f"Warehouse delay {v} min (from your text)")
    for g, kw in {"Orders": "order|delivery data|sla", "Fleet": "fleet|vehicle", "Returns": "return", "Inventory": "inventory|stock", "Suppliers": "supplier"}.items():
        if re.search(r"missing|unavailable|no data|disrupt|lost", tl) and re.search(kw, tl) and not re.search(r"strike", tl): drop.append(g)
    if drop: notes.append(f"Feeds treated as unavailable: {', '.join(drop)}")
    return p, drop, notes
def plan(p, drop):
    s = []
    if p["demand"] > 0: s.append("Pre-position fast-moving stock and open overflow hub capacity for the demand spike")
    if p["drivers"] < 100: s.append("Activate 3PL/gig drivers on priority lanes; consider Executive Override on shift length (see Policy tab)")
    if p["fuel"] > 1: s.append("Consolidate routes and raise load per vehicle to offset fuel cost")
    if p["delay"] > 0: s.append(f"Communicate +{p['delay']} min ETA to customers; fast-track dispatch for premium orders")
    if drop: s.append("Switch to surrogate KPIs, widen decision bands to ±15%, request re-export of missing feeds")
    return s or ["No levers moved - nothing to mitigate"]

# ==== UI ====
st.title("BUDS Round 2.5 · Night Shift"); st.caption("Bengaluru Analytics & Operations Console - Alliance University")
with st.sidebar:
    st.header("Data feeds")
    up = st.file_uploader("CSV / Excel (multi-sheet)", ["csv", "xlsx", "xls"], accept_multiple_files=True)
    dups = st.file_uploader("Policy docs (PDF/TXT)", ["pdf", "txt"], accept_multiple_files=True)
    drop = st.multiselect("Simulate missing feed (C08)", list(GROUPS)); key = st.text_input("Optional AI key (Anthropic) for chat answers", type="password")
    st.caption("No upload = built-in dirty demo data.")
raw = load_files(up) if up else demo_data(); cleaned, log = clean(raw); auto = detect(cleaned)
with st.sidebar.expander("Column mapping (fix if auto-detect is wrong)"):
    opts = ["(none)"] + [f"{s} :: {x}" for s, d in cleaned.items() for x in d.columns]; m = {}
    for r in ROLES:
        cur = f"{auto[r][0]} :: {auto[r][1]}" if auto[r] else "(none)"; ch = st.selectbox(r, opts, index=opts.index(cur), key=f"map_{r}_{len(opts)}")
        m[r] = tuple(ch.split(" :: ", 1)) if ch != "(none)" else None
base = kpis(cleaned, m, drop); sim0 = simulate(base, NEUTRAL)
docs = {f.name: read_doc(f) for f in dups}; docs.update(ref_docs(cleaned)); docs = docs or DEMO_DOCS; secs = parse_docs(docs); conflicts = find_conflicts(secs)
t1, t2, t3, t4, t5, t6, t7 = st.tabs(["Dashboard", "Data Health & Audit", "Ask Your Data", "Policy Q&A", "What-If Sandbox", "Constraint Cards", "Methodology"])

with t1:
    for g in base["missing"]: st.error(f"Missing feed: {g} - fallback surrogates active", icon="🚫")
    fl = lambda k: "⚠ surrogate" if k in base["surrogate"] else ""
    items = [("On-Time Delivery", f"{base['otd']:.1f}%", "OTD", OK if base["otd"] >= 85 else WARN), ("Fleet Utilisation", f"{base['util']:.1f}%", "Util", CYAN), ("Return Rate", f"{base['ret']:.1f}%", "Return", OK if base["ret"] < 8 else BAD),
             ("Holding Cost / mo", inr(base["hold"]), "Holding", CYAN), ("Supplier Score", f"{base['sup']:.1f}", "Supplier", OK if base["sup"] >= 80 else WARN), ("Operating Deficit", inr(sim0["Deficit (cost-rev)"]), "Fleet cost", BAD if sim0["Deficit (cost-rev)"] > 0 else OK)]
    for col, (t, v, k, cl) in zip(st.columns(6), items): col.markdown(card(t, v, cl, fl(k)), unsafe_allow_html=True)
    ex = [(t, v) for t, v in [("Inventory Accuracy", base["inv_acc"]), ("Fleet Availability", base["fleet_avail"]), ("Suppliers Approved", base["sup_ok"])] if v is not None]
    for col, (t, v) in zip(st.columns(3), ex): col.markdown(card(t, f"{v:.1f}%", CYAN), unsafe_allow_html=True)
    st.info(assessment(base)); st.subheader("Recommendations (with evidence)"); recs = recommend(base, cleaned, m)
    for x in recs: st.markdown(f"- {x}")
    st.subheader("Explore any column"); sh = st.selectbox("Sheet", list(cleaned), key="ex_s"); d = cleaned[sh]
    nums = [x for x in d.columns if pd.api.types.is_numeric_dtype(d[x])]; cats = [x for x in d.columns if d[x].nunique() <= 30]
    if nums and cats:
        a, b2 = st.columns(2); y, xg = a.selectbox("Measure", nums), b2.selectbox("Group by", cats); g = d.groupby(xg)[y].sum().reset_index()
        st.plotly_chart(style(go.Figure(go.Bar(x=g[xg].astype(str), y=g[y], marker_color=CYAN)).update_layout(title=f"{y} by {xg}")), width="stretch")
    h = hub_otd(cleaned, m)
    if h is not None: st.plotly_chart(style(go.Figure(go.Bar(x=h.index, y=h.values, marker_color=[OK if v >= 85 else WARN if v >= 50 else BAD for v in h.values])).update_layout(title="On-time delivery % by hub")), width="stretch")
    st.download_button("⬇ Download report", "# BUDS Night Shift report\n\n" + assessment(base) + "\n\n## Recommendations\n" + "\n".join(f"- {x}" for x in recs), "report.md")

with t2:
    st.subheader("Data Health & Audit Trail"); st.caption("Raw uploads are preserved unchanged; every transformation applies only to the cleaned copy.")
    a, b_ = st.columns(2); a.metric("Transformations logged", len(log)); b_.metric("Rows touched", sum(x["Rows affected"] for x in log))
    st.dataframe(pd.DataFrame(log), width="stretch"); st.markdown("**Data structure profile**"); st.dataframe(profile(raw), width="stretch")
    sh = st.selectbox("Compare sheet", list(raw)); c1, c2 = st.columns(2); c1.markdown("**Raw (immutable)**"); c1.dataframe(raw[sh].head(25)); c2.markdown("**Cleaned**"); c2.dataframe(cleaned[sh].head(25))

with t3:
    st.subheader("Ask your data in plain English"); st.caption("Examples: 'total revenue by hub', 'average used hours by hub', 'how many orders by status'")
    q = st.text_input("Question", "total revenue by hub", key="q_data")
    if q:
        t, msg, code = ask(q, cleaned); st.markdown(msg)
        if t is not None:
            st.dataframe(t, width="stretch"); st.caption(f"Calculation used: `{code}`")
            if len(t) > 1: st.plotly_chart(style(go.Figure(go.Bar(x=t.iloc[:, 0].astype(str), y=t.iloc[:, 1], marker_color=CYAN))), width="stretch")
        if key:
            try: st.markdown("**AI analyst answer**"); st.markdown(llm(q, assessment(base) + "\n" + str(base) + "\n" + "\n".join(f"{s}:\n{d.head(15).to_string()}" for s, d in cleaned.items()), key))
            except Exception as e: st.warning(f"AI answer unavailable: {e}")

with t4:
    st.subheader("Rule hierarchy: Executive Overrides > SOP > Provisional Notes")
    if len(conflicts): st.warning("Conflicting rules detected and resolved"); st.dataframe(conflicts, width="stretch")
    else: st.success("No conflicting numeric rules detected.")
    q = st.text_input("Ask a policy / data question", "What is the speed limit in zone A and the driver shift limit at Hub 3?", key="q_pol")
    if q:
        ans, rows = answer(q, secs, cleaned, conflicts); st.markdown(ans)
        if rows is not None: st.markdown("**Dataset evidence (cleaned rows)**"); st.dataframe(rows)
    with st.expander("Parsed policy passages"): st.dataframe(secs[["doc", "section", "tier", "text"]])

with t5:
    st.subheader("Strategy Sandbox"); c = st.columns(4)
    p = dict(demand=c[0].slider("Demand spike %", -30, 100, 0), fuel=c[1].slider("Fuel/Fleet cost ×", .8, 2.0, 1.0, .05), drivers=c[2].slider("Driver availability %", 30, 100, 100), delay=c[3].slider("Warehouse delay (min)", 0, 120, 0))
    res = simulate(base, p); k = st.columns(4)
    k[0].metric("Margin %", f"{res['Margin %']:.1f}", f"{res['Margin %'] - sim0['Margin %']:+.1f} pts"); k[1].metric("Deficit", inr(res["Deficit (cost-rev)"]), inr(res["Deficit (cost-rev)"] - sim0["Deficit (cost-rev)"]), delta_color="inverse")
    k[2].metric("Delay risk %", f"{res['Delay Risk %']:.0f}", f"{res['Delay Risk %'] - sim0['Delay Risk %']:+.0f}", delta_color="inverse"); k[3].metric("Demand served %", f"{res['Demand Served %']:.0f}")
    st.plotly_chart(style(go.Figure([go.Bar(name="Baseline", x=["Revenue", "Op. Cost"], y=[sim0["Revenue"], sim0["Op. Cost"]], marker_color=BORDER), go.Bar(name="Scenario", x=["Revenue", "Op. Cost"], y=[res["Revenue"], res["Op. Cost"]], marker_color=CYAN)])), width="stretch")

with t6:
    st.subheader("Live Constraint Card Simulator"); st.caption("Pick a preset OR type any new card from the judges, e.g. 'Fuel +40% and drivers down 30%, hub outage'")
    pre = st.selectbox("Preset card", list(PRESETS)); txt = st.text_input("Or type a custom card (overrides preset if filled)", "", key="card_txt")
    if st.button("⚡ Simulate Live Constraint Card", type="primary"):
        text = txt or PRESETS[pre]; cp, cdrop, notes = parse_card(text); ab = kpis(cleaned, m, tuple(set(drop) | set(cdrop)))
        before, after = simulate(base, NEUTRAL), simulate(ab, cp); diff = pd.DataFrame({"Before": before, "After": after}); diff["Δ"] = diff.After - diff.Before
        st.markdown(f"#### 🔴 {text}"); st.dataframe(diff.style.format("{:,.1f}"), width="stretch")
        ks = ["Margin %", "Delay Risk %", "Fleet Load %"]; st.plotly_chart(style(go.Figure([go.Bar(name="Before", x=ks, y=[before[x] for x in ks], marker_color=OK), go.Bar(name="After", x=ks, y=[after[x] for x in ks], marker_color=BAD)])), width="stretch")
        if cdrop: st.warning(assessment(ab))
        l, r = st.columns(2); l.markdown("**Mitigation action plan**"); [l.markdown(f"{i}. {s}") for i, s in enumerate(plan(cp, cdrop), 1)]
        r.markdown("**Audit log - assumptions**"); [r.markdown(f"- {s}") for s in notes + [f"Parameters applied: {cp}", "Raw data unchanged; scenario computed on a cleaned copy"]]

with t7:
    st.subheader("How the numbers are calculated")
    st.markdown("""- **On-time delivery** = delivered orders where actual cycle minutes ≤ promised minutes (or delivery date ≤ promised date). Rows with missing/invalid times are excluded from the denominator, never guessed.
- **Fleet utilisation** = used hours ÷ capacity hours, or (if absent) Assigned ÷ operable (Assigned+Available+Standby) vehicle records. **Fleet availability** = operable ÷ all records. **Fleet cost** = order variable cost, else Σ km × cost/km, else runs × average route base cost.
- **Return rate** = share of orders flagged returned. **Holding cost / month** = Σ(qty × unit cost) × 24% ÷ 12 (assumed 24% annual carrying rate).
- **Supplier score** = 60% on-time % + 40% × (100 − defect %); if absent, approval score (Approved 100, Conditional 60, other 30). **Operating deficit** = operating cost − revenue (revenue = delivery fees; cost = variable + hub fixed cost × days + holding + 30% of revenue lost on returns). **Return rate** = return requests ÷ orders when no return flag exists.
- **What-if**: load = demand × utilisation ÷ driver availability; if load > 100% only 1/load of demand is served. Delay risk = (100 − OTD) + 0.35 × delay min + 60 × overload.
- **Policy conflicts**: highest tier wins - Executive > SOP > Provisional.
- **Cleaning**: duplicates dropped, hub IDs standardised to BLR_HUB_NN, bad dates blanked, missing numbers → median, missing text → UNKNOWN, 'Delivered' without delivery date → 'In Transit'. Raw data is never modified.""")
    st.warning("**Limitations**: surrogate values (85% OTD, 6% returns, 70% utilisation, 80 supplier) are benchmarks, not facts; the simulator is a simplified linear model; column auto-detection can be corrected in the sidebar mapping; policy rule reading is keyword-based; figures are indicative (±15%) when feeds are missing.")
