"""BUDS Round 2.5 - Night Shift | Bengaluru Operations Decision-Support Console (v3)
Run:  pip install -r requirements.txt   then   streamlit run app.py
Works on ANY CSV / Excel dataset; understands the Team-17 operations workbooks out of the box."""
import io, os, re, glob, warnings, numpy as np, pandas as pd, streamlit as st
import plotly.graph_objects as go
warnings.filterwarnings("ignore")

# ============ THEME ============
BG, CARD, CARD2, BORDER = "#070B14", "#0F1624", "#141D2F", "#1F2A40"
CYAN, VIOLET, OK, WARN, BAD, MUTED = "#22D3EE", "#8B5CF6", "#34D399", "#FBBF24", "#FB7185", "#94A3B8"
st.set_page_config(page_title="BUDS Night Shift · Decision Console", page_icon="🌙", layout="wide", initial_sidebar_state="expanded")
CSS = f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"],.stApp{{font-family:'Inter',system-ui,-apple-system,'Segoe UI',sans-serif}}
.stApp{{background:radial-gradient(1200px 500px at 85% -10%,#16254a55,transparent),radial-gradient(900px 400px at -10% 0%,#0e3a4a44,transparent),{BG};color:#E2E8F0}}
#MainMenu,footer{{visibility:hidden}} header[data-testid=stHeader]{{background:transparent}}
.block-container{{padding-top:1.2rem;max-width:1500px}}
section[data-testid=stSidebar]{{background:linear-gradient(180deg,{CARD},{BG});border-right:1px solid {BORDER}}}
.hero{{position:relative;overflow:hidden;border:1px solid {BORDER};border-radius:18px;padding:22px 28px;margin-bottom:14px;
 background:linear-gradient(120deg,#0b1f33 0%,#14213f 45%,#2a1b4d 100%)}}
.hero:after{{content:"";position:absolute;right:-60px;top:-60px;width:260px;height:260px;border-radius:50%;background:radial-gradient(circle,{CYAN}55,transparent 70%)}}
.hero h1{{margin:0;font-size:30px;font-weight:800;letter-spacing:-.5px;background:linear-gradient(90deg,#fff,{CYAN} 70%,{VIOLET});-webkit-background-clip:text;-webkit-text-fill-color:transparent}}
.hero p{{margin:4px 0 12px;color:#B6C3DA;font-size:14px}}
.pill{{display:inline-block;margin:0 6px 4px 0;padding:3px 11px;border-radius:999px;font-size:12px;font-weight:600;border:1px solid {BORDER};background:#0b1322cc;color:#CBD5E1}}
.pill.ok{{color:{OK};border-color:{OK}55}} .pill.warn{{color:{WARN};border-color:{WARN}55}} .pill.bad{{color:{BAD};border-color:{BAD}55}} .pill.cy{{color:{CYAN};border-color:{CYAN}55}}
.kpi{{position:relative;background:linear-gradient(160deg,{CARD2},{CARD});border:1px solid {BORDER};border-radius:14px;padding:14px 16px 12px;height:100%;transition:transform .15s,border-color .15s}}
.kpi:hover{{transform:translateY(-2px);border-color:var(--ac)}}
.kpi:before{{content:"";position:absolute;left:0;top:14px;bottom:14px;width:3px;border-radius:3px;background:var(--ac)}}
.kpi h4{{margin:0;min-height:30px;font-size:11px;letter-spacing:.08em;color:{MUTED};text-transform:uppercase;font-weight:600}}
.kpi .v{{margin:4px 0 2px;white-space:nowrap;font-size:24px;font-weight:700;color:#F8FAFC;line-height:1.1}} .kpi .s{{font-size:12px;color:{MUTED}}}
.kpi .tag{{float:right;font-size:10px;padding:1px 7px;border-radius:999px;background:{WARN}22;color:{WARN};font-weight:600}}
.insight{{border:1px solid {BORDER};border-left:3px solid var(--ac);background:{CARD};border-radius:10px;padding:10px 14px;margin:6px 0;font-size:14px}}
.sec{{display:flex;align-items:center;gap:10px;margin:18px 0 6px;font-size:18px;font-weight:700;color:#F1F5F9}}
.sec:after{{content:"";flex:1;height:1px;background:linear-gradient(90deg,{BORDER},transparent)}}
.stTabs [data-baseweb=tab-list]{{gap:6px;border-bottom:1px solid {BORDER}}}
.stTabs [data-baseweb=tab]{{border-radius:10px 10px 0 0;padding:8px 14px;color:{MUTED};font-weight:600}}
.stTabs [aria-selected=true]{{color:{CYAN}!important;background:{CARD}}}
div[data-testid=stMetricValue]{{font-size:1.55rem}} div[data-testid=stMetricLabel] p{{font-size:12px;white-space:nowrap}}
div[data-testid=stMetric]{{background:{CARD};border:1px solid {BORDER};border-radius:12px;padding:10px 14px}}
.stButton>button,.stDownloadButton>button{{border-radius:10px;font-weight:600;border:1px solid {BORDER}}}
.stButton>button[kind=primary]{{background:linear-gradient(90deg,#0891b2,{VIOLET});border:0;color:#fff}}
</style>"""
st.markdown(CSS, unsafe_allow_html=True)

def kpi(title, value, accent=CYAN, sub="", tag=""):
    t = f'<span class="tag">{tag}</span>' if tag else ""
    return f'<div class="kpi" style="--ac:{accent}"><h4>{title}{t}</h4><div class="v">{value}</div><div class="s">{sub}&nbsp;</div></div>'
def section(txt): st.markdown(f'<div class="sec">{txt}</div>', unsafe_allow_html=True)
def insight(txt, ac=CYAN): st.markdown(f'<div class="insight" style="--ac:{ac}">{txt}</div>', unsafe_allow_html=True)
inr = lambda v: f"₹{v:,.0f}"
def style(fig, h=330):
    fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=h, margin=dict(l=10, r=10, t=44, b=10),
                      colorway=[CYAN, VIOLET, OK, WARN, BAD], font=dict(family="Inter, sans-serif", size=12), title_font=dict(size=15), legend=dict(orientation="h", y=1.12))
    fig.update_xaxes(gridcolor="#1b2740", zeroline=False); fig.update_yaxes(gridcolor="#1b2740", zeroline=False)
    return fig
def show(fig, h=330): st.plotly_chart(style(fig, h), width="stretch")

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
        for c in [x for x in df.columns if not pd.api.types.is_numeric_dtype(df[x]) and not pd.api.types.is_datetime64_any_dtype(df[x]) and 1 < df[x].nunique() <= 60 and not re.search(DATERX + r"|(id|key|code|sku)$", x, re.I)]:      # 'ICU' / 'icu ' / 'Icu' -> one spelling
            k = df[c].astype(str).str.strip().str.lower(); k = k.where(df[c].notna())
            if k.nunique() < df[c].nunique():
                canon = df.assign(_k=k).dropna(subset=["_k"]).groupby("_k")[c].agg(lambda z: z.value_counts().index[0]); new = k.map(canon).where(df[c].notna(), df[c]); ch = ((new != df[c]) & df[c].notna()).sum()
                df[c] = new
                if ch: L(name, "Category spelling unified", f"{c}: case/space variants merged into the most common spelling", ch)
        for nc in [x for x in df.columns if re.search(r"_name$", x, re.I) and not pd.api.types.is_numeric_dtype(df[x])]:      # X_Name <-> X_ID restoration (only when mapping is strictly one-to-one)
            ic = re.sub(r"_name$", "_ID", nc, flags=re.I)
            if ic not in df.columns: continue
            for tgt, src_ in ((nc, ic), (ic, nc)):
                pair = df.dropna(subset=[tgt, src_])
                if pair.empty or pair.groupby(src_)[tgt].nunique().max() > 1 or pair.groupby(tgt)[src_].nunique().max() > 1: continue
                ref = pair.groupby(src_)[tgt].agg(lambda z: z.mode().iat[0]); f_ = df[tgt].isna() & df[src_].map(ref).notna()
                if f_.sum(): df.loc[f_, tgt] = df.loc[f_, src_].map(ref); L(name, "Missing value restored", f"{tgt} filled from {src_} (one-to-one mapping)", f_.sum())
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

# ============ ASK YOUR DATA (works offline; optional Claude for richer answers) ============
def ask(q, c):
    ql = q.lower(); toks = {w[:5] for w in re.findall(r"[a-z]+", ql) if len(w) > 2}
    ov = lambda x: sum(t[:5] in toks for t in re.split(r"[_\W]+", x.lower()) if len(t) > 2)
    best = max(c, key=lambda s_: sum(ov(x) for x in c[s_].columns) + 2 * ov(s_)); df = c[best]
    isnum = lambda x: pd.api.types.is_numeric_dtype(df[x]); isdt = lambda x: pd.api.types.is_datetime64_any_dtype(df[x])
    # value filters: any category value spelled out in the question (e.g. 'H02', 'Delivered') becomes a filter
    qn = " " + re.sub(r"[^a-z0-9]+", " ", ql) + " "; flt = {}
    for col in df.columns:
        if isnum(col) or isdt(col) or df[col].nunique() > 60: continue
        for v in df[col].dropna().unique():
            vn = re.sub(r"[^a-z0-9]+", " ", str(v).lower()).strip()
            if len(vn) >= 2 and f" {vn} " in qn and vn not in {"yes", "no", "unknown", "all"}: flt.setdefault(col, []).append(v)
    num = sorted([x for x in df.columns if isnum(x) and ov(x)], key=ov, reverse=True)
    cat = sorted([x for x in df.columns if not isnum(x) and not isdt(x) and ov(x) and not re.search(r"(id|key)$", x, re.I)], key=ov, reverse=True)
    if re.search(r"\bhub", ql) and not cat: cat = [x for x in df.columns if re.search(HUBRX, x, re.I) and not isnum(x)][:1]
    by = re.search(r"\b(?:by|per|each|across)\b", ql)
    if flt and not by: cat = [x for x in cat if x not in flt]          # 'revenue in H02' filters; it does not group by hub
    sub = df
    for col, vals in flt.items(): sub = sub[sub[col].isin(vals)]
    agg = "mean" if re.search(r"averag|mean|typical", ql) else "median" if "median" in ql else "count" if re.search(r"how many|count|number of", ql) else "max" if re.search(r"\bmax|largest|biggest", ql) else "min" if re.search(r"\bmin(?!ute)|smallest", ql) else "sum"
    asc = bool(re.search(r"lowest|least|worst|smallest|bottom", ql)); n = re.search(r"(?:top|bottom|first)\s+(\d+)", ql); n = int(n.group(1)) if n else None
    fdesc = "; ".join(f"{k} = {', '.join(map(str, v))}" for k, v in flt.items())
    base_code = f"{best}[" + " & ".join(f"{best}['{k}'].isin({v})" for k, v in flt.items()) + "]" if flt else best
    if agg == "count" and cat: t = sub[cat[0]].value_counts(ascending=asc).reset_index(); t.columns = [cat[0], "count"]; code = f"{base_code}['{cat[0]}'].value_counts()"
    elif num and cat: t = sub.groupby(cat[0])[num[0]].agg(agg).sort_values(ascending=asc).reset_index(); code = f"{base_code}.groupby('{cat[0]}')['{num[0]}'].{agg}()"
    elif cat: t = sub[cat[0]].value_counts(ascending=asc).reset_index(); t.columns = [cat[0], "count"]; code = f"{base_code}['{cat[0]}'].value_counts()"
    elif num: t = sub[num[0]].agg(["sum", "mean", "median", "min", "max", "count"]).to_frame().T; code = f"{base_code}['{num[0]}'].agg(sum/mean/median/min/max/count)"
    elif flt: t = pd.DataFrame({"filter": [fdesc], "rows": [len(sub)]}); code = f"len({base_code})"
    else: return None, f"Couldn't match your words to a column in **{best}**. Try column names, e.g. *total revenue by hub* or *how many orders by status*.", None
    if n: t = t.head(n)
    msg = f"Answered from sheet **{best}** ({len(sub):,} rows" + (f" after filter: {fdesc}" if flt else "") + ")."
    if len(t) > 1 and t.shape[1] > 1 and not pd.api.types.is_numeric_dtype(t.iloc[:, 0]): msg += f" {'Lowest' if asc else 'Highest'}: **{t.iloc[0, 0]}** ({t.iloc[0, 1]:,.0f})." if float(t.iloc[0, 1]).is_integer() else f" {'Lowest' if asc else 'Highest'}: **{t.iloc[0, 0]}** ({t.iloc[0, 1]:,.2f})."
    return t, msg, code

def llm(q, ctx, key):
    import anthropic
    r = anthropic.Anthropic(api_key=key).messages.create(model="claude-sonnet-4-6", max_tokens=800, messages=[{"role": "user", "content": f"You are an operations analyst. Answer only from this data context, cite numbers, state limits.\n\n{ctx}\n\nQuestion: {q}"}])
    return r.content[0].text

# ============ EXTRA ANALYTICS (v3) ============
def health(raw, log):
    """Data Health Score = mean(Completeness, Uniqueness, Validity), each 0-100, measured on the RAW upload."""
    cells = sum(d.size for d in raw.values()); miss = sum(int(d.isna().sum().sum()) for d in raw.values()); rows = sum(len(d) for d in raw.values())
    dup = sum(int(d.duplicated().sum()) for d in raw.values()); inval = sum(x["Rows affected"] for x in log if re.search(r"invalid|conflict|standardis|normalis|restored", x["Transformation"], re.I))
    comp, uniq, val = 100 * (1 - miss / max(cells, 1)), 100 * (1 - dup / max(rows, 1)), 100 * (1 - min(1, inval / max(rows, 1)))
    return dict(score=(comp + uniq + val) / 3, completeness=comp, uniqueness=uniq, validity=val, missing=miss, dups=dup, fixed=inval, cells=cells, rows=rows)

def sheet_health(raw, cleaned):
    return pd.DataFrame([{"Sheet": s, "Rows (raw)": len(d), "Columns": d.shape[1], "Missing cells %": round(100 * d.isna().sum().sum() / max(d.size, 1), 1), "Exact duplicate rows": int(d.duplicated().sum()),
                          "Rows (cleaned)": len(cleaned[s]) if s in cleaned else None} for s, d in raw.items()])

def daily_trend(c, m):
    om = m.get("order_date")
    if not om or om[0] not in c or om[1] not in c[om[0]] or not pd.api.types.is_datetime64_any_dtype(c[om[0]][om[1]]): return None
    df = c[om[0]]; g = df.assign(_d=df[om[1]].dt.normalize()).dropna(subset=["_d"])
    if g._d.nunique() < 3: return None
    out = g.groupby("_d").size().rename("Orders").to_frame(); x = otd_frame(c, m)
    if x is not None and m.get("promised_min") and m["promised_min"][0] == om[0]:
        d2 = df.loc[x.index, om[1]].dt.normalize(); out["OTD %"] = (x.ok.groupby(d2).mean() * 100)
    return out

def status_mix(c, m):
    s = m.get("status")
    return c[s[0]][s[1]].value_counts() if s and s[0] in c and s[1] in c[s[0]] else None

def sensitivity(b):
    s0 = simulate(b, NEUTRAL); rows = []
    for k, v in {"Demand +20%": dict(demand=20), "Fuel / fleet cost ×1.2": dict(fuel=1.2), "Drivers −20%": dict(drivers=80), "Warehouse delay +30 min": dict(delay=30)}.items():
        p = dict(NEUTRAL); p.update(v); r = simulate(b, p); rows.append({"Shock": k, "Δ Deficit (₹)": r["Deficit (cost-rev)"] - s0["Deficit (cost-rev)"], "Δ Delay risk (pts)": r["Delay Risk %"] - s0["Delay Risk %"]})
    return pd.DataFrame(rows)

def combine(stack):
    p = dict(NEUTRAL); drop = set()
    for s in stack:
        p["demand"] += s["p"]["demand"]; p["fuel"] *= s["p"]["fuel"]; p["drivers"] = p["drivers"] * s["p"]["drivers"] / 100; p["delay"] += s["p"]["delay"]; drop |= set(s["drop"])
    p["drivers"] = max(5, p["drivers"]); return p, tuple(drop)

def severity(before, after, levers_moved=True):
    """Severity reflects the CHANGE a card causes (not the already-weak baseline)."""
    if not levers_moved: return ("INFO · data / rule card", CYAN)
    d = (after["Deficit (cost-rev)"] - before["Deficit (cost-rev)"]) / max(abs(before["Deficit (cost-rev)"]), 1) * 100
    s = max(0, after["Delay Risk %"] - before["Delay Risk %"]) + max(0, d) * .3 + max(0, after["Fleet Load %"] - 100) * .5
    return ("HIGH", BAD) if s >= 20 else ("MEDIUM", WARN) if s >= 5 else ("LOW", OK)

def card_title(t):
    t = t.strip(); mm = re.search(r"C\d{2}\s*[—–-]\s*[^\n]+", t)
    return mm.group(0).strip() if mm else (t.splitlines()[0] if t else "Card")[:90]

def build_xlsx(base, sim0, recs, log, cleaned, history, conflicts, hs):
    bio = io.BytesIO(); lab = {"otd": "On-time delivery %", "util": "Fleet utilisation %", "ret": "Return rate %", "hold": "Holding cost / month (INR)", "sup": "Supplier score", "rev": "Revenue (INR)", "fleet_cost": "Fleet cost (INR)", "fixed": "Hub fixed cost (INR)", "inv_acc": "Inventory accuracy %", "fleet_avail": "Fleet availability %", "sup_ok": "Suppliers approved %", "n_orders": "Orders analysed"}
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        pd.DataFrame([{"KPI": lab[k], "Value": (round(float(base[k]), 2) if base.get(k) is not None else None), "Surrogate?": "yes" if any(s_.lower()[:3] == k[:3] for s_ in base["surrogate"]) else ""} for k in lab]).to_excel(w, sheet_name="KPIs", index=False)
        pd.DataFrame({"Recommendation": [re.sub(r"\*\*", "", r) for r in recs]}).to_excel(w, sheet_name="Recommendations", index=False)
        pd.DataFrame([{"Health score": round(hs["score"], 1), "Completeness": round(hs["completeness"], 1), "Uniqueness": round(hs["uniqueness"], 1), "Validity": round(hs["validity"], 1)}]).to_excel(w, sheet_name="Data_Health", index=False)
        pd.DataFrame(log).to_excel(w, sheet_name="Audit_Log", index=False)
        if len(conflicts): conflicts.to_excel(w, sheet_name="Rule_Conflicts", index=False)
        if history: pd.DataFrame(history).to_excel(w, sheet_name="Scenario_History", index=False)
        for s_, d in cleaned.items(): d.to_excel(w, sheet_name=f"Clean_{s_}"[:31], index=False)
    return bio.getvalue()

# ============ GENERIC MODE (any dataset, no operations vocabulary needed) ============
def _main_sheet(c): return max(c, key=lambda k: len(c[k]) * c[k].shape[1])
def _nums(d): return [x for x in d.columns if pd.api.types.is_numeric_dtype(d[x]) and d[x].nunique() > 10 and not re.search(r"(id|key|index|code)$", x, re.I)]
def _cats(d): return [x for x in d.columns if not pd.api.types.is_numeric_dtype(d[x]) and not pd.api.types.is_datetime64_any_dtype(d[x]) and 2 <= d[x].nunique() <= 15 and not re.search(r"(id|key)$", x, re.I)]
def generic_kpis(c):
    d = c[_main_sheet(c)]; out = [("Rows", f"{len(d):,}", f"{len(c)} sheet(s)"), ("Columns", f"{d.shape[1]}", "in main sheet"), ("Completeness", f"{100 * (1 - d.isna().mean().mean()):.1f}%", "non-empty cells")]
    for x in _nums(d)[:3]: out.append((f"Total {x}"[:26], f"{d[x].sum():,.0f}", f"mean {d[x].mean():,.1f} · median {d[x].median():,.1f}"))
    return out
def generic_recs(c):
    d = c[_main_sheet(c)]; nums, cats, r = _nums(d), _cats(d), []; best = None
    for cat in cats:
        for num in nums[:4]:
            g = d.groupby(cat)[num].mean().dropna()
            if len(g) > 1 and d[num].mean():
                sp = (g.max() - g.min()) / abs(d[num].mean())
                if best is None or sp > best[0]: best = (sp, cat, num, g)
    if best:
        sp, cat, num, g = best; ov = d[num].mean()
        r.append(f"**Biggest performance gap: {cat}.** Average {num} is highest for **{g.idxmax()}** ({g.max():,.1f}, {100 * (g.max() / ov - 1):+.0f}% vs overall) and lowest for **{g.idxmin()}** ({g.min():,.1f}, {100 * (g.min() / ov - 1):+.0f}%). Investigate what drives the difference before reallocating resources (association, not proof of cause).")
    dts = [x for x in d.columns if pd.api.types.is_datetime64_any_dtype(d[x]) and d[x].notna().sum() > 20]
    if dts and nums:
        t = d.dropna(subset=[dts[0]]).sort_values(dts[0]); half = len(t) // 2; a_, b_ = t[nums[0]].iloc[:half].mean(), t[nums[0]].iloc[half:].mean()
        if a_: r.append(f"**Trend in {nums[0]}:** second half of the period averages {b_:,.1f} vs {a_:,.1f} in the first half ({100 * (b_ / a_ - 1):+.1f}%).")
    bad = [f"{k} ({v:.0%})" for k, v in d.isna().mean().items() if v > .1]
    if bad: r.append(f"**Data completeness risk:** {', '.join(bad[:4])} have more than 10% missing values - fix capture at source before relying on these fields.")
    for k in nums[:6]:
        o = int((((d[k] - d[k].mean()) / (d[k].std() or 1)).abs() > 3).sum())
        if o: r.append(f"**{k}** has {o} extreme outlier(s) (>3σ) - confirm whether they are entry errors or genuine spikes."); break
    return r or ["Not enough structure detected for automatic recommendations - use Ask Your Data or Auto Insights."]

# ============ CACHED PIPELINE (re-runs only when files change) ============
class _Up(io.BytesIO):
    def __init__(s, name, data): super().__init__(data); s.name = name
@st.cache_data(show_spinner="Reading & cleaning data…")
def pipeline(blobs):
    raw = load_files([_Up(n, b) for n, b in blobs]) if blobs else demo_data(); cl, lg = clean(raw); return raw, cl, lg
@st.cache_data(show_spinner=False)
def read_docs(blobs): return {n: read_doc(_Up(n, b)) for n, b in blobs}

# ============ AUTO INSIGHTS (any dataset, no column names needed) ============
def insights(c):
    out = []
    for s_, df in c.items():
        out.append(f"**{s_}**: {len(df):,} rows × {df.shape[1]} columns")
        bad = [f"{k} ({v:.0%})" for k, v in df.isna().mean().items() if v > .2]
        if bad: out.append(f"- Heavily incomplete columns: {', '.join(bad[:5])}")
        for k in [k for k in df.columns if pd.api.types.is_numeric_dtype(df[k]) and df[k].nunique() > 10][:8]:
            o = int((((df[k] - df[k].mean()) / (df[k].std() or 1)).abs() > 3).sum())
            if o: out.append(f"- {k}: {o} extreme outliers (>3σ) - entry errors or genuine spikes, worth checking")
        for k in [k for k in df.columns if 2 <= df[k].nunique() <= 15 and not pd.api.types.is_numeric_dtype(df[k])][:6]:
            vc = df[k].value_counts(normalize=True)
            if vc.iloc[0] > .4: out.append(f"- {k}: '{vc.index[0]}' dominates with {vc.iloc[0]:.0%} of rows")
        num = df.select_dtypes("number").loc[:, lambda d: d.nunique() > 10]
        if num.shape[1] > 1:
            cm = num.corr().abs().where(~np.eye(num.shape[1], dtype=bool)).stack()
            if len(cm) and cm.max() > .5: a, b = cm.idxmax(); out.append(f"- Strongest relationship: {a} vs {b} (|r| = {cm.max():.2f}) - correlation, not proof of cause")
    return out

DEMO_DOCS = {"SOP_Operations.txt": "1. Driver Shift Policy\nDrivers shall not exceed 9 hours per shift. Standard Operating Procedure.\n2. Route Speed\nVehicles must not exceed 40 km/h in zone A routes. Standard Operating Procedure.",
             "Executive_Override.txt": "1. Night Shift Directive\nExecutive Override: drivers may extend to 11 hours per shift during hub outages.\n2. Route Speed\nExecutive Override: speed limit in zone A is 30 km/h during monsoon.",
             "Provisional_Notes.txt": "1. Draft Speed Note\nProvisional note: zone A speed limit may rise to 50 km/h. Not yet approved."}
TIERS = [("Executive", 4, r"executive|override|directive|ceo|constraint"), ("Unverified", 1, r"\bunverified|not approved|unapproved|rejected|blocked|suspended"), ("Provisional", 2, r"provisional|draft|conditional|pending"), ("SOP", 3, r"standard operating|sop|policy|confirmed|approved")]
TOPICS = {"speed": r"speed|km/h", "shift": r"shift|hours", "return": r"return|refund", "hub": r"hub|outage", "route": r"route|vehicle|restrict", "ontime": r"on-?time|sla|promis", "supplier": r"supplier|approv|lead time"}
def ref_docs(c):
    """Reference sheets (routes, service rules, supplier rules) become policy passages; status columns drive the tier."""
    ok = lambda v: v not in ("", "UNKNOWN", "nan", "NaT")
    return {f"{s_}.sheet": "\n".join(" | ".join(f"{k}: {v}" for k, v in zip(df.columns, r) if ok(v)) for r in df.fillna("").astype(str).itertuples(index=False)) for s_, df in c.items() if re.search(r"route|_ref|reference|suppl", s_, re.I)}
def read_doc(f):
    n = f.name.lower()
    if n.endswith(".pdf"):
        from pypdf import PdfReader
        return "\n".join(p.extract_text() or "" for p in PdfReader(f).pages)
    if n.endswith(".docx"):
        import zipfile; x = zipfile.ZipFile(f).read("word/document.xml").decode("utf8", "ignore")
        return re.sub(r"<[^>]+>", "", re.sub(r"</w:p>|<w:br/>", "\n", x))
    return f.read().decode("utf-8", "ignore")
def parse_docs(docs):
    """Text docs: tier from file name / line keywords. Reference SHEETS: tier comes only from the row's status field
    (Confirmed/Approved = SOP, Provisional/Conditional = Provisional, Unverified/Not approved = Unverified) - never from free-text notes."""
    secs, title = [], "Preamble"; tierof = lambda txt, d: next(((n, w) for n, w, rx in TIERS if re.search(rx, txt, re.I)), d)
    for dn, txt in docs.items():
        is_sheet = dn.endswith(".sheet"); tier = ("SOP", 3) if is_sheet else tierof(dn, ("SOP", 2))
        for ln in txt.splitlines():
            if not is_sheet and re.match(r"^(\d+[.)]|#+|section)\s*\S", ln, re.I) and len(ln) < 70: title = ln.strip(); continue
            if not ln.strip(): continue
            if is_sheet: st_ = re.search(r"status[^:|]*:\s*([^|]+)", ln, re.I); t2 = tierof(st_.group(1), tier) if st_ else tier; title = ln.split("|")[0].strip()
            else: t2 = tierof(ln, tier)
            secs.append(dict(doc=dn, section=title, text=ln.strip(), tier=t2[0], w=t2[1]))
    return pd.DataFrame(secs, columns=["doc", "section", "text", "tier", "w"])
def ref_audit(secs):
    """Reference-sheet entries that are NOT confirmed: shown to management, never used as hard constraints."""
    x = secs[secs.doc.str.endswith(".sheet") & secs.tier.isin(["Provisional", "Unverified"])]
    return x[["doc", "section", "tier", "text"]].rename(columns={"doc": "Source", "section": "Entry", "tier": "Reliability", "text": "Details"})
def find_conflicts(secs):
    rows = []; full = secs; secs = secs[~secs.doc.str.endswith(".sheet")]
    for tp, rx in TOPICS.items():
        g = secs[secs.text.str.contains(rx, case=False, regex=True)]
        g = g.assign(val=g.text.str.extract(r"(\d+(?:\.\d+)?)\s*(?:km/h|hours|hrs|min(?:ute)?s?|%|days|units|km)", flags=re.I)[0]).dropna(subset=["val"])
        if g.val.nunique() > 1:
            win = g.sort_values("w", ascending=False, kind="stable").iloc[0]
            rows.append(dict(Topic=tp, Competing="; ".join(f"{r.val} ({r.tier}, {r.doc})" for r in g.itertuples()), Winner=f"{win.val} - {win.tier} ({win.doc} / {win.section})", Rule="Executive > SOP > Provisional"))
    # entity cross-check: a text document quotes numbers for a keyed record (RTE-03, SUP-02 ...) that disagree with the reference sheet row
    allp = full.copy(); sheet = allp[allp.doc.str.endswith(".sheet")]; text = allp[~allp.doc.str.endswith(".sheet")]; keyrx = r"\b[A-Z]{2,4}-\d{1,3}\b"; num = lambda t: {f"{float(x):g}" for x in re.findall(r"\d+(?:\.\d+)?", re.sub(keyrx, "", t))}
    for r in text.itertuples():
        for k in sorted(set(re.findall(keyrx, r.text))):
            hit = sheet[sheet.text.str.contains(rf"\b{re.escape(k)}\b", regex=True)]
            if hit.empty: continue
            sr = hit.iloc[0]; dn, sn = num(r.text), num(sr.text)
            if dn and (dn - sn):
                win = r if r.w > sr.w else sr; rows.append(dict(Topic=k, Competing=f"{', '.join(sorted(dn - sn))} ({r.tier}, {r.doc}) vs sheet row ({sr.tier}, {sr.doc})", Winner=f"{win.tier} source ({win.doc})" + ("" if r.w > sr.w else " - sheet record kept; document value needs confirmation"), Rule="Executive > SOP > Provisional"))
    return pd.DataFrame(rows, columns=["Topic", "Competing", "Winner", "Rule"])
def answer(q, secs, c, conf):
    stem = lambda t: {w[:5] for w in re.findall(r"[a-z0-9/]+", t.lower()) if w not in {"the", "a", "is", "what", "of", "in", "for", "and", "to", "at", "which", "have", "has", "or"}}; tok = stem(q)
    if secs.empty: return "No policy documents loaded.", None
    secs = secs.assign(score=secs.text.map(lambda t: len(tok & stem(t))) + secs.w * .1); top = secs[secs.score >= 1.1].sort_values("score", ascending=False).head(3)
    if top.empty: return "No supporting passage found in the loaded policy documents.", None
    out = "**Evidence-backed answer**\n\n" + "\n".join(f"- {r.text}  \n  _Source: {r.doc} › {r.section} [{r.tier}]_" for r in top.itertuples())
    weak = top[top.tier.isin(["Provisional", "Unverified"])]
    if len(weak): out += f"\n\n⚠️ **Reliability caveat:** {len(weak)} of these passage(s) come from **{', '.join(sorted(set(weak.tier)))}** sources - verify before applying as a hard rule."
    for r in conf.itertuples():
        if re.search(TOPICS.get(r.Topic, re.escape(r.Topic)), q, re.I): out += f"\n\n⚠️ **Conflict on '{r.Topic}'** - resolved by hierarchy: **{r.Winner}**."
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
HERE = os.path.dirname(os.path.abspath(__file__)); SAMPLE = os.path.join(HERE, "sample_data")
s_data = sorted(glob.glob(os.path.join(SAMPLE, "*War_Room*.xlsx"))); s_cards = sorted(glob.glob(os.path.join(SAMPLE, "*_C[0-9][0-9].docx")))
_blob = lambda files: tuple((f.name, f.getvalue()) for f in files)
_pblob = lambda paths: tuple((os.path.basename(p), open(p, "rb").read()) for p in paths)

with st.sidebar:
    st.markdown(f"<div style='font-size:20px;font-weight:800;color:#F1F5F9'>🌙 Night Shift</div><div style='color:{MUTED};font-size:12px;margin-bottom:10px'>BUDS Round 2.5 · Decision Console</div>", unsafe_allow_html=True)
    with st.expander("📁 Data feeds", expanded=True):
        up = st.file_uploader("CSV / Excel (multi-sheet OK)", ["csv", "xlsx", "xls"], accept_multiple_files=True)
        use_sample = st.checkbox("Load bundled Team-17 sample + C02/C08/C11 cards", value=False, disabled=not s_data, help="One-click demo using the files shipped in sample_data/") if s_data else False
        st.caption("No upload = built-in dirty demo data." if not use_sample else "Using bundled sample workbook.")
    with st.expander("📚 Policy docs & judge cards", expanded=True):
        dups = st.file_uploader("Policy / reference docs (PDF, TXT, DOCX)", ["pdf", "txt", "docx"], accept_multiple_files=True)
        cards_up = st.file_uploader("📩 Constraint cards from judges", ["pdf", "txt", "docx"], accept_multiple_files=True)
    with st.expander("🛠 Scenario & AI options"):
        drop = st.multiselect("Simulate missing feed (C08)", list(GROUPS), help="Treat a whole data source as unavailable for ALL tabs")
        key = st.text_input("Optional Anthropic API key (richer chat answers)", type="password", value=os.environ.get("ANTHROPIC_API_KEY", ""))

blobs = _blob(up) if up else (_pblob(s_data) if use_sample else ())
raw, cleaned, log = pipeline(blobs); auto = detect(cleaned)
with st.sidebar.expander("🧭 Column mapping (fix if auto-detect is wrong)"):
    opts = ["(none)"] + [f"{s} :: {x}" for s, d in cleaned.items() for x in d.columns]; m = {}
    for r in ROLES:
        cur = f"{auto[r][0]} :: {auto[r][1]}" if auto[r] else "(none)"; ch = st.selectbox(r, opts, index=opts.index(cur), key=f"map_{r}_{len(opts)}_{hash(tuple(opts)) % 9999}")
        m[r] = tuple(ch.split(" :: ", 1)) if ch != "(none)" else None
base = kpis(cleaned, m, drop); sim0 = simulate(base, NEUTRAL); hs = health(raw, log)
card_blobs = _blob(cards_up) if cards_up else (_pblob(s_cards) if use_sample and s_cards else ())
cards = {k: v.strip() for k, v in read_docs(card_blobs).items()} if card_blobs else {}; docs = dict(read_docs(_blob(dups))) if dups else {}
docs.update(ref_docs(cleaned)); docs.update({f"Executive_Constraint_{k}": v for k, v in cards.items()}); docs = docs or DEMO_DOCS
for k_, d_ in (("history", []), ("stack", [])): st.session_state.setdefault(k_, d_)
generic = len(base["missing"]) >= 4
secs = parse_docs(docs); conflicts = find_conflicts(secs); recs = generic_recs(cleaned) if generic else recommend(base, cleaned, m)

# ---------- hero ----------
n_rows = sum(len(d) for d in cleaned.values()); hcls = "ok" if hs["score"] >= 90 else "warn" if hs["score"] >= 75 else "bad"
mode = "Live data" if blobs else "Demo data"
st.markdown(f"""<div class="hero"><h1>Bengaluru Operations · Night Shift Console</h1>
<p>Upload anything → clean it with a full audit trail → KPIs, evidence-based recommendations, plain-English Q&A, policy reasoning and live what-if simulation.</p>
<span class="pill cy">● {mode}</span><span class="pill">{len(cleaned)} sheets · {n_rows:,} rows</span><span class="pill {hcls}">Data health {hs['score']:.0f}/100</span>
<span class="pill">{len(log)} cleaning steps logged</span><span class="pill {'cy' if generic else 'warn' if base['surrogate'] or base['missing'] else 'ok'}">{'🧭 Generic dataset mode' if generic else '⚠ ' + str(len(base['surrogate'])) + ' surrogate KPI(s)' if base['surrogate'] else 'All KPIs from real data'}</span>
{''.join(f'<span class="pill bad">Feed off: {g}</span>' for g in drop)}</div>""", unsafe_allow_html=True)

T = st.tabs(["📊 Command Center", "🧪 Data Health", "💬 Ask Your Data", "📚 Policy Q&A", "🎛️ What-If", "⚡ Constraint Cards", "🔎 Auto Insights", "📘 Methodology"])

def gauge(title, val, lo_bad, lo_ok, h=190):
    f = go.Figure(go.Indicator(mode="gauge+number", value=val, number=dict(suffix="%", font=dict(size=28, color="#F8FAFC")), title=dict(text=title, font=dict(size=13, color=MUTED)),
        gauge=dict(axis=dict(range=[0, 100], tickcolor="#334155"), bar=dict(color=CYAN, thickness=.28), bgcolor="#0b1322", borderwidth=0,
                   steps=[dict(range=[0, lo_bad], color="#3b1a26"), dict(range=[lo_bad, lo_ok], color="#3a3018"), dict(range=[lo_ok, 100], color="#10302a")])))
    return style(f, h).update_layout(margin=dict(l=20, r=20, t=40, b=0))

# ============ 1 COMMAND CENTER ============
with T[0]:
    if 3 <= len(base["missing"]) < 4: st.warning("Few operational columns were recognised. Use the sidebar **Column mapping**, or open **Auto Insights** - it works on any data.")
    for g in ([] if generic else base["missing"]): st.error(f"Missing feed: {g} — fallback surrogates active", icon="🚫")
    items = []
    if generic:
        st.info("🧭 **Generic dataset mode** — no operations feeds (orders, fleet, inventory, returns, suppliers) were recognised, so operations KPIs are replaced by dataset-agnostic KPIs and recommendations. Ask Your Data, Auto Insights and What-If still work; use the sidebar **Column mapping** to teach the app your columns.")
        gk = generic_kpis(cleaned)
        for col, (t, v, sub) in zip(st.columns(len(gk)), gk): col.markdown(kpi(t, v, CYAN, sub), unsafe_allow_html=True)
        section("🎯 Recommendations, with evidence")
        for r_ in recs: insight(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", r_), WARN if "risk" in r_ or "outlier" in r_ else CYAN)
    else:
        sc = lambda k: "surrogate" if k in base["surrogate"] else ""; dfc = sim0["Deficit (cost-rev)"]
        items = [("On-Time Delivery", f"{base['otd']:.1f}%", "OTD", OK if base["otd"] >= 85 else BAD if base["otd"] < 60 else WARN, "target ≥ 85%"), ("Fleet Utilisation", f"{base['util']:.1f}%", "Util", CYAN, "sweet spot 65–90%"),
                 ("Return Rate", f"{base['ret']:.1f}%", "Return", OK if base["ret"] < 8 else BAD, "alert > 8%"), ("Holding Cost / mo", inr(base["hold"]), "Holding", VIOLET, "24% annual carrying"),
                 ("Supplier Score", f"{base['sup']:.1f}", "Supplier", OK if base["sup"] >= 80 else WARN, "target ≥ 80"), ("Operating Deficit", inr(dfc), "Fleet cost", BAD if dfc > 0 else OK, "cost − revenue")]
        for col, (t, v, k, cl, sub) in zip(st.columns(6), items): col.markdown(kpi(t, v, cl, sub, sc(k)), unsafe_allow_html=True)
        ex = [(t, v, sub) for t, v, sub in [("Inventory Accuracy", base["inv_acc"], "system = physical"), ("Fleet Availability", base["fleet_avail"], "operable vehicles"), ("Suppliers Approved", base["sup_ok"], "fully approved")] if v is not None]
        if ex:
            st.write(""); 
            for col, (t, v, sub) in zip(st.columns(max(3, len(ex))), ex): col.markdown(kpi(t, f"{v:.1f}%", OK if v >= 80 else WARN if v >= 50 else BAD, sub), unsafe_allow_html=True)
        g1, g2, g3 = st.columns(3)
        with g1: show(gauge("On-time delivery", base["otd"], 60, 85), 190)
        with g2: show(gauge("Fleet utilisation", base["util"], 50, 65), 190)
        with g3: show(gauge("Data health score", hs["score"], 75, 90), 190)
        st.info(assessment(base))
        section("🎯 Recommendations, with evidence")
        for r_ in recs:
            ac = BAD if re.search(r"deficit|unreliable|Raise on-time|Fleet is stretched", r_) else WARN if re.search(r"⚠|under-used|weak|high|low|Only", r_) else CYAN
            insight(re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", r_), ac)
        section("📈 Operational view")
        a, b2 = st.columns(2); h = hub_otd(cleaned, m)
        if h is not None: 
            with a: show(go.Figure(go.Bar(x=h.index, y=h.values, marker_color=[OK if v >= 85 else WARN if v >= 50 else BAD for v in h.values], text=[f"{v:.0f}%" for v in h.values], textposition="outside")).update_layout(title="On-time delivery % by hub", yaxis_range=[0, 105]))
        sm = status_mix(cleaned, m)
        if sm is not None:
            with b2: show(go.Figure(go.Pie(labels=sm.index.astype(str), values=sm.values, hole=.62, marker=dict(colors=[CYAN, VIOLET, WARN, BAD, OK, MUTED]), textinfo="none", hovertemplate="%{label}: %{value} (%{percent})")).update_layout(title="Order status mix", legend=dict(orientation="v", x=1.0, y=.5)))
        tr = daily_trend(cleaned, m)
        if tr is not None:
            f = go.Figure(go.Bar(x=tr.index, y=tr["Orders"], name="Orders", marker_color=CYAN, opacity=.75))
            if "OTD %" in tr: f.add_trace(go.Scatter(x=tr.index, y=tr["OTD %"], name="OTD %", yaxis="y2", line=dict(color=WARN, width=2)))
            show(f.update_layout(title="Daily orders & on-time %", yaxis2=dict(overlaying="y", side="right", range=[0, 100], gridcolor="rgba(0,0,0,0)")), 300)
    section("🧩 Explore any column")
    sh = st.selectbox("Sheet", list(cleaned), key="ex_s"); d = cleaned[sh]; nums = [x for x in d.columns if pd.api.types.is_numeric_dtype(d[x])]; cats = [x for x in d.columns if d[x].nunique() <= 30]
    if nums and cats:
        a, b2, c3 = st.columns(3); y, xg, ag = a.selectbox("Measure", nums), b2.selectbox("Group by", cats), c3.selectbox("Aggregate", ["sum", "mean", "count", "max"])
        g = d.groupby(xg)[y].agg(ag).reset_index(); show(go.Figure(go.Bar(x=g[xg].astype(str), y=g[y], marker_color=CYAN)).update_layout(title=f"{ag.title()} of {y} by {xg}"))
    st.write(""); d1, d2 = st.columns(2)
    rep = "# BUDS Night Shift report\n\n" + assessment(base) + "\n\n## KPIs\n" + "\n".join(f"- {t}: {v}" for t, v, *_ in (items or generic_kpis(cleaned))) + "\n\n## Recommendations\n" + "\n".join(f"- {x}" for x in recs)
    d1.download_button("⬇ Management report (.md)", rep, "report.md", use_container_width=True)
    d2.download_button("⬇ Full audit pack (.xlsx: KPIs, log, cleaned data, scenarios)", build_xlsx(base, sim0, recs, log, cleaned, st.session_state.history, conflicts, hs), "night_shift_audit_pack.xlsx", use_container_width=True)

# ============ 2 DATA HEALTH ============
with T[1]:
    section("🧪 Data Health & Audit Trail"); st.caption("Raw uploads are preserved unchanged; every transformation applies only to the cleaned copy.")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi("Health score", f"{hs['score']:.0f}/100", OK if hs["score"] >= 90 else WARN if hs["score"] >= 75 else BAD, "mean of 3 dimensions"), unsafe_allow_html=True)
    c2.markdown(kpi("Completeness", f"{hs['completeness']:.1f}%", CYAN, f"{hs['missing']:,} empty cells"), unsafe_allow_html=True)
    c3.markdown(kpi("Uniqueness", f"{hs['uniqueness']:.1f}%", VIOLET, f"{hs['dups']:,} duplicate rows"), unsafe_allow_html=True)
    c4.markdown(kpi("Validity", f"{hs['validity']:.1f}%", WARN, f"{hs['fixed']:,} rows repaired / flagged"), unsafe_allow_html=True)
    st.write(""); st.dataframe(sheet_health(raw, cleaned), width="stretch", hide_index=True)
    L = pd.DataFrame(log); a, b_ = st.columns(2); a.metric("Transformations logged", len(log)); b_.metric("Rows touched", int(L["Rows affected"].sum()) if len(L) else 0)
    if len(L):
        agg = L.groupby("Transformation")["Rows affected"].sum().sort_values()
        show(go.Figure(go.Bar(x=agg.values, y=agg.index, orientation="h", marker_color=VIOLET)).update_layout(title="Rows affected by transformation type"), max(260, 28 * len(agg) + 80))
    st.markdown("**Audit log**"); st.dataframe(L, width="stretch", hide_index=True)
    with st.expander("Data structure profile (types, missing, unique)"): st.dataframe(profile(raw), width="stretch", hide_index=True)
    section("🔬 Raw vs cleaned"); sh = st.selectbox("Compare sheet", list(raw)); c1, c2 = st.columns(2); c1.markdown("**Raw (immutable)**"); c1.dataframe(raw[sh].head(25)); c2.markdown("**Cleaned**"); c2.dataframe(cleaned[sh].head(25))
    st.download_button(f"⬇ Cleaned '{sh}' (CSV)", cleaned[sh].to_csv(index=False), f"clean_{sh}.csv")

# ============ 3 ASK YOUR DATA ============
with T[2]:
    section("💬 Ask your data in plain English"); st.caption("Understands filters (e.g. 'H02', 'Delivered'), 'top 3', average / count / max / min. Add an API key in the sidebar for richer narrative answers.")
    st.session_state.setdefault("q_data", "total delivery fee by hub")
    ex = ["total delivery fee by hub", "average actual cycle time by hub", "how many orders by order status", "top 3 order value by product category", "average distance for H02"]
    for col, e in zip(st.columns(len(ex)), ex): col.button(e, key=f"chip_{e}", on_click=lambda e=e: st.session_state.update(q_data=e), use_container_width=True)
    q = st.text_input("Question", key="q_data")
    if q:
        t, msg, code = ask(q, cleaned); st.markdown(msg)
        if t is not None:
            a, b2 = st.columns([1, 1.3]); a.dataframe(t, width="stretch", hide_index=True); a.caption(f"Calculation used: `{code}`")
            if len(t) > 1 and t.shape[1] > 1: 
                with b2: show(go.Figure(go.Bar(x=t.iloc[:, 0].astype(str), y=t.iloc[:, 1], marker_color=CYAN)), 300)
        if key:
            try: st.markdown("**AI analyst answer**"); st.markdown(llm(q, assessment(base) + "\n" + str(base) + "\n" + "\n".join(f"{s}:\n{d.head(15).to_string()}" for s, d in cleaned.items()), key))
            except Exception as e: st.warning(f"AI answer unavailable: {e}")

# ============ 4 POLICY Q&A ============
with T[3]:
    section("📚 Rule hierarchy: Executive Override > SOP > Provisional"); st.caption(f"{len(secs)} passages parsed from {len(docs)} document(s) — includes route/service reference sheets and any uploaded constraint cards.")
    if len(conflicts): st.warning("Conflicting rules detected and resolved"); st.dataframe(conflicts, width="stretch", hide_index=True)
    else: st.success("No conflicting numeric rules detected.")
    aud = ref_audit(secs)
    if len(aud): st.warning(f"{len(aud)} reference entr{'y is' if len(aud) == 1 else 'ies are'} provisional / unverified — surfaced for review, **not** used as hard constraints"); st.dataframe(aud, width="stretch", hide_index=True)
    q = st.text_input("Ask a policy / data question", "What is the speed limit in zone A and the driver shift limit at Hub 3?" if docs is DEMO_DOCS else "Which routes have vehicle restrictions and is the restriction status confirmed?", key="q_pol")
    if q:
        ans, rows = answer(q, secs, cleaned, conflicts); st.markdown(ans)
        if rows is not None: st.markdown("**Dataset evidence (cleaned rows)**"); st.dataframe(rows)
    with st.expander("Parsed policy passages"): st.dataframe(secs[["doc", "section", "tier", "text"]], hide_index=True)

# ============ 5 WHAT-IF ============
with T[4]:
    section("🎛️ Strategy sandbox"); c = st.columns(4)
    p = dict(demand=c[0].slider("Demand spike %", -30, 100, 0), fuel=c[1].slider("Fuel / fleet cost ×", .8, 2.0, 1.0, .05), drivers=c[2].slider("Driver availability %", 30, 100, 100), delay=c[3].slider("Warehouse delay (min)", 0, 120, 0))
    res = simulate(base, p); k = st.columns(5)
    k[0].metric("Margin %", f"{res['Margin %']:.1f}", f"{res['Margin %'] - sim0['Margin %']:+.1f} pts"); k[1].metric("Deficit", inr(res["Deficit (cost-rev)"]), inr(res["Deficit (cost-rev)"] - sim0["Deficit (cost-rev)"]), delta_color="inverse")
    k[2].metric("Delay risk %", f"{res['Delay Risk %']:.0f}", f"{res['Delay Risk %'] - sim0['Delay Risk %']:+.0f}", delta_color="inverse"); k[3].metric("Demand served %", f"{res['Demand Served %']:.0f}")
    sat = (p["drivers"] / 100) / max(base["util"] / 100, .01) * 100 - 100; k[4].metric("Fleet saturates at demand", f"{sat:+.0f}%", help="Demand change at which fleet load reaches 100% given current driver availability")
    a, b2 = st.columns(2)
    with a: show(go.Figure([go.Bar(name="Baseline", x=["Revenue", "Op. Cost"], y=[sim0["Revenue"], sim0["Op. Cost"]], marker_color=BORDER), go.Bar(name="Scenario", x=["Revenue", "Op. Cost"], y=[res["Revenue"], res["Op. Cost"]], marker_color=CYAN)]).update_layout(title="Revenue vs cost"))
    sn = sensitivity(base).sort_values("Δ Deficit (₹)")
    with b2: show(go.Figure(go.Bar(x=sn["Δ Deficit (₹)"], y=sn["Shock"], orientation="h", marker_color=[BAD if v > 0 else OK for v in sn["Δ Deficit (₹)"]])).update_layout(title="Sensitivity: which shock hurts most? (Δ deficit)"))
    st.caption("Sensitivity applies one standard shock at a time to the live baseline, so you can rank which risk to hedge first. A zero bar means that lever has no effect at current utilisation (spare fleet capacity absorbs it).")

# ============ 6 CONSTRAINT CARDS ============
def render(text, cp, cdrop, notes):
    ab = kpis(cleaned, m, tuple(set(drop) | set(cdrop))); before, after = simulate(base, NEUTRAL), simulate(ab, cp); diff = pd.DataFrame({"Before": before, "After": after}); diff["Δ"] = diff.After - diff.Before
    moved = cp != NEUTRAL or bool(cdrop) or bool(drop); sev, scol = severity(before, after, moved); st.markdown(f"#### 🔴 {card_title(text)}  <span class='pill' style='color:{scol};border-color:{scol}'>SEVERITY: {sev}</span>", unsafe_allow_html=True)
    a, b2 = st.columns([1, 1.2]); a.dataframe(diff.style.format("{:,.1f}"), width="stretch")
    ks = ["Margin %", "Delay Risk %", "Fleet Load %"]
    with b2: show(go.Figure([go.Bar(name="Before", x=ks, y=[before[x] for x in ks], marker_color=OK), go.Bar(name="After", x=ks, y=[after[x] for x in ks], marker_color=BAD)]), 280)
    if cdrop or drop: st.warning(assessment(ab))
    l, r = st.columns(2); l.markdown("**Mitigation action plan**"); [l.markdown(f"{i}. {s_}") for i, s_ in enumerate(plan(cp, cdrop), 1)]
    r.markdown("**Audit log — assumptions**"); [r.markdown(f"- {s_}") for s_ in notes + [f"Final lever settings: {cp}", "Raw data unchanged; scenario computed on a cleaned copy"]]
    st.session_state.history.append({"Card": card_title(text)[:60], "Severity": sev, "Margin % after": round(after["Margin %"], 1), "Delay risk % after": round(after["Delay Risk %"]), "Levers": str({k_: round(v, 2) for k_, v in cp.items()}), "Feeds lost": ", ".join(cdrop) or "-"})

with T[5]:
    section("⚡ Live constraint card simulator")
    st.caption("New card from the judges? Upload it in the sidebar (DOCX / PDF / TXT), pick a preset or type it. The app proposes lever settings — adjust them, then simulate. **Stack** several cards to see their combined impact.")
    srcs = ["Preset"] + [f"Uploaded: {k}" for k in cards] + ["Type my own"]; src = st.selectbox("Card source", srcs, index=1 if cards else 0)
    if src == "Preset": pre = st.selectbox("Preset card", list(PRESETS)); text = PRESETS[pre]
    elif src == "Type my own": text = st.text_area("Describe the new condition", "Fuel +40% and drivers down 30%, hub outage", key="card_txt")
    else: text = cards[src.replace("Uploaded: ", "")]; st.text_area("Card text (read-only)", text[:1500], disabled=True)
    cp0, cdrop0, notes0 = parse_card(text); k_ = abs(hash(text)) % 10**6; cl = st.columns(4)
    cp = dict(demand=cl[0].slider("Demand %", -50, 150, int(max(-50, min(150, cp0["demand"]))), key=f"d{k_}"), fuel=cl[1].slider("Fuel / fleet cost ×", .5, 3.0, float(max(.5, min(3.0, cp0["fuel"]))), .05, key=f"f{k_}"),
              drivers=cl[2].slider("Driver availability %", 5, 100, int(max(5, min(100, cp0["drivers"]))), key=f"v{k_}"), delay=cl[3].slider("Warehouse delay (min)", 0, 240, int(min(240, cp0["delay"])), key=f"y{k_}"))
    cdrop = st.multiselect("Data feeds unavailable under this card", list(GROUPS), default=cdrop0, key=f"g{k_}")
    if re.search(r"data source|source .*unavailable|becomes unavailable", text, re.I) and not cdrop: st.info("This card says a data source is unavailable but doesn't name it — pick the affected feed above and the KPIs will fall back to surrogates with an assessment note.")
    dq = [w for w in ("duplicate", "inconsisten", "invalid date", "missing value", "conflicting status") if w in text.lower()]
    if dq: st.info(f"Data-integrity constraint detected ({', '.join(dq)}): the cleaning engine already handles this — see **Data Health**. Raw data stays untouched.")
    if re.search(r"provisional|reference documents?|documents may disagree|route restriction|supplier rules", text, re.I): st.info("Document/rule constraint detected — see the **Policy Q&A** tab for hierarchy-based conflict resolution.")
    b1, b2, b3 = st.columns([1.3, 1, 1]); notes = notes0 + ([f"Feeds treated as unavailable: {', '.join(cdrop)}"] if cdrop else [])
    if b1.button("⚡ Simulate this card", type="primary", use_container_width=True): render(text, cp, cdrop, notes)
    if b2.button("➕ Add to stack", use_container_width=True): st.session_state.stack.append(dict(text=card_title(text)[:80], p=cp, drop=list(cdrop))); st.toast("Card added to stack")
    if b3.button("🗑 Clear stack", use_container_width=True): st.session_state.stack = []
    if st.session_state.stack:
        section(f"🧱 Stacked cards ({len(st.session_state.stack)}) — combined impact")
        st.dataframe(pd.DataFrame([{"Card": s["text"], "Demand %": s["p"]["demand"], "Fuel ×": round(s["p"]["fuel"], 2), "Drivers %": s["p"]["drivers"], "Delay min": s["p"]["delay"], "Feeds lost": ", ".join(s["drop"]) or "-"} for s in st.session_state.stack]), width="stretch", hide_index=True)
        if st.button("⚡ Simulate combined stack", type="primary"): sp, sd = combine(st.session_state.stack); render("COMBINED: " + " + ".join(s["text"][:25] for s in st.session_state.stack), sp, list(sd), ["Cards combined: demand and delay add, fuel multiplies, driver availability multiplies"])
    if st.session_state.history:
        with st.expander(f"Constraint card history ({len(st.session_state.history)})"): st.dataframe(pd.DataFrame(st.session_state.history), width="stretch", hide_index=True)

# ============ 7 AUTO INSIGHTS ============
with T[6]:
    section("🔎 Auto insights — works on any dataset"); st.caption("No column names needed: automatic checks for gaps, outliers, dominant categories and relationships.")
    for x in insights(cleaned): insight(x.lstrip("- ").replace("**", ""), VIOLET if x.startswith("-") else CYAN)
    sh = st.selectbox("Sheet to chart", list(cleaned), key="ai_s"); d = cleaned[sh]; nm = [x for x in d.columns if pd.api.types.is_numeric_dtype(d[x])]
    if nm:
        col = st.selectbox("Distribution of", nm, key="ai_c"); show(go.Figure(go.Histogram(x=d[col], marker_color=CYAN)).update_layout(title=f"Distribution of {col}"))
    if len(nm) > 2:
        cm = d[nm[:12]].corr(); show(go.Figure(go.Heatmap(z=cm.values, x=cm.columns, y=cm.index, colorscale="Tealgrn", zmin=-1, zmax=1)).update_layout(title="Correlation map"), 430)

# ============ 8 METHODOLOGY ============
with T[7]:
    section("📘 How the numbers are calculated")
    st.markdown("""- **On-time delivery** = delivered orders where actual cycle minutes ≤ promised minutes (or delivery date ≤ promised date). Rows with missing/invalid times are excluded from the denominator, never guessed.
- **Fleet utilisation** = used hours ÷ capacity hours, or (if absent) Assigned ÷ operable (Assigned + Available + Standby) vehicle records. **Fleet availability** = operable ÷ all records. **Fleet cost** = order variable cost, else Σ km × cost/km, else runs × average route base cost.
- **Return rate** = share of orders flagged returned (or return requests ÷ orders when no flag exists). **Holding cost / month** = Σ(qty × unit cost) × 24% ÷ 12 (assumed annual carrying rate).
- **Supplier score** = 60% on-time % + 40% × (100 − defect %); if absent, approval score (Approved 100, Conditional 60, other 30). **Operating deficit** = operating cost − revenue (revenue = delivery fees; cost = variable + hub fixed cost × days + holding + 30% of revenue lost on returns).
- **What-if**: load = demand × utilisation ÷ driver availability; if load > 100% only 1/load of demand is served. Delay risk = (100 − OTD) + 0.35 × delay min + 60 × overload. **Saturation point** = driver availability ÷ utilisation − 1.
- **Sensitivity** applies one standard shock (demand +20%, fuel ×1.2, drivers −20%, +30 min delay) to the live baseline and ranks the change in deficit.
- **Stacked cards**: demand % and delay minutes add; fuel multiplier and driver availability multiply; lost feeds are unioned.
- **Policy conflicts**: highest tier wins — Executive Override > SOP > Provisional.
- **Data health score** = mean of Completeness (1 − empty cells ÷ cells), Uniqueness (1 − exact duplicate rows ÷ rows) and Validity (1 − repaired/flagged rows ÷ rows), measured on the raw upload.
- **Cleaning**: duplicates dropped, hub IDs standardised, bad dates blanked, missing numbers → median (except time/distance/rating fields, which stay blank), missing text → UNKNOWN, 'Delivered' without delivery date → 'In Transit'. **Raw data is never modified.**""")
    st.warning("**Limitations**: surrogate values (85% OTD, 6% returns, 70% utilisation, 80 supplier) are benchmarks, not facts; the simulator is a simplified linear model; column auto-detection can be corrected in the sidebar mapping; policy reading is keyword-based; figures are indicative (±15%) when feeds are missing; NL Q&A without an API key is rule-based and may miss complex questions.")
