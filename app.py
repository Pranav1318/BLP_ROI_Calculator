"""
Isarwerk Präzision GmbH: what is AP automation worth to the CFO?

An interactive value calculator (fictional case, not a BLP quotation).
Model logic mirrors the Excel file "Isarwerk_AP_ROI_With_Payback.xlsx".

Run locally:   streamlit run app.py
"""

from collections import Counter

import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# Page setup and look
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Isarwerk: what is AP automation worth?",
    page_icon="📈",
    layout="wide",
)

INK, TEAL, TEAL_LIGHT = "#16324A", "#1B8A98", "#7FBFC7"
GREEN, AMBER, CORAL = "#3FA37A", "#F2B84B", "#E8765F"
MUTED, LINE = "#5F7380", "#DCE7EA"

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap');
    html, body, [class*="css"], .stMarkdown, .stMetric {{ font-family: 'DM Sans', sans-serif; }}
    .stApp {{ background: #F7FAFB; }}
    h1, h2, h3 {{ color: {INK}; letter-spacing: -0.01em; }}
    div[data-testid="stMetric"] {{
        background: #FFFFFF; border: 1px solid {LINE}; border-radius: 10px;
        padding: 14px 16px; box-shadow: 0 1px 4px rgba(120,150,160,.15);
    }}
    div[data-testid="stMetricValue"] {{ color: {TEAL}; font-weight: 700; }}
    .verdict {{
        background: #E3F2F4; border-left: 5px solid {TEAL}; border-radius: 6px;
        padding: 14px 18px; color: {INK}; font-size: 1.05rem; margin: 6px 0 14px 0;
    }}
    .verdict.warn {{ background: #FCE8E3; border-left-color: {CORAL}; }}
    .pill {{ display:inline-block; padding:2px 10px; border-radius: 12px; font-size: .8rem; font-weight:700; }}
    .pill.High {{ background:#E3F4EC; color:#2E8560; }}
    .pill.Medium {{ background:#FDF3DC; color:#B7841A; }}
    .pill.Low {{ background:#FCE8E3; color:#C4573F; }}
    table.assump {{ width:100%; border-collapse: collapse; font-size: .93rem; background:#fff; }}
    table.assump th {{ text-align:left; color:{MUTED}; font-weight:500; padding:7px 9px; border-bottom:1px solid {LINE}; }}
    table.assump td {{ padding:7px 9px; border-bottom:1px solid #EEF3F5; color:{INK}; }}
    table.assump td.num, table.assump th.num {{ text-align:right; }}
    table.assump tr.grp td {{ background:#EAF4F5; color:{TEAL}; font-weight:700; font-size:.8rem;
        letter-spacing:.06em; text-transform:uppercase; }}
    table.assump a {{ color:{TEAL}; text-decoration:none; }}
    .nochange {{ color:{MUTED}; font-style:italic; }}
    .uc {{ background:#fff; border:1px solid {LINE}; border-radius:12px; padding:14px 16px; height:100%; }}
    .uc.sel {{ border:2px solid {TEAL}; background:#E3F2F4; box-shadow:0 2px 10px rgba(27,138,152,.25); }}
    .uc.dim {{ opacity:.55; }}
    .uc h4 {{ margin:0 0 4px 0; color:{INK}; font-size:1rem; }}
    .uc .big {{ color:{TEAL}; font-size:1.5rem; font-weight:700; line-height:1.1; }}
    .uc .small {{ color:{MUTED}; font-size:.82rem; }}
    .uc .row {{ display:flex; justify-content:space-between; font-size:.88rem; padding:2px 0; color:{INK}; }}
    .chip {{ display:inline-block; padding:3px 11px; border-radius:14px; font-size:.82rem; font-weight:600; margin:3px 4px 3px 0; }}
    .chip.done {{ background:#E3F4EC; color:#2E8560; }}
    .chip.todo {{ background:#FDF3DC; color:#B7841A; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Model (mirrors the Excel workbook)
# ----------------------------------------------------------------------------
USE_CASES = [  # invoice streams modelled in the Excel workbook
    ("Direct material", "Parts and components bought under purchase orders. Plausibly the largest stream in manufacturing."),
    ("Indirect / services", "Indirect goods and purchased services under POs. More accounting review per invoice."),
    ("Freight / logistics", "Many carrier and shipping bills, with checks of rates and shipments."),
    ("No-PO & other", "Costs without a purchase order. Needs GL coding and approvals, so more touch time."),
]

DEFAULTS = {
    "scenario": "Base",
    "time_cut": 60,          # % of manual effort removed
    "share": 50,             # % of freed time that turns into real cost avoided
    "vol_0": 24600, "vol_1": 16000, "vol_2": 12000, "vol_3": 8000,
    "min_0": 10, "min_1": 11, "min_2": 11, "min_3": 15,
    "labour": 45,            # EUR per hour, employer cost
    "prod_hours": 1650,      # productive hours per FTE per year
    "platform": 300,         # BLP platform fee, EUR per month
    "usage": 0.32,           # BLP usage fee, EUR per invoice
    "d_erp": 3, "d_wf": 4, "d_test": 2, "d_go": 2,   # rollout consulting days
    "hrs_day": 8,            # chargeable hours per consulting day
    "hourly": 137.5,         # EUR per consulting hour (ISR published support rate)
}
REVENUE = 320_000_000        # fictional, fixed

SCENARIOS = {  # only these two levers move between scenarios in the Excel model
    "Conservative": {"time_cut": 50, "share": 40},
    "Base": {"time_cut": 60, "share": 50},
    "Optimistic": {"time_cut": 70, "share": 60},
}


def build_params(g):
    """g: function name -> value (session state or defaults)."""
    days = g("d_erp") + g("d_wf") + g("d_test") + g("d_go")
    return {
        "vols": [g(f"vol_{i}") for i in range(4)],
        "mins": [g(f"min_{i}") for i in range(4)],
        "time_cut": g("time_cut") / 100,
        "share": g("share") / 100,
        "labour": g("labour"),
        "prod_hours": g("prod_hours"),
        "platform": g("platform"),
        "usage": g("usage"),
        "days": [g("d_erp"), g("d_wf"), g("d_test"), g("d_go")],
        "hrs_day": g("hrs_day"),
        "hourly": g("hourly"),
        "setup": days * g("hrs_day") * g("hourly"),
    }


def calc(p):
    """Return all model outputs for one set of parameters."""
    invoices = sum(p["vols"])
    hours = sum(v * m for v, m in zip(p["vols"], p["mins"])) / 60
    freed = hours * p["time_cut"]
    cash = freed * p["labour"] * p["share"]
    software = 12 * p["platform"] + invoices * p["usage"]
    cost1 = software + p["setup"]
    net = cash - cost1
    monthly_net = (cash - software) / 12
    payback = p["setup"] / monthly_net if monthly_net > 0 else None
    cover = cost1 / (cash / 12) if cash > 0 else None   # month when savings reach year-1 cost
    roi = net / cost1 if cost1 else 0
    fte = freed / p["prod_hours"]
    denom = hours * p["time_cut"] * p["labour"]
    breakeven_share = cost1 / denom if denom else None
    return dict(invoices=invoices, hours=hours, freed=freed, cash=cash, software=software,
                setup=p["setup"], cost1=cost1, net=net, payback=payback, cover=cover, roi=roi,
                fte=fte, breakeven_share=breakeven_share)


def variant(p, **changes):
    q = dict(p)
    q.update(changes)
    return q


def eur(x):
    return f"€{x:,.0f}"


def eur_k(x):
    return f"€{x / 1000:,.1f}k"


# ----------------------------------------------------------------------------
# Sidebar: inputs
# ----------------------------------------------------------------------------
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)


def apply_scenario():
    sc = st.session_state["scenario"]
    if sc in SCENARIOS:
        st.session_state["time_cut"] = SCENARIOS[sc]["time_cut"]
        st.session_state["share"] = SCENARIOS[sc]["share"]


def mark_custom():
    st.session_state["scenario"] = "Custom"


def reset_all():
    for k, v in DEFAULTS.items():
        st.session_state[k] = v


with st.sidebar:
    st.header("Your assumptions")
    st.radio("Scenario", ["Conservative", "Base", "Optimistic", "Custom"],
             key="scenario", on_change=apply_scenario, horizontal=True,
             help="Scenarios change two levers: time cut and cash share. Everything else stays the same.")
    st.slider("Manual effort removed by automation (%)", 20, 90, step=5, key="time_cut", on_change=mark_custom,
              help="Share of today's manual invoice handling time that automation takes over.")
    st.slider("Freed time that becomes real savings (%)", 0, 100, step=5, key="share", on_change=mark_custom,
              help="Freed hours only save cash if they avoid hiring, overtime or temp cover.")

    with st.expander("Invoice volume and handling time"):
        for i, (name, _) in enumerate(USE_CASES):
            c1, c2 = st.columns(2)
            c1.number_input(f"{name}: invoices / year", min_value=0, step=500, key=f"vol_{i}")
            c2.number_input("Minutes each", min_value=1, max_value=60, key=f"min_{i}")

    with st.expander("Labour cost and capacity"):
        st.slider("Employer cost per hour (€)", 30, 70, key="labour")
        st.slider("Productive hours per FTE per year", 1200, 1800, step=10, key="prod_hours")

    with st.expander("BLP price"):
        st.slider("Platform fee (€ / month)", 0, 1000, step=25, key="platform")
        st.slider("Usage fee (€ / invoice)", 0.0, 1.0, step=0.01, key="usage")

    with st.expander("One-time rollout"):
        r1, r2 = st.columns(2)
        r1.number_input("ERP integration (days)", min_value=0, max_value=30, key="d_erp")
        r2.number_input("Workflow setup (days)", min_value=0, max_value=30, key="d_wf")
        r1.number_input("Testing + training (days)", min_value=0, max_value=30, key="d_test")
        r2.number_input("Go-live help (days)", min_value=0, max_value=30, key="d_go")
        st.slider("Chargeable hours per day", 4, 10, key="hrs_day")
        st.number_input("Consulting rate (€ / hour)", min_value=50.0, max_value=300.0, step=2.5, key="hourly")

    st.button("Reset to base case", on_click=reset_all, width='stretch')
    st.caption("Fictional case. Not a BLP quotation.")

# ----------------------------------------------------------------------------
# Calculations
# ----------------------------------------------------------------------------
P = build_params(lambda k: st.session_state[k])
R = calc(P)
BASE_P = build_params(lambda k: DEFAULTS[k])
BASE_R = calc(BASE_P)

# ----------------------------------------------------------------------------
# Header and verdict
# ----------------------------------------------------------------------------
st.title("What is AP automation worth to Isarwerk?")
st.markdown(
    "Isarwerk Präzision GmbH is a fictional €320m German manufacturer that processes about "
    f"**{R['invoices']:,.0f} supplier invoices a year**. Change the assumptions on the left and see the "
    "year-1 result, what changes for the business, how sure we are, and what happens if we are wrong."
)

if R["net"] > 0 and R["payback"] is not None:
    be = R["breakeven_share"]
    st.markdown(
        f"<div class='verdict'><b>At these settings, automation nets {eur(R['net'])} in year 1</b> and pays back "
        f"the setup in {R['payback']:.1f} months. It still breaks even if only <b>{be * 100:.0f}%</b> of the "
        f"freed time turns into real savings. That is enough to justify a deeper study with real data.</div>",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f"<div class='verdict warn'><b>At these settings, year 1 is {eur(R['net'])}.</b> Savings do not cover the "
        f"software and setup cost. Raise the time cut or cash share, or lower the BLP price, to see what changes the result.</div>",
        unsafe_allow_html=True,
    )

k1, k2, k3, k4 = st.columns(4)
k1.metric("Cash savings per year", eur_k(R["cash"]), f"{(R['cash'] - BASE_R['cash']) / 1000:+,.1f}k vs base")
k2.metric("Year-1 cost (software + setup)", eur_k(R["cost1"]), f"{(R['cost1'] - BASE_R['cost1']) / 1000:+,.1f}k vs base",
          delta_color="inverse")
k3.metric("Year-1 net value", eur_k(R["net"]), f"{(R['net'] - BASE_R['net']) / 1000:+,.1f}k vs base")
k4.metric("Payback on setup", f"{R['payback']:.1f} mo" if R["payback"] else "no payback",
          None if R["payback"] is None or BASE_R["payback"] is None
          else f"{R['payback'] - BASE_R['payback']:+.1f} months vs base", delta_color="inverse")

s1, s2, s3 = st.columns(3)
s1.metric("Manual hours today", f"{R['hours']:,.0f} h")
s2.metric("Hours freed", f"{R['freed']:,.0f} h", f"{R['fte']:.1f} FTE of capacity", delta_color="off")
s3.metric("Year-1 return on cost", f"{R['roi'] * 100:,.0f}%")

# ----------------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["The result", "What changes for the CFO", "Use cases", "All 21 assumptions",
     "What if we are wrong?", "What we need next"]
)

LAYOUT = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
              font=dict(family="DM Sans, sans-serif", color=INK), margin=dict(l=10, r=10, t=40, b=10), height=360)


def style(fig, **kw):
    fig.update_layout(**{**LAYOUT, **kw})


# ---- Tab 1: result -----------------------------------------------------------
with tab1:
    left, right = st.columns(2)

    with left:
        st.subheader("From savings to year-1 net value")
        wf = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute", "relative", "relative", "total"],
            x=["Cash savings", "Software (year 1)", "Setup (one-off)", "Year-1 net"],
            y=[R["cash"], -R["software"], -R["setup"], 0],
            text=[eur_k(R["cash"]), "−" + eur_k(R["software"]), "−" + eur_k(R["setup"]), eur_k(R["net"])],
            textposition="outside",
            increasing=dict(marker=dict(color=TEAL)),
            decreasing=dict(marker=dict(color=CORAL)),
            totals=dict(marker=dict(color=INK)),
            connector=dict(line=dict(color=LINE)),
        ))
        style(wf, showlegend=False, yaxis=dict(visible=False))
        st.plotly_chart(wf, width='stretch')

    with right:
        st.subheader("Three scenarios side by side")
        st.caption("Same volumes, prices and costs as in the sidebar. Only time cut and cash share change.")
        rows = []
        for name, lv in SCENARIOS.items():
            r = calc(variant(P, time_cut=lv["time_cut"] / 100, share=lv["share"] / 100))
            rows.append((name, r))
        cmp_fig = go.Figure(go.Bar(
            x=[n for n, _ in rows], y=[r["net"] for _, r in rows],
            marker_color=[TEAL_LIGHT, TEAL, TEAL_LIGHT],
            text=[f"{eur_k(r['net'])}<br>payback {r['payback']:.1f} mo" if r["payback"] else eur_k(r["net"]) for _, r in rows],
            textposition="outside",
        ))
        style(cmp_fig, showlegend=False, yaxis=dict(visible=False))
        st.plotly_chart(cmp_fig, width='stretch')

# ---- Tab 2: what changes for the CFO ---------------------------------------------
with tab2:
    st.subheader("What changes in operations and in the P&L")
    st.markdown(
        "A side-by-side view of today versus the automated process. Hours and capacity are the operational "
        "change. Cash savings are only the share of freed time that turns into avoided cost."
    )
    hours_b = R["hours"]
    hours_a = hours_b * (1 - P["time_cut"])
    fte_b, fte_a = hours_b / P["prod_hours"], hours_a / P["prod_hours"]
    min_b = hours_b * 60 / R["invoices"]
    min_a = min_b * (1 - P["time_cut"])
    labour_b = hours_b * P["labour"]                         # manual handling labour cost today
    labour_a = hours_a * P["labour"]                         # remaining manual labour
    freed_value = R["freed"] * P["labour"]                   # value of freed hours
    redeployed = freed_value - R["cash"]                     # freed but still paid: capacity, not cash
    cost_a = labour_b - R["cash"] + R["software"]            # AP processing cost, cash view
    cpi_b, cpi_a = labour_b / R["invoices"], cost_a / R["invoices"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Manual AP hours per year", f"{hours_b:,.0f} → {hours_a:,.0f}", f"-{R['freed']:,.0f} h", delta_color="inverse")
    m2.metric("AP cost per invoice", f"€{cpi_b:,.2f} → €{cpi_a:,.2f}", f"-€{cpi_b - cpi_a:,.2f}", delta_color="inverse")
    m3.metric("AP capacity freed", f"{R['fte']:.1f} FTE", f"{fte_b:.1f} → {fte_a:.1f} FTE workload", delta_color="off")

    c_chart, c_table = st.columns([2, 3])
    with c_chart:
        bars = go.Figure()
        bars.add_trace(go.Bar(name="Manual labour", x=["Today", "With automation"], y=[labour_b, labour_a],
                              marker_color=INK, text=[eur_k(labour_b), eur_k(labour_a)], textposition="inside"))
        bars.add_trace(go.Bar(name="Freed capacity, still paid", x=["Today", "With automation"], y=[0, redeployed],
                              marker_color=AMBER, text=["", eur_k(redeployed)], textposition="inside"))
        bars.add_trace(go.Bar(name="Software", x=["Today", "With automation"], y=[0, R["software"]],
                              marker_color=TEAL_LIGHT, text=["", eur_k(R["software"])], textposition="inside"))
        bars.add_annotation(x="With automation", y=cost_a, text=f"<b>−{eur_k(R['cash'] - R['software'])}</b> a year",
                            showarrow=False, yshift=14, font=dict(color=TEAL, size=13))
        style(bars, barmode="stack", legend=dict(orientation="h", y=-0.2), yaxis=dict(visible=False), height=400,
              title=dict(text="Where the AP money goes (€ per year)", font=dict(size=14), x=0))
        st.plotly_chart(bars, width='stretch')
        st.caption("Amber = hours freed but not turned into cash. Still on payroll, available for other work.")

    with c_table:
        def chg(a, b, fmt):
            d = b - a
            return f"{'+' if d > 0 else '−'}{fmt(abs(d))}"
        rows_html = [
            ("Revenue", "€320M", "€320M",
             "<span class='nochange'>No change modelled</span>"),
            ("Manual AP hours per year", f"{hours_b:,.0f}", f"{hours_a:,.0f}", chg(hours_b, hours_a, lambda x: f"{x:,.0f} h")),
            ("AP workload (FTE)", f"{fte_b:.1f}", f"{fte_a:.1f}", chg(fte_b, fte_a, lambda x: f"{x:.1f} FTE")),
            ("Touch time per invoice", f"{min_b:.1f} min", f"{min_a:.1f} min", chg(min_b, min_a, lambda x: f"{x:.1f} min")),
            ("AP processing cost, cash view", eur_k(labour_b), eur_k(cost_a), chg(labour_b, cost_a, eur_k)),
            ("AP cost per invoice", f"€{cpi_b:,.2f}", f"€{cpi_a:,.2f}", chg(cpi_b, cpi_a, lambda x: f"€{x:,.2f}")),
            ("Software (per year)", "–", eur_k(R["software"]), f"+{eur_k(R['software'])}"),
            ("One-off setup (year 1)", "–", eur_k(R["setup"]), f"+{eur_k(R['setup'])}"),
            ("<b>Year-1 net value</b>", "–", f"<b>{eur_k(R['net'])}</b>", f"<b>+{eur_k(R['net'])}</b>"),
        ]
        html = "<table class='assump'><tr><th>Area</th><th class='num'>Today</th><th class='num'>With automation</th><th class='num'>Change</th></tr>"
        for a, b, c, d in rows_html:
            html += f"<tr><td>{a}</td><td class='num'>{b}</td><td class='num'>{c}</td><td class='num'>{d}</td></tr>"
        html += "</table>"
        st.markdown(html, unsafe_allow_html=True)
        st.caption("Cash view: today's manual labour cost, minus cash savings, plus software. "
                   "Setup is a one-off and shown separately.")

    st.info(
        "**Revenue:** no revenue effect is modelled. Revenue stays at €320M, so this is a cost and capacity case. "
        "Possible upsides such as faster approvals or early-payment discounts are not counted and would need to be verified."
    )

# ---- Tab 3: use cases ---------------------------------------------------------
with tab3:
    st.subheader("Four invoice use cases, and what we have already done")
    st.markdown("Each invoice stream is modelled with its own volume and touch time. "
                "Pick one to highlight it and see where the savings come from.")
    names = [n for n, _ in USE_CASES]
    pick = st.radio("Highlight a use case", ["All"] + names, horizontal=True, label_visibility="collapsed", key="uc_pick")

    uc = []
    for i, (name, blurb) in enumerate(USE_CASES):
        vol, mins = P["vols"][i], P["mins"][i]
        hrs = vol * mins / 60
        freed_i = hrs * P["time_cut"]
        cash_i = freed_i * P["labour"] * P["share"]
        uc.append(dict(name=name, blurb=blurb, vol=vol, mins=mins, hrs=hrs, freed=freed_i, cash=cash_i,
                       share=vol / R["invoices"]))

    cols = st.columns(4)
    for col, u in zip(cols, uc):
        cls = "uc" + (" sel" if pick == u["name"] else (" dim" if pick != "All" else ""))
        col.markdown(
            f"<div class='{cls}'><h4>{u['name']}</h4><div class='small'>{u['blurb']}</div><br>"
            f"<div class='big'>{eur_k(u['cash'])}</div><div class='small'>cash savings per year</div><br>"
            f"<div class='row'><span>Invoices</span><b>{u['vol']:,} ({u['share'] * 100:.0f}%)</b></div>"
            f"<div class='row'><span>Touch time</span><b>{u['mins']} min</b></div>"
            f"<div class='row'><span>Manual hours</span><b>{u['hrs']:,.0f} h</b></div>"
            f"<div class='row'><span>Hours freed</span><b>{u['freed']:,.0f} h</b></div></div>",
            unsafe_allow_html=True,
        )

    st.markdown("&nbsp;")
    uc_fig = go.Figure(go.Bar(
        y=[u["name"] for u in uc][::-1], x=[u["cash"] for u in uc][::-1], orientation="h",
        marker_color=[TEAL if (pick in ("All", u["name"])) else "#CFE3E6" for u in uc][::-1],
        text=[f"{eur_k(u['cash'])}  ·  {u['cash'] / R['cash'] * 100:.0f}% of savings" if R["cash"] else "" for u in uc][::-1],
        textposition="auto",
    ))
    style(uc_fig, showlegend=False, xaxis=dict(visible=False), height=260,
          title=dict(text="Cash savings by use case (€ per year)", font=dict(size=14), x=0))
    st.plotly_chart(uc_fig, width='stretch')

    st.markdown("---")
    st.subheader("Already proven by BLP: Migros")
    p1, p2 = st.columns([3, 2])
    with p1:
        st.markdown(
            "<div class='uc'><h4>Switzerland's largest retailer runs BLP in accounts payable</h4>"
            "<div class='row'><span>Selection</span><b>Selected after onsite tests with 10 and 300 invoices</b></div>"
            "<div class='row'><span>Speed</span><b>First Migros company live 7–8 weeks after signing</b></div>"
            "<div class='row'><span>SAP integration</span><b>About 3 days</b></div>"
            "<div class='row'><span>Effort</span><b>30–40% fewer resources needed</b></div>"
            "<div class='row'><span>Since then</span><b>Rolling out to further areas and processes</b></div></div>",
            unsafe_allow_html=True,
        )
    with p2:
        st.markdown(
            "**Proof and sources**\n\n"
            "- [BLP customer page: Migros](https://blp.digital/resources/our-customers/migros)\n"
            "- [IHL Group analysis of the Migros rollout](https://www.ihlservices.com/news/analyst-corner/2026/04/when-ai-agents-meet-real-enterprise-complexity/)\n"
            "- [BLP funding news, with Migros quote](https://pulse2.com/blp-digital-50-million-investment-from-goldman-sachs-alternatives-to-scale-agentic-ai-for-core-finance-processes/)"
        )
        st.caption("How we use it: Migros supports a comparable 3-day SAP integration, not a guaranteed project duration. "
                   "Its 30–40% resource reduction is different from our 50–70% touch-time scenarios, so it does not validate them.")

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**What we have already done**")
        for t in ["Four invoice use cases with volume and touch time",
                  "Pricing checked against Klippa and ISR public prices",
                  "Three scenarios: Conservative, Base, Optimistic",
                  "Freed capacity and cash savings kept separate",
                  "Confidence rating for all 21 inputs",
                  "Stress test: what if we are wrong?"]:
            st.markdown(f"<span class='chip done'>✓ {t}</span>", unsafe_allow_html=True)
    with d2:
        st.markdown("**Still to do before this is a forecast**")
        for t in ["SAP invoice counts by category",
                  "Timed samples of real invoices in AP",
                  "A firm BLP commercial offer",
                  "Finance plan for turning freed time into cost avoidance"]:
            st.markdown(f"<span class='chip todo'>○ {t}</span>", unsafe_allow_html=True)

# ---- Tab 4: all 21 assumptions -------------------------------------------------
with tab4:
    st.subheader("All 21 rated assumptions in this study")
    IW = ("IW Consult (invoice time study)", "https://www.iwkoeln.de/fileadmin/user_upload/Studien/Gutachten/PDF/2025/IWC-Studie-Wenn-sich-Digitalisierung-rechnet.pdf")
    AEO = ("Aequo AP (modelled touch time)", "https://www.aequoap.com/research/invoice-processing-time-benchmarks")
    USF = ("USFarathane (60k invoices/yr)", "https://documentpeople.dk/en/case-studies/usfarathane/")
    MIG = ("BLP Migros case", "https://blp.digital/resources/our-customers/migros")
    KLI = ("Klippa prices", "https://www.klippa.com/en/invoice-processing/prices/")
    ISR = ("ISR offer", "https://isr.de/en/technologies/isr/accounts-payable-flow/offer/")
    NOSRC = ("Planning assumption", None)
    G1, G2, G3, G4, G5, G6 = ("Invoice volume and touch time", "Labour and capacity", "Savings levers",
                              "BLP price (assumed)", "One-time rollout effort", "Public price references")
    A = [
        (G1, "Direct material invoices (PO)", f"{P['vols'][0]:,} / yr · {P['mins'][0]} min", "Low", IW),
        (G1, "Indirect purchases / services (PO)", f"{P['vols'][1]:,} / yr · {P['mins'][1]} min", "Low", AEO),
        (G1, "Freight and logistics", f"{P['vols'][2]:,} / yr · {P['mins'][2]} min", "Low",
         ("GoComet (workflow only)", "https://www.gocomet.com/products/freight-invoice-reconciliation")),
        (G1, "Non-PO and other expenses", f"{P['vols'][3]:,} / yr · {P['mins'][3]} min", "Low", AEO),
        (G1, "Total annual invoices", f"{R['invoices']:,}", "Medium",
         USF),
        (G2, "Employer labour cost", f"€{P['labour']} / hour", "Medium",
         ("Destatis 2025", "https://www.destatis.de/DE/Presse/Pressemitteilungen/2026/04/PD26_148_624.html")),
        (G2, "Productive AP hours per FTE", f"{P['prod_hours']:,} h / year", "Low", NOSRC),
        (G3, "Manual touch-time reduction", f"{P['time_cut'] * 100:.0f}%  (scenarios 50 / 60 / 70%)", "Low",
         ("BLP Migros (resource KPI only)", "https://blp.digital/resources/our-customers/migros")),
        (G3, "Freed hours converted into cash", f"{P['share'] * 100:.0f}%  (scenarios 40 / 50 / 60%)", "Low", NOSRC),
        (G4, "Platform fee", f"€{P['platform']:,} / month", "Low", KLI),
        (G4, "Usage fee on every invoice", f"€{P['usage']:.2f} / invoice", "Low", KLI),
        (G4, "Total one-time implementation", eur(P["setup"]), "Low", ISR),
        (G5, "ERP integration + connector validation", f"{P['days'][0]} days", "Low", MIG),
        (G5, "Workflow setup + configuration", f"{P['days'][1]} days", "Low", NOSRC),
        (G5, "User testing + training", f"{P['days'][2]} days", "Low",
         ("ISR (training scope, not days)", "https://isr.de/en/technologies/isr/accounts-payable-flow/offer/")),
        (G5, "Go-live assistance", f"{P['days'][3]} days", "Low", NOSRC),
        (G5, "Consultant hours per day", f"{P['hrs_day']} h", "Medium", NOSRC),
        (G6, "Klippa Premium plan (up to 12k invoices)", "€275 / month", "High", KLI),
        (G6, "Klippa extra-invoice price", "from €0.28 / invoice", "High", KLI),
        (G6, "ISR standard rollout package", "€19,500 one-time", "High", ISR),
        (G6, "ISR published support rate", f"€{P['hourly']:,.2f} / hour", "High", ISR),
    ]
    counts = Counter(c for _, _, _, c, _ in A)

    t_left, t_right = st.columns([1, 2])
    with t_left:
        donut = go.Figure(go.Pie(labels=["Low", "Medium", "High"],
                                 values=[counts["Low"], counts["Medium"], counts["High"]], hole=0.62,
                                 marker=dict(colors=[CORAL, AMBER, GREEN], line=dict(color="white", width=2)),
                                 textinfo="value", sort=False))
        style(donut, height=260, legend=dict(orientation="h", y=-0.05),
              annotations=[dict(text=f"{counts['Low']} of {len(A)}<br>are Low", x=0.5, y=0.5, showarrow=False,
                                font=dict(size=16, color=INK))])
        st.plotly_chart(donut, width='stretch')
    with t_right:
        st.markdown(
            f"The model rests on **{len(A)} rated inputs: {counts['Low']} Low, {counts['Medium']} Medium, "
            f"{counts['High']} High**. Values update with the sidebar.\n\n"
            "- <span class='pill High'>High</span> a public, listed fact\n"
            "- <span class='pill Medium'>Medium</span> backed by a published comparable\n"
            "- <span class='pill Low'>Low</span> an unverified estimate of mine\n\n"
            "Calculated totals and fixed fictional company data (revenue, employees, AP staff, ERP) are not rated and are not counted.",
            unsafe_allow_html=True,
        )

    html = ("<table class='assump'><tr><th>#</th><th>Assumption</th><th>Value now</th>"
            "<th>Confidence</th><th>Source</th></tr>")
    last, n = None, 0
    for grp, label, val, conf, (src, url) in A:
        if grp != last:
            html += f"<tr class='grp'><td colspan='5'>{grp}</td></tr>"
            last = grp
        n += 1
        link = f"<a href='{url}' target='_blank'>{src}</a>" if url else f"<span class='nochange'>{src}</span>"
        html += (f"<tr><td>{n}</td><td>{label}</td><td><b>{val}</b></td>"
                 f"<td><span class='pill {conf}'>{conf}</span></td><td>{link}</td></tr>")
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

# ---- Tab 5: what if wrong ------------------------------------------------------
with tab5:
    st.subheader("If we are wrong, does the case still hold?")
    st.markdown(
        "Each bar is the **cash savings in year 1 if only that one input is worse**. The dashed line is the "
        f"**total year-1 cost ({eur_k(R['cost1'])}: software + setup)**. A bar that reaches past the dashed line "
        "means the savings still cover the cost. The gap is the net value."
    )

    tests = {
        "Minutes per invoice 25% lower": dict(mins=[m * 0.75 for m in P["mins"]]),
        "Cash share 10 points lower": dict(share=max(P["share"] - 0.10, 0)),
        "Time cut 10 points lower": dict(time_cut=max(P["time_cut"] - 0.10, 0)),
        "Labour cost €5 per hour lower": dict(labour=max(P["labour"] - 5, 0)),
        "Rollout costs €19.5k (ISR package)": dict(setup=19500),
    }
    cases = {name: calc(variant(P, **ch)) for name, ch in tests.items()}

    st.markdown("**Stress test: combine problems**")
    chosen = st.multiselect(
        "Pick the problems that happen together",
        list(tests), default=["Minutes per invoice 25% lower", "Cash share 10 points lower"],
    )
    combined = {}
    for name in chosen:
        combined.update(tests[name])
    stress = calc(variant(P, **combined)) if chosen else R

    names_w = list(cases) + ["Stress test (your selection)"]
    rs = list(cases.values()) + [stress]
    order = sorted(range(len(rs)), key=lambda i: rs[i]["cash"])
    is_stress = [names_w[i].startswith("Stress") for i in order]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=[names_w[i] for i in order], x=[rs[i]["cash"] for i in order], orientation="h",
        marker_color=[CORAL if st_ else TEAL_LIGHT for st_ in is_stress],
        text=[f"{eur_k(rs[i]['cash'])} savings  →  net {eur_k(rs[i]['net'])}" for i in order],
        textposition="inside", insidetextanchor="start", name="Cash savings in year 1",
    ))
    diff = [i for i in order if abs(rs[i]["cost1"] - R["cost1"]) > 1]
    if diff:
        fig.add_trace(go.Scatter(
            x=[rs[i]["cost1"] for i in diff], y=[names_w[i] for i in diff], mode="markers", name="Higher cost in this case",
            marker=dict(symbol="diamond", size=13, color="white", line=dict(color=CORAL, width=3)),
            hovertemplate="Year-1 cost %{x:,.0f}<extra></extra>"))
    fig.add_vline(x=R["cost1"], line_dash="dash", line_width=3, line_color=CORAL,
                  annotation_text=f"Total year-1 cost {eur_k(R['cost1'])}", annotation_position="top",
                  annotation_font=dict(color=CORAL, size=12))
    top = max(rs[i]["cash"] for i in order) * 1.08
    style(fig, legend=dict(orientation="h", y=-0.18), height=400, margin=dict(l=10, r=10, t=50, b=10),
          xaxis=dict(title="Cash savings in year 1 (€)", tickformat=",.0f", range=[0, top]))
    st.plotly_chart(fig, width='stretch')

    held = stress["net"] > 0
    ratio = stress["cash"] / stress["cost1"] if stress["cost1"] else 0
    w1, w2 = st.columns(2)
    w1.metric("Stress-test year-1 net", eur_k(stress["net"]), f"{(stress['net'] - R['net']) / 1000:+,.1f}k vs current")
    w2.metric("Does the case still hold?", "Yes" if held else "No",
              f"savings are {ratio:.1f}x the cost" if held else f"savings cover only {ratio * 100:.0f}% of the cost",
              delta_color="normal" if held else "inverse")

    st.markdown("---")
    st.subheader("Cash share × time cut: year-1 net value (€k)")
    st.caption("These are the two levers that move between scenarios. Green is positive, red is negative.")
    cuts = [40, 50, 60, 70, 80]
    shares = [30, 40, 50, 60, 70]
    z = [[calc(variant(P, time_cut=c / 100, share=s / 100))["net"] / 1000 for c in cuts] for s in shares]
    heat = go.Figure(go.Heatmap(
        z=z, x=[f"{c}%" for c in cuts], y=[f"{s}%" for s in shares],
        colorscale=[[0, "#F4B9AC"], [0.35, "#FFFFFF"], [1, "#6FC3A0"]], zmid=0,
        text=[[f"{v:,.0f}" for v in row] for row in z], texttemplate="%{text}", showscale=False,
    ))
    style(heat, height=340, xaxis=dict(title="Manual effort removed"),
          yaxis=dict(title="Freed time that becomes savings"))
    st.plotly_chart(heat, width='stretch')

# ---- Tab 6: next steps ---------------------------------------------------------
with tab6:
    st.subheader("What it takes to make this estimate credible")
    n1, n2 = st.columns(2)
    with n1:
        st.markdown("#### Data")
        st.write("SAP invoice counts by category for the last 12 months, to replace the assumed 60,600 and the category split.")
        st.markdown("#### Access")
        st.write("Timed, random samples of real invoices in AP, to replace the assumed 10–15 minutes each.")
    with n2:
        st.markdown("#### People")
        st.write("The AP lead for workload, Finance for a cost-avoidance or hiring plan, and BLP for a firm commercial offer.")
        st.markdown("#### Time")
        st.write("About 1–2 weeks to collect the data and issue a revised model.")

    st.markdown("---")
    st.markdown(
        "**Numbers I trust least:** how much freed time turns into real savings, minutes per invoice, and the "
        "manual time reduction. They are the ones the stress test moves most."
    )

st.caption(
    "Illustrative manufacturing case for a course assignment. All company data is fictional. "
    "Klippa and ISR pricing are public reference points only. Not a BLP quotation."
)
