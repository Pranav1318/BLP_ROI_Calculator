"""
Isarwerk Präzision GmbH: what is AP automation worth to the CFO?

An interactive value calculator (fictional case, not a BLP quotation).
Model logic mirrors the Excel file "Isarwerk_AP_ROI_With_Payback.xlsx".

Run locally:   streamlit run app.py
"""

import pandas as pd
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
    table.assump {{ width:100%; border-collapse: collapse; font-size: .95rem; }}
    table.assump th {{ text-align:left; color:{MUTED}; font-weight:500; padding:6px 8px; border-bottom:1px solid {LINE}; }}
    table.assump td {{ padding:7px 8px; border-bottom:1px solid #EEF3F5; color:{INK}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Model (mirrors the Excel workbook)
# ----------------------------------------------------------------------------
CATEGORIES = ["Direct material", "Indirect / services", "Freight / logistics", "No-PO & other"]

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
    "days": 11,              # rollout days
    "day_rate": 1100,        # EUR per consulting day
}

SCENARIOS = {  # only these two levers move between scenarios in the Excel model
    "Conservative": {"time_cut": 50, "share": 40},
    "Base": {"time_cut": 60, "share": 50},
    "Optimistic": {"time_cut": 70, "share": 60},
}


def params_from_state():
    s = st.session_state
    return {
        "vols": [s[f"vol_{i}"] for i in range(4)],
        "mins": [s[f"min_{i}"] for i in range(4)],
        "time_cut": s["time_cut"] / 100,
        "share": s["share"] / 100,
        "labour": s["labour"],
        "prod_hours": s["prod_hours"],
        "platform": s["platform"],
        "usage": s["usage"],
        "setup": s["days"] * s["day_rate"],
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
    roi = net / cost1 if cost1 else 0
    fte = freed / p["prod_hours"]
    denom = hours * p["time_cut"] * p["labour"]
    breakeven_share = cost1 / denom if denom else None
    return dict(invoices=invoices, hours=hours, freed=freed, cash=cash, software=software,
                setup=p["setup"], cost1=cost1, net=net, payback=payback, roi=roi, fte=fte,
                breakeven_share=breakeven_share)


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
        for i, name in enumerate(CATEGORIES):
            c1, c2 = st.columns(2)
            c1.number_input(f"{name}: invoices / year", min_value=0, step=500, key=f"vol_{i}")
            c2.number_input("Minutes each", min_value=1, max_value=60, key=f"min_{i}")

    with st.expander("Labour cost and capacity"):
        st.slider("Employer cost per hour (€)", 30, 70, key="labour")
        st.slider("Productive hours per FTE per year", 1200, 1800, step=10, key="prod_hours")

    with st.expander("BLP price and rollout"):
        st.slider("Platform fee (€ / month)", 0, 1000, step=25, key="platform")
        st.slider("Usage fee (€ / invoice)", 0.0, 1.0, step=0.01, key="usage")
        st.slider("Rollout effort (consulting days)", 3, 40, key="days")
        st.number_input("Day rate (€)", min_value=500, max_value=2500, step=50, key="day_rate")

    st.button("Reset to base case", on_click=reset_all, width='stretch')
    st.caption("Fictional case. Not a BLP quotation.")

# ----------------------------------------------------------------------------
# Calculations
# ----------------------------------------------------------------------------
P = params_from_state()
R = calc(P)
BASE_P = {
    "vols": [DEFAULTS[f"vol_{i}"] for i in range(4)],
    "mins": [DEFAULTS[f"min_{i}"] for i in range(4)],
    "time_cut": DEFAULTS["time_cut"] / 100, "share": DEFAULTS["share"] / 100,
    "labour": DEFAULTS["labour"], "prod_hours": DEFAULTS["prod_hours"],
    "platform": DEFAULTS["platform"], "usage": DEFAULTS["usage"],
    "setup": DEFAULTS["days"] * DEFAULTS["day_rate"],
}
BASE_R = calc(BASE_P)

# ----------------------------------------------------------------------------
# Header and verdict
# ----------------------------------------------------------------------------
st.title("What is AP automation worth to Isarwerk?")
st.markdown(
    "Isarwerk Präzision GmbH is a fictional €320m German manufacturer that processes about "
    f"**{R['invoices']:,.0f} supplier invoices a year**. Change the assumptions on the left and see the "
    "year-1 result, how sure we are, and what happens if we are wrong."
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
k4.metric("Payback on setup cost", f"{R['payback']:.1f} months" if R["payback"] else "no payback",
          None if R["payback"] is None or BASE_R["payback"] is None
          else f"{R['payback'] - BASE_R['payback']:+.1f} months vs base", delta_color="inverse")

s1, s2, s3 = st.columns(3)
s1.metric("Manual hours today", f"{R['hours']:,.0f} h")
s2.metric("Hours freed", f"{R['freed']:,.0f} h", f"{R['fte']:.1f} FTE of capacity", delta_color="off")
s3.metric("Year-1 return on cost", f"{R['roi'] * 100:,.0f}%")

# ----------------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["The result", "How sure are we?", "What if we are wrong?", "What we need next"]
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
        st.subheader("When does it pay back?")
        months = list(range(0, 13))
        monthly = (R["cash"] - R["software"]) / 12
        cum = [-R["setup"] + m * monthly for m in months]
        line = go.Figure()
        line.add_trace(go.Scatter(x=months, y=cum, mode="lines+markers", line=dict(color=TEAL, width=3),
                                  marker=dict(size=6), name="Cumulative net"))
        line.add_hline(y=0, line_dash="dash", line_color=MUTED)
        style(line, showlegend=False,
                           xaxis=dict(title="Month after go-live", dtick=1),
                           yaxis=dict(title="Cumulative net (€)", tickformat=",.0f"))
        st.plotly_chart(line, width='stretch')

    st.subheader("Three scenarios side by side")
    st.caption("Same invoice volumes, prices and costs as in the sidebar. Only time cut and cash share change.")
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
    style(cmp_fig, showlegend=False, yaxis=dict(visible=False), height=320)
    st.plotly_chart(cmp_fig, width='stretch')

# ---- Tab 2: confidence --------------------------------------------------------
with tab2:
    st.subheader("Most numbers are estimates, and the model says so")
    st.markdown(
        "Each assumption carries a confidence level from the Excel model. "
        "**High** is a public, listed fact. **Medium** is backed by a published comparable. "
        "**Low** is my own unverified estimate."
    )
    assumptions = [
        ("Annual invoices", f"{R['invoices']:,.0f}", "Medium"),
        ("Manual minutes per invoice", f"{min(P['mins'])}–{max(P['mins'])} min", "Low"),
        ("Employer labour cost", f"€{P['labour']} / hour", "Medium"),
        ("Productive hours per FTE", f"{P['prod_hours']:,} h", "Low"),
        ("Manual effort removed", f"{P['time_cut'] * 100:.0f}%", "Low"),
        ("Freed time that becomes savings", f"{P['share'] * 100:.0f}%", "Low"),
        ("BLP platform fee (assumed)", f"€{P['platform']} / month", "Low"),
        ("BLP usage fee (assumed)", f"€{P['usage']:.2f} / invoice", "Low"),
        ("One-off rollout cost", eur(P["setup"]), "Low"),
        ("Klippa list price (reference)", "€275 / month", "High"),
    ]
    c_left, c_right = st.columns([3, 2])
    with c_left:
        html = "<table class='assump'><tr><th>Assumption</th><th>Value now</th><th>Confidence</th></tr>"
        for a, v, c in assumptions:
            html += f"<tr><td>{a}</td><td><b>{v}</b></td><td><span class='pill {c}'>{c}</span></td></tr>"
        html += "</table>"
        st.markdown(html, unsafe_allow_html=True)
    with c_right:
        counts = {"Low": 14, "Medium": 3, "High": 4}
        donut = go.Figure(go.Pie(labels=list(counts), values=list(counts.values()), hole=0.62,
                                 marker=dict(colors=[CORAL, AMBER, GREEN], line=dict(color="white", width=2)),
                                 textinfo="value", sort=False))
        style(donut, height=300, legend=dict(orientation="h", y=-0.05),
                            annotations=[dict(text="14 of 21<br>are Low", x=0.5, y=0.5, showarrow=False,
                                              font=dict(size=16, color=INK))])
        st.plotly_chart(donut, width='stretch')
        st.caption("21 rated inputs in the Excel model. Calculated totals and fixed fictional inputs are excluded.")

# ---- Tab 3: what if wrong ------------------------------------------------------
with tab3:
    st.subheader("How much does it hurt if one assumption is worse than we think?")
    st.markdown(
        "Each bar shows the **year-1 net value if only that one input is worse**. Everything else stays at your "
        "current settings. The dashed line is your current result. A bar that stops short of the line shows how much value is lost."
    )

    tests = {
        "Minutes per invoice 25% lower": dict(mins=[m * 0.75 for m in P["mins"]]),
        "Cash share 10 points lower": dict(share=max(P["share"] - 0.10, 0)),
        "Time cut 10 points lower": dict(time_cut=max(P["time_cut"] - 0.10, 0)),
        "Labour cost €5 per hour lower": dict(labour=max(P["labour"] - 5, 0)),
        "Rollout costs €19.5k (ISR package)": dict(setup=19500),
    }
    results = {name: calc(variant(P, **ch))["net"] for name, ch in tests.items()}

    st.markdown("**Stress test: combine problems**")
    chosen = st.multiselect(
        "Pick the problems that happen together",
        list(tests), default=["Minutes per invoice 25% lower", "Cash share 10 points lower"],
    )
    combined = {}
    for name in chosen:
        combined.update(tests[name])
    stress_net = calc(variant(P, **combined))["net"] if chosen else R["net"]
    stress_label = "Stress test"
    names = list(results) + [stress_label]
    values = list(results.values()) + [stress_net]
    order = sorted(range(len(names)), key=lambda i: values[i])
    fig = go.Figure(go.Bar(
        y=[names[i] for i in order], x=[values[i] for i in order], orientation="h",
        marker_color=[CORAL if names[i] == stress_label else TEAL_LIGHT for i in order],
        text=[eur_k(values[i]) for i in order], textposition="inside", insidetextanchor="end",
    ))
    fig.add_vline(x=R["net"], line_dash="dash", line_color=TEAL,
                  annotation_text=f"Current {eur_k(R['net'])}", annotation_position="top")
    style(fig, showlegend=False, xaxis=dict(title="Year-1 net value (€)", tickformat=",.0f", rangemode="tozero"),
                      height=380, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, width='stretch')

    m1, m2 = st.columns(2)
    m1.metric("Stress-test year-1 net", eur_k(stress_net), f"{(stress_net - R['net']) / 1000:+,.1f}k vs current")
    m2.metric("Still positive?", "Yes" if stress_net > 0 else "No",
              "case holds" if stress_net > 0 else "case breaks", delta_color="normal" if stress_net > 0 else "inverse")

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

# ---- Tab 4: next steps ---------------------------------------------------------
with tab4:
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
    "Klippa pricing is a public reference point only. Not a BLP quotation."
)
