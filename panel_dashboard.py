import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import timedelta, date

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="First Pass Quality Dashboard",
    page_icon="🏭",
    layout="wide",
)

NAVY     = "#000E57"
BLUE     = "#0F3CFF"
TINT     = "#E7ECFF"
OFFWHITE = "#FFFEF5"
WARMGRAY = "#ABA18B"
RED      = "#C0392B"
GREEN    = "#27AE60"

_SB_URL  = 'https://ptbhguthosenkffjhbry.supabase.co'
_SB_ANON = ('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9'
             '.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB0YmhndXRob3NlbmtmZmpoYnJ5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ5MDkzNzAsImV4cCI6MjEwMDQ4NTM3MH0'
             '.y71yXWsIUjzFBqC4itzZ36w9ixw_QDej2RaBPh8MLPI')

st.markdown(f"""
<style>
  .main {{ background-color: {OFFWHITE}; }}
  h1, h2, h3 {{ color: {NAVY}; font-family: Georgia, serif; }}
  .kpi-box {{
      background: {NAVY}; color: white; border-radius: 8px;
      padding: 18px 12px; text-align: center; margin: 4px;
  }}
  .kpi-label {{ font-size: 11px; opacity: 0.75; font-family: Arial;
                letter-spacing:1px; text-transform:uppercase; }}
  .kpi-value {{ font-size: 30px; font-weight: bold; font-family: Georgia, serif; }}
  .kpi-sub   {{ font-size: 11px; opacity: 0.65; font-family: Arial; }}
  .section-header {{
      background: {NAVY}; color: white; padding: 8px 16px;
      border-radius: 4px; font-family: Georgia; font-size: 15px;
      font-weight: bold; margin: 20px 0 10px 0;
  }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fiscal_quarter_label(dt):
    """Return label like 'FY25 Q1' for fiscal year starting Oct 1."""
    p = dt.to_period("Q-SEP")
    return f"FY{str(p.year)[2:]} Q{p.quarter}"


def short_name(full):
    full = str(full).strip()
    if not full or full == "nan":
        return ""
    if "," in full:
        last, first = full.split(",", 1)
        first, last = first.strip(), last.strip()
    else:
        parts = full.strip().split()
        first = parts[0]
        last = parts[-1] if len(parts) > 1 else ""
    return f"{first} {last[0]}." if last else first


def extract_all_initials(text, known):
    """Return list of known initials found in text.
    Handles (AB), (CO, IM), and (AB) (QP) patterns."""
    if not isinstance(text, str):
        return []
    results = []
    for group in re.findall(r'\(([^)]+)\)', text):
        for init in re.findall(r'[A-Z]{2,3}', group):
            if init in known:
                results.append(init)
    return results


@st.cache_data(ttl=300, show_spinner=False)
def load_time_data():
    import urllib.request as _ur, json as _js
    url = (_SB_URL + "/rest/v1/panel_time_summary"
           "?select=*&order=layout_started_at.desc.nullslast&limit=10000")
    req = _ur.Request(url, headers={
        'apikey': _SB_ANON,
        'Authorization': f'Bearer {_SB_ANON}',
        'Accept': 'application/json',
    })
    with _ur.urlopen(req, timeout=20) as r:
        return pd.DataFrame(_js.loads(r.read()))


@st.cache_data(ttl=60, show_spinner=False)
def _load_debug_sample():
    """Fetch 8 raw rows for debugging — shows exact API values before any conversion."""
    import urllib.request as _ur, json as _js
    url = (_SB_URL + "/rest/v1/panel_time_summary"
           "?select=panel_number,job_number,layout_started_at,is_complete"
           "&order=layout_started_at.desc.nullslast&limit=8")
    req = _ur.Request(url, headers={
        'apikey': _SB_ANON,
        'Authorization': f'Bearer {_SB_ANON}',
        'Accept': 'application/json',
    })
    with _ur.urlopen(req, timeout=10) as r:
        return _js.loads(r.read())


# ── Load & process data ───────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    xl = pd.read_excel(file_bytes, sheet_name=None)

    dr = xl.get("Daily Results") if "Daily Results" in xl else xl.get("old daily results")
    if dr is None:
        st.error("Could not find 'Daily Results' sheet."); st.stop()

    dr.columns = (
        ["Panel","Date","Pass_Fail","Reason","Line",
         "Emp","Routing_Hours","Tech","Layout","Wire","FA","Sales_Order",
         "C1","C2","C3"][:len(dr.columns)]
    )
    dr = dr[["Panel","Date","Pass_Fail","Reason","Line",
             "Routing_Hours","Tech","Layout","Wire","FA","Sales_Order"]].copy()
    dr["Date"] = pd.to_datetime(dr["Date"], errors="coerce")
    dr = dr[dr["Date"].notna()].copy()
    dr["Pass_Fail"] = (
        dr["Pass_Fail"].astype(str).str.strip().str.upper()
        .replace({"FALL":"FAIL","PAS":"PASS","FAIL ":"FAIL"})
    )
    for col in ["Layout","Wire","Tech","FA"]:
        dr[col] = dr[col].astype(str).str.strip().str.upper().replace("NAN","")

    # Names
    names_df = xl.get("Names", pd.DataFrame(columns=["Name","Initials"]))
    init_map = {
        str(r["Initials"]).strip(): short_name(r["Name"])
        for _, r in names_df.iterrows() if pd.notna(r["Initials"])
    }
    known_inits = set(init_map.keys())

    def name(init):
        i = str(init).strip().upper() if pd.notna(init) else ""
        return init_map.get(i, i) if i else ""

    # Extract ALL initials from reason (anywhere, multiple)
    dr["Fail_Inits"] = dr["Reason"].map(
        lambda x: extract_all_initials(x, known_inits)
    )
    dr["Layout_Name"] = dr["Layout"].map(name)
    dr["Wire_Name"]   = dr["Wire"].map(name)
    dr["FA_Name"]     = dr["FA"].map(name)
    dr["Is_Fail"]  = (dr["Pass_Fail"] == "FAIL").astype(int)
    dr["Is_Pass"]  = (dr["Pass_Fail"] == "PASS").astype(int)
    dr["Month"]    = dr["Date"].dt.to_period("M").astype(str)
    dr["Quarter"]  = dr["Date"].apply(fiscal_quarter_label)
    dr["Week"]     = dr["Date"].dt.to_period("W-SUN").apply(
        lambda p: p.start_time.strftime("%Y-%m-%d"))
    dr["MonthLabel"] = dr["Date"].dt.strftime("%b %Y")

    # Exploded fail attribution (one row per attributed person per fail)
    fail_rows = dr[dr["Is_Fail"] == 1].copy()
    exploded = fail_rows.explode("Fail_Inits")
    exploded = exploded[exploded["Fail_Inits"].notna() & (exploded["Fail_Inits"] != "")]
    exploded["Fail_Name"] = exploded["Fail_Inits"].map(lambda x: init_map.get(x, x))

    return dr, exploded, init_map, known_inits


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<h2 style='color:{NAVY}; font-family:Georgia'>⚙ Controls</h2>",
                unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload First Pass Q Report (.xlsx)", type=["xlsx"])

st.markdown(
    f"<h1 style='color:{NAVY}; font-family:Georgia; margin-bottom:0'>"
    f"First Pass Quality Dashboard</h1>", unsafe_allow_html=True)
st.markdown(
    f"<p style='color:{WARMGRAY}; font-family:Arial; font-size:13px; margin-top:0'>"
    f"Kennesaw Panel Assembly · First Pass Yield & Technician Performance</p>",
    unsafe_allow_html=True)

if not uploaded:
    st.markdown("---")
    c1, c2, c3 = st.columns([1,2,1])
    with c2:
        st.markdown(f"""
        <div style='text-align:center; padding:60px 20px; background:{TINT};
                    border-radius:12px; border: 2px dashed {BLUE}'>
            <div style='font-size:48px'>📂</div>
            <h3 style='color:{NAVY}; font-family:Georgia'>Upload your report to get started</h3>
            <p style='color:{WARMGRAY}; font-family:Arial'>
                Download the <b>First Pass Q Report</b> from Excel Online,<br>
                then drag it onto the uploader in the sidebar.
            </p>
        </div>""", unsafe_allow_html=True)
    st.stop()

with st.spinner("Processing report…"):
    df, fail_exploded, init_map, known_inits = load_data(uploaded)

MAX_DATE = df["Date"].max()
MIN_DATE = df["Date"].min()
quarters_avail = sorted(df["Quarter"].unique())

# ── Period selector ───────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    period_choice = st.selectbox("Time period", [
        "Last 7 Days",
        "Last 30 Days",
        "Last 90 Days",
        "Last 6 Months",
        "Last Year",
        "By Quarter",
        "All Time",
    ], index=2)

    selected_q = None
    if period_choice == "By Quarter":
        selected_q = st.selectbox("Quarter", quarters_avail[::-1])

    st.divider()
    st.caption(f"📅 {MIN_DATE.strftime('%b %d, %Y')} – {MAX_DATE.strftime('%b %d, %Y')}")
    st.caption(f"📋 {len(df):,} total records")

# Apply period filter
period_days = {
    "Last 7 Days": 7, "Last 30 Days": 30, "Last 90 Days": 90,
    "Last 6 Months": 182, "Last Year": 365,
}
if period_choice in period_days:
    cutoff = MAX_DATE - timedelta(days=period_days[period_choice])
    dff  = df[df["Date"] >= cutoff].copy()
    dff_fails = fail_exploded[fail_exploded["Date"] >= cutoff].copy()
    period_label = f"{period_choice}  ({cutoff.strftime('%b %d')} – {MAX_DATE.strftime('%b %d, %Y')})"
elif period_choice == "By Quarter":
    dff  = df[df["Quarter"] == selected_q].copy()
    dff_fails = fail_exploded[fail_exploded["Quarter"] == selected_q].copy()
    period_label = selected_q
else:  # All Time
    dff  = df.copy()
    dff_fails = fail_exploded.copy()
    period_label = "All Time"

st.markdown(
    f"<p style='color:{BLUE}; font-family:Arial; font-weight:bold; font-size:13px'>"
    f"Period: {period_label}</p>", unsafe_allow_html=True)

# ── KPI tiles ─────────────────────────────────────────────────────────────────
total   = len(dff)
passes  = dff["Is_Pass"].sum()
fails   = dff["Is_Fail"].sum()
fpy     = passes / total * 100 if total else 0
days    = max((dff["Date"].max() - dff["Date"].min()).days, 1)
avg_day = total / days

attr_counts = dff_fails["Fail_Name"].value_counts()
top_fail_person = attr_counts.index[0] if len(attr_counts) else "—"
top_fail_count  = attr_counts.iloc[0] if len(attr_counts) else 0

c1, c2, c3, c4, c5 = st.columns(5)
for col, label, value, sub in [
    (c1, "Total Panels",     f"{total:,}",        f"{days} days"),
    (c2, "First Pass Yield", f"{fpy:.1f}%",        f"{passes:,} passed"),
    (c3, "Total Fails",      f"{fails:,}",          f"{fails/total*100:.1f}% rate" if total else ""),
    (c4, "Avg / Day",        f"{avg_day:.1f}",      "panels per day"),
    (c5, "Most Fails From",  top_fail_person,       f"{top_fail_count} attributed fails"),
]:
    col.markdown(f"""<div class='kpi-box'>
        <div class='kpi-label'>{label}</div>
        <div class='kpi-value'>{value}</div>
        <div class='kpi-sub'>{sub}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📛 Fail Attribution", "🔧 Layout by Person", "🔌 Wire by Person",
     "✅ Final by Person", "📈 Trends", "⏱ Time Analysis"])


# ══ TAB 1: Fail Attribution ═══════════════════════════════════════════════════
with tab1:
    st.markdown(f"<div class='section-header'>Fail Attribution — {period_label}</div>",
                unsafe_allow_html=True)
    st.caption("Click a bar to see that person's full fail list.")

    if dff_fails.empty:
        st.info("No attributed fails in this period.")
    else:
        fa_table = (
            dff_fails.groupby("Fail_Name")
            .agg(Fails=("Is_Fail","sum"))
            .reset_index()
            .sort_values("Fails", ascending=False)
        )
        fa_table["% of Fails"] = (fa_table["Fails"] / fails * 100).round(1)

        def top_issue(person_name):
            reasons = dff_fails[dff_fails["Fail_Name"] == person_name]["Reason"].dropna()
            if reasons.empty: return "—"
            clean = reasons.str.replace(r'\s*\([A-Z]{2,3}\)', '', regex=True).str.strip()
            return clean.value_counts().idxmax()[:70] if len(clean) else "—"

        fa_table["Top Issue"] = fa_table["Fail_Name"].map(top_issue)
        fa_table.columns = ["Person","Fails Caused","% of Fails","Most Common Issue"]

        col_chart, col_table = st.columns([3, 2])

        with col_chart:
            fig = px.bar(
                fa_table, x="Fails Caused", y="Person", orientation="h",
                color="Fails Caused",
                color_continuous_scale=[[0,TINT],[0.5,BLUE],[1,NAVY]],
                text="Fails Caused",
                title="Fails Attributed by Person",
                custom_data=["Person"]
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(
                plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
                font=dict(family="Arial", color=NAVY),
                title_font=dict(family="Georgia", color=NAVY, size=15),
                coloraxis_showscale=False,
                yaxis=dict(autorange="reversed"),
                height=max(300, len(fa_table)*42),
                margin=dict(l=10, r=50, t=40, b=20)
            )

            # Clickable bar chart
            event = st.plotly_chart(
                fig, use_container_width=True,
                on_select="rerun", key="fail_bar_chart"
            )

        with col_table:
            st.dataframe(
                fa_table.style.format({"% of Fails": "{:.1f}%"}),
                use_container_width=True, hide_index=True
            )

        # ── Drilldown on click ────────────────────────────────────────────────
        selected_points = event.selection.get("points", []) if event.selection else []
        if selected_points:
            clicked_person = selected_points[0].get("customdata", [None])[0]
            if clicked_person:
                st.markdown(
                    f"<div class='section-header'>📋 {clicked_person} — All Attributed Fails</div>",
                    unsafe_allow_html=True)
                person_fails = dff_fails[dff_fails["Fail_Name"] == clicked_person][
                    ["Date","Sales_Order","Panel","Line","Reason"]
                ].sort_values("Date", ascending=False).copy()
                person_fails["Date"] = person_fails["Date"].dt.strftime("%Y-%m-%d")
                person_fails.rename(columns={"Sales_Order": "Job #"}, inplace=True)
                person_fails["Reason"] = person_fails["Reason"].str.strip()
                st.dataframe(person_fails, use_container_width=True, hide_index=True)

        # ── Unattributed fails ────────────────────────────────────────────────
        fail_rows_period = dff[dff["Is_Fail"] == 1]
        unattr = fail_rows_period[fail_rows_period["Fail_Inits"].map(len) == 0]
        if len(unattr):
            with st.expander(f"⚠️ {len(unattr)} fails with no initials found in reason"):
                show = unattr[["Date","Sales_Order","Panel","Line","Reason"]].sort_values(
                    "Date", ascending=False).copy()
                show["Date"] = show["Date"].dt.strftime("%Y-%m-%d")
                show.rename(columns={"Sales_Order": "Job #"}, inplace=True)
                st.dataframe(show, use_container_width=True, hide_index=True)


# ══ TABS 2 & 3: Layout / Wire by Person ═══════════════════════════════════════
def person_panels_tab(role_col, role_label):
    st.markdown(
        f"<div class='section-header'>{role_label} Panels by Person — {period_label}</div>",
        unsafe_allow_html=True)

    sub = dff[dff[role_col] != ""].copy()
    if sub.empty:
        st.info("No data."); return

    group_by = st.radio("Group by", ["Day","Week","Month"], horizontal=True,
                        index=2, key=f"grp_{role_col}")

    if group_by == "Month":
        sub["Period"] = sub["MonthLabel"]
        period_order = sub.groupby("Period")["Date"].min().sort_values().index.tolist()
    elif group_by == "Week":
        sub["Period"] = "Wk " + sub["Week"]
        period_order = sorted(sub["Period"].unique())
    else:
        sub["Period"] = sub["Date"].dt.strftime("%b %d")
        period_order = sub.groupby("Period")["Date"].min().sort_values().index.tolist()

    summary = (
        sub.groupby([role_col,"Period"])
        .agg(Panels=("Panel","count"), Fails=("Is_Fail","sum"))
        .reset_index()
    )
    summary["FPY%"] = (
        (summary["Panels"] - summary["Fails"]) / summary["Panels"] * 100
    ).round(1)
    summary["Person"] = summary[role_col].map(lambda x: init_map.get(x, x))

    pivot = summary.pivot_table(
        index=role_col, columns="Period", values="Panels", fill_value=0)
    pivot = pivot.reindex(columns=[c for c in period_order if c in pivot.columns])
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)

    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.bar(
            summary, x="Period", y="Panels", color="Person",
            category_orders={"Period": period_order},
            barmode="group",
            title=f"{role_label} Panels per {group_by} by Person",
            color_discrete_sequence=px.colors.qualitative.Bold,
            custom_data=["Person"]
        )
        fig.update_layout(
            plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
            font=dict(family="Arial", color=NAVY),
            title_font=dict(family="Georgia", color=NAVY, size=15),
            legend=dict(title=None),
            height=400, margin=dict(l=10, r=10, t=40, b=20)
        )
        ev_grp = st.plotly_chart(fig, use_container_width=True,
                                 on_select="rerun", key=f"grp_chart_{role_col}")

    with col_table:
        display = pivot.reset_index()
        display[role_col] = display[role_col].map(lambda x: init_map.get(x, x))
        display.rename(columns={role_col: "Person"}, inplace=True)
        st.dataframe(display, use_container_width=True, hide_index=True)

    # FPY by person — only count fails where THIS person's initials are in the reason
    fpy_rows = []
    for init in sub[role_col].unique():
        person_sub = sub[sub[role_col] == init]
        n = len(person_sub)
        my_fails = person_sub[
            (person_sub["Is_Fail"] == 1) &
            person_sub["Fail_Inits"].apply(
                lambda inits: isinstance(inits, list) and init in inits)
        ].shape[0]
        fpy_rows.append({
            role_col: init,
            "Panels": n,
            "Fails": my_fails,
            "FPY%": round((n - my_fails) / n * 100, 1) if n else 100.0,
            "Person": init_map.get(init, init),
        })
    fpy_p = pd.DataFrame(fpy_rows).sort_values("FPY%", ascending=True)

    fpy_tooltip = (
        f"First Pass Yield (FPY) = panels that passed inspection on the first "
        f"attempt ÷ total panels built, shown per {role_label.lower()} technician. "
        f"A panel 'fails' if it required any rework or correction before passing. "
        f"Higher is better — 90% is the target."
    )

    fig2 = px.bar(
        fpy_p, x="FPY%", y="Person", orientation="h",
        text="FPY%",
        color="FPY%",
        color_continuous_scale=[[0,RED],[0.5,WARMGRAY],[1,GREEN]],
        range_color=[70,100],
        custom_data=["Person"]
    )
    fig2.add_vline(x=90, line_dash="dash", line_color=NAVY,
                   annotation_text="90% target")
    fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig2.add_annotation(
        x=0, y=1.12,
        xref="paper", yref="paper",
        text=f"First Pass Yield % — by {role_label} Person  <b>ⓘ</b>",
        font=dict(size=15, color=NAVY, family="Georgia"),
        showarrow=False,
        xanchor="left",
        hovertext=fpy_tooltip,
        hoverlabel=dict(
            bgcolor=NAVY, font_color="white",
            font_size=12, font_family="Arial",
            bordercolor=BLUE
        )
    )
    fig2.update_layout(
        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
        font=dict(family="Arial", color=NAVY),
        coloraxis_showscale=False,
        yaxis=dict(autorange="reversed"),
        height=max(300, len(fpy_p)*42),
        margin=dict(l=10, r=70, t=55, b=20)
    )
    ev_fpy = st.plotly_chart(fig2, use_container_width=True,
                             on_select="rerun", key=f"fpy_chart_{role_col}")

    # ── Drilldown: panels for clicked person ──────────────────────────────────
    clicked_person = None
    for ev in [ev_grp, ev_fpy]:
        pts = ev.selection.get("points", []) if ev.selection else []
        if pts:
            cd = pts[0].get("customdata")
            clicked_person = cd[0] if cd else None
            break

    if clicked_person:
        st.markdown(
            f"<div class='section-header'>📋 {clicked_person} — {role_label} Panels</div>",
            unsafe_allow_html=True)
        name_col = f"{role_col}_Name"

        # Map full name back to initials so we can check Fail_Inits
        clicked_init = next((k for k, v in init_map.items() if v == clicked_person), None)

        raw = dff[dff[name_col] == clicked_person][
            ["Date","Sales_Order","Pass_Fail","Panel","Line","Reason","Fail_Inits",
             "Layout","Wire","FA"]
        ].sort_values("Date", ascending=False).copy()

        def attr_label(row):
            if row["Pass_Fail"] != "FAIL":
                return ""
            inits = row["Fail_Inits"]
            if not isinstance(inits, list) or not inits:
                return "— unattributed"
            return ", ".join(init_map.get(i, i) for i in inits)

        def attr_role(row):
            """Which role(s) are attributed for this fail."""
            inits = row["Fail_Inits"]
            if row["Pass_Fail"] != "FAIL" or not isinstance(inits, list) or not inits:
                return ""
            roles = []
            for init in inits:
                lo = str(row.get("Layout") or "").strip()
                wi = str(row.get("Wire")   or "").strip()
                fa = str(row.get("FA")     or "").strip()
                if lo and lo == init:
                    roles.append("Layout")
                if wi and wi == init:
                    roles.append("Wire")
                if fa and fa == init:
                    roles.append("Final")
            return ", ".join(dict.fromkeys(roles)) if roles else ""

        raw["Failure Caused By"] = raw.apply(attr_label, axis=1)
        raw["Role"] = raw.apply(attr_role, axis=1)
        # True when THIS person (in their role for this tab) caused the fail
        raw["_my_fail"] = raw.apply(
            lambda r: (
                r["Pass_Fail"] == "FAIL" and
                isinstance(r["Fail_Inits"], list) and
                clicked_init is not None and
                clicked_init in r["Fail_Inits"]
            ), axis=1
        )

        raw["Date"] = raw["Date"].dt.strftime("%Y-%m-%d")
        raw.rename(columns={"Sales_Order": "Job #", "Pass_Fail": "Pass/Fail"}, inplace=True)

        display = raw[
            ["Date","Job #","Pass/Fail","Panel","Line","Reason","Role","Failure Caused By"]
        ].reset_index(drop=True)
        my_fail_mask = raw["_my_fail"].reset_index(drop=True)

        def _hl_role_fails(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for i in df.index:
                if my_fail_mask.iloc[i]:
                    styles.iloc[i] = "background-color: #fff0f0"
            return styles

        st.caption(
            f"**Role** = which role caused each fail. "
            f"Rows highlighted red = fail attributed to {clicked_person} in the {role_label} role."
        )
        st.dataframe(
            display.style.apply(_hl_role_fails, axis=None),
            use_container_width=True, hide_index=True
        )

with tab2:
    person_panels_tab("Layout", "Layout")

with tab3:
    person_panels_tab("Wire", "Wire")

with tab4:
    person_panels_tab("FA", "Final")


# ══ TAB 5: Trends ════════════════════════════════════════════════════════════
with tab5:

    # ── Panels built chart ────────────────────────────────────────────────────
    st.markdown(f"<div class='section-header'>Panels Built — {period_label}</div>",
                unsafe_allow_html=True)

    grp_trends = st.radio("Group by", ["Day","Week","Month"], horizontal=True,
                          index=1, key="grp_trends")

    td = dff.copy()
    if grp_trends == "Month":
        td["_p"] = td["Date"].dt.to_period("M")
    elif grp_trends == "Week":
        td["_p"] = td["Date"].dt.to_period("W-SUN")
    else:
        td["_p"] = td["Date"].dt.to_period("D")

    trend = (
        td.groupby("_p")
        .agg(Panels=("Panel","count"), Fails=("Is_Fail","sum"))
        .reset_index()
    )
    trend["_p_start"] = trend["_p"].apply(lambda p: p.start_time)
    trend["Label"] = trend["_p"].astype(str)
    if grp_trends == "Month":
        trend["Label"] = trend["_p_start"].dt.strftime("%b %Y")
    elif grp_trends == "Week":
        trend["Label"] = trend["_p_start"].dt.strftime("Wk %b %d")
    else:
        trend["Label"] = trend["_p_start"].dt.strftime("%b %d")
    trend = trend.sort_values("_p_start")

    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(
        x=trend["Label"], y=trend["Panels"],
        marker_color=BLUE, name="Panels",
        text=trend["Panels"], textposition="outside"
    ))
    fig_trend.update_layout(
        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
        font=dict(family="Arial", color=NAVY),
        height=380,
        margin=dict(l=10, r=10, t=20, b=20),
        yaxis_title="Panels Built",
        xaxis=dict(tickangle=-35 if grp_trends == "Day" else 0)
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # ── Period comparison ─────────────────────────────────────────────────────
    st.markdown("<div class='section-header'>Period Comparison</div>",
                unsafe_allow_html=True)

    compare_by = st.radio("Compare by", ["Quarter","Month"], horizontal=True,
                          key="compare_by")

    if compare_by == "Quarter":
        comp = (
            df.groupby("Quarter")
            .agg(Panels=("Panel","count"), Fails=("Is_Fail","sum"))
            .reset_index()
            .rename(columns={"Quarter": "Period"})
            .sort_values("Period")
        )
    else:
        comp = (
            df.groupby(["Month","MonthLabel"])
            .agg(Panels=("Panel","count"), Fails=("Is_Fail","sum"))
            .reset_index()
            .sort_values("Month")
            .rename(columns={"MonthLabel": "Period"})
        )

    comp["Pass Rate%"] = ((comp["Panels"] - comp["Fails"]) / comp["Panels"] * 100).round(1)
    comp["Fail Rate%"] = (comp["Fails"] / comp["Panels"] * 100).round(1)

    fig_comp = px.bar(
        comp, x="Period", y="Panels", text="Panels",
        title=f"Total Panels by {compare_by}",
        color_discrete_sequence=[BLUE]
    )
    fig_comp.update_traces(textposition="outside")
    fig_comp.update_layout(
        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
        font=dict(family="Arial", color=NAVY),
        title_font=dict(family="Georgia", color=NAVY, size=14),
        height=320, margin=dict(l=10, r=10, t=40, b=20),
        xaxis=dict(tickangle=-35 if compare_by == "Month" else 0)
    )
    st.plotly_chart(fig_comp, use_container_width=True)

    fig_fail = go.Figure()
    fig_fail.add_trace(go.Scatter(
        x=comp["Period"], y=comp["Fail Rate%"],
        mode="lines+markers+text",
        line=dict(color=RED, width=2),
        marker=dict(size=7, color=RED),
        text=comp["Fail Rate%"].map(lambda v: f"{v:.1f}%"),
        textposition="top center",
        textfont=dict(color=RED, size=11),
        fill="tozeroy",
        fillcolor="rgba(192,57,43,0.08)",
    ))
    fig_fail.update_layout(
        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
        font=dict(family="Arial", color=NAVY),
        title=dict(text=f"Fail Rate % by {compare_by}",
                   font=dict(family="Georgia", color=NAVY, size=14)),
        height=280, margin=dict(l=10, r=10, t=40, b=20),
        yaxis=dict(title="Fail Rate %", ticksuffix="%", rangemode="tozero"),
        xaxis=dict(tickangle=-35 if compare_by == "Month" else 0),
        showlegend=False,
    )
    st.plotly_chart(fig_fail, use_container_width=True)

    display_comp = comp[["Period","Panels","Fails","Pass Rate%","Fail Rate%"]].copy()
    if compare_by == "Month":
        display_comp = display_comp.drop(columns=["Month"], errors="ignore")
    st.dataframe(display_comp, use_container_width=True, hide_index=True)

    # ── Master person summary ─────────────────────────────────────────────────
    st.markdown("<div class='section-header'>Person Summary</div>",
                unsafe_allow_html=True)
    st.caption("Panels and fail rate for each person by their role on the panel.")

    ps_col1, ps_col2 = st.columns([1, 2])
    with ps_col1:
        ps_grp = st.selectbox("Group by", ["Day","Week","Month","Quarter","All Time"],
                              index=2, key="ps_grp")
    with ps_col2:
        if ps_grp == "Day":
            days_avail = sorted(df["Date"].dt.strftime("%Y-%m-%d").unique())
            ps_sel = st.selectbox("Select day", days_avail[::-1], key="ps_day")
            ps_df = df[df["Date"].dt.strftime("%Y-%m-%d") == ps_sel]
            ps_label = ps_sel
        elif ps_grp == "Week":
            weeks_avail = sorted(df["Week"].unique())
            ps_sel = st.selectbox("Select week starting", weeks_avail[::-1], key="ps_week")
            ps_df = df[df["Week"] == ps_sel]
            ps_label = f"Week of {ps_sel}"
        elif ps_grp == "Month":
            months_avail = df.groupby("Month")["MonthLabel"].first()
            month_options = months_avail.index.tolist()[::-1]
            month_labels  = months_avail.values.tolist()[::-1]
            ps_sel = st.selectbox("Select month", month_options,
                                  format_func=lambda m: months_avail.get(m, m),
                                  key="ps_month")
            ps_df = df[df["Month"] == ps_sel]
            ps_label = months_avail.get(ps_sel, ps_sel)
        elif ps_grp == "Quarter":
            qs_avail = sorted(df["Quarter"].unique())[::-1]
            ps_sel = st.selectbox("Select quarter", qs_avail, key="ps_quarter")
            ps_df = df[df["Quarter"] == ps_sel]
            ps_label = ps_sel
        else:
            ps_df = df
            ps_label = "All Time"
            st.empty()

    all_inits = {
        str(v) for col in ["Layout","Wire","FA"]
        for v in ps_df[col].unique()
        if pd.notna(v) and str(v).strip() not in ("", "NAN", "nan", "NaN")
    }

    master_rows = []
    for init in sorted(all_inits):
        disp = init_map.get(init, init)
        row = {"Person": disp}
        for col, label in [("Layout","Layout"), ("Wire","Wire"), ("FA","Final")]:
            sub = ps_df[ps_df[col] == init]
            n = len(sub)
            # Only count fails where this person's initial is in the reason
            my_fails = sub[
                (sub["Is_Fail"] == 1) &
                sub["Fail_Inits"].apply(
                    lambda inits: isinstance(inits, list) and init in inits)
            ].shape[0]
            fr = my_fails / n * 100 if n else None
            row[f"{label} Panels"] = n if n else 0
            row[f"{label} Fail%"]  = round(fr, 1) if fr is not None else None
        master_rows.append(row)

    master_df = pd.DataFrame(master_rows)
    master_df = master_df[master_df[["Layout Panels","Wire Panels","Final Panels"]].sum(axis=1) > 0]
    master_df = master_df.sort_values("Layout Panels", ascending=False)

    def color_fail(val):
        if pd.isna(val) or val == "":
            return ""
        if val >= 15:
            return f"color: {RED}; font-weight:bold"
        if val >= 8:
            return f"color: darkorange; font-weight:bold"
        return f"color: {GREEN}"

    st.dataframe(
        master_df.style.map(
            color_fail,
            subset=["Layout Fail%","Wire Fail%","Final Fail%"]
        ).format(
            {"Layout Fail%": "{:.1f}%", "Wire Fail%": "{:.1f}%", "Final Fail%": "{:.1f}%"},
            na_rep="—"
        ),
        use_container_width=True,
        hide_index=True
    )

# ══ TAB 6: Time Analysis ══════════════════════════════════════════════════════
with tab6:
    st.markdown(
        "<div class='section-header'>⏱ Time Analysis — Actual vs Routing Hours</div>",
        unsafe_allow_html=True)

    c_ref, c_info, _ = st.columns([1, 3, 2])
    with c_ref:
        if st.button("🔄 Refresh tracker data", key="ta_refresh"):
            load_time_data.clear()
            st.rerun()

    try:
        with st.spinner("Loading tracker data…"):
            tm = load_time_data()
            # Auto-clear stale cache if too few rows (e.g. cached before GRANT was applied)
            if len(tm) < 10 and not tm.empty:
                load_time_data.clear()
                tm = load_time_data()
        ta_err = None
    except Exception as _e:
        tm = pd.DataFrame()
        ta_err = str(_e)

    with c_info:
        if not ta_err and not tm.empty:
            st.caption(f"📊 {len(tm)} panels loaded from Supabase")

    # ── DEBUG (temporary) ──
    with st.expander("🔍 Debug — raw API values (click to expand)"):
        try:
            _dbg = _load_debug_sample()
            st.caption("These are the exact values the Supabase REST API returns before any Python processing:")
            for _r in _dbg:
                st.write(f"panel={_r.get('panel_number')} | layout_started_at={_r.get('layout_started_at')!r} | is_complete={_r.get('is_complete')!r}")
        except Exception as _de:
            st.write(f"Debug fetch error: {_de}")
        if not ta_err and not tm.empty:
            _db_null = tm['layout_started_at'].isna().sum()
            st.write(f"layout_started_at: {len(tm) - _db_null} non-null, {_db_null} null")
            st.write(f"is_complete True count: {tm['is_complete'].isin([True,'true','True','TRUE',1]).sum()}")

    if ta_err:
        st.error(f"Could not reach Supabase: {ta_err}")
    elif tm.empty:
        st.info("No tracker data yet — start scanning panels in the tracker app.")
    else:
        # ── Type conversions ──
        for _c in ['layout_secs', 'wire_secs', 'final_secs', 'tech_secs']:
            tm[_c] = pd.to_numeric(tm[_c], errors='coerce').fillna(0)
        tm['actual_total_hours'] = pd.to_numeric(tm['actual_total_hours'], errors='coerce')
        tm['routing_hours']      = pd.to_numeric(tm['routing_hours'],      errors='coerce')
        tm['layout_started_at']  = pd.to_datetime(tm['layout_started_at'], errors='coerce', utc=True)
        tm['tech_stopped_at']    = pd.to_datetime(tm['tech_stopped_at'],   errors='coerce', utc=True)
        tm['layout_hrs'] = tm['layout_secs'] / 3600
        tm['wire_hrs']   = tm['wire_secs']   / 3600
        tm['final_hrs']  = tm['final_secs']  / 3600
        tm['tech_hrs']   = tm['tech_secs']   / 3600
        try:
            tm['date_built'] = tm['layout_started_at'].dt.tz_convert('US/Eastern').dt.date
        except Exception:
            tm['date_built'] = tm['layout_started_at'].dt.date
        tm['panel_family'] = tm['panel_number'].astype(str).str.split('-').str[0]
        tm['is_over_time']  = tm['is_over_time'].fillna(False).astype(bool)

        # ── Join with Excel Pass/Fail ──
        tm['job_tail'] = tm['job_number'].astype(str).str.strip().str[-4:]
        fp_join = df[['Panel', 'Sales_Order', 'Pass_Fail', 'Is_Fail']].copy()
        fp_join['job_tail'] = fp_join['Sales_Order'].astype(str).str.strip().str[-4:]
        tm = tm.merge(
            fp_join.drop_duplicates(['Panel', 'job_tail']),
            left_on=['panel_number', 'job_tail'],
            right_on=['Panel', 'job_tail'], how='left'
        )

        # ── Date filter ──
        ta_max_date = pd.to_datetime(tm['date_built']).max()
        ta_min_date = pd.to_datetime(tm['date_built']).min()

        dp1, dp2, dp3 = st.columns([2, 2, 2])
        with dp1:
            ta_period = st.selectbox("Time period", [
                "Last 7 Days", "Last 30 Days", "Last 90 Days",
                "Last 6 Months", "Last Year", "By Month", "All Time",
            ], index=2, key="ta_period")
        with dp2:
            if ta_period == "By Month":
                ta_months = tm.dropna(subset=['date_built'])
                ta_months['_m'] = pd.to_datetime(ta_months['date_built']).dt.to_period('M')
                ta_month_opts = sorted(ta_months['_m'].unique().astype(str), reverse=True)
                ta_sel_month = st.selectbox("Month", ta_month_opts, key="ta_month")
            else:
                st.empty()
        with dp3:
            complete_only = st.checkbox("Completed panels only", value=True, key="ta_complete")

        _ta_period_days = {
            "Last 7 Days": 7, "Last 30 Days": 30, "Last 90 Days": 90,
            "Last 6 Months": 182, "Last Year": 365,
        }
        tmf = tm.copy()
        if ta_period in _ta_period_days:
            _ta_cutoff = ta_max_date - pd.Timedelta(days=_ta_period_days[ta_period])
            tmf = tmf[pd.to_datetime(tmf['date_built']) >= pd.to_datetime(_ta_cutoff)]
        elif ta_period == "By Month":
            tmf = tmf[
                pd.to_datetime(tmf['date_built']).dt.to_period('M').astype(str) == ta_sel_month
            ]

        if complete_only:
            tmf = tmf[tmf['is_complete'].isin([True, 'true', 'True', 'TRUE', 1])]

        # ── Content filters ──
        all_ta_workers = sorted({
            w for col in ['layout_worker', 'wire_worker', 'final_worker', 'tech_worker']
            for w in tmf[col].dropna().unique() if str(w).strip()
        })
        all_ta_panels = sorted(tmf['panel_number'].dropna().unique())

        fc1, fc2 = st.columns([2, 2])
        with fc1:
            sel_workers = st.multiselect("Filter by worker", all_ta_workers, key="ta_workers")
        with fc2:
            sel_panels  = st.multiselect("Panel number", all_ta_panels, key="ta_panels")

        if sel_panels:
            tmf = tmf[tmf['panel_number'].isin(sel_panels)]
        if sel_workers:
            mask = (
                tmf['layout_worker'].isin(sel_workers) |
                tmf['wire_worker'].isin(sel_workers)   |
                tmf['final_worker'].isin(sel_workers)  |
                tmf['tech_worker'].isin(sel_workers)
            )
            tmf = tmf[mask]

        if tmf.empty:
            st.info("No data matches the selected filters.")
        else:
            # ── KPIs ──
            n_panels      = len(tmf)
            n_with_rh     = int(tmf['routing_hours'].notna().sum())
            n_over        = int(tmf['is_over_time'].sum())
            pct_over      = n_over / n_with_rh * 100 if n_with_rh else 0
            sum_actual    = tmf['actual_total_hours'].sum()
            sum_routing   = tmf['routing_hours'].sum()

            kc1, kc2, kc3, kc4 = st.columns(4)
            for _col, _lbl, _val, _sub in [
                (kc1, "Panels Tracked",      f"{n_panels:,}",
                 f"{n_with_rh} with routing target"),
                (kc2, "Over Time",           f"{n_over:,}",
                 f"{pct_over:.1f}% of those with target"),
                (kc3, "Actual Working Time",
                 f"{int(sum_actual)}h {int((sum_actual % 1)*60)}m" if sum_actual else "—",
                 "summed wall-clock across panels"),
                (kc4, "Target Time",
                 f"{int(sum_routing)}h {int((sum_routing % 1)*60)}m" if sum_routing else "—",
                 "summed routing hours across panels"),
            ]:
                _col.markdown(f"""<div class='kpi-box'>
                    <div class='kpi-label'>{_lbl}</div>
                    <div class='kpi-value'>{_val}</div>
                    <div class='kpi-sub'>{_sub}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Scatter: actual vs routing ──
            scatter_df = tmf[
                tmf['routing_hours'].notna() & tmf['actual_total_hours'].notna()
            ].copy()
            scatter_df['Status'] = scatter_df['is_over_time'].map(
                {True: 'Over Time', False: 'On Time'})

            ch1, ch2 = st.columns([3, 2])
            with ch1:
                if not scatter_df.empty:
                    mx = max(scatter_df['routing_hours'].max(),
                             scatter_df['actual_total_hours'].max()) * 1.1
                    fig_sc = px.scatter(
                        scatter_df,
                        x='routing_hours', y='actual_total_hours',
                        color='Status',
                        color_discrete_map={'Over Time': RED, 'On Time': GREEN},
                        hover_data={
                            'panel_number': True, 'job_number': True,
                            'routing_hours': ':.3f', 'actual_total_hours': ':.3f',
                            'Status': False,
                        },
                        title='Actual vs Routing Hours',
                        labels={'routing_hours': 'Target (h)',
                                'actual_total_hours': 'Actual (h)',
                                'panel_number': 'Panel',
                                'job_number': 'Job'},
                    )
                    fig_sc.add_shape(type='line', x0=0, y0=0, x1=mx, y1=mx,
                                     line=dict(color=WARMGRAY, dash='dash', width=1))
                    fig_sc.add_annotation(
                        x=mx * 0.88, y=mx * 0.88,
                        text="on target", showarrow=False,
                        font=dict(color=WARMGRAY, size=10))
                    fig_sc.update_layout(
                        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
                        font=dict(family='Arial', color=NAVY),
                        title_font=dict(family='Georgia', color=NAVY, size=14),
                        legend=dict(title=None),
                        height=400, margin=dict(l=10, r=10, t=40, b=20),
                    )
                    st.plotly_chart(fig_sc, use_container_width=True)
                else:
                    st.info("No completed panels with both actual and routing hours.")

            with ch2:
                stage_avgs = []
                for _s, _col in [('Layout','layout_hrs'),('Wire','wire_hrs'),
                                  ('Final','final_hrs'),('Tech','tech_hrs')]:
                    vals = tmf[_col].replace(0, pd.NA).dropna()
                    if len(vals):
                        stage_avgs.append({'Stage': _s, 'Avg (h)': vals.mean()})
                if stage_avgs:
                    sa_df = pd.DataFrame(stage_avgs)
                    fig_st = px.bar(
                        sa_df, x='Avg (h)', y='Stage', orientation='h',
                        text=sa_df['Avg (h)'].map(lambda v: f"{v:.2f}h"),
                        color='Stage',
                        color_discrete_sequence=[BLUE, NAVY, '#4A90D9', '#7FAFD9'],
                        title='Avg Hours per Stage',
                    )
                    fig_st.update_traces(textposition='outside')
                    fig_st.update_layout(
                        plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
                        font=dict(family='Arial', color=NAVY),
                        title_font=dict(family='Georgia', color=NAVY, size=14),
                        showlegend=False,
                        height=300, margin=dict(l=10, r=60, t=40, b=20),
                    )
                    st.plotly_chart(fig_st, use_container_width=True)

            # ── Worker breakdown ──
            st.markdown("<div class='section-header'>Worker Time Breakdown</div>",
                        unsafe_allow_html=True)

            worker_rows = []
            for _stage, _wc, _hc in [
                ('Layout', 'layout_worker', 'layout_hrs'),
                ('Wire',   'wire_worker',   'wire_hrs'),
                ('Final',  'final_worker',  'final_hrs'),
                ('Tech',   'tech_worker',   'tech_hrs'),
            ]:
                sub = tmf[tmf[_wc].notna() & (tmf[_wc] != '') & (tmf[_hc] > 0)]
                for worker, grp in sub.groupby(_wc):
                    hrs = grp[_hc]
                    worker_rows.append({
                        'Stage': _stage, 'Worker': worker,
                        'Panels': len(grp),
                        'Avg': hrs.mean(),
                        'Min': hrs.min(),
                        'Max': hrs.max(),
                    })

            if worker_rows:
                wdf = pd.DataFrame(worker_rows).sort_values(['Stage', 'Worker'])
                fig_w = px.bar(
                    wdf, x='Worker', y='Avg', color='Stage', barmode='group',
                    text=wdf['Avg'].map(lambda v: f"{int(v*60)}m"),
                    color_discrete_sequence=[BLUE, NAVY, '#4A90D9', '#7FAFD9'],
                    title='Avg Stage Time by Worker',
                    labels={'Avg': 'Avg Hours'},
                )
                fig_w.update_traces(textposition='outside')
                fig_w.update_layout(
                    plot_bgcolor=OFFWHITE, paper_bgcolor=OFFWHITE,
                    font=dict(family='Arial', color=NAVY),
                    title_font=dict(family='Georgia', color=NAVY, size=14),
                    height=360, margin=dict(l=10, r=10, t=40, b=20),
                )
                st.plotly_chart(fig_w, use_container_width=True)
                # Format table columns as H:M:S
                wdf_display = wdf.copy()
                for _c in ['Avg', 'Min', 'Max']:
                    wdf_display[_c] = wdf_display[_c].map(
                        lambda h: f"{int(h*3600)//3600}:{(int(h*3600)%3600)//60:02d}:{int(h*3600)%60:02d}"
                    )
                wdf_display.rename(columns={'Avg':'Avg H:M:S','Min':'Min H:M:S','Max':'Max H:M:S'},
                                   inplace=True)
                st.dataframe(wdf_display, use_container_width=True, hide_index=True)
            else:
                st.info("No stage-level worker data yet.")

            # ── Time Breakdown ──
            def _hms(h):
                if pd.isna(h) or h == 0:
                    return '—'
                s = int(round(h * 3600))
                return f"{s//3600}:{(s%3600)//60:02d}:{s%60:02d}"

            st.markdown("<div class='section-header'>Time Breakdown</div>",
                        unsafe_allow_html=True)

            completed = tmf[tmf['actual_total_hours'].notna()].copy()

            if completed.empty:
                st.info("No completed panels to analyze.")
            else:
                def _hms_signed(h):
                    sign = '▲ ' if h > 0 else ('▼ ' if h < 0 else '')
                    return sign + _hms(abs(h))

                avg_layout = completed['layout_hrs'].replace(0, pd.NA).mean()
                avg_wire   = completed['wire_hrs'].replace(0, pd.NA).mean()
                avg_final  = completed['final_hrs'].replace(0, pd.NA).mean()
                avg_tech   = completed['tech_hrs'].replace(0, pd.NA).mean()
                avg_total  = completed['actual_total_hours'].mean()

                unique_targets = completed['routing_hours'].dropna().unique()
                single_target  = unique_targets[0] if len(unique_targets) == 1 else None
                target_label   = _hms(single_target) if single_target else '— (mixed)'
                delta          = avg_total - single_target if single_target else None
                is_over        = delta is not None and delta > (1 / 60)

                # Metric cards row
                mc1, mc2, mc3, mc4 = st.columns(4)
                for _col, _lbl, _val in [
                    (mc1, 'Avg Layout',  _hms(avg_layout) if pd.notna(avg_layout) else '—'),
                    (mc2, 'Avg Wire',    _hms(avg_wire)   if pd.notna(avg_wire)   else '—'),
                    (mc3, 'Avg Final',   _hms(avg_final)  if pd.notna(avg_final)  else '—'),
                    (mc4, 'Avg Tech',    _hms(avg_tech)   if pd.notna(avg_tech)   else '—'),
                ]:
                    _col.markdown(f"""<div class='kpi-box'>
                        <div class='kpi-label'>{_lbl}</div>
                        <div class='kpi-value' style='font-size:22px'>{_val}</div>
                    </div>""", unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # Target / Actual / Delta summary
                ta1, ta2, ta3 = st.columns(3)
                ta_color = RED if is_over else GREEN
                ta_bg    = '#fff0f0' if is_over else '#f0fff4'
                for _col, _lbl, _val, _color, _bg in [
                    (ta1, 'Target (routing hours)', target_label, NAVY, TINT),
                    (ta2, 'Avg Completion',
                     _hms(avg_total) if pd.notna(avg_total) else '—', ta_color, ta_bg),
                    (ta3, 'Delta',
                     _hms_signed(delta) if delta is not None else '—', ta_color, ta_bg),
                ]:
                    _col.markdown(f"""<div class='kpi-box' style='background:{_bg};color:{_color}'>
                        <div class='kpi-label' style='color:{_color}'>{_lbl}</div>
                        <div class='kpi-value' style='font-size:22px;color:{_color}'>{_val}</div>
                    </div>""", unsafe_allow_html=True)

                # ── Quickest team ──
                st.markdown("<br>", unsafe_allow_html=True)
                team_cols = ['layout_worker', 'wire_worker', 'final_worker', 'tech_worker']
                team_data = completed.dropna(subset=['actual_total_hours']).copy()
                team_data['Team'] = (
                    team_data['layout_worker'].fillna('—') + ' / ' +
                    team_data['wire_worker'].fillna('—')   + ' / ' +
                    team_data['final_worker'].fillna('—')  + ' / ' +
                    team_data['tech_worker'].fillna('—')
                )
                team_summary = (
                    team_data.groupby('Team')
                    .agg(
                        Panels=('actual_total_hours', 'count'),
                        Avg_h=('actual_total_hours', 'mean'),
                        layout_h=('layout_hrs', 'mean'),
                        wire_h=('wire_hrs', 'mean'),
                        final_h=('final_hrs', 'mean'),
                        tech_h=('tech_hrs', 'mean'),
                    )
                    .reset_index()
                    .sort_values('Avg_h')
                )

                if len(team_summary) > 0:
                    st.markdown(
                        "<div style='font-family:Georgia;font-size:14px;font-weight:bold;"
                        f"color:{NAVY};margin:8px 0 4px'>Quickest Teams</div>",
                        unsafe_allow_html=True)
                    st.caption(
                        "Layout by / Wire by / Final by / Tech by — sorted by avg completion time. "
                        "Green = fastest.")

                    ts_display = team_summary.copy()
                    ts_display['Avg Total'] = ts_display['Avg_h'].map(_hms)
                    ts_display['Layout']    = ts_display['layout_h'].map(_hms)
                    ts_display['Wire']      = ts_display['wire_h'].map(_hms)
                    ts_display['Final']     = ts_display['final_h'].map(_hms)
                    ts_display['Tech']      = ts_display['tech_h'].map(_hms)
                    if single_target:
                        ts_display['vs Target'] = ts_display['Avg_h'].map(
                            lambda h: _hms_signed(h - single_target))
                    ts_display.drop(columns=['Avg_h','layout_h','wire_h','final_h','tech_h'],
                                    inplace=True)
                    ts_display.rename(columns={'Panels': '# Panels'}, inplace=True)

                    n_teams = len(team_summary)
                    # Rank-based colors: green (fastest) → red (slowest)
                    rank_colors = ['#f0fff4','#f7fff0','#fffde7','#fff3e0','#fff0f0']

                    def _hl_team(row):
                        rank = ts_display.index.get_loc(row.name)
                        idx  = min(rank, len(rank_colors) - 1) if n_teams > 1 else 0
                        if n_teams > 1:
                            idx = round(rank / (n_teams - 1) * (len(rank_colors) - 1))
                        bg = rank_colors[idx]
                        return [f'background-color: {bg}'] * len(row)

                    st.dataframe(
                        ts_display.style.apply(_hl_team, axis=1),
                        use_container_width=True, hide_index=True,
                    )

            # ── Panel detail table ──
            st.markdown("<div class='section-header'>Panel Detail</div>",
                        unsafe_allow_html=True)

            detail = tmf[[c for c in [
                'panel_number', 'job_number', 'panel_seq',
                'routing_hours', 'actual_total_hours',
                'layout_hrs', 'layout_worker',
                'wire_hrs',   'wire_worker',
                'final_hrs',  'final_worker',
                'tech_hrs',   'tech_worker',
                'is_over_time', 'Pass_Fail', 'date_built',
            ] if c in tmf.columns]].copy()

            detail['delta_h'] = (
                detail['actual_total_hours'] - detail['routing_hours']
            )
            detail = detail.sort_values(
                ['job_number', 'panel_seq'], na_position='last')

            for _hcol, _label in [
                ('layout_hrs', 'Layout H:M:S'),
                ('wire_hrs',   'Wire H:M:S'),
                ('final_hrs',  'Final H:M:S'),
                ('tech_hrs',   'Tech H:M:S'),
            ]:
                if _hcol in detail.columns:
                    detail[_label] = detail[_hcol].map(_hms)
                    detail.drop(columns=[_hcol], inplace=True)

            detail['Target'] = detail['routing_hours'].map(
                lambda h: _hms(h) if pd.notna(h) else '—')
            detail['Actual'] = detail['actual_total_hours'].map(
                lambda h: _hms(h) if pd.notna(h) else '—')
            detail['Delta'] = detail['delta_h'].map(
                lambda h: ('▲ ' if h > 0 else '') + _hms(abs(h))
                if pd.notna(h) else '—')

            detail.drop(columns=['routing_hours', 'actual_total_hours', 'delta_h'],
                        inplace=True)
            detail.rename(columns={
                'panel_number':  'Panel #',
                'job_number':    'Job #',
                'panel_seq':     'Seq',
                'layout_worker': 'Layout by',
                'wire_worker':   'Wire by',
                'final_worker':  'Final by',
                'tech_worker':   'Tech by',
                'is_over_time':  'Over?',
                'Pass_Fail':     'Q Result',
                'date_built':    'Date',
            }, inplace=True)

            # Column order: identity cols, then paired stage cols, then summary
            _id_cols    = [c for c in ['Panel #','Job #','Seq','Date','Q Result','Over?']
                           if c in detail.columns]
            _stage_cols = ['Layout by', 'Layout H:M:S',
                           'Wire by',   'Wire H:M:S',
                           'Final by',  'Final H:M:S',
                           'Tech by',   'Tech H:M:S']
            _stage_cols = [c for c in _stage_cols if c in detail.columns]
            _sum_cols   = [c for c in ['Target','Actual','Delta']
                           if c in detail.columns]
            detail = detail[_id_cols + _stage_cols + _sum_cols]

            def _hl_over(row):
                bg = 'background-color: #fff0f0' if row.get('Over?') else ''
                return [bg] * len(row)

            st.dataframe(
                detail.style.apply(_hl_over, axis=1),
                use_container_width=True, hide_index=True,
            )


# ── Version history (bottom-left) ─────────────────────────────────────────────
st.markdown("""
<style>
.version-history {
    position: fixed;
    bottom: 12px;
    left: 12px;
    z-index: 9999;
    font-family: Arial, sans-serif;
    font-size: 10px;
    color: #aaa;
    line-height: 1.5;
    max-width: 260px;
}
.version-history summary {
    cursor: pointer;
    font-size: 10px;
    color: #bbb;
    user-select: none;
}
.version-history summary:hover { color: #888; }
</style>
<div class="version-history">
  <details>
    <summary>v1.12 — changelog</summary>
    <b>v1.12</b> Time Analysis tab: actual vs routing hours, worker breakdown, over-time flag<br>
    <b>v1.11</b> Only attributed fails count against FPY% &amp; fail rates per person<br>
    <b>v1.10</b> Person Summary: own Day/Week/Month/Quarter filter<br>
    <b>v1.9</b> Failure Caused By column in drilldown; attribution bug fix<br>
    <b>v1.8</b> Click bar in Layout/Wire/Final to see panel list<br>
    <b>v1.7</b> Job # added to fail drill-downs<br>
    <b>v1.6</b> Final by Person tab added<br>
    <b>v1.5</b> Trends rework: panels chart, QoQ/MoM toggle, master table<br>
    <b>v1.4</b> FPY% chart title hover tooltip<br>
    <b>v1.3</b> Fiscal quarters starting Oct 1 (FY## Q#)<br>
    <b>v1.2</b> Handle comma-separated initials e.g. (CO, IM)<br>
    <b>v1.1</b> Multi-initials, click-to-drilldown, expanded period options<br>
    <b>v1.0</b> Initial release
  </details>
</div>
""", unsafe_allow_html=True)
