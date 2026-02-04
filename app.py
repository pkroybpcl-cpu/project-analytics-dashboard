import os
from datetime import date
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# =========================
# Page config + theme
# =========================

st.set_page_config(page_title="Project Analytics Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<style>
 .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }

 /* Metric card styling */
 [data-testid="stMetric"] {
    background: linear-gradient(135deg, rgba(255,255,255,0.95), rgba(248,250,252,0.95));
    border: 2px solid rgba(59, 130, 246, 0.3);
    border-radius: 12px;
    padding: 16px 18px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
 }

 [data-testid="stMetric"]:hover {
    border-color: rgba(59, 130, 246, 0.6);
    box-shadow: 0 6px 16px rgba(59, 130, 246, 0.15);
    transform: translateY(-2px);
    transition: all 0.3s ease;
 }

 [data-testid="stMetricLabel"] {
    font-weight: 600;
    font-size: 0.9rem;
    color: rgba(49, 51, 63, 0.7);
 }

 [data-testid="stMetricValue"] {
    font-size: 1.8rem;
    font-weight: 700;
    color: #1e293b;
 }
            
 .card {
    background: rgba(255,255,255,0.92);
    border: 1px solid rgba(49, 51, 63, 0.10);
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.06);
}

.sb { font-weight: 700; margin: 8px 0 6px; }
</style>
""", unsafe_allow_html=True)
st.title("Project Analytics Dashboard")
st.markdown("<h4 style='margin-top: -20px; color: #666;'>⏱️ Real-time Gantt Visualization • 📈 Progress Tracking • ⚠️ Lag Analysis</h4>", unsafe_allow_html=True)

# =========================
# Sidebar controls
# =========================

st.sidebar.header("📂 Data Source")
default_path = "sample.xlsx"
path = st.sidebar.text_input("📄 Excel file path (.xlsx)", value=default_path)

sheet_name = st.sidebar.text_input("📑 Sheet name (optional)", value="")

st.sidebar.header("🔄 Refresh")
auto_refresh = st.sidebar.checkbox("⏱️ Auto-refresh", value=True)
refresh_seconds = st.sidebar.slider("⏲️ Refresh interval (seconds)", 2, 30, 5)

if auto_refresh:
    st_autorefresh(interval=refresh_seconds * 1000, key="gantt_refresh")

st.sidebar.header("🎨 Visualization")
color_by = st.sidebar.selectbox("🎯 Color Gantt by", ["project", "vendor", "status"], index=0)

st.sidebar.header("🔍 Filters")
vendor_filter = st.sidebar.text_input("🏷️ Vendor contains (optional)", "")
project_filter = st.sidebar.text_input("📁 Project contains (optional)", "")

st.sidebar.header("📅 Lag Definition")
today = st.sidebar.date_input("📆 Today (for lag calculations)", value=date.today())

# =========================
# Helpers
# =========================

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Handle trailing spaces and case; match your file reliably.
    """
# Helpers
# =========================

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Handle trailing spaces and case; match your file reliably.
    """
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]  # strips trailing spaces like "activity "
    return df

def coerce_percent(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, str):
        s = x.strip().replace("%", "")
        try:
            return float(s)
        except Exception:
            return np.nan
    try:
        return float(x)
    except Exception:
        return np.nan

def load_excel(file_path: str, sheet: str | None) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=(sheet if sheet else 0), engine="openpyxl")
    df = normalize_columns(df)

    # Expected columns (after stripping/lowering)
    # vendor, project, activity, start, end, %complete
    required = ["vendor", "project", "activity", "start", "end", "%complete"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing expected columns: {missing}\n"
            f"Found columns: {list(df.columns)}\n"
            f"Tip: Ensure headers are vendor, project, activity, start, end, %complete"
        )

    # Forward-fill vendor & project (common when cells were merged)
    df["vendor"] = df["vendor"].ffill()
    df["project"] = df["project"].ffill()

    # Clean types
    df["activity"] = df["activity"].astype(str).str.strip()
    df["start"] = pd.to_datetime(df["start"], errors="coerce")
    df["end"] = pd.to_datetime(df["end"], errors="coerce")
    df["pct_complete"] = df["%complete"].apply(coerce_percent)

    # Drop rows missing essential fields
    df = df.dropna(subset=["vendor", "project", "activity", "start", "end"])
    df["pct_complete"] = df["pct_complete"].fillna(0).clip(0, 100)

    # Fix inverted dates
    swapped = df["end"] < df["start"]
    if swapped.any():
        df.loc[swapped, ["start", "end"]] = df.loc[swapped, ["end", "start"]].values

    return df

def add_schedule_metrics(df: pd.DataFrame, today_dt: pd.Timestamp) -> pd.DataFrame:
    df = df.copy()

    # Durations
    df["duration_days"] = (df["end"] - df["start"]).dt.days.clip(lower=0)

    # Actual progress
    df["actual_progress"] = (df["pct_complete"] / 100.0).clip(0, 1)

    # Planned progress based on time elapsed between start and end (no baseline needed)
    # planned = (today-start)/(end-start), clipped to [0,1]
    denom = (df["end"] - df["start"]).dt.total_seconds()
    numer = (today_dt - df["start"]).dt.total_seconds()

    planned = np.where(denom > 0, numer / denom, np.where(today_dt >= df["end"], 1.0, 0.0))
    df["planned_progress"] = np.clip(planned, 0, 1)

    # Progress gap: positive = behind schedule (planned > actual)
    df["behind_by_pct_points"] = (df["planned_progress"] - df["actual_progress"]) * 100.0

    # Lag / overdue days: only if incomplete AND today > end
    df["overdue_days"] = np.where(
        (df["pct_complete"] < 100) & (today_dt > df["end"]),
        (today_dt - df["end"]).dt.days,
        0
    ).astype(int)

    # Status
    df["status"] = np.select(
        [
            df["pct_complete"] >= 100,
            df["overdue_days"] > 0,
            df["behind_by_pct_points"] > 5,  # threshold you can tune
        ],
        [
            "Completed",
            "Overdue",
            "Behind",
        ],
        default="On track"
    )

    return df

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    out = df
    if vendor_filter.strip():
        out = out[out["vendor"].astype(str).str.lower().str.contains(vendor_filter.strip().lower(), na=False)]
    if project_filter.strip():
        out = out[out["project"].astype(str).str.lower().str.contains(project_filter.strip().lower(), na=False)]
    return out

def weighted_project_summary(df: pd.DataFrame) -> pd.DataFrame:
    # Weight by duration; fallback weight=1 if duration=0
    w = df["duration_days"].replace(0, 1)

    def wavg(series):
        return np.average(series, weights=w.loc[series.index])

    grp = df.groupby(["vendor", "project"], dropna=False)

    summary = grp.apply(lambda g: pd.Series({
        "activities": len(g),
        "overdue_activities": int((g["overdue_days"] > 0).sum()),
        "max_overdue_days": int(g["overdue_days"].max()),
        "avg_overdue_days_overdue_only": float(
            g.loc[g["overdue_days"] > 0, "overdue_days"].mean()
        ) if (g["overdue_days"] > 0).any() else 0.0,
        "weighted_actual_%": float(
            np.average(g["actual_progress"], weights=g["duration_days"].replace(0, 1)) * 100.0
        ),
        "weighted_planned_%": float(
            np.average(g["planned_progress"], weights=g["duration_days"].replace(0, 1)) * 100.0
        ),
        "weighted_behind_%pts": float(
            (np.average(g["planned_progress"], weights=g["duration_days"].replace(0, 1))
            - np.average(g["actual_progress"], weights=g["duration_days"].replace(0, 1))) * 100.0
        ),
        "start_min": g["start"].min(),
        "end_max": g["end"].max(),
    })).reset_index()

     # Nice formatting columns
    summary = summary.sort_values(["max_overdue_days", "weighted_behind_%pts"], ascending=[False, False])
    return summary


# =========================
# Main load + compute
# =========================

# Top status bar with file info on left and refresh controls on right
status_left, status_right = st.columns([2, 1])

with status_left:
    st.subheader("Status")
    st.write(f"*File:* {path}")
    if not os.path.exists(path):
        st.warning("File not found. Put it at the path above or change the path.")

with status_right:
    if os.path.exists(path):
        st.write("")  # Spacer to align with subheader
        st.write(f"*Last modified:* {pd.to_datetime(os.path.getmtime(path), unit='s').strftime('%Y-%m-%d %H:%M:%S')}")
        if st.button("🔄 Refresh now"):
            st.rerun()

st.divider()


# Main content area
try:
    df_raw = load_excel(path, sheet_name.strip() or None)
    df = add_schedule_metrics(df_raw, pd.Timestamp(today))
    df = apply_filters(df)

    if df.empty:
        st.warning("No rows after filters.")
        st.stop()

    # =========================
    # KPIs
    # =========================

    total = len(df)
    overdue_count = int((df["overdue_days"] > 0).sum())
    behind_count = int((df["status"] == "Behind").sum())
    max_overdue = int(df["overdue_days"].max())
    avg_behind = float(df["behind_by_pct_points"].clip(lower=0).mean())

    k1, k2, k3, k4 = st.columns(4)

    k1.metric("Activities", f"{total}")
    k2.metric("Overdue activities", f"{overdue_count}")
    k3.metric("Behind (not overdue)", f"{behind_count}")
    k4.metric("Max overdue (days)", f"{max_overdue}")

    st.caption(
        "Definitions: *Overdue* = %Complete < 100 and Today > End date. "
        "*Behind* = planned progress based on dates is > actual progress by > 5 percentage points (tunable)."
    )

    st.divider()


    # =========================
    # Tabs
    # =========================

    tab_gantt, tab_risks, tab_summary = st.tabs([
        "📊 Gantt Chart",
        "⚠️ Risks & Delays",
        "📌 Project Summary"
    ])


    # =========================
    # Gantt tab
    # =========================

    with tab_gantt:
        st.subheader("Gantt Chart")

        gantt_df = df.copy()
        gantt_df["label"] = gantt_df["project"].astype(str) + " • " + gantt_df["activity"].astype(str)

        fig = px.timeline(
            gantt_df,
            x_start="start",
            x_end="end",
            y="label",
            color=color_by,
            hover_data=["vendor", "project", "activity", "pct_complete", "status", "overdue_days", "behind_by_pct_points"],
            template="plotly_white",
        )

        fig.update_yaxes(autorange="reversed", showgrid=True, gridwidth=1, gridcolor="rgba(200,200,200,0.3)")
        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="rgba(200,200,200,0.3)")

        fig.update_layout(
            height=max(560, 26 * len(gantt_df)),
            margin=dict(l=10, r=10, t=40, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            hoverlabel=dict(namelength=-1),
            xaxis=dict(
                showgrid=True,
                gridwidth=1,
                gridcolor="rgba(200,200,200,0.3)",
            ),
            yaxis=dict(
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(200,200,200,0.3)',
            ),
        )

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


    # =========================
    # Risks tab (Overdue + Behind Schedule)
    # =========================

    with tab_risks:
        c1, c2 = st.columns(2)

        with c1:
            st.subheader("Overdue days by activity (only overdue)")
            overdue = df[df["overdue_days"] > 0].copy()
            if overdue.empty:
                st.info("No overdue activities based on the selected 'Today' date.")
            else:
                overdue["label"] = overdue["project"].astype(str) + " • " + overdue["activity"].astype(str)
                fig_overdue = px.bar(
                    overdue.sort_values("overdue_days", ascending=False),
                    x="overdue_days",
                    y="label",
                    hover_data=["vendor", "project", "pct_complete", "end"],
                    orientation="h",
                    template="plotly_white",
                )

                fig_overdue.update_layout(
                    height=max(480, 26 * len(overdue)),
                    margin=dict(l=10, r=10, t=40, b=10),
                )

                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.plotly_chart(fig_overdue, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.subheader("Behind schedule (% points) by activity")
            behind = df.copy()
            behind["behind_pos"] = behind["behind_by_pct_points"].clip(lower=0)
            behind["label"] = behind["project"].astype(str) + " • " + behind["activity"].astype(str)

            fig_behind = px.bar(
                behind.sort_values("behind_pos", ascending=False),
                x="behind_pos",
                y="label",
                hover_data=["vendor", "project", "pct_complete", "start", "end"],
                orientation="h",
                template="plotly_white",
            )

            fig_behind.update_layout(
                height=max(480, 26 * len(behind)),
                margin=dict(l=10, r=10, t=40, b=10),
            )

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.plotly_chart(fig_behind, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)


    # =========================
    # Project summary tab
    # =========================

    with tab_summary:
        st.subheader("Project summary (weighted by activity duration)")
        summary = weighted_project_summary(df)

        # Make nicer display
        display = summary.copy()
        display["weighted_actual_%"] = display["weighted_actual_%"].round(1)
        display["weighted_planned_%"] = display["weighted_planned_%"].round(1)
        display["weighted_behind_%pts"] = display["weighted_behind_%pts"].round(1)
        display["avg_overdue_days_overdue_only"] = display["avg_overdue_days_overdue_only"].round(1)
        display["start_min"] = display["start_min"].dt.date
        display["end_max"] = display["end_max"].dt.date

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.dataframe(display, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.download_button(
            "⬇️ Download summary (CSV)",
            data=display.to_csv(index=False).encode("utf-8"),
            file_name="project_summary.csv",
            mime="text/csv",
        )

        with st.expander("📋 Raw data with computed fields"):
            st.dataframe(
                df[
                    [
                        "vendor",
                        "project",
                        "activity",
                        "start",
                        "end",
                        "pct_complete",
                        "status",
                        "overdue_days",
                        "behind_by_pct_points",
                        "planned_progress",
                        "actual_progress",
                    ]
                ].sort_values(["project", "start"]),
                use_container_width=True,
            )

except Exception as e:
    st.error(str(e))
    st.info(
    "Your Excel must have columns: vendor, project, activity, start, end, %complete "
    "(trailing spaces are okay). Vendor/Project can be blank in rows (merged-cell style) - it will forward-fill."
    )
