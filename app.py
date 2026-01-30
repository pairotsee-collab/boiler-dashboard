# -*- coding: utf-8 -*-
"""
Boiler Fuel–Energy Dashboard (Realtime) — Cloud Ready + OneDrive/SharePoint
- แหล่งข้อมูลรองรับ 4 โหมด:
  1) อัปโหลดไฟล์ Excel (.xlsx)
  2) URL สาธารณะ (OneDrive/SharePoint/GitHub) → direct download URL
  3) OneDrive/SharePoint (Graph API) → ไม่ต้อง public link (ต้องใส่ secrets)
  4) พาธภายใน (ใช้บน LAN/On-Prem เท่านั้น)
"""
import os
import re
import base64
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

try:
    import requests
except Exception:
    requests = None

# ---------------- Page setup ----------------
st.set_page_config(page_title="Boiler Fuel–Energy Dashboard (Realtime)", layout="wide")
st.title("📊 Boiler Fuel–Energy Dashboard (Realtime)")
st.caption("Feed = Cost fuel/ตันบรรจุ • Steam/Feed = น้ำ m3/ตันบรรจุ • Fuel/ตันบรรจุ = ΣFuel/Σตัน • Cloud Ready + OneDrive")

# ---------------- Defaults ----------------
DEFAULT_TARGETS = {
    "cost_baht_per_ton_feed": 75.0,
    "steam_per_feed": 0.18,
    "cost_baht_per_ton_steam": 420.0,
}
FUEL_ORDER = ["woodchip_kg", "cashew_shell_kg", "furniture_wood_kg"]
FUEL_COLORS = {
    "woodchip_kg": "#4A90E2",
    "cashew_shell_kg": "#E74C3C",
    "furniture_wood_kg": "#82C6FF",
}

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("⚙️ ตั้งค่า/แหล่งข้อมูล")

    data_source = st.radio(
        "เลือกแหล่งข้อมูล",
        options=[
            "อัปโหลด Excel (.xlsx)",
            "URL สาธารณะ (OneDrive/SharePoint/GitHub)",
            "OneDrive/SharePoint (Graph API)",
            "พาธภายใน (LAN เท่านั้น)",
        ],
        index=0,
        help="โหมดคลาวด์แนะนำ 'อัปโหลด' หรือ 'URL สาธารณะ'; ข้อมูลอ่อนไหวใช้ Graph API",
    )

    file_obj = None
    direct_url = ""
    graph_share_url = ""
    file_path = None

    if data_source == "อัปโหลด Excel (.xlsx)":
        file_obj = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"], accept_multiple_files=False)
    elif data_source == "URL สาธารณะ (OneDrive/SharePoint/GitHub)":
        direct_url = st.text_input(
            "วางลิงก์ดาวน์โหลดตรง (direct download URL)",
            placeholder="https://.../Excel.xlsx?download=1",
            help="OneDrive/SharePoint: ตั้งแชร์เป็น Anyone แล้วต่อท้าย ?download=1 หรือใช้ลิงก์ download ของ GitHub raw",
        )
    elif data_source == "OneDrive/SharePoint (Graph API)":
        graph_share_url = st.text_input(
            "วางลิงก์ Share (ไม่ต้อง public)",
            placeholder="ลิงก์จากปุ่ม Share ของ OneDrive/SharePoint",
            help="ต้องตั้งค่า TENANT_ID/CLIENT_ID/CLIENT_SECRET ใน Secrets ก่อน",
        )
    else:
        file_path = st.text_input(
            "พาธไฟล์ภายใน (ใช้ได้เฉพาะรันใน LAN)",
            value=r"Fuel Dashboard Boiler banbung.xlsx",
            help=r"ตัวอย่าง: D:\\Boiler\\data.xlsx หรือ \\SERVER01\\boiler\\fuel.xlsx",
        )

    auto_refresh = st.checkbox("รีเฟรชอัตโนมัติ", value=True)
    refresh_sec = st.number_input("ช่วงเวลารีเฟรช (วินาที)", min_value=2, max_value=120, value=5, step=1)
    zero_is_missing = st.checkbox("ถือว่า 0 = ไม่มีข้อมูล (เฉพาะ KPI บางตัว)", value=True)

    st.divider()
    st.subheader("🎯 เป้าหมาย")
    targets = {
        "cost_baht_per_ton_feed": st.number_input("Baht/Ton feed", value=float(DEFAULT_TARGETS["cost_baht_per_ton_feed"]), step=1.0),
        "steam_per_feed": st.number_input("Ton steam/Ton feed", value=float(DEFAULT_TARGETS["steam_per_feed"]), step=0.01, format="%0.3f"),
        "cost_baht_per_ton_steam": st.number_input("Baht/Ton steam", value=float(DEFAULT_TARGETS["cost_baht_per_ton_steam"]), step=1.0),
    }

    st.subheader("🗓️ รูปแบบแกนเวลา")
    time_grain = st.radio("ความละเอียดของเวลา", options=["รายวัน", "รายสัปดาห์", "รายเดือน"], index=0, horizontal=True)
    date_tick_fmt = st.selectbox(
        "รูปแบบแสดงผลวันที่",
        options=["%d/%m", "%-d %b", "%d %b %Y", "%b %Y"],
        index=1 if time_grain == "รายวัน" else (3 if time_grain == "รายเดือน" else 2),
        help="เลือกรูปแบบให้เหมาะกับความละเอียดเวลา",
    )
    tick_angle = st.slider("หมุนป้ายวันที่ (องศา)", min_value=0, max_value=90, value=45, step=5)
    tick_every = st.number_input(
        "แสดงป้ายทุก N หน่วยเวลา", min_value=1, max_value=31,
        value=2 if time_grain == "รายวัน" else 1, step=1,
        help="รายวัน N=2 = วันเว้นวัน / รายเดือน N=2 = 2 เดือนครั้ง",
    )
    show_spike = st.checkbox("แสดงเส้นชี้ตำแหน่ง (Spike line)", value=True)

    st.caption(f"⏱️ เวลาเซิร์ฟเวอร์: {datetime.now():%H:%M:%S}")

# Auto-refresh
if auto_refresh:
    components.html(
        f"""
        <script>setTimeout(function(){{window.location.reload(1);}}, {int(refresh_sec)*1000});</script>
        """,
        height=0,
    )

# ---------------- Helpers ----------------
PARENS_MAP = str.maketrans({"（": "(", "）": ")", "【": "[", "】": "]"})

def normalize_col(name: str) -> str:
    name = (name or "").translate(PARENS_MAP)
    name = re.sub(r"\s+", " ", str(name).strip())
    return name

RAW_TO_STD = {
    "Cost engergy (Baht/Ton feed) Target 75": "cost_baht_per_ton_feed_orig",
    "Cost energy (Baht/Ton feed) Target 75": "cost_baht_per_ton_feed_orig",
    "Cost engergy (Ton steam/Ton feed)Target 0.18": "steam_per_feed_orig",
    "Cost energy (Ton steam/Ton feed) Target 0.18": "steam_per_feed_orig",
    "Cost engergy(Baht/Ton steam)Target 420": "cost_baht_per_ton_steam",
    "Cost energy (Baht/Ton steam) Target 420": "cost_baht_per_ton_steam",
    "ไม้สับ (กก.)": "woodchip_kg",
    "เปลือกมะม่วงหิมพานต์ (กก.)": "cashew_shell_kg",
    "ไม้เฟอร์นิเจิร์บด (กก.)": "furniture_wood_kg",
    "ยอดบรรจุ(ตัน)": "packed_ton",
    "Cost fuel (Baht)": "cost_fuel_baht",
    "ใช้น้ำ m3": "water_m3",
    "น้ำ m3": "water_m3",
    "ปริมาณน้ำ (m3)": "water_m3",
    "Water m3": "water_m3",
    "Water (m3)": "water_m3",
    "usage water m3": "water_m3",
}

@st.cache_data(show_spinner=False)
def _read_excel_bytes(xbytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(xbytes), header=0, engine="openpyxl").dropna(how="all")
    df.columns = [normalize_col(c) for c in df.columns]

    # rename
    rename_map = {}
    for raw, std in RAW_TO_STD.items():
        key = normalize_col(raw)
        if key in df.columns:
            rename_map[key] = std
    df = df.rename(columns=rename_map)

    # types & dates
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    for c in [c for c in df.columns if c != "Date"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.loc[:, ~df.columns.duplicated()].copy().reset_index(drop=True)
    if "Date" in df.columns:
        df = df[~df["Date"].isna()].copy()

    # computed columns
    if {"cost_fuel_baht", "packed_ton"}.issubset(df.columns):
        with np.errstate(divide="ignore", invalid="ignore"):
            df["cost_baht_per_ton_feed"] = df["cost_fuel_baht"] / df["packed_ton"]
    else:
        df["cost_baht_per_ton_feed"] = df.get("cost_baht_per_ton_feed_orig", np.nan)

    if {"water_m3", "packed_ton"}.issubset(df.columns):
        with np.errstate(divide="ignore", invalid="ignore"):
            df["steam_per_feed"] = df["water_m3"] / df["packed_ton"]
        df.attrs["spf_from_water"] = True
    else:
        df["steam_per_feed"] = df.get("steam_per_feed_orig", np.nan)
        df.attrs["spf_from_water"] = False

    if "cost_baht_per_ton_steam" not in df.columns:
        df["cost_baht_per_ton_steam"] = np.nan
    need_fill = df["cost_baht_per_ton_steam"].isna()
    if {"cost_baht_per_ton_feed", "steam_per_feed"}.issubset(df.columns):
        with np.errstate(divide="ignore", invalid="ignore"):
            calc = df["cost_baht_per_ton_feed"] / df["steam_per_feed"]
        df.loc[need_fill, "cost_baht_per_ton_steam"] = calc[need_fill]
        df.attrs["steam_cost_fallback"] = True
    else:
        df.attrs["steam_cost_fallback"] = False

    return df

# --- Graph API helpers ---

def _get_graph_token() -> str:
    if not st.secrets.get("TENANT_ID"):
        raise RuntimeError("ยังไม่ได้ตั้ง TENANT_ID/CLIENT_ID/CLIENT_SECRET ใน Secrets")
    tenant = st.secrets["TENANT_ID"]
    client_id = st.secrets["CLIENT_ID"]
    client_secret = st.secrets["CLIENT_SECRET"]
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(token_url, data=data, timeout=20)
    r.raise_for_status()
    return r.json()["access_token"]


def _download_from_graph_share_link(share_url: str) -> bytes:
    if not share_url:
        raise ValueError("ต้องระบุลิงก์ Share ของ OneDrive/SharePoint")
    enc = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    api = f"https://graph.microsoft.com/v1.0/shares/u!{enc}/driveItem/content"
    token = _get_graph_token()
    r = requests.get(api, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.content

# --- Loader ---

def load_data(data_source: str, file_obj, direct_url: str, graph_share_url: str, file_path: str) -> pd.DataFrame:
    # 1) Upload
    if data_source == "อัปโหลด Excel (.xlsx)":
        if not file_obj:
            st.info("อัปโหลดไฟล์ Excel เพื่อเริ่มต้น"); st.stop()
        return _read_excel_bytes(file_obj.read())

    # 2) Direct URL (OneDrive/SharePoint/GitHub)
    elif data_source == "URL สาธารณะ (OneDrive/SharePoint/GitHub)":
        if not direct_url:
            st.info("วางลิงก์ดาวน์โหลดตรง (direct download URL)"); st.stop()
        if requests is None:
            st.error("ไม่พบไลบรารี requests"); st.stop()
        try:
            r = requests.get(direct_url, timeout=30)
            r.raise_for_status()
        except Exception as e:
            st.error(f"ดาวน์โหลดไฟล์จาก URL ไม่สำเร็จ: {e}"); st.stop()
        return _read_excel_bytes(r.content)

    # 3) Graph API
    elif data_source == "OneDrive/SharePoint (Graph API)":
        if requests is None:
            st.error("ไม่พบไลบรารี requests"); st.stop()
        try:
            xbytes = _download_from_graph_share_link(graph_share_url)
        except Exception as e:
            st.error(f"ดึงไฟล์ผ่าน Graph ไม่สำเร็จ: {e}"); st.stop()
        return _read_excel_bytes(xbytes)

    # 4) LAN path
    else:
        if not file_path or not os.path.exists(file_path):
            st.error(f"ไม่พบไฟล์: {file_path}"); st.stop()
        with open(file_path, "rb") as f:
            return _read_excel_bytes(f.read())

# ---------------- Load ----------------
try:
    df = load_data(data_source, file_obj, direct_url, graph_share_url, file_path)
    st.success("โหลดข้อมูลสำเร็จ")
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}")
    st.stop()

# 0 -> NaN สำหรับบาง KPI
for k in ["steam_per_feed", "cost_baht_per_ton_steam"]:
    if k in df.columns and zero_is_missing:
        df.loc[df[k] == 0, k] = np.nan

# Date filter
if "Date" not in df.columns or df["Date"].dropna().empty:
    st.warning("ไม่มีคอลัมน์ Date หรือไม่มีข้อมูลวันที่ที่ใช้ได้"); st.stop()

min_d, max_d = pd.to_datetime(df["Date"]).min(), pd.to_datetime(df["Date"]).max()
start_d, end_d = st.slider(
    "ช่วงวันที่",
    min_value=min_d.to_pydatetime(),
    max_value=max_d.to_pydatetime(),
    value=(min_d.to_pydatetime(), max_d.to_pydatetime()),
)

df_f = df[(df["Date"] >= pd.to_datetime(start_d)) & (df["Date"] <= pd.to_datetime(end_d))].copy()

# Notices
if getattr(df, "attrs", {}).get("spf_from_water", False):
    st.caption("ℹ️ Ton steam/Ton feed = ใช้น้ำ m3 ÷ ยอดบรรจุ(ตัน)")
if getattr(df, "attrs", {}).get("steam_cost_fallback", False):
    st.info("คำนวณค่า **Baht/Ton steam** จากสูตร *Baht/Ton feed ÷ Ton steam/Ton feed* (fallback)")

# ---------------- KPI ----------------
col1, col2, col3, _ = st.columns(4)

# feed cost (weighted)
if {"cost_fuel_baht", "packed_ton"}.issubset(df_f.columns) and df_f["packed_ton"].sum() > 0:
    feed_avg = df_f["cost_fuel_baht"].sum() / df_f["packed_ton"].sum()
else:
    feed_avg = df_f.get("cost_baht_per_ton_feed", pd.Series(dtype=float)).mean()
col1.metric("เฉลี่ย: Baht/Ton feed", "-" if pd.isna(feed_avg) else f"{feed_avg:,.0f}", None if pd.isna(feed_avg) else f"{(feed_avg - targets['cost_baht_per_ton_feed']):+.0f}")

# steam/feed (weighted)
if {"water_m3", "packed_ton"}.issubset(df_f.columns) and df_f["packed_ton"].sum() > 0:
    spf_avg = df_f["water_m3"].sum() / df_f["packed_ton"].sum()
else:
    spf_avg = df_f.get("steam_per_feed", pd.Series(dtype=float)).mean()
col2.metric("เฉลี่ย: Ton steam/Ton feed", "-" if pd.isna(spf_avg) else f"{spf_avg:,.2f}", None if pd.isna(spf_avg) else f"{(spf_avg - targets['steam_per_feed']):+.2f}")

# steam cost
steam_avg = df_f.get("cost_baht_per_ton_steam", pd.Series(dtype=float)).mean()
col3.metric("เฉลี่ย: Baht/Ton steam", "-" if pd.isna(steam_avg) else f"{steam_avg:,.0f}", None if pd.isna(steam_avg) else f"{(steam_avg - targets['cost_baht_per_ton_steam']):+.0f}")

# ---------------- Totals ----------------
st.markdown("### ข้อมูลกระบวนการผลิตไอน้ำ")
tot1, tot2, tot3 = st.columns(3)
packed_sum2 = df_f["packed_ton"].sum() if "packed_ton" in df_f.columns else np.nan
cost_sum2 = df_f["cost_fuel_baht"].sum() if "cost_fuel_baht" in df_f.columns else np.nan
water_sum2 = df_f["water_m3"].sum() if "water_m3" in df_f.columns else np.nan

tot1.metric("ยอดบรรจุรวม (ตัน)", "-" if pd.isna(packed_sum2) else f"{packed_sum2:,.0f}")
tot2.metric("รวม Cost fuel (Baht)", "-" if pd.isna(cost_sum2) else f"{cost_sum2:,.0f}")
tot3.metric("รวมใช้น้ำ (m³)", "-" if pd.isna(water_sum2) else f"{water_sum2:,.0f}")

st.divider()

# ---------------- Aggregations & Charts ----------------

def _aggregate_by_grain(df_in: pd.DataFrame, grain: str) -> pd.DataFrame:
    df_k = df_in.copy()
    if grain == "รายเดือน":
        df_k["Period"] = df_k["Date"].dt.to_period("M").dt.to_timestamp(); how = "mean"
    elif grain == "รายสัปดาห์":
        df_k["Period"] = df_k["Date"].dt.to_period("W-MON").dt.start_time; how = "mean"
    else:
        df_k["Period"] = df_k["Date"].dt.floor("D"); how = "mean"

    if {"cost_fuel_baht", "packed_ton"}.issubset(df_k.columns):
        with np.errstate(divide="ignore", invalid="ignore"):
            df_k["cost_baht_per_ton_feed_calc"] = df_k["cost_fuel_baht"] / df_k["packed_ton"]
    else:
        df_k["cost_baht_per_ton_feed_calc"] = df_k.get("cost_baht_per_ton_feed", np.nan)

    if {"water_m3", "packed_ton"}.issubset(df_k.columns):
        with np.errstate(divide="ignore", invalid="ignore"):
            df_k["steam_per_feed_calc"] = df_k["water_m3"] / df_k["packed_ton"]
    else:
        df_k["steam_per_feed_calc"] = df_k.get("steam_per_feed", np.nan)

    agg_cols = [c for c in ["cost_baht_per_ton_feed_calc", "steam_per_feed_calc", "cost_baht_per_ton_steam"] if c in df_k.columns]
    if not agg_cols:
        return df_k

    df_out = df_k.groupby("Period", as_index=False)[agg_cols].agg(how)
    return df_out.sort_values("Period")


def _make_bar(fig_df: pd.DataFrame, y_col: str, title: str):
    fig = px.bar(fig_df, x="Period", y=y_col, title=title)

    tkey_map = {
        "cost_baht_per_ton_feed_calc": "cost_baht_per_ton_feed",
        "steam_per_feed_calc": "steam_per_feed",
    }
    tkey = tkey_map.get(y_col, y_col)

    if tkey in targets and targets[tkey] is not None:
        fig.add_hline(
            y=targets[tkey], line_dash="dash", line_color="red",
            annotation_text="Target", annotation_position="top left",
        )

    fig.update_traces(marker_color="#2E86C1")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))

    if time_grain == "รายวัน":
        dtick = f"D{int(tick_every)}"
    elif time_grain == "รายสัปดาห์":
        dtick = f"D{int(7 * tick_every)}"
    else:
        dtick = f"M{int(tick_every)}"

    fig.update_xaxes(
        tickformat=date_tick_fmt, tickangle=tick_angle, tickmode="auto", dtick=dtick,
        tick0=fig_df["Period"].min(), ticks="outside", showgrid=False,
    )

    fig.update_traces(
        hovertemplate=("<b>%{x|%d %b %Y}</b><br>" + title + ": %{y:,.2f}<extra></extra>")
    )
    fig.update_layout(
        hovermode="x unified", xaxis_showspikes=show_spike, xaxis_spikemode="across",
        xaxis_spikecolor="#999", xaxis_spikethickness=1,
    )
    return fig

# Charts
_dfk = _aggregate_by_grain(df_f, time_grain)
cols = st.columns(3)
chart_meta = [
    ("cost_baht_per_ton_feed_calc", "Baht/Ton feed"),
    ("steam_per_feed_calc", "Ton steam/Ton feed"),
    ("cost_baht_per_ton_steam", "Baht/Ton steam"),
]
for c, (k, title) in zip(cols, chart_meta):
    with c:
        if k in _dfk.columns and not _dfk[k].dropna().empty:
            st.plotly_chart(_make_bar(_dfk, k, title), use_container_width=True)
        else:
            st.info(f"ไม่มีข้อมูลสำหรับ {title}")

# ---------------- Fuel Mix ----------------
st.markdown("## สัดส่วนเชื้อเพลิงและปริมาณการใช้")
colA, colB = st.columns([1.2, 1])
with colA:
    fuels = [c for c in FUEL_ORDER if c in df_f.columns]
    if fuels:
        df_m = df_f.melt(id_vars=["Date"], value_vars=fuels, var_name="Fuel", value_name="kg")
        df_m["Fuel"] = pd.Categorical(df_m["Fuel"], categories=FUEL_ORDER, ordered=True)
        fig = px.bar(
            df_m, x="Date", y="kg", color="Fuel", barmode="stack",
            title="การใช้เชื้อเพลิงรายวัน", color_discrete_map=FUEL_COLORS, category_orders={"Fuel": FUEL_ORDER},
        )
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=60, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ไม่มีคอลัมน์เชื้อเพลิงที่รองรับในไฟล์นี้")

with colB:
    if set(FUEL_ORDER).issubset(df_f.columns):
        total = df_f[FUEL_ORDER].sum()
        pie_df = total.reset_index(); pie_df.columns = ["Fuel", "kg"]
        legend_order = ["furniture_wood_kg", "woodchip_kg", "cashew_shell_kg"]
        pie_df["Fuel"] = pd.Categorical(pie_df["Fuel"], categories=legend_order, ordered=True)
        fig = px.pie(
            pie_df.sort_values("Fuel"), names="Fuel", values="kg",
            title="สัดส่วนเชื้อเพลิง (ช่วงที่เลือก)", color="Fuel", color_discrete_map=FUEL_COLORS, hole=0,
        )
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=60, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ยังคำนวณสัดส่วนไม่ได้ เพราะคอลัมน์เชื้อเพลิงไม่ครบ")

# ---------------- Download cleaned data ----------------
with st.expander("ดาวน์โหลดข้อมูลที่ทำความสะอาดแล้ว"):
    def make_xlsx_bytes(df_to_save):
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            df_to_save.to_excel(writer, index=False, sheet_name="data")
        bio.seek(0)
        return bio

    st.download_button("ดาวน์โหลดข้อมูลทั้งหมด (Excel)", data=make_xlsx_bytes(df), file_name="fuel_dashboard_clean.xlsx")
    st.download_button("ดาวน์โหลดข้อมูลช่วงที่เลือก (Excel)", data=make_xlsx_bytes(df_f), file_name="fuel_dashboard_filtered.xlsx")
