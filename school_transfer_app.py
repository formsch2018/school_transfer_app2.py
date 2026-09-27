# -*- coding: utf-8 -*-
"""
نظام إدارة مسارات الطلبة - واجهة محسنة بتبويبات
"""
import streamlit as st
import pandas as pd
import folium
from folium import plugins
from streamlit_folium import st_folium
import requests
import numpy as np
from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import sys
import io
from math import radians, cos, sin, asin, sqrt

# ==================== تنسيقات CSS ====================
st.markdown("""
<style>
    /* الخلفية العامة */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
    }

    /* العنوان الرئيسي */
    h1 {
        background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
        color: white !important;
        padding: 25px 40px;
        border-radius: 15px;
        text-align: center;
        font-size: 2.2em !important;
        font-weight: 900;
        box-shadow: 0 6px 20px rgba(0,0,0,0.25);
        margin-bottom: 30px !important;
        letter-spacing: 1px;
    }

    /* الشريط الجانبي */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #e3f2fd 0%, #bbdefb 100%) !important;
        border-right: 3px solid #64b5f6;
    }

    section[data-testid="stSidebar"] > div {
        padding: 20px;
    }

    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label {
        color: #1565c0 !important;
        font-weight: 700 !important;
    }

    /* التبويبات الرئيسية */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0;
        background: #34495e;
        border-radius: 12px 12px 0 0;
        padding: 8px 8px 0 8px;
    }

    .stTabs [data-baseweb="tab"] {
        background: #7f8c8d;
        color: white;
        border-radius: 10px 10px 0 0;
        padding: 18px 45px;
        font-weight: 900;
        font-size: 1.5em;
        margin-right: 4px;
        border: 2px solid #95a5a6;
        border-bottom: none;
        transition: all 0.3s ease;
    }

    .stTabs [data-baseweb="tab"]:hover {
        background: #95a5a6;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #3498db 0%, #2980b9 100%) !important;
        color: white !important;
        border-color: #2980b9 !important;
        transform: translateY(-3px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }

    /* بطاقات الإحصائيات */
    div[data-testid="stMetric"] {
        background: white;
        padding: 20px;
        border-radius: 15px;
        border: 2px solid #ecf0f1;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
    }

    div[data-testid="stMetric"]:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        border-color: #3498db;
    }

    div[data-testid="stMetric"] label {
        color: #7f8c8d;
        font-size: 0.85em;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #2c3e50;
        font-size: 2em;
        font-weight: 900;
    }

    /* جدول البيانات */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        border: 2px solid #ecf0f1;
    }

    /* الأزرار */
    .stButton > button {
        background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
        color: white;
        font-weight: 700;
        border-radius: 10px;
        padding: 12px 25px;
        border: none;
        box-shadow: 0 3px 10px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
    }

    .stButton > button:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.3);
    }

    /* صناديق الرسائل */
    .stSuccess {
        background: linear-gradient(135deg, #d5f4e6 0%, #abebc6 100%);
        border-left: 5px solid #27ae60;
        border-radius: 10px;
        padding: 15px 20px;
        color: #1e8449;
        font-weight: 600;
    }

    .stInfo {
        background: linear-gradient(135deg, #d6eaf8 0%, #aed6f1 100%);
        border-left: 5px solid #3498db;
        border-radius: 10px;
        padding: 15px 20px;
        color: #1a5276;
        font-weight: 600;
    }

    .stWarning {
        background: linear-gradient(135deg, #fdebd0 0%, #f9e79f 100%);
        border-left: 5px solid #f39c12;
        border-radius: 10px;
        padding: 15px 20px;
        color: #9a7d0a;
        font-weight: 600;
    }

    /* خط الفاصل */
    hr {
        border: none;
        height: 3px;
        background: linear-gradient(90deg, transparent 0%, #3498db 20%, #3498db 80%, transparent 100%);
        margin: 25px 0;
    }
</style>
""", unsafe_allow_html=True)

# ==================== إعداد الصفحة ====================
st.set_page_config(page_title="نظام إدارة مسارات الطلبة", page_icon=" ", layout="wide")

# ==================== المسار الافتراضي للملف ====================
DEFAULT_FILE = (
    r"C:\Users\SAHA\Desktop\1مسارات الطلبة\مسارات الطلبة"
    r"\1تنسيق الزيارات\1تنسيق الزيارات\تنسيق الزيارات1.xlsx"
)

# ==================== ألوان الجهات ====================
ENTITY_COLORS = {
    "مأمون": "#e74c3c",
    "قسم": "#3498db",
}


# ==================== تحميل البيانات ====================
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df


def haversine(lon1, lat1, lon2, lat2):
    """حساب المسافة المستقيمة بين نقطتين بالكيلومتر"""
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * asin(sqrt(a))
    r = 6371
    return c * r


@st.cache_data(show_spinner=False)
def get_route_osrm_cached(lon1, lat1, lon2, lat2):
    """حساب المسافة عبر الشوارع باستخدام OSRM (مع تخزين مؤقت)"""
    url = (
        f"https://router.project-osrm.org/route/v1/driving/"
        f"{lon1},{lat1};{lon2},{lat2}"
    )
    params = {"overview": "full", "geometries": "geojson", "steps": "false"}
    try:
        response = requests.get(url, params=params, timeout=15)
        if response.status_code == 200:
            data = response.json()
            if data.get("code") == "Ok" and data.get("routes"):
                route = data["routes"][0]
                return {
                    "distance_km": route["distance"] / 1000,
                    "duration_min": route["duration"] / 60,
                    "geometry": route["geometry"]["coordinates"],
                }
    except Exception:
        pass
    return None


def calculate_all_routes(df):
    """حساب المسافات لجميع الأزواج باستخدام ThreadPoolExecutor"""
    results = []
    total = len(df)
    progress_bar = st.progress(0)
    status_text = st.empty()

    def process_row(idx, row):
        return idx, get_route_osrm_cached(row["x1"], row["y1"], row["x2"], row["y2"])

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {
            executor.submit(process_row, idx, row): idx
            for idx, row in df.iterrows()
        }
        for i, future in enumerate(as_completed(futures)):
            idx, route = future.result()
            if route:
                results.append({"idx": idx, **route})
            progress_bar.progress((i + 1) / total)
            status_text.text(f"تم حساب {i + 1} من {total}")

    progress_bar.empty()
    status_text.empty()
    return results


def offset_point(lat, lon, offset_meters=500):
    """إزاحة نقطة بمقدار معين بالمتري"""
    lat_offset = offset_meters / 111320
    lon_offset = offset_meters / (111320 * abs(cos(radians(lat))))
    return lat + lat_offset, lon + lon_offset


def create_map(df, routes_data=None, map_type="قمر اصطناعي", center_lat=None, center_lon=None, zoom_level=9):
    """إنشاء خريطة Folium مع النقاط والخطوط"""
    if center_lat is None:
        center_lat = (df["y1"].mean() + df["y2"].mean()) / 2
    if center_lon is None:
        center_lon = (df["x1"].mean() + df["x2"].mean()) / 2

    m = folium.Map(location=[center_lat, center_lon], zoom_start=zoom_level)

    # إضافة نوع الخريطة المحدد فقط
    if map_type == "قمر اصطناعي":
        folium.TileLayer("Esri World Imagery", name="قمر اصطناعي").add_to(m)
    elif map_type == "تضاريس":
        folium.TileLayer(
            tiles="https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png",
            name="تضاريس",
            attr="OpenTopoMap",
        ).add_to(m)
    else:
        folium.TileLayer("OpenStreetMap", name="عادية").add_to(m)

    folium.LayerControl().add_to(m)

    # أدوات التنقل
    plugins.MeasureControl(position="topleft").add_to(m)
    plugins.MousePosition(position="bottomright").add_to(m)
    plugins.Fullscreen(position="topright").add_to(m)

    for _, row in df.iterrows():
        entity = row.get("الجهه", "")
        color = ENTITY_COLORS.get(entity, "#95a5a6")

        # حساب الإزاحة - إزاحة كبيرة في الاتجاهين
        lat1, lon1 = row["y1"], row["x1"]
        lat2, lon2 = row["y2"], row["x2"]
        mid_lat = (lat1 + lat2) / 2
        mid_lon = (lon1 + lon2) / 2
        offset_lat, offset_lon = offset_point(mid_lat, mid_lon, 500)

        # المدرسة المنقول إليها
        popup_dest = f"""
        <div style="font-family: Arial; min-width: 200px;">
            <h4 style="margin: 0; color: #27ae60;">  المدرسة المنقول إليها</h4>
            <hr style="margin: 5px 0;">
            <p><b>الاسم:</b> {row.get('اسم المدرسه المنقول اليها', 'غير محدد')}</p>
            <p><b>الجهة:</b> {row.get('الجهه', 'غير محدد')}</p>
            <p><b>المديرية:</b> {row.get('مديريات وزاره المنقول اليها', 'غير محدد')}</p>
            <p><b>أدنى صف:</b> {row.get('ادنى صف', 'غير محدد')}</p>
            <p><b>أعلى صف:</b> {row.get('اعلى صف', 'غير محدد')}</p>
        </div>
        """
        folium.Marker(
            location=[row["y1"], row["x1"]],
            popup=folium.Popup(popup_dest, max_width=300),
            icon=folium.Icon(color="green", icon="school", prefix="fa"),
        ).add_to(m)

        # المدرسة المنقولة
        popup_src = f"""
        <div style="font-family: Arial; min-width: 200px;">
            <h4 style="margin: 0; color: #e67e22;">  المدرسة المنقولة</h4>
            <hr style="margin: 5px 0;">
            <p><b>الاسم:</b> {row.get('اسم المدرسه المنقوله', 'غير محدد')}</p>
            <p><b>الجهة:</b> {row.get('الجهه', 'غير محدد')}</p>
            <p><b>المديرية:</b> {row.get('مديريات المنقولة', 'غير محدد')}</p>
            <p><b>أدنى صف:</b> {row.get('ادنى صف.1', 'غير محدد')}</p>
            <p><b>أعلى صف:</b> {row.get('اعلى صف.1', 'غير محدد')}</p>
        </div>
        """
        folium.Marker(
            location=[row["y2"], row["x2"]],
            popup=folium.Popup(popup_src, max_width=300),
            icon=folium.Icon(color="orange", icon="school", prefix="fa"),
        ).add_to(m)

        # الخط بينهما - لون واحد غامق حسب الجهة
        if entity == "مأمون":
            color_main = "#c0392b"      # أحمر غامق
        else:
            color_main = "#1a5276"      # أزرق غامق

        # حساب الإزاحة العمودية على اتجاه الخط
        lat1, lon1 = row["y1"], row["x1"]
        lat2, lon2 = row["y2"], row["x2"]
        mid_lat = (lat1 + lat2) / 2
        mid_lon = (lon1 + lon2) / 2

        # حساب زاوية الخط والإزاحة العمودية
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        angle = np.arctan2(dlon, dlat)

        # إزاحة عمودية على الخط (300 متر)
        perp_angle = angle + np.pi / 2
        offset_dist = 300  # متر

        # حساب الإزاحة بالاتجاه العمودي
        lat_offset = (offset_dist / 111320) * np.cos(perp_angle)
        lon_offset = (offset_dist / (111320 * abs(np.cos(np.radians(mid_lat))))) * np.sin(perp_angle)

        if routes_data:
            route = next(
                (r for r in routes_data if r["idx"] == row.name), None
            )
            if route:
                locations = [
                    [coord[1], coord[0]] for coord in route["geometry"]
                ]
                # إزاحة المسار بالكامل
                shifted_locations = [
                    [coord[1] + lat_offset, coord[0] + lon_offset]
                    for coord in route["geometry"]
                ]
                # خط واحد غامق مع إزاحة
                folium.PolyLine(
                    locations=shifted_locations,
                    color=color_main,
                    weight=5,
                    opacity=1.0,
                    popup=f"الجهة: {entity} | المسافة: {route['distance_km']:.2f} كم",
                ).add_to(m)
            else:
                folium.PolyLine(
                    locations=[
                        [row["y2"] + lat_offset, row["x2"] + lon_offset],
                        [row["y1"] + lat_offset, row["x1"] + lon_offset],
                    ],
                    color=color_main,
                    weight=5,
                    opacity=1.0,
                    dash_array="5, 10",
                ).add_to(m)
        else:
            folium.PolyLine(
                locations=[
                    [row["y2"] + lat_offset, row["x2"] + lon_offset],
                    [row["y1"] + lat_offset, row["x1"] + lon_offset],
                ],
                color=color_main,
                weight=5,
                opacity=1.0,
            ).add_to(m)

    # مفتاح الخريطة
    legend_html = """
    <div style="position: fixed; bottom: 50px; left: 50px; z-index: 1000;
                background-color: white; padding: 10px; border: 2px solid grey;
                border-radius: 5px; font-family: Arial;">
        <p><b>مفتاح الخريطة</b></p>
        <p>  <span style="color: #e74c3c;">━</span> مأمون</p>
        <p>  <span style="color: #3498db;">━</span> قسم</p>
        <p>  <span style="color: green;">●</span> المدرسة المنقول إليها</p>
        <p>  <span style="color: orange;">●</span> المدرسة المنقولة</p>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    return m


def add_permanent_labels(m, df, routes_data=None):
    """إضافة أسماء المدارس بشكل مثبت على الخريطة + المسافة"""
    for _, row in df.iterrows():
        entity = row.get("الجهه", "")
        color_code = "#e74c3c" if entity == "مأمون" else "#3498db"

        if routes_data:
            route = next((r for r in routes_data if r["idx"] == row.name), None)
            dist_text = f"{route['distance_km']:.1f} km" if route else ""
        else:
            dist_text = ""

        # اسم المدرسة المنقول إليها
        dest_name = str(row.get("اسم المدرسه المنقول اليها", ""))
        if len(dest_name) > 30:
            dest_name = dest_name[:30] + "..."
        folium.Marker(
            location=[row["y1"], row["x1"]],
            icon=folium.DivIcon(
                html=f'<div style="font-size: 13px; font-weight: 900; color: {color_code}; text-shadow: 2px 2px 2px white, -2px -2px 2px white, 2px -2px 2px white, -2px 2px 2px white; white-space: nowrap; transform: translate(14px, -22px); background: rgba(255,255,255,0.7); padding: 1px 4px; border-radius: 4px;">{dest_name}</div>',
                icon_size=(0, 0),
                icon_anchor=(0, 0),
            ),
        ).add_to(m)

        # اسم المدرسة المنقولة
        src_name = str(row.get("اسم المدرسه المنقوله", ""))
        if len(src_name) > 30:
            src_name = src_name[:30] + "..."
        folium.Marker(
            location=[row["y2"], row["x2"]],
            icon=folium.DivIcon(
                html=f'<div style="font-size: 13px; font-weight: 900; color: {color_code}; text-shadow: 2px 2px 2px white, -2px -2px 2px white, 2px -2px 2px white, -2px 2px 2px white; white-space: nowrap; transform: translate(14px, 8px); background: rgba(255,255,255,0.7); padding: 1px 4px; border-radius: 4px;">{src_name}</div>',
                icon_size=(0, 0),
                icon_anchor=(0, 0),
            ),
        ).add_to(m)

        # المسافة بجانب الخط
        if dist_text:
            mid_lat = (row["y1"] + row["y2"]) / 2
            mid_lon = (row["x1"] + row["x2"]) / 2

            # حساب الإزاحة العمودية - 50 متر فقط بجانب الخط
            lat1, lon1 = row["y1"], row["x1"]
            lat2, lon2 = row["y2"], row["x2"]
            dlat = lat2 - lat1
            dlon = lon2 - lon1
            angle = np.arctan2(dlon, dlat)
            perp_angle = angle + np.pi / 2
            offset_dist = 50

            lat_offset = (offset_dist / 111320) * np.cos(perp_angle)
            lon_offset = (offset_dist / (111320 * abs(np.cos(np.radians(mid_lat))))) * np.sin(perp_angle)

            folium.Marker(
                location=[mid_lat + lat_offset, mid_lon + lon_offset],
                icon=folium.DivIcon(
                    html=f'<div style="font-size: 14px; font-weight: 900; color: #000; background: rgba(255,255,255,0.95); padding: 4px 10px; border-radius: 10px; border: 2px solid #333; white-space: nowrap; transform: translate(-50%, -50%); box-shadow: 0 2px 6px rgba(0,0,0,0.3);">{dist_text}</div>',
                    icon_size=(0, 0),
                    icon_anchor=(0, 0),
                ),
            ).add_to(m)

    return m


# ==================== الواجهة الرئيسية ====================
def main():
    st.title(" ️ نظام إدارة مسارات الطلبة")

    # تحميل الملف
    uploaded_file = st.sidebar.file_uploader("  تحميل ملف إكسل", type=["xlsx", "xls"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
    elif os.path.exists(DEFAULT_FILE):
        df = load_data(DEFAULT_FILE)
    else:
        st.error("  يرجى تحميل ملف إكسل")
        return

    # تعبئة القيم المفقودة فقط في الأعمدة غير الإحداثية
    coord_cols = ["x1", "y1", "x2", "y2"]
    non_coord_cols = [c for c in df.columns if c not in coord_cols]
    df[non_coord_cols] = df[non_coord_cols].fillna("غير محدد")

    # تحويل الإحداثيات إلى أرقام وإزالة الصفوف غير الصالحة
    for col in coord_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=coord_cols).reset_index(drop=True)

    st.sidebar.success(f"✅ تم تحميل {len(df)} سجل")

    # ==================== الفلاتر في الشريط الجانبي ====================
    st.sidebar.markdown("---")
    st.sidebar.subheader("  الفلاتر")

    entity_values = df["الجهه"].unique().tolist()
    selected_entity = st.sidebar.multiselect(
        "  **الجهة**",
        entity_values,
        default=entity_values,
        key="filter_entity",
    )

    transport_values = df["النقل"].unique().tolist()
    selected_transport = st.sidebar.multiselect(
        "  **النقل**",
        transport_values,
        default=transport_values,
        key="filter_transport",
    )

    # تصفية البيانات حسب الجهة والنقل أولاً
    temp_df = df.copy()
    temp_df = temp_df[temp_df["الجهه"].isin(selected_entity)]
    temp_df = temp_df[temp_df["النقل"].isin(selected_transport)]

    # قوائم المدارس المحدثة حسب الفلاتر
    src_options = ["  جميع المدارس المنقولة"] + sorted(temp_df["اسم المدرسه المنقوله"].unique().tolist())
    if "filter_src_school" in st.session_state:
        if st.session_state["filter_src_school"] not in src_options:
            st.session_state["filter_src_school"] = "  جميع المدارس المنقولة"
    selected_src = st.sidebar.selectbox(
        "  **المدرسة المنقولة**",
        src_options,
        key="filter_src_school",
    )

    dest_options = ["  جميع المدارس المنقول إليها"] + sorted(temp_df["اسم المدرسه المنقول اليها"].unique().tolist())
    if "filter_dest_school" in st.session_state:
        if st.session_state["filter_dest_school"] not in dest_options:
            st.session_state["filter_dest_school"] = "  جميع المدارس المنقول إليها"
    selected_dest = st.sidebar.selectbox(
        "  **المدرسة المنقول إليها**",
        dest_options,
        key="filter_dest_school",
    )

    # زر مسح الفلاتر
    if st.sidebar.button("  مسح جميع الفلاتر"):
        st.session_state.clear()
        st.rerun()

    # ==================== فلتر نوع الخريطة ====================
    st.sidebar.markdown("---")
    st.sidebar.subheader("  نوع عرض الخريطة")

    map_type = st.sidebar.radio(
        "اختر نوع الخريطة:",
        ["قمر اصطناعي", "تضاريس", "عادية"],
        key="map_type",
        horizontal=True,
    )

    st.sidebar.info(f"  نوع الخريطة الحالي: **{map_type}**")

    # ==================== تطبيق الفلاتر ====================
    filtered_df = temp_df.copy()
    if selected_src != "  جميع المدارس المنقولة":
        filtered_df = filtered_df[filtered_df["اسم المدرسه المنقوله"] == selected_src]
    if selected_dest != "  جميع المدارس المنقول إليها":
        filtered_df = filtered_df[filtered_df["اسم المدرسه المنقول اليها"] == selected_dest]

    # ==================== التبويبات الرئيسية ====================
    tab_data, tab_map = st.tabs([
        "  البيانات التفصيلية للمدارس",
        "  الخريطة التفصيلية للمدارس",
    ])

    # ==================== تبويب البيانات ====================
    with tab_data:
        st.subheader(f"  النتائج ({len(filtered_df)} سجل)")

        # إحصائيات سريعة
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("إجمالي السجلات", len(filtered_df))
        c2.metric("مأمون", len(filtered_df[filtered_df["الجهه"] == "مأمون"]))
        c3.metric("قسم", len(filtered_df[filtered_df["الجهه"] == "قسم"]))

        if len(filtered_df) > 0:
            avg_dist = filtered_df.apply(
                lambda r: haversine(r["x1"], r["y1"], r["x2"], r["y2"]), axis=1
            ).mean()
            c4.metric("متوسط المسافة المستقيمة", f"{avg_dist:.1f} كم")

        st.dataframe(filtered_df, use_container_width=True)

        if len(filtered_df) > 0:
            csv = filtered_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "  تحميل النتائج كملف CSV",
                csv,
                "filtered_results.csv",
                "text/csv",
            )

    # ==================== تبويب الخريطة ====================
    with tab_map:
        st.subheader("  الخريطة التفاعلية")

        if len(filtered_df) > 0:
            map_df = filtered_df.copy()

            # حساب مركز المدارس المختارة للتكبير التلقائي
            center_lat = (map_df["y1"].mean() + map_df["y2"].mean()) / 2
            center_lon = (map_df["x1"].mean() + map_df["x2"].mean()) / 2

            # حساب مستوى التكبير المناسب حسب عدد المدارس
            num_schools = len(map_df)
            if num_schools <= 3:
                zoom_level = 14
            elif num_schools <= 10:
                zoom_level = 12
            elif num_schools <= 30:
                zoom_level = 10
            elif num_schools <= 100:
                zoom_level = 9
            else:
                zoom_level = 8

            show_routes = st.checkbox(
                "  عرض المسارات عبر الشوارع (شبكة الطريق)",
                value=True,
            )

            routes_data = None
            if show_routes:
                with st.spinner("جاري حساب المسارات عبر الشوارع... يرجى الانتظار"):
                    routes_data = calculate_all_routes(map_df)
                    st.success(f"تم حساب {len(routes_data)} مسار عبر الشوارع")

            m = create_map(
                map_df,
                routes_data,
                map_type=map_type,
                center_lat=center_lat,
                center_lon=center_lon,
                zoom_level=zoom_level,
            )
            m = add_permanent_labels(m, map_df, routes_data)
            map_data = st_folium(m, width=1200, height=600, returned_objects=["last_object_clicked"])

            # جدول شامل لجميع المسارات
            st.markdown("---")
            st.subheader("  بيانات جميع المسارات")

            # إنشاء جدول شامل
            all_routes_df = map_df.copy()

            # إضافة بيانات المسارات
            if routes_data:
                routes_df = pd.DataFrame(routes_data)
                all_routes_df = all_routes_df.merge(
                    routes_df[["idx", "distance_km", "duration_min"]],
                    left_index=True,
                    right_on="idx",
                    how="left",
                )

            # حساب المسافة المستقيمة
            all_routes_df["المسافة المستقيمة (كم)"] = all_routes_df.apply(
                lambda r: haversine(r["x1"], r["y1"], r["x2"], r["y2"]), axis=1
            ).round(2)

            # تحديد الأعمدة للعرض
            display_columns = [
                "اسم المدرسه المنقوله",
                "رمز المؤسسه المنقوله",
                "الجهه",
                "مديريات المنقولة",
                "الملكية2",
                "الطاقه الاستيعابيه",
                "مجموع الطلاب",
                "ادنى صف.1",
                "اعلى صف.1",
                "اسم المدرسه المنقول اليها",
                "رمز المؤسسه المنقول اليها",
                "مديريات وزاره المنقول اليها",
                "الملكية1",
                "ادنى صف",
                "اعلى صف",
                "النقل",
                "المبرر",
                "المسافة المستقيمة (كم)",
            ]

            # إضافة أعمدة المسافة عبر الشوارع إذا كانت موجودة
            if "distance_km" in all_routes_df.columns:
                all_routes_df["المسافة عبر الشوارع (كم)"] = all_routes_df["distance_km"].round(2)
                display_columns.append("المسافة عبر الشوارع (كم)")

            if "duration_min" in all_routes_df.columns:
                all_routes_df["الزمن المقدر (دقيقة)"] = all_routes_df["duration_min"].round(1)
                display_columns.append("الزمن المقدر (دقيقة)")

            # عرض الجدول
            st.dataframe(
                all_routes_df[display_columns],
                use_container_width=True,
                height=400,
            )

            # إحصائيات سريعة
            st.markdown("###  إحصائيات سريعة")
            sc1, sc2, sc3, sc4 = st.columns(4)
            sc1.metric("إجمالي المسارات", len(all_routes_df))
            sc2.metric("مأمون", len(all_routes_df[all_routes_df["الجهه"] == "مأمون"]))
            sc3.metric("قسم", len(all_routes_df[all_routes_df["الجهه"] == "قسم"]))

            if "distance_km" in all_routes_df.columns:
                valid_distances = all_routes_df["distance_km"].dropna()
                if len(valid_distances) > 0:
                    sc4.metric("متوسط المسافة عبر الشوارع", f"{valid_distances.mean():.1f} كم")

            # جدول المسافات المحسوبة
            if routes_data:
                st.markdown("---")
                st.subheader("  جدول المسافات عبر الشوارع")

                results_df = pd.DataFrame(routes_data)
                final_df = map_df.reset_index().merge(
                    results_df, left_on="index", right_on="idx", how="left"
                )
                final_df["straight_distance_km"] = final_df.apply(
                    lambda row: haversine(row["x1"], row["y1"], row["x2"], row["y2"]),
                    axis=1,
                )

                display_cols = [
                    "index",
                    "straight_distance_km",
                    "distance_km",
                    "duration_min",
                ]
                st.dataframe(final_df[display_cols])

                # إحصائيات المسافات
                st.markdown("###  إحصائيات المسافات")
                sc1, sc2, sc3 = st.columns(3)
                valid_routes = final_df.dropna(subset=["distance_km"])
                if len(valid_routes) > 0:
                    sc1.metric(
                        "متوسط المسافة عبر الشوارع",
                        f"{valid_routes['distance_km'].mean():.1f} كم",
                    )
                    sc2.metric(
                        "أقصى مسافة",
                        f"{valid_routes['distance_km'].max():.1f} كم",
                    )
                    sc3.metric(
                        "أدنى مسافة",
                        f"{valid_routes['distance_km'].min():.1f} كم",
                    )
        else:
            st.warning("  لا توجد بيانات مطابقة للفلاتر المحددة")


if __name__ == "__main__":
    main()
