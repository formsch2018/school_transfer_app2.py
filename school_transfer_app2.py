import streamlit as st
import pandas as pd
import numpy as np
import folium
from streamlit_folium import st_folium
import requests
import polyline

st.set_page_config(
    page_title="نظام دمج ونقل المدارس وإدارة الأسطول",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        .main { 
            background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 50%, #7dd3fc 100%);
            color: #0f172a;
        }
        h1, h2, h3, h4, h5, h6 { 
            color: #0f172a !important; 
            font-weight: 800; 
        }
        p, span, label, div, .stMarkdown {
            color: #1e293b !important;
        }
        .welcome-card {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(12px);
            padding: 35px;
            border-radius: 24px;
            box-shadow: 0 15px 35px rgba(14, 165, 233, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.9);
            border-right: 8px solid #0284c7;
            margin-bottom: 30px;
        }
        .hero-title {
            color: #0f172a !important;
            font-size: 2.4rem;
            font-weight: 900;
            margin-bottom: 12px;
        }
        .hero-subtitle {
            color: #334155 !important;
            font-size: 1.15rem;
            line-height: 1.7;
            font-weight: 600;
        }
        div[data-testid="stMetric"] { 
            background: rgba(255, 255, 255, 0.9) !important; 
            padding: 22px !important; 
            border-radius: 18px !important; 
            box-shadow: 0 8px 20px rgba(14, 165, 233, 0.12) !important; 
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
        }
        div[data-testid="stMetric"] label {
            color: #334155 !important;
            font-size: 1rem !important;
            font-weight: 700 !important;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: #0f172a !important;
            font-size: 1.8rem !important;
            font-weight: 900 !important;
        }
        .stTabs [data-baseweb="tab-list"] { 
            gap: 12px; 
            background-color: transparent;
            padding: 10px 0;
        }
        .stTabs [data-baseweb="tab"] { 
            background-color: rgba(255, 255, 255, 0.7); 
            border-radius: 14px; 
            padding: 12px 26px; 
            font-weight: bold; 
            color: #0f172a; 
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.8);
        }
        .stTabs [aria-selected="true"] { 
            background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important; 
            color: white !important; 
            box-shadow: 0 8px 20px rgba(2, 132, 199, 0.35) !important;
            border: none !important;
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f0f9ff 0%, #e0f2fe 100%);
            color: #0f172a;
            border-left: 1px solid rgba(2, 132, 199, 0.1);
        }
        [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] span {
            color: #0f172a !important;
            font-weight: 600;
        }
        .legend-box {
            background: rgba(255, 255, 255, 0.95);
            padding: 18px 22px;
            border-radius: 16px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            border: 1px solid #bae6fd;
            margin-top: 15px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-around;
            flex-wrap: wrap;
            gap: 15px;
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.95rem;
            font-weight: bold;
            color: #0f172a;
        }
        .legend-color {
            width: 18px;
            height: 18px;
            border-radius: 4px;
            display: inline-block;
        }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://img.icons8.com/color/96/school-bus.png", width=75)
    st.header("إعدادات الملف والبيانات")
    
    default_file = "تنسيق الزيارات1.xlsx"
    uploaded_file = st.file_uploader("قم برفع ملف الإكسل (أو استخدام الافتراضي):", type=["xlsx", "xls"])
    target_file = uploaded_file if uploaded_file is not None else default_file

    st.markdown("---")
    st.markdown("### 🗺️ إعدادات مسارات الطرق")
    use_actual_roads = st.checkbox("🚗 رسم المسارات حسب الشوارع الرئيسية وأقصر الطرق", value=True)

    st.markdown("---")
    st.markdown("### 📌 خيارات العرض")
    show_home = st.checkbox("🏠 العودة للصفحة الرئيسية", value=True)

def calculate_buses_detailed(count):
    if count <= 0:
        return 0, 0, 0
    buses_55 = (count + 54) // 55
    buses_25 = (count + 24) // 25
    buses_16 = (count + 15) // 16
    return buses_55, buses_25, buses_16

def highlight_by_side(row):
    side_val = str(row.get('الجهة', '')).strip().lower()
    if 'مامون' in side_val or 'مأمون' in side_val:
        return ['background-color: #fee2e2; color: #0f172a; font-weight: 600;'] * len(row)
    elif 'قسم' in side_val:
        return ['background-color: #e0f2fe; color: #0f172a; font-weight: 600;'] * len(row)
    else:
        return ['background-color: #f1f5f9; color: #0f172a; font-weight: 600;'] * len(row)

@st.cache_data(ttl=3600)
def get_road_route_coordinates(lat1, lon1, lat2, lon2):
    try:
        url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=polyline"
        response = requests.get(url, timeout=3)
        if response.status_code == 200:
            data = response.json()
            if "routes" in data and len(data["routes"]) > 0:
                encoded_polyline = data["routes"][0]["geometry"]
                decoded_coords = polyline.decode(encoded_polyline)
                if decoded_coords:
                    return decoded_coords
    except Exception:
        pass
    return [[lat1, lon1], [lat2, lon2]]

try:
    xls = pd.ExcelFile(target_file)
    sheet_names = xls.sheet_names
    
    selected_sheet = st.sidebar.selectbox("اختر ورقة العمل (Sheet):", sheet_names, index=sheet_names.index('Sheet1') if 'Sheet1' in sheet_names else 0)
    df = pd.read_excel(target_file, sheet_name=selected_sheet)
    
    df.columns = [str(c).strip() for c in df.columns]
    
    students_col = next((c for c in df.columns if 'مجموع' in str(c) or 'طالب' in str(c) or 'طلاب' in str(c) or 'الطاقة' in str(c)), None)
    if not students_col:
        df['مجموع الطلاب'] = 0
        students_col = 'مجموع الطلاب'

    side_col = None
    for col in df.columns:
        c_clean = str(col).strip().replace('ة', 'ه')
        if 'جهه' in c_clean or c_clean == 'ب' or col == df.columns[1]:
            side_col = col
            break
    if not side_col:
        side_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    moved_school_col = None
    to_school_col = None

    for col in df.columns:
        c_str = str(col).strip()
        if ('المدرسة المنقوله' in c_str or 'المدرسة المنقولة' in c_str) and 'رمز' not in c_str:
            moved_school_col = col
            break

    for col in df.columns:
        c_str = str(col).strip()
        if ('المنقول اليها' in c_str or 'المنقول إليها' in c_str) and 'رمز' not in c_str:
            to_school_col = col
            break

    if not moved_school_col:
        moved_school_col = df.columns[16] if len(df.columns) > 16 else df.columns[0]
    if not to_school_col:
        to_school_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

    ownership_col = next((c for c in df.columns if 'الملكية2' in str(c) or 'ملكية2' in str(c)), None)
    if not ownership_col:
        ownership_col = next((c for c in df.columns if 'الملكية' in str(c) or 'ملكية' in str(c)), df.columns[0])

    min_grade_col = next((c for c in df.columns if 'أدنى صف' in str(c) or 'ادنى صف' in str(c)), None)
    max_grade_col = next((c for c in df.columns if 'أعلى صف' in str(c) or 'اعلى صف' in str(c)), None)

    if len(df.columns) > 0:
        processed_df = df.copy()
        
        processed_df['المدرسة المنقولة'] = processed_df[moved_school_col].fillna("").astype(str).str.strip()
        processed_df['المدرسة المنقول اليها'] = processed_df[to_school_col].fillna("").astype(str).str.strip()
        
        processed_df = processed_df[
            (processed_df['المدرسة المنقولة'] != "") & 
            (processed_df['المدرسة المنقولة'].str.lower() != "nan") & 
            (processed_df['المدرسة المنقولة'].str.lower() != "nat")
        ]

        processed_df['الجهة'] = processed_df[side_col].fillna("غير محدد").astype(str).str.strip() if side_col else "غير محدد"
        
        raw_ownership = processed_df[ownership_col].fillna("غير متوفر").astype(str).str.strip() if ownership_col else "غير متوفر"
        def normalize_ownership(val):
            v = str(val).strip()
            if 'مستأجر' in v or 'مستاجر' in v or 'إيجار' in v or 'ايجار' in v:
                return 'مستأجر'
            elif 'ملك' in v or 'حكومي' in v or 'ملكه' in v or 'ملكية' in v:
                return 'ملك'
            elif v == '' or v.lower() == 'nan' or v.lower() == 'nat':
                return 'غير متوفر'
            return v

        processed_df['الملكية2'] = raw_ownership.apply(normalize_ownership)
        
        processed_df['أدنى صف'] = processed_df[min_grade_col].fillna("غير متوفر").astype(str).str.strip() if min_grade_col else "غير متوفر"
        processed_df['أعلى صف'] = processed_df[max_grade_col].fillna("غير متوفر").astype(str).str.strip() if max_grade_col else "غير متوفر"

        def clean_side_name(val):
            v_str = str(val).strip()
            if 'مامون' in v_str.lower() or 'مأمون' in v_str.lower():
                return 'مأمون'
            elif 'قسم' in v_str.lower():
                return 'قسم'
            return v_str

        processed_df['الجهة'] = processed_df['الجهة'].apply(clean_side_name)

        col_lon_from = next((c for c in processed_df.columns if str(c).lower() == 'x2'), 'x2')
        col_lat_from = next((c for c in processed_df.columns if str(c).lower() == 'y2'), 'y2')
        
        col_lon_to = next((c for c in processed_df.columns if str(c).lower() == 'x1'), 'x1')
        col_lat_to = next((c for c in processed_df.columns if str(c).lower() == 'y1'), 'y1')

        for c in [col_lon_from, col_lat_from, col_lon_to, col_lat_to]:
            if c in processed_df.columns:
                processed_df[c] = pd.to_numeric(processed_df[c], errors='coerce')

        processed_df['المسافة_رقمية'] = 0.0
        distance_col_name = next((c for c in processed_df.columns if 'distance' in str(c).lower() or 'المسافة' in str(c)), None)
        if distance_col_name:
            processed_df['المسافة_رقمية'] = pd.to_numeric(processed_df[distance_col_name], errors='coerce').fillna(0)

        processed_df[students_col] = pd.to_numeric(processed_df[students_col], errors='coerce').fillna(0)
        total_all_schools_count = processed_df['المدرسة المنقولة'].nunique()

        if show_home:
            st.markdown("""
                <div class="welcome-card">
                    <div class="hero-title">🏫 نظام تخطيط نقل ودمج المدارس وإدارة الأسطول</div>
                    <div class="hero-subtitle">
                        لوحة المؤشرات الإحصائية تعتمد على احتساب عدد المدارس المنقولة الفريدة وإجمالي الطلبة والسجلات.
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            total_students_moved = processed_df[students_col].sum()

            c1, c2 = st.columns(2)
            with c1:
                st.metric("🏫 إجمالي المدارس المنقولة الفريدة", f"{total_all_schools_count} مدرسة")
            with c2:
                st.metric("👥 إجمالي الطلبة المنقولين", f"{int(total_students_moved)} طالب")
            
            st.markdown("---")

        tab1, tab2, tab3 = st.tabs([
            "📊 لوحة مؤشرات النقل العامة والملكية (حسب الجهة)", 
            "🛰️ خريطة عرض المسارات (التفاعل المباشر بالنقر)", 
            "📋 جدول تفاصيل خطط النقل والأسطول والمباني"
        ])
        
        with tab1:
            st.subheader("📊 لوحة مؤشرات النقل وإحصائيات الملكية الحقيقية (حسب الجهة)")
            
            st.markdown("""
                <div class="legend-box">
                    <div class="legend-item"><span class="legend-color" style="background-color: #ef4444;"></span> جهة مأمون (لون أحمر)</div>
                    <div class="legend-item"><span class="legend-color" style="background-color: #3b82f6;"></span> جهة القسم (لون أزرق)</div>
                </div>
            """, unsafe_allow_html=True)
            
            default_sides = ['مأمون', 'قسم']
            existing_sides = processed_df['الجهة'].unique()
            all_sides = sorted(list(set(default_sides).intersection(existing_sides).union(existing_sides)))
            all_ownership_types = ['ملك', 'مستأجر', 'غير متوفر']
            
            for side_name in all_sides:
                side_subset = processed_df[processed_df['الجهة'] == side_name]
                side_schools_count = side_subset['المدرسة المنقولة'].nunique()
                side_total_students = side_subset[students_col].sum()
                
                st.markdown(f"### 🏷️ الجهة: `{side_name}` (عدد المدارس المنقولة: {side_schools_count} | إجمالي الطلبة: {int(side_total_students)})")
                
                ownership_counts = side_subset.groupby('الملكية2')['المدرسة المنقولة'].nunique()
                ownership_counts = ownership_counts.reindex(all_ownership_types, fill_value=0)
                
                if not ownership_counts.empty:
                    cols_own = st.columns(len(ownership_counts))
                    for i, (own_type, count_val) in enumerate(ownership_counts.items()):
                        with cols_own[i]:
                            st.metric(f"عدد المدارس ({own_type})", f"{count_val} مدرسة")
                else:
                    st.info("لا توجد بيانات ملكية مسجلة لهذه الجهة.")
                
                st.markdown("---")

            st.markdown("### 🏛️ إجمالي عدد المدارس المنقولة الفريدة حسب تصنيف الملكية (لكافة الجهات):")
            global_ownership_counts = processed_df.groupby('الملكية2')['المدرسة المنقولة'].nunique()
            global_ownership_counts = global_ownership_counts.reindex(all_ownership_types, fill_value=0)
            if not global_ownership_counts.empty:
                g_cols = st.columns(len(global_ownership_counts))
                for i, (own_type, count_val) in enumerate(global_ownership_counts.items()):
                    with g_cols[i]:
                        st.metric(f"إجمالي المدارس ({own_type})", f"{count_val} مدرسة")

            st.markdown("---")
            st.markdown("### 📋 جدول البيانات الإحصائية المفصلة (كافة السجلات):")
            summary_table = processed_df.copy()
            bus_calc_summary = summary_table[students_col].apply(lambda x: pd.Series(calculate_buses_detailed(int(x))))
            summary_table[['باصات كبيرة (55)', 'باصات كوستر (25)', 'باصات صغيرة (16)']] = bus_calc_summary
            
            summary_cols_show = ['الجهة', 'المدرسة المنقولة', 'الملكية2', 'أدنى صف', 'أعلى صف', 'المدرسة المنقول اليها', students_col, 'باصات كبيرة (55)', 'باصات كوستر (25)', 'باصات صغيرة (16)']
            if distance_col_name:
                summary_cols_show.append(distance_col_name)
            else:
                summary_cols_show.append('المسافة_رقمية')
                
            st.dataframe(summary_table[summary_cols_show].style.apply(highlight_by_side, axis=1), use_container_width=True)

        with tab2:
            st.subheader("🛰️ خريطة الأقمار الصناعية التفاعلية (مع إزاحة المسارات لتظهر بجانب بعضها)")
            
            target_school_col = 'المدرسة المنقولة'
            unique_moved_schools = sorted([str(x).strip() for x in processed_df[target_school_col].dropna().unique() if str(x).strip() != ''])
            
            if "selected_schools_state" not in st.session_state:
                st.session_state.selected_schools_state = []

            selected_schools_filter = st.multiselect(
                "🔍 اختر اسم المدرسة المنقولة:",
                unique_moved_schools,
                default=st.session_state.selected_schools_state,
                key="tab2_school_multiselect"
            )
            
            if selected_schools_filter != st.session_state.selected_schools_state:
                st.session_state.selected_schools_state = selected_schools_filter

            valid_map_df = processed_df.dropna(subset=[col_lat_from, col_lon_from])
            if len(valid_map_df) > 0:
                m_lat = valid_map_df[col_lat_from].mean()
                m_lon = valid_map_df[col_lon_from].mean()
                
                transfer_map = folium.Map(
                    location=[m_lat, m_lon], 
                    zoom_start=11, 
                    tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
                    attr='Esri'
                )
                
                active_selected = st.session_state.selected_schools_state

                for _, row in valid_map_df.iterrows():
                    from_name = str(row.get(target_school_col, ''))
                    to_name = str(row.get('المدرسة المنقول اليها', ''))
                    students_cnt = row.get(students_col, 0)
                    distance_val = row['المسافة_رقمية']
                    side_val = str(row.get('الجهة', '')).strip()
                    ownership_val = str(row.get('الملكية2', 'غير متوفر'))
                    min_grade_val = str(row.get('أدنى صف', 'غير متوفر'))
                    max_grade_val = str(row.get('أعلى صف', 'غير متوفر'))
                    
                    is_mamoun = 'مامون' in side_val.lower() or 'مأمون' in side_val.lower()
                    route_line_color = '#ef4444' if is_mamoun else '#3b82f6'
                    
                    # إزاحة طفيفة لإحداثيات العرض لضمان عدم تطابق المسارات والرموز فوق بعضها
                    offset = 0.00012 if is_mamoun else -0.00012
                    f_lat = row[col_lat_from] + offset
                    f_lon = row[col_lon_from] + offset
                    
                    show_on_map = (len(active_selected) == 0) or (from_name in active_selected) or (to_name in active_selected)
                    
                    if show_on_map:
                        popup_from_html = f"""
                        <div style="font-size: 11px; width: 230px; font-family: Tahoma, sans-serif; line-height: 1.4;">
                            <b>🏫 المدرسة المنقولة:</b><br>{from_name}<br>
                            <b>🎯 المنقول إليها:</b><br>{to_name}<br>
                            <b>📍 الجهة:</b> {side_val}<br>
                            <b>🏢 الملكية2:</b> {ownership_val}<br>
                            <b>📉 أدنى صف:</b> {min_grade_val}<br>
                            <b>📈 أعلى صف:</b> {max_grade_val}<br>
                            <b>👥 عدد الطلاب:</b> {students_cnt}<br><br>
                        </div>
                        """
                        
                        icon_html = f"""
                        <div style="font-size: 13px; font-weight: 900; color: #0f172a; 
                                    background-color: rgba(255, 255, 255, 0.95); padding: 3px 6px; 
                                    border-radius: 6px; border: 2px solid {route_line_color}; white-space: nowrap;
                                    transform: translate(-50%, -30px); box-shadow: 0 4px 10px rgba(0,0,0,0.3);
                                    font-family: Tahoma, sans-serif;">
                            🏫 {from_name} ({side_val})
                        </div>
                        """
                        
                        folium.Marker(
                            location=[f_lat, f_lon],
                            popup=folium.Popup(popup_from_html, max_width=250),
                            tooltip=f"تفاصيل: {from_name} ({side_val})",
                            icon=folium.DivIcon(html=icon_html)
                        ).add_to(transfer_map)
                        
                        if col_lat_to in row and col_lon_to in row and pd.notnull(row[col_lat_to]) and pd.notnull(row[col_lon_to]):
                            t_lat = row[col_lat_to] + offset
                            t_lon = row[col_lon_to] + offset

                            popup_to_html = f"""
                            <div style="font-size: 11px; width: 230px; font-family: Tahoma, sans-serif; line-height: 1.4;">
                                <b>🎯 المدرسة المنقول إليها:</b><br>{to_name}<br>
                                <b>🏫 المدرسة المنقولة:</b><br>{from_name}<br>
                                <b>📍 الجهة:</b> {side_val}<br>
                                <b>👥 الطلبة:</b> {students_cnt}<br>
                                <b>📏 المسافة:</b> {distance_val}
                            </div>
                            """
                            
                            icon_to_html = f"""
                            <div style="font-size: 13px; font-weight: 900; color: #1e3a8a; 
                                        background-color: rgba(254, 240, 138, 0.95); padding: 3px 6px; 
                                        border-radius: 6px; border: 2px solid {route_line_color}; white-space: nowrap;
                                        transform: translate(-50%, -30px); box-shadow: 0 4px 10px rgba(0,0,0,0.3);
                                        font-family: Tahoma, sans-serif;">
                                🎯 {to_name}
                            </div>
                            """
                            
                            folium.Marker(
                                location=[t_lat, t_lon],
                                popup=folium.Popup(popup_to_html, max_width=250),
                                tooltip=f"المدرسة المنقول إليها: {to_name}",
                                icon=folium.DivIcon(html=icon_to_html)
                            ).add_to(transfer_map)

                            if use_actual_roads:
                                raw_route_coords = get_road_route_coordinates(row[col_lat_from], row[col_lon_from], row[col_lat_to], row[col_lon_to])
                                route_coords = [[pt[0] + offset, pt[1] + offset] for pt in raw_route_coords]
                            else:
                                route_coords = [[f_lat, f_lon], [t_lat, t_lon]]

                            folium.PolyLine(
                                locations=route_coords,
                                color=route_line_color, weight=5, opacity=0.85,
                                tooltip=f"طريق ({side_val}) من: {from_name} ➔ إلى: {to_name}"
                            ).add_to(transfer_map)
                
                st_folium(transfer_map, width=1200, height=500, key="interactive_map")
            else:
                st.warning("⚠️ لا توجد إحداثيات كافية لعرض الخريطة.")

        with tab3:
            st.subheader("📋 الجدول التفصيلي للأسطول والمباني والصفوف الدراسية")
            fleet_table = processed_df.copy()
            bus_calc = fleet_table[students_col].apply(lambda x: pd.Series(calculate_buses_detailed(int(x))))
            fleet_table[['باصات كبيرة (55)', 'باصات كوستر (25)', 'باصات صغيرة (16)']] = bus_calc
            
            cols_to_show = ['الجهة', 'المدرسة المنقولة', 'الملكية2', 'أدنى صف', 'أعلى صف', 'المدرسة المنقول اليها', students_col, 'باصات كبيرة (55)', 'باصات كوستر (25)', 'باصات صغيرة (16)']
            if distance_col_name:
                cols_to_show.append(distance_col_name)
            
            st.dataframe(fleet_table[cols_to_show].style.apply(highlight_by_side, axis=1), use_container_width=True)

    else:
        st.error("❌ ملف الإكسل فارغ أو تعذر قراءة البيانات منه.")

except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة الملف: {e}")