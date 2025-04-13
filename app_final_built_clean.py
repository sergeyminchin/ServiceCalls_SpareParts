
import streamlit as st
import pandas as pd
from io import BytesIO
import datetime as dt
from PIL import Image

st.set_page_config(page_title="ניתוח קריאות שירות וחלקים", page_icon="politex.ico", layout="centered")
logo = Image.open("logo.png")
st.image(logo, width=200)

st.title(" אפליקציית ניתוח קריאות שירות וחלקים")

# --- Upload section
st.header("1️⃣ העלה קבצים")
service_file = st.file_uploader("📄 קובץ קריאות שירות (Excel)", type=['xlsx'], key='service')
parts_file = st.file_uploader("🔧 קובץ חלקים (Excel)", type=['xlsx'], key='parts')

if service_file and parts_file:
    st.success("✔️ הקבצים נטענו בהצלחה")

    service_df_temp = pd.read_excel(service_file)
    service_df_temp['ת. פתיחה'] = pd.to_datetime(service_df_temp['ת. פתיחה'], errors='coerce')
    min_date, max_date = service_df_temp['ת. פתיחה'].min(), service_df_temp['ת. פתיחה'].max()

    st.header("2️⃣ בחר טווח תאריכים לניתוח")
    filter_type = st.selectbox("בחר שיטת סינון", ["רבעון", "חודש", "טווח מותאם אישית"])

    if filter_type == "רבעון":
        quarter = st.selectbox("בחר רבעון", ["Q1", "Q2", "Q3", "Q4"])
        q_map = {"Q1": (1, 3), "Q2": (4, 6), "Q3": (7, 9), "Q4": (10, 12)}
        start_month, end_month = q_map[quarter]
        start_date = dt.datetime(min_date.year, start_month, 1)
        end_date = dt.datetime(min_date.year, end_month, 28) + pd.offsets.MonthEnd(0)
    elif filter_type == "חודש":
        month = st.selectbox("בחר חודש", list(range(1, 13)))
        start_date = dt.datetime(min_date.year, month, 1)
        end_date = dt.datetime(min_date.year, month, 28) + pd.offsets.MonthEnd(0)
    else:
        start_date = st.date_input("מתאריך", value=min_date.date(), min_value=min_date.date(), max_value=max_date.date())
        end_date = st.date_input("עד תאריך", value=max_date.date(), min_value=min_date.date(), max_value=max_date.date())
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

    st.markdown(f"🔍 *מנתח בין:* `{start_date.date()}` ועד `{end_date.date()}`")

    if st.button("🚀 בצע ניתוח"):
        with st.spinner("מריץ ניתוחים..."):

            service_df = service_df_temp.copy()
            parts_df = pd.read_excel(parts_file)
            service_df = service_df[(service_df['ת. פתיחה'] >= start_date) & (service_df['ת. פתיחה'] <= end_date)]

            results = {}
            results["כמות קריאות"] = pd.DataFrame({'סה"כ קריאות': [service_df.shape[0]]})
            results["התפלגות סוגי קריאה"] = service_df['סוג קריאה'].value_counts().reset_index().rename(columns={'index': 'סוג קריאה', 'סוג קריאה': 'כמות'})
            tech_df = service_df[service_df['סוג קריאה'] == 'ביקור טכני']
            results["ביקורים טכניים לפי דגם"] = tech_df['דגם'].value_counts().reset_index().rename(columns={'index': 'דגם', 'דגם': 'כמות ביקורים'})

            if 'דגם' in service_df.columns and 'תאור תקלה' in service_df.columns:
                issues_by_model = (
                    service_df[['דגם', 'תאור תקלה']]
                    .dropna(subset=['דגם', 'תאור תקלה'])
                    .groupby(['דגם', 'תאור תקלה'])
                    .size()
                    .reset_index(name='כמות')
                    .sort_values(by=['דגם', 'כמות'], ascending=[True, False])
                )
                results["תקלות לפי דגם"] = issues_by_model
            else:
                results["תקלות לפי דגם"] = pd.DataFrame(columns=["דגם", "תאור תקלה", "כמות"])

            results["צירופי תקלה ופעולה"] = service_df.groupby(['תאור תקלה', 'תאור קוד פעולה']).size().reset_index(name='כמות').sort_values(by='כמות', ascending=False)
            results["קריאות לפי טכנאי וסוג קריאה"] = service_df.groupby(['לטיפול', 'סוג קריאה']).size().reset_index(name='כמות קריאות')
            results["קריאות לפי אתר"] = service_df['תאור האתר'].value_counts().reset_index().rename(columns={'index': 'תאור האתר', 'תאור האתר': 'כמות קריאות'})

            tech_df = tech_df.copy()
            tech_df['ת. פתיחה'] = pd.to_datetime(tech_df['ת. פתיחה'])
            tech_df.sort_values(by=["מס' מכשיר", "ת. פתיחה"], inplace=True)
            tech_df['prev_open'] = tech_df.groupby("מס' מכשיר")['ת. פתיחה'].shift(1)
            tech_df['prev_technician'] = tech_df.groupby("מס' מכשיר")['לטיפול'].shift(1)
            tech_df['days_diff'] = (tech_df['ת. פתיחה'] - tech_df['prev_open']).dt.days
            tech_df['חוזרת'] = tech_df['days_diff'].apply(lambda x: x <= 30 if pd.notnull(x) else False)
            repeated = tech_df[tech_df['חוזרת']]
            repeat_count = repeated['prev_technician'].value_counts().reset_index()
            repeat_count.columns = ['טכנאי', 'קריאות חוזרות']
            total = tech_df['לטיפול'].value_counts().reset_index()
            total.columns = ['טכנאי', 'סה"כ ביקורים']
            repeat_stats = pd.merge(repeat_count, total, on='טכנאי', how='left')
            repeat_stats['אחוז חוזרות'] = (repeat_stats['קריאות חוזרות'] / repeat_stats['סה"כ ביקורים'] * 100).round(2)
            results["קריאות חוזרות לפי טכנאי"] = repeat_stats

            results["חלקים הכי נפוצים"] = parts_df['תאור מוצר - חלק'].value_counts().reset_index().rename(columns={'index': 'תיאור חלק', 'תאור מוצר - חלק': 'כמות החלפות'})
            results['חלקים לפי מק"ט בטיפול'] = parts_df['מק"ט בטיפול'].value_counts().reset_index().rename(columns={'index': 'מק"ט בטיפול', 'מק"ט בטיפול': 'כמות החלפות'})
            results["חלקים לפי דגם מכונה"] = parts_df.groupby(['מק"ט בטיפול', 'תאור מוצר - חלק']).size().reset_index(name='כמות החלפות')

            
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in results.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        output.seek(0)

        # קביעת שם קובץ מותאם לפי בחירה
        if filter_type == "רבעון":
            period = f"{quarter}_{start_date.year}"
        elif filter_type == "חודש":
            period = f"חודש_{start_date.month}_{start_date.year}"
        else:
            period = f"{start_date.date()}_עד_{end_date.date()}"
        filename = f'Service Calls & Parts Report_{period}.xlsx'

        st.success("🎉 הניתוח הסתיים! ניתן להוריד את הקובץ:")
        st.download_button('📥 Excel הורד דוח', output, file_name=filename)

