
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import plotly.express as px

st.set_page_config(page_title="期末考试监考安排系统", layout="wide")
st.title("📘 期末考试监考安排系统（全功能版）")

# 上传或默认加载数据
def load_data():
    if "teachers_df" not in st.session_state:
        st.session_state.teachers_df = pd.read_excel("teachers_template.xlsx", sheet_name="教师名单")
    if "rooms_df" not in st.session_state:
        st.session_state.rooms_df = pd.read_excel("rooms_template.xlsx", sheet_name="教室列表")
    if "subjects_df" not in st.session_state:
        st.session_state.subjects_df = pd.read_excel("exam_subject_class_template.xlsx", sheet_name="科目")
    if "classes_df" not in st.session_state:
        st.session_state.classes_df = pd.read_excel("exam_subject_class_template.xlsx", sheet_name="班级")

def generate_5min_intervals(start="08:00", end="18:00"):
    tlist = []
    s = datetime.strptime(start, "%H:%M")
    e = datetime.strptime(end, "%H:%M")
    while s < e:
        tlist.append(s.strftime("%H:%M"))
        s += timedelta(minutes=5)
    return tlist

if "schedule_df" not in st.session_state:
    st.session_state.schedule_df = pd.DataFrame(columns=["考试时间段", "教室编号", "监考老师1", "监考老师2", "科目", "班级"])

load_data()

st.sidebar.header("🗂 上传模板文件（可选）")
up_teachers = st.sidebar.file_uploader("教师名单", type=["xlsx"], key="upload_teachers")
if up_teachers:
    st.session_state.teachers_df = pd.read_excel(up_teachers, sheet_name="教师名单")

up_rooms = st.sidebar.file_uploader("教室列表", type=["xlsx"], key="upload_rooms")
if up_rooms:
    st.session_state.rooms_df = pd.read_excel(up_rooms, sheet_name="教室列表")

up_subjects = st.sidebar.file_uploader("考试科目与班级", type=["xlsx"], key="upload_subjects")
if up_subjects:
    st.session_state.subjects_df = pd.read_excel(up_subjects, sheet_name="科目")
    st.session_state.classes_df = pd.read_excel(up_subjects, sheet_name="班级")

teachers_df = st.session_state.teachers_df
rooms_df = st.session_state.rooms_df
subjects_df = st.session_state.subjects_df
classes_df = st.session_state.classes_df
schedule_df = st.session_state.schedule_df

st.sidebar.header("🧩 添加监考安排")
day_label = st.sidebar.selectbox("考试日", ["Day1", "Day2", "Day3"])
time_start = st.sidebar.selectbox("开始时间", generate_5min_intervals())
time_end = st.sidebar.selectbox("结束时间", generate_5min_intervals())
exam_time = f"{day_label} {time_start}-{time_end}"

used_rooms = schedule_df[schedule_df["考试时间段"] == exam_time]["教室编号"].tolist()
available_rooms = [r for r in rooms_df["教室编号"] if r not in used_rooms]
room = st.sidebar.selectbox("教室编号", available_rooms)

used_teachers = schedule_df[schedule_df["考试时间段"] == exam_time][["监考老师1", "监考老师2"]].values.flatten().tolist()
available_teachers = [t for t in teachers_df["教师编号"] if t not in used_teachers]
teacher1 = st.sidebar.selectbox("监考老师 1", [""] + available_teachers)
teacher2 = st.sidebar.selectbox("监考老师 2（可选）", [""] + [t for t in available_teachers if t != teacher1])

subject = st.sidebar.selectbox("考试科目", subjects_df["科目名称"])
class_ = st.sidebar.selectbox("考试班级", classes_df["班级名称"])

if st.sidebar.button("➕ 添加安排"):
    new_row = {
        "考试时间段": exam_time,
        "教室编号": room,
        "监考老师1": teacher1,
        "监考老师2": teacher2,
        "科目": subject,
        "班级": class_
    }
    st.session_state.schedule_df = pd.concat([schedule_df, pd.DataFrame([new_row])], ignore_index=True)
    st.experimental_rerun()

if st.sidebar.button("🧹 清空所有安排"):
    st.session_state.schedule_df = pd.DataFrame(columns=schedule_df.columns)
    st.experimental_rerun()

st.subheader("📋 当前监考安排（可删除）")
for idx, row in schedule_df.iterrows():
    cols = st.columns([10, 1])
    with cols[0]:
        st.write(row.to_frame().T.reset_index(drop=True))
    with cols[1]:
        if st.button("❌ 删除", key=f"del_{idx}"):
            st.session_state.schedule_df = st.session_state.schedule_df.drop(index=idx).reset_index(drop=True)
            st.experimental_rerun()

st.subheader("📊 教师监考次数分布")
all_teachers = schedule_df["监考老师1"].tolist() + schedule_df["监考老师2"].dropna().tolist()
if all_teachers:
    count_df = pd.Series(all_teachers).value_counts().reset_index()
    count_df.columns = ["教师编号", "监考次数"]
    fig = px.bar(count_df, x="教师编号", y="监考次数", title="教师监考频次统计")
    st.plotly_chart(fig, use_container_width=True)

st.subheader("📈 教室使用频率")
if not schedule_df.empty:
    room_count = schedule_df["教室编号"].value_counts().reset_index()
    room_count.columns = ["教室编号", "使用次数"]
    st.dataframe(room_count)

# 导出
def convert_df(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="监考安排")
    return output.getvalue()

st.download_button(
    label="📥 导出监考安排为 Excel",
    data=convert_df(schedule_df),
    file_name="监考安排_导出.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Made for Lifeng ✨ | ChatGPT Streamlit App Plus ✨")
