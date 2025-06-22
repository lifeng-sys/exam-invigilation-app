
import streamlit as st
import pandas as pd
from datetime import datetime, time
from io import BytesIO

st.set_page_config(page_title="期末考试监考安排系统", layout="wide")
st.title("📘 期末考试监考安排系统（支持固定文件导入与自定义时间段）")

# 固定文件名读取
@st.cache_data
def load_data():
    teachers_df = pd.read_excel("teachers_template.xlsx", sheet_name="教师名单")
    rooms_df = pd.read_excel("rooms_template.xlsx", sheet_name="教室列表")
    subjects_df = pd.read_excel("exam_subject_class_template.xlsx", sheet_name="科目")
    classes_df = pd.read_excel("exam_subject_class_template.xlsx", sheet_name="班级")
    return teachers_df, rooms_df, subjects_df, classes_df

teachers_df, rooms_df, subjects_df, classes_df = load_data()

# 初始化安排表
if 'schedule_df' not in st.session_state:
    st.session_state.schedule_df = pd.DataFrame(columns=[
        "考试日期", "开始时间", "结束时间", "考试科目", "考试班级",
        "教室编号", "监考老师1", "监考老师2", "冲突检查结果"
    ])

# 用户输入模块（自定义时间 + 科目班级）
st.sidebar.header("🧩 安排考试")
exam_date = st.sidebar.date_input("考试日期", datetime.today())
start_time = st.sidebar.time_input("开始时间", time(9, 0))
end_time = st.sidebar.time_input("结束时间", time(11, 0))

subject = st.sidebar.selectbox("考试科目", subjects_df["科目名称"].tolist())
clazz = st.sidebar.selectbox("考试班级", classes_df["班级名称"].tolist())
room = st.sidebar.selectbox("选择教室", rooms_df["教室编号"].tolist())
teacher1 = st.sidebar.selectbox("监考老师 1", [""] + teachers_df["教师编号"].tolist())
teacher2 = st.sidebar.selectbox("监考老师 2（可选）", [""] + teachers_df["教师编号"].tolist())

# 添加安排
if st.sidebar.button("➕ 添加安排"):
    if start_time >= end_time:
        st.sidebar.error("结束时间必须晚于开始时间")
    else:
        new_row = {
            "考试日期": exam_date.strftime("%Y-%m-%d"),
            "开始时间": start_time.strftime("%H:%M"),
            "结束时间": end_time.strftime("%H:%M"),
            "考试科目": subject,
            "考试班级": clazz,
            "教室编号": room,
            "监考老师1": teacher1,
            "监考老师2": teacher2,
            "冲突检查结果": "待检测"
        }
        st.session_state.schedule_df = pd.concat([st.session_state.schedule_df, pd.DataFrame([new_row])], ignore_index=True)

# 冲突检测函数
def check_conflicts(df):
    df = df.copy()
    df["冲突检查结果"] = "✅ 正常"
    for i, row in df.iterrows():
        date = row["考试日期"]
        s1 = datetime.strptime(row["开始时间"], "%H:%M").time()
        e1 = datetime.strptime(row["结束时间"], "%H:%M").time()
        room_id = row["教室编号"]
        t1 = row["监考老师1"]
        t2 = row["监考老师2"]

        # 教室冲突
        overlap = df[(df.index != i) & (df["考试日期"] == date) & (df["教室编号"] == room_id)]
        for _, r in overlap.iterrows():
            s2 = datetime.strptime(r["开始时间"], "%H:%M").time()
            e2 = datetime.strptime(r["结束时间"], "%H:%M").time()
            if not (e1 <= s2 or s1 >= e2):
                df.at[i, "冲突检查结果"] = "❌ 教室冲突"

        # 教师冲突
        for teacher in [t1, t2]:
            if not teacher:
                continue
            overlap = df[(df.index != i) & (df["考试日期"] == date) & ((df["监考老师1"] == teacher) | (df["监考老师2"] == teacher))]
            for _, r in overlap.iterrows():
                s2 = datetime.strptime(r["开始时间"], "%H:%M").time()
                e2 = datetime.strptime(r["结束时间"], "%H:%M").time()
                if not (e1 <= s2 or s1 >= e2):
                    df.at[i, "冲突检查结果"] = f"❌ 教师{teacher}冲突"
    return df

# 应用冲突检查
st.session_state.schedule_df = check_conflicts(st.session_state.schedule_df)

# 显示安排表
st.subheader("📋 当前监考安排")
st.dataframe(st.session_state.schedule_df, use_container_width=True)

# 统计监考次数
def get_teacher_stats(df):
    t1_list = df["监考老师1"].dropna().tolist()
    t2_list = df["监考老师2"].dropna().tolist()
    all_teachers = t1_list + t2_list
    stats = pd.Series(all_teachers).value_counts().reset_index()
    stats.columns = ["教师编号", "监考次数"]
    return stats.sort_values("教师编号")

st.subheader("📊 教师监考次数统计")
teacher_stats = get_teacher_stats(st.session_state.schedule_df)
st.dataframe(teacher_stats, use_container_width=True)

# 导出安排
def convert_df(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="监考安排")
    return output.getvalue()

st.download_button(
    label="📥 导出监考安排为 Excel",
    data=convert_df(st.session_state.schedule_df),
    file_name='监考安排_导出.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

st.caption("Made for Lifeng ✨ | ChatGPT Custom Streamlit App")
