import streamlit as st
import pandas as pd
import utils
import os

st.set_page_config(page_title="期末考试监考排班系统", layout="wide")
st.title("期末考试监考排班系统")

# 自动创建templates及模板文件（仅首次部署时执行一次）
if not os.path.exists('templates'):
    os.makedirs('templates')
if not os.path.exists('templates/teacher_template.xlsx'):
    pd.DataFrame({'姓名': ['张三', '李四', '王五']}).to_excel('templates/teacher_template.xlsx', index=False)
if not os.path.exists('templates/room_template.xlsx'):
    pd.DataFrame({'名称': ['教学楼A101', '教学楼A102', '教学楼A201'], '是否大教室': ['否', '是', '否']}).to_excel('templates/room_template.xlsx', index=False)
if not os.path.exists('templates/exam_template.xlsx'):
    pd.DataFrame({'考试名称': ['期末考试1', '期末考试2', '期末考试3'],
                  '班级': ['环境231+232', '化学221', '环境231'],
                  '科目': ['高等数学', '有机化学', '物理']}).to_excel('templates/exam_template.xlsx', index=False)

# ===== 1. 上传数据 & 下载模板 =====
st.sidebar.header("1. 上传基础数据")
with st.sidebar.expander("下载模板"):
    st.markdown("下载并填写模板，避免格式错误。")
    with open("templates/teacher_template.xlsx", "rb") as f:
        st.download_button("下载老师模板", f, file_name="老师模板.xlsx")
    with open("templates/room_template.xlsx", "rb") as f:
        st.download_button("下载教室模板", f, file_name="教室模板.xlsx")
    with open("templates/exam_template.xlsx", "rb") as f:
        st.download_button("下载考试模板", f, file_name="考试模板.xlsx")

uploaded_teachers = st.sidebar.file_uploader("上传老师表", type="xlsx")
uploaded_rooms = st.sidebar.file_uploader("上传教室表", type="xlsx")
uploaded_exams = st.sidebar.file_uploader("上传考试表", type="xlsx")

@st.cache_data(persist=True, show_spinner=False)
def load_excel(file):
    return pd.read_excel(file) if file else None

teachers = load_excel(uploaded_teachers)
rooms = load_excel(uploaded_rooms)
exams = load_excel(uploaded_exams)

if 'teachers' not in st.session_state and teachers is not None:
    st.session_state['teachers'] = teachers
if 'rooms' not in st.session_state and rooms is not None:
    st.session_state['rooms'] = rooms
if 'exams' not in st.session_state and exams is not None:
    st.session_state['exams'] = exams

# ===== 3. 数据校验/去重/查错/标红冲突 =====
st.header("2. 数据校验与冲突检查")
if all([teachers is not None, rooms is not None, exams is not None]):
    check_report = utils.data_check(teachers, rooms, exams)
    if check_report['error']:
        st.error("存在以下数据问题：")
        st.write(check_report['message'])
    else:
        st.success("数据校验通过，无明显格式或冲突错误。")
else:
    st.info("请上传全部基础数据。")

# ===== 4. 一键智能排班 =====
st.header("3. 一键智能排班")
if st.button("一键智能排班"):
    if all([teachers is not None, rooms is not None, exams is not None]):
        with st.spinner("智能排班中，请稍候……"):
            schedule, warn_msg = utils.auto_schedule(
                teachers, rooms, exams,
                start_time="2024-07-01 08:30",
                slot_duration=120,
                daily_slots=3
            )
            st.session_state['schedule'] = schedule
            st.session_state['warn_msg'] = warn_msg
            st.success("智能排班完成！")
            if warn_msg:
                st.warning("预警：" + warn_msg)
    else:
        st.warning("请先上传完整数据表。")

# ===== 5. 排班结果多维度展示、筛选、导出/打印 =====
st.header("4. 排班结果浏览/筛选/导出/打印")
if 'schedule' in st.session_state:
    schedule = st.session_state['schedule']
    col1, col2, col3, col4 = st.columns(4)
    teacher_filter = col1.selectbox("按老师筛选", ["全部"] + sorted(set(sum([x.split('、') for x in schedule['监考老师'] if pd.notnull(x)], []))))
    room_filter = col2.selectbox("按教室筛选", ["全部"] + schedule['教室'].unique().tolist())
    day_filter = col3.selectbox("按日期筛选", ["全部"] + sorted(schedule['考试时间'].apply(lambda x: str(x)[:10]).unique()))
    class_filter = col4.selectbox("按班级筛选", ["全部"] + sorted(set(sum([str(x).split('+') for x in schedule['班级']], []))))

    filtered = schedule.copy()
    if teacher_filter != "全部":
        filtered = filtered[filtered['监考老师'].str.contains(teacher_filter)]
    if room_filter != "全部":
        filtered = filtered[filtered['教室'] == room_filter]
    if day_filter != "全部":
        filtered = filtered[filtered['考试时间'].str[:10] == day_filter]
    if class_filter != "全部":
        filtered = filtered[filtered['班级'].str.contains(class_filter)]

    st.dataframe(filtered, use_container_width=True)

    if st.session_state.get('warn_msg'):
        st.warning("预警：" + st.session_state['warn_msg'])

    st.download_button(
        "导出排班结果Excel",
        utils.to_excel(schedule),
        file_name="排班结果.xlsx"
    )

    st.subheader("打印安排表（预览）")
    html_print = utils.generate_print_html(schedule)
    st.components.v1.html(html_print, height=800, scrolling=True)

# ===== 6. 监考工作统计 =====
st.header("5. 监考工作统计")
if 'schedule' in st.session_state:
    stats_teacher = utils.teacher_statistics(schedule)
    st.subheader("每位老师监考总次数")
    st.dataframe(stats_teacher)
    stats_room = utils.room_statistics(schedule)
    st.subheader("每个教室每天考试安排统计")
    st.dataframe(stats_room)
    abnormal = utils.find_abnormal(schedule)
    st.subheader("异常监考筛查")
    st.dataframe(abnormal)

st.caption("开发者：Lifeng | 代码适配本地与GitHub/Streamlit | 数据仅本地保存")
