import os
import pandas as pd

base_dir = 'exam_invigilation_scheduler'
template_dir = os.path.join(base_dir, 'templates')
os.makedirs(template_dir, exist_ok=True)

# 1. 老师模板
df_teacher = pd.DataFrame({'姓名': ['张三', '李四', '王五']})
df_teacher.to_excel(os.path.join(template_dir, 'teacher_template.xlsx'), index=False)

# 2. 教室模板
df_room = pd.DataFrame({
    '名称': ['教学楼A101', '教学楼A102', '教学楼A201'],
    '是否大教室': ['否', '是', '否']
})
df_room.to_excel(os.path.join(template_dir, 'room_template.xlsx'), index=False)

# 3. 考试安排模板
df_exam = pd.DataFrame({
    '考试名称': ['期末考试1', '期末考试2', '期末考试3'],
    '班级': ['环境231+232', '化学221', '环境231'],
    '科目': ['高等数学', '有机化学', '物理']
})
df_exam.to_excel(os.path.join(template_dir, 'exam_template.xlsx'), index=False)

# 4. main.py
main_py = import streamlit as st
import pandas as pd
import utils
import io

st.set_page_config(page_title="期末考试监考排班系统", layout="wide")

st.title("期末考试监考排班系统")

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

# ===== 2. 数据读取与缓存，自动保存/续作 =====
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

# ===== 5. 排班结果多维度展示、筛选、修改 =====
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

    # 一键打印
    st.subheader("打印安排表（预览）")
    html_print = utils.generate_print_html(schedule)
    st.components.v1.html(html_print, height=800, scrolling=True)

    # 如需PDF导出，取消注释
    # st.download_button("下载打印版PDF", utils.to_pdf(html_print), file_name="打印安排表.pdf")

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

with open(os.path.join(base_dir, 'main.py'), 'w', encoding='utf-8') as f:
    f.write(main_py)

# 5. utils.py
utils_py = import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def data_check(teachers, rooms, exams):
    msg = []
    # 检查空字段
    for df, name in zip([teachers, rooms, exams], ['老师', '教室', '考试']):
        if df.isnull().any().any():
            msg.append(f"{name}表存在空值，请检查。")
    # 去重
    if teachers['姓名'].duplicated().any():
        msg.append("老师表有重复姓名。")
    if rooms['名称'].duplicated().any():
        msg.append("教室表有重复名称。")
    # 字段检查
    for col in ['考试名称', '班级', '科目']:
        if col not in exams.columns:
            msg.append(f"考试表缺少字段：{col}")
    if len(msg) > 0:
        return {'error': True, 'message': '\n'.join(msg)}
    return {'error': False, 'message': '数据校验无问题。'}

def auto_schedule(teachers, rooms, exams, 
                  start_time="2024-07-01 08:30", 
                  slot_duration=120, 
                  daily_slots=3):
    teachers = teachers['姓名'].tolist()
    rooms = rooms.copy()
    exams = exams.copy()
    rooms['is_big'] = rooms['是否大教室'].apply(lambda x: str(x).strip() == "是")
    n_teachers = len(teachers)
    n_rooms = len(rooms)
    slot_dt = datetime.strptime(start_time, "%Y-%m-%d %H:%M")
    exam_slots = []
    i = 0
    while len(exam_slots) < len(exams):
        for slot in range(daily_slots):
            time_str = (slot_dt + timedelta(minutes=slot*slot_duration)).strftime("%Y-%m-%d %H:%M")
            for room in rooms['名称']:
                exam_slots.append({"考试时间": time_str, "教室": room})
                if len(exam_slots) == len(exams):
                    break
            if len(exam_slots) == len(exams):
                break
        slot_dt += timedelta(days=1)
    exams = exams.reset_index(drop=True)
    for i, slot in enumerate(exam_slots):
        exams.loc[i, '考试时间'] = slot['考试时间']
        exams.loc[i, '教室'] = slot['教室']
    exams['需要监考人数'] = exams['教室'].map(dict(zip(rooms['名称'], rooms['is_big'].apply(lambda x: 2 if x else 1))))
    schedule = []
    teacher_workload = {t:0 for t in teachers}
    time_teacher_busy = {}
    for idx, row in exams.iterrows():
        exam_time = row['考试时间']
        classroom = row['教室']
        n_invigilators = int(row['需要监考人数'])
        # 找所有此时间段未分配的老师
        available = [t for t in teachers if not time_teacher_busy.get((exam_time, t), False)]
        available = sorted(available, key=lambda x: teacher_workload[x])
        chosen = []
        for t in available:
            if t not in chosen:
                chosen.append(t)
            if len(chosen) == n_invigilators:
                break
        warn = ""
        if len(chosen) < n_invigilators:
            warn = f"【空岗预警】{exam_time} {classroom} 监考缺口：{n_invigilators-len(chosen)}人"
        for t in chosen:
            teacher_workload[t] += 1
            time_teacher_busy[(exam_time, t)] = True
        schedule.append({
            "考试名称": row['考试名称'],
            "班级": row['班级'],
            "科目": row['科目'],
            "考试时间": exam_time,
            "教室": classroom,
            "监考老师": "、".join(chosen),
            "监考人数": n_invigilators,
            "空岗预警": warn
        })
    schedule_df = pd.DataFrame(schedule)
    warn_msg = "；".join(schedule_df[schedule_df['空岗预警'] != ""].空岗预警.tolist())
    return schedule_df, warn_msg

def to_excel(schedule):
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        schedule.to_excel(writer, index=False)
    return output.getvalue()

def generate_print_html(schedule_df):
    html = """
    <html>
    <head>
    <style>
    @media print {
      .pagebreak { page-break-before: always; }
    }
    table {
      border-collapse: collapse; width: 100%; margin-bottom: 30px;
      font-size: 15px;
    }
    th, td {
      border: 1px solid #444; padding: 7px 10px; text-align: center;
    }
    th { background: #f0f0f0; }
    h2 { margin: 20px 0 10px 0;}
    </style>
    </head>
    <body>
    """
    for idx, group in schedule_df.groupby(['考试时间', '教室']):
        t, room = idx
        html += f'<div class="pagebreak"><h2>考试安排表</h2>'
        html += f"<b>考试时间：</b>{t}　　<b>教室：</b>{room}<br><br>"
        html += "<table><tr>"
        for col in ['考试名称', '班级', '科目', '监考老师', '监考人数']:
            html += f"<th>{col}</th>"
        html += "</tr>"
        for _, row in group.iterrows():
            html += "<tr>"
            for col in ['考试名称', '班级', '科目', '监考老师', '监考人数']:
                html += f"<td>{row[col]}</td>"
            html += "</tr>"
        html += "</table></div>"
    html += "</body></html>"
    return html

def teacher_statistics(schedule):
    stat = []
    for t in set(sum([x.split('、') for x in schedule['监考老师'] if pd.notnull(x)], [])):
        cnt = sum([t in x for x in schedule['监考老师'] if pd.notnull(x)])
        stat.append({"老师": t, "监考次数": cnt})
    return pd.DataFrame(stat).sort_values("监考次数", ascending=False)

def room_statistics(schedule):
    schedule['考试日期'] = schedule['考试时间'].apply(lambda x: str(x)[:10])
    return schedule.groupby(['教室', '考试日期']).size().reset_index(name='考试场次')

def find_abnormal(schedule):
    # 一天内同一老师多次、超标、重复等
    schedule['考试日期'] = schedule['考试时间'].apply(lambda x: str(x)[:10])
    data = []
    for t in set(sum([x.split('、') for x in schedule['监考老师'] if pd.notnull(x)], [])):
        df_t = schedule[schedule['监考老师'].str.contains(t)]
        cnt_per_day = df_t.groupby('考试日期').size()
        for d, c in cnt_per_day.items():
            if c > 1:
                data.append({'老师': t, '日期': d, '当天监考次数': c, '备注': "一天多场"})
    return pd.DataFrame(data)

with open(os.path.join(base_dir, 'utils.py'), 'w', encoding='utf-8') as f:
    f.write(utils_py)

# 6. requirements.txt
with open(os.path.join(base_dir, 'requirements.txt'), 'w', encoding='utf-8') as f:
    f.write("streamlit\npandas\nopenpyxl\nxlsxwriter\n")

print(f"✅ 已在 {base_dir}/ 目录下生成全部文件和Excel模板，可直接运行 Streamlit！")
print("请进入该目录后执行：\n\n    pip install -r requirements.txt\n    streamlit run main.py\n")
