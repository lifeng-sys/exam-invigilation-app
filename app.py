import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="自动排考系统（含统考功能）", layout="wide")
st.title("期末考试自动排考系统 Demo - 支持统考同时间段")

# ===== 模板数据 =====
def get_exam_template():
    df = pd.DataFrame({
        "班级": ["生技231", "工程232", "生技231"],
        "科目": ["数学A", "数学A", "生物技术"],
        "是否统考": ["是", "是", "否"],
        "考试类型": ["笔试", "笔试", "笔试"],
    })
    return df

def get_rooms_template():
    df = pd.DataFrame({
        "教室编号": ["一教101", "机房202", "一教102"],
        "是否大教室": ["是", "否", "否"],
        "是否为机房": ["否", "是", "否"],
    })
    return df

def get_teachers_template():
    df = pd.DataFrame({
        "姓名": ["张三", "李四", "王五", "赵六"],
        "可用时段（可选）": ["全天", "全天", "全天", "全天"],
    })
    return df

def get_timeslots_template():
    df = pd.DataFrame({
        "日期": ["2024-06-20", "2024-06-20", "2024-06-21", "2024-06-21"],
        "时间段": ["8:30-9:50", "10:10-11:30", "8:30-9:50", "10:10-11:30"],
    })
    return df

# ===== 模板下载按钮 =====
with st.sidebar:
    st.subheader("下载模板（先下载再录入数据）")
    st.download_button("考试安排模板", get_exam_template().to_csv(index=False).encode(), "exam_subject_class_template.csv", "text/csv")
    st.download_button("教室模板", get_rooms_template().to_csv(index=False).encode(), "rooms_template.csv", "text/csv")
    st.download_button("教师模板", get_teachers_template().to_csv(index=False).encode(), "teachers_template.csv", "text/csv")
    st.download_button("时间段模板", get_timeslots_template().to_csv(index=False).encode(), "timeslots_template.csv", "text/csv")

# ===== 上传文件 =====
st.header("1. 上传数据")
uploaded_exam = st.file_uploader("上传考试安排表", type=["csv", "xlsx"])
uploaded_rooms = st.file_uploader("上传教室表", type=["csv", "xlsx"])
uploaded_teachers = st.file_uploader("上传教师表", type=["csv", "xlsx"])
uploaded_timeslots = st.file_uploader("上传可用考试时间段表", type=["csv", "xlsx"])

def load_file(f):
    if f is None:
        return None
    ext = f.name.split('.')[-1].lower()
    if ext in ["xlsx", "xls"]:
        return pd.read_excel(f)
    elif ext == "csv":
        try:
            return pd.read_csv(f, encoding="utf-8")
        except UnicodeDecodeError:
            try:
                return pd.read_csv(f, encoding="gbk")
            except UnicodeDecodeError:
                return pd.read_csv(f, encoding="latin1")
    return None

exam_df = load_file(uploaded_exam)
rooms_df = load_file(uploaded_rooms)
teachers_df = load_file(uploaded_teachers)
timeslots_df = load_file(uploaded_timeslots)

if exam_df is not None:
    st.success("已加载考试安排表")
    st.dataframe(exam_df)
if rooms_df is not None:
    st.success("已加载教室表")
    st.dataframe(rooms_df)
if teachers_df is not None:
    st.success("已加载教师表")
    st.dataframe(teachers_df)
if timeslots_df is not None:
    st.success("已加载时间段表")
    st.dataframe(timeslots_df)

# ===== 排考主逻辑 =====
st.header("2. 自动排考（含统考同时间段功能）")

def auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df):
    df = exam_df.copy()
    df["日期"] = ""
    df["时间段"] = ""
    df["分配教室"] = ""
    df["监考老师1"] = ""
    df["监考老师2"] = ""
    used = set()  # (日期,时间段,教室)
    used_teacher = set()  # (日期,时间段,教师)

    teachers = teachers_df["姓名"].tolist()
    teacher_idx = 0

    # Step1: 统考，按科目+类型分组，组内所有考试分配同一时间段
    tka = df[df["是否统考"]=="是"].groupby(["科目","考试类型"])
    assigned_times = set()

    for (subject, etype), group in tka:
        # 找到第一个未用过的时段
        slot = None
        for i, row in timeslots_df.iterrows():
            key = (row["日期"], row["时间段"])
            if key not in assigned_times:
                slot = key
                assigned_times.add(key)
                break
        if slot is None:
            for idx in group.index:
                df.at[idx, "备注"] = "无可用时间"
            continue
        # 给这个组所有班级分配同一个时间段
        avail_rooms = rooms_df.copy()
        for idx in group.index:
            # 机考仅机房
            r = avail_rooms
            if etype == "机考":
                r = r[r["是否为机房"]=="是"]
            else:
                r = r[r["是否为机房"]=="否"]
            # 找可用教室
            this_room = None
            for _, room_row in r.iterrows():
                room_key = (slot[0], slot[1], room_row["教室编号"])
                if room_key not in used:
                    this_room = room_row["教室编号"]
                    used.add(room_key)
                    break
            if this_room is None:
                df.at[idx, "备注"] = "无可用教室"
                continue
            # 分配老师
            tcnt = 2 if room_row["是否大教室"]=="是" else 1
            assigned_teachers = []
            for _ in range(tcnt):
                for _ in range(len(teachers)):
                    t = teachers[teacher_idx % len(teachers)]
                    teacher_idx += 1
                    if (slot[0], slot[1], t) not in used_teacher:
                        assigned_teachers.append(t)
                        used_teacher.add((slot[0], slot[1], t))
                        break
            # 写入
            df.at[idx, "日期"] = slot[0]
            df.at[idx, "时间段"] = slot[1]
            df.at[idx, "分配教室"] = this_room
            df.at[idx, "监考老师1"] = assigned_teachers[0] if assigned_teachers else ""
            df.at[idx, "监考老师2"] = assigned_teachers[1] if tcnt==2 and len(assigned_teachers)>1 else ""
    # Step2: 普通考试，逐一分配未用时段
    for i, row in df[df["是否统考"]!="是"].iterrows():
        etype = row["考试类型"]
        slot = None
        for idx, ts in timeslots_df.iterrows():
            key = (ts["日期"], ts["时间段"])
            # 教室可用
            r = rooms_df
            if etype == "机考":
                r = r[r["是否为机房"]=="是"]
            else:
                r = r[r["是否为机房"]=="否"]
            for _, room_row in r.iterrows():
                room_key = (key[0], key[1], room_row["教室编号"])
                if room_key in used:
                    continue
                # 老师可用
                tcnt = 2 if room_row["是否大教室"]=="是" else 1
                assigned_teachers = []
                for _ in range(tcnt):
                    for _ in range(len(teachers)):
                        t = teachers[teacher_idx % len(teachers)]
                        teacher_idx += 1
                        if (key[0], key[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            break
                if len(assigned_teachers) < tcnt:
                    continue
                # 分配
                slot = key
                used.add(room_key)
                for t in assigned_teachers:
                    used_teacher.add((key[0], key[1], t))
                df.at[i, "日期"] = key[0]
                df.at[i, "时间段"] = key[1]
                df.at[i, "分配教室"] = room_row["教室编号"]
                df.at[i, "监考老师1"] = assigned_teachers[0]
                df.at[i, "监考老师2"] = assigned_teachers[1] if tcnt==2 else ""
                break
            if slot:
                break
        if slot is None:
            df.at[i, "备注"] = "无可用时间或教室/老师"
    return df

if st.button("一键自动排考（含统考同时间段功能）"):
    if (exam_df is not None) and (rooms_df is not None) and (teachers_df is not None) and (timeslots_df is not None):
        result = auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df)
        st.success("排考完成，以下为排考结果：")
        st.dataframe(result)
        towrite = io.BytesIO()
        result.to_excel(towrite, index=False)
        towrite.seek(0)
        st.download_button(
            "下载自动排考结果Excel",
            data=towrite,
            file_name="auto_exam_schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("请确保所有4个表都已上传！")

st.markdown("---")
st.markdown("""
**使用说明**  
1. 下载4个模板（考试安排、教室、教师、考试时间段），录入后上传  
2. 支持“统考科目”同时间段安排，普通科目自动排时间  
3. 自动检测时间、教室、监考老师冲突
""")
