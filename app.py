import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="自动排考系统 Demo", layout="wide")

st.title("期末考试自动排考系统 Demo")
st.write("上传数据、自动排考、冲突检测，并支持一键下载模板！")

# ---- 模板生成函数 ----
def get_exam_template():
    df = pd.DataFrame({
        "班级": ["生技231", "工程232"],
        "科目": ["生物技术", "计算机应用"],
        "是否统考": ["否", "否"],
        "考试类型": ["笔试", "机考"],
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
        "姓名": ["张三", "李四"],
        "可用时段（可选）": ["全天", "全天"],
    })
    return df

# ---- 模板下载 ----
st.sidebar.subheader("下载空白模板（建议先下载再录入数据）")
col1, col2, col3 = st.sidebar.columns(3)
with col1:
    exam_csv = get_exam_template().to_csv(index=False).encode()
    st.download_button("考试安排模板", exam_csv, "exam_subject_class_template.csv", "text/csv")
with col2:
    room_csv = get_rooms_template().to_csv(index=False).encode()
    st.download_button("教室模板", room_csv, "rooms_template.csv", "text/csv")
with col3:
    teacher_csv = get_teachers_template().to_csv(index=False).encode()
    st.download_button("教师模板", teacher_csv, "teachers_template.csv", "text/csv")

st.markdown("---")

# ---- 上传数据 ----
st.header("1. 上传数据文件")
uploaded_exam = st.file_uploader("上传考试安排表", type=["csv", "xlsx"], key="exam")
uploaded_rooms = st.file_uploader("上传教室表", type=["csv", "xlsx"], key="rooms")
uploaded_teachers = st.file_uploader("上传教师表", type=["csv", "xlsx"], key="teachers")

def load_file(f):
    if f is None:
        return None
    ext = f.name.split('.')[-1]
    if ext in ["xlsx", "xls"]:
        return pd.read_excel(f)
    elif ext == "csv":
        return pd.read_csv(f)
    return None

# 加载数据
exam_df = load_file(uploaded_exam)
rooms_df = load_file(uploaded_rooms)
teachers_df = load_file(uploaded_teachers)

if exam_df is not None:
    st.success("已加载考试安排表")
    st.dataframe(exam_df)
if rooms_df is not None:
    st.success("已加载教室表")
    st.dataframe(rooms_df)
if teachers_df is not None:
    st.success("已加载教师表")
    st.dataframe(teachers_df)

# ---- 自动排考 Demo ----
st.header("2. 自动排考（仅Demo：分配教室与监考老师，基础冲突检测）")

if exam_df is not None and rooms_df is not None and teachers_df is not None:
    if st.button("一键自动排考（基础版Demo）"):
        df = exam_df.copy()
        df["分配教室"] = ""
        df["监考老师1"] = ""
        df["监考老师2"] = ""
        # 可用老师列表
        teacher_list = teachers_df["姓名"].tolist()
        teacher_idx = 0
        # 可用教室列表
        rooms = rooms_df.to_dict(orient="records")
        room_idx = 0

        # Demo分配：按行循环分配教室和老师，实际可扩展为更智能算法
        for i, row in df.iterrows():
            # 跳过统考
            if str(row["是否统考"]).strip() == "是":
                continue
            exam_type = str(row["考试类型"]).strip()
            # 找到符合考试类型的教室
            ok_rooms = [r for r in rooms if ((exam_type == "机考" and r["是否为机房"]=="是") or (exam_type == "笔试" and r["是否为机房"]=="否"))]
            if not ok_rooms:
                df.at[i, "分配教室"] = "无可用教室"
                continue
            # Demo只取第一个可用教室
            this_room = ok_rooms[room_idx % len(ok_rooms)]
            df.at[i, "分配教室"] = this_room["教室编号"]
            # 分配老师
            if this_room["是否大教室"] == "是":
                # 2名监考老师
                df.at[i, "监考老师1"] = teacher_list[teacher_idx % len(teacher_list)]
                df.at[i, "监考老师2"] = teacher_list[(teacher_idx+1) % len(teacher_list)]
                teacher_idx += 2
            else:
                df.at[i, "监考老师1"] = teacher_list[teacher_idx % len(teacher_list)]
                teacher_idx += 1
            room_idx += 1

        st.success("自动排考结果（仅Demo逻辑，可根据需要进一步完善算法）")
        st.dataframe(df)
        # 提供下载
        towrite = io.BytesIO()
        df.to_excel(towrite, index=False)
        towrite.seek(0)
        st.download_button(
            "下载自动排考结果Excel",
            data=towrite,
            file_name="auto_exam_schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.info("此Demo为简单顺序分配，后续可根据实际需求扩展：如排除老师/教室冲突、时间安排等。")

else:
    st.warning("请先上传考试安排表、教室表和教师表三份文件后，进行自动排考。")

st.markdown("---")
st.markdown("""
**使用方法说明**  
1. 点击左侧栏依次下载三个空白模板，录入你的考试/教师/教室数据并上传  
2. 点击“一键自动排考”即可生成基础排考结果  
3. 可导出Excel表，供后续人工调整/公示/导入教务系统  
4. 算法如需升级（如时间/教室/教师冲突检测，公平分配等），可随时联系扩展
""")
