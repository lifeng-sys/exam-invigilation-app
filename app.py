import streamlit as st
import pandas as pd
import io
import plotly.figure_factory as ff
from collections import defaultdict

st.set_page_config(page_title="升级版自动排考系统", layout="wide")
st.title("期末考试自动排考系统 - 升级版 Demo")

# ======================
# 1. 数据上传与参数配置
# ======================
st.header("1. 数据上传和参数设置")
colA, colB = st.columns(2)
with colA:
    max_per_day = st.number_input("老师每日最大监考场次数", min_value=1, max_value=10, value=3)
with colB:
    balance_mode = st.selectbox("老师负载均衡优先级", ["最少总场次优先", "随机", "顺序"], index=0)

exam_file = st.file_uploader("上传考试安排表", type=["csv", "xlsx"])
rooms_file = st.file_uploader("上传教室表", type=["csv", "xlsx"])
teachers_file = st.file_uploader("上传教师表", type=["csv", "xlsx"])
timeslots_file = st.file_uploader("上传可用考试时间段表", type=["csv", "xlsx"])

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

exam_df = load_file(exam_file)
rooms_df = load_file(rooms_file)
teachers_df = load_file(teachers_file)
timeslots_df = load_file(timeslots_file)

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

# ======================
# 2. 排考主逻辑
# ======================
st.header("2. 一键自动排考")

def auto_schedule_with_balance(exam_df, rooms_df, teachers_df, timeslots_df, max_per_day, balance_mode):
    df = exam_df.copy()
    df["日期"] = ""
    df["时间段"] = ""
    df["分配教室"] = ""
    df["监考老师1"] = ""
    df["监考老师2"] = ""
    df["备注"] = ""

    # 记录老师每日监考场次
    teacher_stats = defaultdict(lambda: defaultdict(int))  # teacher_stats[teacher][date] = count
    teacher_total = defaultdict(int)  # teacher_total[teacher] = count

    used = set()  # (日期,时间段,教室)
    used_teacher = set()  # (日期,时间段,教师)
    used_class = set()  # (日期,时间段,班级)

    teachers = teachers_df["姓名"].tolist()

    # 排统考：同科目同类型必须同一时间段
    tka = df[df["是否统考"]=="是"].groupby(["科目","考试类型"])
    assigned_times = set()

    for (subject, etype), group in tka:
        # 找未用的时段
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
        avail_rooms = rooms_df.copy()
        for idx in group.index:
            # 机考只用机房
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
            # 负载均衡选人
            candidate_teachers = sorted(
                teachers,
                key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
            )
            for t in candidate_teachers:
                if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                    assigned_teachers.append(t)
                    used_teacher.add((slot[0], slot[1], t))
                    teacher_stats[t][slot[0]] += 1
                    teacher_total[t] += 1
                    if len(assigned_teachers) == tcnt:
                        break
            if len(assigned_teachers) < tcnt:
                df.at[idx, "备注"] = "无可用老师"
                continue
            df.at[idx, "日期"] = slot[0]
            df.at[idx, "时间段"] = slot[1]
            df.at[idx, "分配教室"] = this_room
            df.at[idx, "监考老师1"] = assigned_teachers[0]
            df.at[idx, "监考老师2"] = assigned_teachers[1] if tcnt==2 else ""
            used_class.add((slot[0], slot[1], df.at[idx, "班级"]))
    # 普通考试
    for i, row in df[df["是否统考"]!="是"].iterrows():
        etype = row["考试类型"]
        assigned = False
        for idx, ts in timeslots_df.iterrows():
            key = (ts["日期"], ts["时间段"])
            r = rooms_df
            if etype == "机考":
                r = r[r["是否为机房"]=="是"]
            else:
                r = r[r["是否为机房"]=="否"]
            for _, room_row in r.iterrows():
                room_key = (key[0], key[1], room_row["教室编号"])
                if room_key in used:
                    continue
                # 检查班级冲突
                if (key[0], key[1], row["班级"]) in used_class:
                    continue
                # 分配老师
                tcnt = 2 if room_row["是否大教室"]=="是" else 1
                candidate_teachers = sorted(
                    teachers,
                    key=lambda t: (teacher_stats[t][key[0]], teacher_total[t])
                )
                assigned_teachers = []
                for t in candidate_teachers:
                    if teacher_stats[t][key[0]] < max_per_day and (key[0], key[1], t) not in used_teacher:
                        assigned_teachers.append(t)
                        # 注意：不能提前+1，后面统一加
                        if len(assigned_teachers) == tcnt:
                            break
                if len(assigned_teachers) < tcnt:
                    continue
                # 分配
                df.at[i, "日期"] = key[0]
                df.at[i, "时间段"] = key[1]
                df.at[i, "分配教室"] = room_row["教室编号"]
                df.at[i, "监考老师1"] = assigned_teachers[0]
                df.at[i, "监考老师2"] = assigned_teachers[1] if tcnt==2 else ""
                used.add(room_key)
                for t in assigned_teachers:
                    used_teacher.add((key[0], key[1], t))
                    teacher_stats[t][key[0]] += 1
                    teacher_total[t] += 1
                used_class.add((key[0], key[1], row["班级"]))
                assigned = True
                break
            if assigned:
                break
        if not assigned:
            df.at[i, "备注"] = "无可用时段/教室/老师"

    return df, teacher_stats, teacher_total

if st.button("一键自动排考"):
    if (exam_df is not None) and (rooms_df is not None) and (teachers_df is not None) and (timeslots_df is not None):
        schedule, teacher_stats, teacher_total = auto_schedule_with_balance(
            exam_df, rooms_df, teachers_df, timeslots_df, max_per_day, balance_mode
        )
        st.success("排考完成！可在下方调整或导出。")
        # 3. 冲突高亮
        st.dataframe(schedule.style.applymap(lambda x: 'background-color: #ffa' if str(x).startswith('无可用') else ''))
        # 4. 微调：用data_editor
        st.markdown("### 4. 结果微调")
        editable_schedule = st.data_editor(schedule, num_rows="dynamic", use_container_width=True, key="edit_sched")
        # 6. 一键导出（多视图切换）
        st.markdown("### 6. 导出/筛选多视图")
        tab1, tab2, tab3 = st.tabs(["按教师", "按教室", "按班级"])
        with tab1:
            teacher_sel = st.selectbox("选择教师", ["全部"]+list(teachers_df["姓名"]))
            df_view = editable_schedule if teacher_sel=="全部" else editable_schedule[
                (editable_schedule["监考老师1"]==teacher_sel) | (editable_schedule["监考老师2"]==teacher_sel)
            ]
            st.dataframe(df_view)
            st.download_button("导出本视图Excel", df_view.to_excel(index=False), file_name="exam_by_teacher.xlsx")
        with tab2:
            room_sel = st.selectbox("选择教室", ["全部"]+list(rooms_df["教室编号"]))
            df_view = editable_schedule if room_sel=="全部" else editable_schedule[editable_schedule["分配教室"]==room_sel]
            st.dataframe(df_view)
            st.download_button("导出本视图Excel", df_view.to_excel(index=False), file_name="exam_by_room.xlsx")
        with tab3:
            class_sel = st.selectbox("选择班级", ["全部"]+list(exam_df["班级"].unique()))
            df_view = editable_schedule if class_sel=="全部" else editable_schedule[editable_schedule["班级"]==class_sel]
            st.dataframe(df_view)
            st.download_button("导出本视图Excel", df_view.to_excel(index=False), file_name="exam_by_class.xlsx")
        # 5. 甘特图可视化
        st.markdown("### 5. 排考甘特图可视化（按教室）")
        try:
            gantt_data = []
            for _, row in editable_schedule.iterrows():
                if pd.isna(row["分配教室"]) or pd.isna(row["日期"]) or pd.isna(row["时间段"]): continue
                gantt_data.append(dict(
                    Task=row["分配教室"],
                    Start=f"{row['日期']} {row['时间段'].split('-')[0]}",
                    Finish=f"{row['日期']} {row['时间段'].split('-')[1]}",
                    Resource=row["科目"]+"("+str(row["班级"])+")"
                ))
            if gantt_data:
                fig = ff.create_gantt(gantt_data, group_tasks=True, index_col='Resource', show_colorbar=True, title="考场安排甘特图")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无数据可展示甘特图")
        except Exception as e:
            st.error(f"甘特图生成失败: {e}")
        # 监考负载统计
        st.markdown("#### 监考老师工作量统计")
        stat_table = pd.DataFrame({
            "教师": list(teacher_total.keys()),
            "总监考场次": list(teacher_total.values()),
        }).sort_values("总监考场次", ascending=False)
        st.dataframe(stat_table)
        st.download_button("导出监考工作量表", stat_table.to_excel(index=False), file_name="teacher_stat.xlsx")
    else:
        st.warning("请先上传全部4个基础表格！")

st.markdown("---")
st.markdown("""
**功能说明**  
- 自动均衡分配老师负载、每日最多3场  
- 检查所有老师/教室/班级的同时间冲突，结果高亮备注  
- 可页面直接微调，随时导出Excel  
- 可按教师/教室/班级筛选视图与导出  
- 甘特图显示整体教室排班  
- 实时统计监考老师负载  
""")
