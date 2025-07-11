import streamlit as st
import pandas as pd
import io
from collections import defaultdict

st.set_page_config(page_title="智能排考系统", layout="wide")
st.title("期末考试智能排考系统（教室轮转分配+指定场次+监考老师不足备注）")

# ========== 1. 数据上传与参数设置区 ==========
st.header("步骤1：上传表格（支持指定考试场次）")

specified_file = st.file_uploader("外部指定考试场次表（可选，.xlsx）", type=["xlsx"])
exam_file = st.file_uploader("考试安排表（.xlsx）", type=["xlsx"])
rooms_file = st.file_uploader("教室表（.xlsx）", type=["xlsx"])
teachers_file = st.file_uploader("教师表（.xlsx）", type=["xlsx"])
timeslots_file = st.file_uploader("考试时间段表（.xlsx）", type=["xlsx"])

max_per_day = st.number_input("每位老师每天最多监考场数", min_value=1, max_value=10, value=3)
st.info(
    "所有教室均匀轮转分配，指定场次优先排，监考老师不够时自动备注。每个考试项目所有班级同一时间考试，不跨时间段顺延。"
)

def load_xlsx(f):
    if f is not None:
        return pd.read_excel(f)
    return None

specified_df = load_xlsx(specified_file)
exam_df = load_xlsx(exam_file)
rooms_df = load_xlsx(rooms_file)
teachers_df = load_xlsx(teachers_file)
timeslots_df = load_xlsx(timeslots_file)

def show_table(name, df):
    if df is not None:
        st.success(f"已加载{name}")
        st.dataframe(df)

show_table("外部指定考试场次表", specified_df)
show_table("考试安排表", exam_df)
show_table("教室表", rooms_df)
show_table("教师表", teachers_df)
show_table("考试时间段表", timeslots_df)

if teachers_df is not None:
    teachers = teachers_df["姓名"].tolist()
else:
    teachers = []

# ========== 2. 自动排考 & 指定场次分配 ==========
st.header("步骤2：一键自动排考（教室均匀轮转分配）")

# 教室均匀轮转分配辅助
room_use_count = defaultdict(int)
def get_balanced_rooms(rooms_df, used, slot, etype, is_big=None):
    # etype: "机考" or not
    if etype == "机考":
        filtered = rooms_df[(rooms_df["是否为机房"] == "是")]
    else:
        filtered = rooms_df[(rooms_df["是否为机房"] == "否")]
    if is_big is not None:
        filtered = filtered[filtered["是否大教室"] == ("是" if is_big else "否")]
    room_list = []
    for _, row in filtered.iterrows():
        room_key = (slot[0], slot[1], row["教室编号"])
        if room_key not in used:
            room_list.append(row)
    room_list = sorted(room_list, key=lambda r: room_use_count[r["教室编号"]])
    return room_list

def assign_specified_monitor(specified_df, teachers, teacher_stats, teacher_total, used_teacher, max_per_day):
    rows = []
    if specified_df is None:
        return [], teacher_stats, teacher_total, used_teacher
    for _, row in specified_df.iterrows():
        tcnt = int(row.get("需监考老师数", 1))
        # 选择合适老师且无冲突
        assigned_teachers = []
        candidate_teachers = sorted(
            teachers,
            key=lambda t: (teacher_stats[t][row["日期"]], teacher_total[t])
        )
        for t in candidate_teachers:
            if teacher_stats[t][row["日期"]] < max_per_day and (row["日期"], row["时间段"], t) not in used_teacher:
                assigned_teachers.append(t)
                if len(assigned_teachers) == tcnt:
                    break
        # 补足空缺
        while len(assigned_teachers) < tcnt:
            assigned_teachers.append("")
        # 记录冲突
        for t in assigned_teachers:
            if t:
                used_teacher.add((row["日期"], row["时间段"], t))
                teacher_stats[t][row["日期"]] += 1
                teacher_total[t] += 1
        more_remark = str(row.get("备注", "")) + " 指定场次"
        if any(t == "" for t in assigned_teachers):
            more_remark += "（监考老师待补）"
        rows.append({
            "班级": row["班级"],
            "科目": row["科目"],
            "考试类型": row["考试类型"],
            "日期": row["日期"],
            "时间段": row["时间段"],
            "分配教室": row["教室"],
            "监考老师1": assigned_teachers[0],
            "监考老师2": assigned_teachers[1] if tcnt == 2 else "",
            "备注": more_remark
        })
    return rows, teacher_stats, teacher_total, used_teacher

def auto_schedule(exam_df, rooms_df, teachers, teacher_stats, teacher_total, used_teacher, timeslots_df, max_per_day):
    schedule_rows = []
    used = set()           # (日期,时间段,教室)
    used_class = set()     # (日期,时间段,班级)
    used_slots = set()
    assigned_time_by_proj = dict()  # (科目,考试类型) -> (日期,时间段)
    if exam_df is None:
        return []
    proj_groups = exam_df.groupby(["科目", "考试类型"])
    for (subject, etype), group in proj_groups:
        # 1. 找一个未占用的时间段
        slot = None
        for _, ts in timeslots_df.iterrows():
            key = (ts["日期"], ts["时间段"])
            if key not in used_slots:
                slot = key
                used_slots.add(key)
                assigned_time_by_proj[(subject, etype)] = key
                break
        if slot is None:
            for _, row in group.iterrows():
                schedule_rows.append({
                    **row, "日期": "", "时间段": "", "分配教室": "", "监考老师1": "", "监考老师2": "", "备注": "无可用时间"
                })
            continue
        # 2. 同组考试全部安排在slot
        class_rows = list(group.iterrows())
        assigned_classes = 0
        # ---单双号逻辑---
        for idx, exam_row in class_rows:
            if str(exam_row.get("分单双号", "")).strip() == "是":
                # 优先均匀分配大教室
                big_rooms = get_balanced_rooms(rooms_df, used, slot, etype, is_big=True)
                assigned = False
                for room_row in big_rooms:
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    tcnt = 2
                    assigned_teachers = []
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                    )
                    for t in candidate_teachers:
                        if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers) == tcnt:
                                break
                    while len(assigned_teachers) < tcnt:
                        assigned_teachers.append("")
                    used.add(room_key)
                    room_use_count[room_row["教室编号"]] += 1
                    for t in assigned_teachers:
                        if t:
                            used_teacher.add((slot[0], slot[1], t))
                            teacher_stats[t][slot[0]] += 1
                            teacher_total[t] += 1
                    remark = "大教室，整班不分单双号（统一考试项目同时间段）"
                    if any(t == "" for t in assigned_teachers):
                        remark += "（监考老师待补）"
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0], "监考老师2": assigned_teachers[1],
                        "备注": remark
                    })
                    used_class.add((slot[0], slot[1], exam_row["班级"]))
                    assigned = True
                    assigned_classes += 1
                    break
                if assigned: continue
                # 拆单双号，均匀分配普通教室
                norm_rooms = get_balanced_rooms(rooms_df, used, slot, etype, is_big=False)
                if len(norm_rooms) < 2:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无足够教室分单双号（统一考试项目同时间段）"
                    })
                    continue
                tgroup = []
                candidate_teachers = sorted(
                    teachers,
                    key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                )
                for t in candidate_teachers:
                    if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                        tgroup.append(t)
                    if len(tgroup) == 2:
                        break
                while len(tgroup) < 2:
                    tgroup.append("")
                names = [f"{exam_row['班级']}(单)", f"{exam_row['班级']}(双)"]
                for j in range(2):
                    room_row = norm_rooms[j]
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    used.add(room_key)
                    room_use_count[room_row["教室编号"]] += 1
                    if tgroup[j]:
                        used_teacher.add((slot[0], slot[1], tgroup[j]))
                        teacher_stats[tgroup[j]][slot[0]] += 1
                        teacher_total[tgroup[j]] += 1
                    remark = "分单双号考场（统一考试项目同时间段）"
                    if tgroup[j] == "":
                        remark += "（监考老师待补）"
                    schedule_rows.append({
                        **exam_row, "班级": names[j], "日期": slot[0], "时间段": slot[1], "分配教室": room_row["教室编号"],
                        "监考老师1": tgroup[j], "监考老师2": "",
                        "备注": remark
                    })
                    used_class.add((slot[0], slot[1], names[j]))
                assigned_classes += 1
            else:
                # 普通班级，均匀分配所有合格教室
                avail_rooms = get_balanced_rooms(rooms_df, used, slot, etype)
                assigned = False
                for room_row in avail_rooms:
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    tcnt = 2 if room_row["是否大教室"] == "是" else 1
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                    )
                    assigned_teachers = []
                    for t in candidate_teachers:
                        if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers) == tcnt:
                                break
                    while len(assigned_teachers) < tcnt:
                        assigned_teachers.append("")
                    used.add(room_key)
                    room_use_count[room_row["教室编号"]] += 1
                    for t in assigned_teachers:
                        if t:
                            used_teacher.add((slot[0], slot[1], t))
                            teacher_stats[t][slot[0]] += 1
                            teacher_total[t] += 1
                    remark = "统一考试项目同时间段"
                    if any(t == "" for t in assigned_teachers):
                        remark += "（监考老师待补）"
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0],
                        "监考老师2": assigned_teachers[1] if tcnt == 2 else "",
                        "备注": remark
                    })
                    used_class.add((slot[0], slot[1], exam_row["班级"]))
                    assigned = True
                    assigned_classes += 1
                    break
                if not assigned:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无可用教室/老师（统一考试项目同时间段）"
                    })
    return schedule_rows

if st.button("一键自动排考"):
    if teachers and (rooms_df is not None) and (timeslots_df is not None):
        teacher_stats = defaultdict(lambda: defaultdict(int))
        teacher_total = defaultdict(int)
        used_teacher = set()
        room_use_count.clear()
        # 1. 指定场次优先分配
        specified_sched_rows, teacher_stats, teacher_total, used_teacher = assign_specified_monitor(
            specified_df, teachers, teacher_stats, teacher_total, used_teacher, max_per_day)
        # 2. 自动排考（轮转教室分配）
        auto_sched_rows = auto_schedule(
            exam_df, rooms_df, teachers, teacher_stats, teacher_total, used_teacher, timeslots_df, max_per_day)
        # 3. 合并
        all_sched_df = pd.DataFrame(specified_sched_rows + auto_sched_rows)
        st.success("排考完成！下方可调整、筛选、导出、统计。")
        def highlight_row(r):
            if "单双号" in str(r["备注"]): return ['background-color: #ffd;']*len(r)
            if "大教室" in str(r["备注"]): return ['background-color: #e6f7ff']*len(r)
            if "指定场次" in str(r["备注"]): return ['background-color: #ffeaea']*len(r)
            if "待补" in str(r["备注"]): return ['background-color: #fdd']*len(r)
            if "无可用" in str(r["备注"]): return ['background-color: #faa']*len(r)
            return ['']*len(r)
        st.subheader("全排考表（可编辑微调后导出）")
        st.dataframe(all_sched_df.style.apply(highlight_row, axis=1), use_container_width=True)
        output = io.BytesIO()
        all_sched_df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("导出完整排考Excel", output, file_name="exam_schedule_all.xlsx")
        st.subheader("筛选视图与导出")
        tab1, tab2, tab3 = st.tabs(["按教师", "按教室", "按班级"])
        with tab1:
            tsel = st.selectbox("教师筛选", ["全部"] + teachers)
            dfv = all_sched_df if tsel == "全部" else all_sched_df[
                (all_sched_df["监考老师1"] == tsel) | (all_sched_df["监考老师2"] == tsel)]
            st.dataframe(dfv)
            output2 = io.BytesIO()
            dfv.to_excel(output2, index=False)
            output2.seek(0)
            st.download_button("导出教师排表", output2, file_name="exam_by_teacher.xlsx")
        with tab2:
            room_list = list(all_sched_df["分配教室"].dropna().unique())
            rsel = st.selectbox("教室筛选", ["全部"] + room_list)
            dfv = all_sched_df if rsel == "全部" else all_sched_df[all_sched_df["分配教室"] == rsel]
            st.dataframe(dfv)
            output3 = io.BytesIO()
            dfv.to_excel(output3, index=False)
            output3.seek(0)
            st.download_button("导出教室排表", output3, file_name="exam_by_room.xlsx")
        with tab3:
            class_list = list(all_sched_df["班级"].dropna().unique())
            csel = st.selectbox("班级筛选", ["全部"] + class_list)
            dfv = all_sched_df if csel == "全部" else all_sched_df[all_sched_df["班级"] == csel]
            st.dataframe(dfv)
            output4 = io.BytesIO()
            dfv.to_excel(output4, index=False)
            output4.seek(0)
            st.download_button("导出班级排表", output4, file_name="exam_by_class.xlsx")
        st.subheader("监考老师工作量统计表")
        teacher_stat = pd.DataFrame({
            "教师": list(teacher_total.keys()),
            "总监考场次": list(teacher_total.values()),
        }).sort_values("总监考场次", ascending=False)
        st.dataframe(teacher_stat)
        output5 = io.BytesIO()
        teacher_stat.to_excel(output5, index=False)
        output5.seek(0)
        st.download_button("导出监考老师工作量表", output5, file_name="teacher_stat.xlsx")
    else:
        st.warning("请先上传教师、教室、时间段和其他必需表格！")
