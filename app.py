import streamlit as st
import pandas as pd
import io
import plotly.figure_factory as ff
from collections import defaultdict

st.set_page_config(page_title="智能排考系统", layout="wide")
st.title("期末考试智能排考系统（考试项目同时间段+傻瓜式操作）")

# ========== 1. 数据上传与参数设置区 ==========
st.header("步骤1：上传四类模板表格")

colA, colB = st.columns(2)
with colA:
    max_per_day = st.number_input("每位老师每天最多监考场数", min_value=1, max_value=10, value=3)
with colB:
    st.info("所有同一考试项目（科目+考试类型）将自动安排在同一时间段！如需分单双号请填写‘是’，大教室自动合并考场。")

exam_file = st.file_uploader("考试安排表（.xlsx）", type=["xlsx"])
rooms_file = st.file_uploader("教室表（.xlsx）", type=["xlsx"])
teachers_file = st.file_uploader("教师表（.xlsx）", type=["xlsx"])
timeslots_file = st.file_uploader("考试时间段表（.xlsx）", type=["xlsx"])

def load_xlsx(f):
    if f is not None:
        return pd.read_excel(f)
    return None

exam_df = load_xlsx(exam_file)
rooms_df = load_xlsx(rooms_file)
teachers_df = load_xlsx(teachers_file)
timeslots_df = load_xlsx(timeslots_file)

def show_table(name, df):
    if df is not None:
        st.success(f"已加载{name}")
        st.dataframe(df)

show_table("考试安排表", exam_df)
show_table("教室表", rooms_df)
show_table("教师表", teachers_df)
show_table("考试时间段表", timeslots_df)

# ========== 2. 主排考逻辑：考试项目分组，全部同时间 ==========
st.header("步骤2：一键自动排考（同一考试项目同一时间段）")

def auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df, max_per_day):
    # 复制原始表
    schedule_rows = []
    teachers = teachers_df["姓名"].tolist()
    teacher_stats = defaultdict(lambda: defaultdict(int))  # 老师当天分配场次数
    teacher_total = defaultdict(int)                      # 老师总场次数
    used = set()           # (日期,时间段,教室)
    used_teacher = set()   # (日期,时间段,教师)
    used_class = set()     # (日期,时间段,班级)
    assigned_time_by_proj = dict()  # (科目,考试类型) -> (日期,时间段)

    # --- 按考试项目分组，同组所有考试排同一时间 ---
    proj_groups = exam_df.groupby(["科目", "考试类型"])
    used_slots = set()
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
        for _, exam_row in group.iterrows():
            # 单双号逻辑
            if str(exam_row.get("分单双号", "")).strip() == "是":
                # 先尝试大教室整体排
                assigned = False
                r = rooms_df[rooms_df["是否大教室"]=="是"]
                if etype == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                for _, room_row in r.iterrows():
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    if room_key in used: continue
                    tcnt = 2
                    assigned_teachers = []
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                    )
                    for t in candidate_teachers:
                        if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers)==tcnt: break
                    if len(assigned_teachers)<tcnt: continue
                    # 分配
                    used.add(room_key)
                    for t in assigned_teachers:
                        used_teacher.add((slot[0], slot[1], t))
                        teacher_stats[t][slot[0]] += 1
                        teacher_total[t] += 1
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0], "监考老师2": assigned_teachers[1],
                        "备注": "大教室，整班不分单双号（统一考试项目同时间段）"
                    })
                    used_class.add((slot[0], slot[1], exam_row["班级"]))
                    assigned = True
                    break
                if assigned: continue
                # 没大教室，拆单双号
                r = rooms_df[rooms_df["是否大教室"]=="否"]
                if etype == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                rooms_iter = list(r.iterrows())
                if len(rooms_iter)<2:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无足够教室分单双号（统一考试项目同时间段）"
                    })
                    continue
                # 分配两个教室
                room_rows = []
                room_keys = []
                for _, room_row in rooms_iter:
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    if room_key not in used:
                        room_rows.append(room_row)
                        room_keys.append(room_key)
                    if len(room_rows)==2: break
                if len(room_rows)<2:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无足够教室分单双号（统一考试项目同时间段）"
                    })
                    continue
                # 分配老师
                tgroup = []
                candidate_teachers = sorted(
                    teachers,
                    key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                )
                for t in candidate_teachers:
                    if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                        tgroup.append(t)
                    if len(tgroup)==2: break
                if len(tgroup)<2:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无足够老师分单双号（统一考试项目同时间段）"
                    })
                    continue
                # 写入单双号
                names = [f"{exam_row['班级']}(单)", f"{exam_row['班级']}(双)"]
                for j in range(2):
                    used.add(room_keys[j])
                    used_teacher.add((slot[0], slot[1], tgroup[j]))
                    teacher_stats[tgroup[j]][slot[0]] += 1
                    teacher_total[tgroup[j]] += 1
                    schedule_rows.append({
                        **exam_row, "班级": names[j], "日期": slot[0], "时间段": slot[1], "分配教室": room_rows[j]["教室编号"],
                        "监考老师1": tgroup[j], "监考老师2": "",
                        "备注": "分单双号考场（统一考试项目同时间段）"
                    })
                    used_class.add((slot[0], slot[1], names[j]))
            else:
                # 不分单双号，普通排考
                assigned = False
                r = rooms_df
                if etype == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                for _, room_row in r.iterrows():
                    room_key = (slot[0], slot[1], room_row["教室编号"])
                    if room_key in used: continue
                    tcnt = 2 if room_row["是否大教室"]=="是" else 1
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][slot[0]], teacher_total[t])
                    )
                    assigned_teachers = []
                    for t in candidate_teachers:
                        if teacher_stats[t][slot[0]] < max_per_day and (slot[0], slot[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers)==tcnt: break
                    if len(assigned_teachers)<tcnt: continue
                    # 分配
                    used.add(room_key)
                    for t in assigned_teachers:
                        used_teacher.add((slot[0], slot[1], t))
                        teacher_stats[t][slot[0]] += 1
                        teacher_total[t] += 1
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0],
                        "监考老师2": assigned_teachers[1] if tcnt==2 else "",
                        "备注": "统一考试项目同时间段"
                    })
                    used_class.add((slot[0], slot[1], exam_row["班级"]))
                    assigned = True
                    break
                if not assigned:
                    schedule_rows.append({
                        **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "",
                        "监考老师1": "", "监考老师2": "",
                        "备注": "无可用教室/老师（统一考试项目同时间段）"
                    })
    # 输出DataFrame
    schedule_df = pd.DataFrame(schedule_rows)
    return schedule_df, teacher_total

# ========== 3. 自动排考、结果展示与美化 ==========
if st.button("一键自动排考"):
    if exam_df is not None and rooms_df is not None and teachers_df is not None and timeslots_df is not None:
        sched, teacher_total = auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df, max_per_day)
        st.success("排考完成！下方可调整、筛选、导出、可视化")
        # 冲突/单双号行高亮
        def highlight_row(r):
            if "单双号" in str(r["备注"]): return ['background-color: #ffd;']*len(r)
            if "大教室" in str(r["备注"]): return ['background-color: #e6f7ff']*len(r)
            if "统一考试项目同时间段" in str(r["备注"]): return ['background-color: #f7e5fa']*len(r)
            if "无可用" in str(r["备注"]): return ['background-color: #faa']*len(r)
            return ['']*len(r)
        st.subheader("全排考表（可编辑微调后导出）")
        st.dataframe(sched.style.apply(highlight_row, axis=1), use_container_width=True)
        # Excel导出标准写法
        output = io.BytesIO()
        sched.to_excel(output, index=False)
        output.seek(0)
        st.download_button("导出完整排考Excel", output, file_name="exam_schedule_all.xlsx")
        # 可视化与分视图
        st.subheader("筛选视图与导出")
        tab1, tab2, tab3 = st.tabs(["按教师", "按教室", "按班级"])
        with tab1:
            tsel = st.selectbox("教师筛选", ["全部"]+teachers)
            dfv = sched if tsel=="全部" else sched[(sched["监考老师1"]==tsel)|(sched["监考老师2"]==tsel)]
            st.dataframe(dfv)
            output2 = io.BytesIO()
            dfv.to_excel(output2, index=False)
            output2.seek(0)
            st.download_button("导出教师排表", output2, file_name="exam_by_teacher.xlsx")
        with tab2:
            rsel = st.selectbox("教室筛选", ["全部"]+list(rooms_df["教室编号"]))
            dfv = sched if rsel=="全部" else sched[sched["分配教室"]==rsel]
            st.dataframe(dfv)
            output3 = io.BytesIO()
            dfv.to_excel(output3, index=False)
            output3.seek(0)
            st.download_button("导出教室排表", output3, file_name="exam_by_room.xlsx")
        with tab3:
            csel = st.selectbox("班级筛选", ["全部"]+list(sched["班级"].unique()))
            dfv = sched if csel=="全部" else sched[sched["班级"]==csel]
            st.dataframe(dfv)
            output4 = io.BytesIO()
            dfv.to_excel(output4, index=False)
            output4.seek(0)
            st.download_button("导出班级排表", output4, file_name="exam_by_class.xlsx")
        # 甘特图可视化
        st.subheader("教室使用甘特图可视化")
        try:
            gantt_data = []
            for _, row in sched.iterrows():
                if pd.isna(row["分配教室"]) or pd.isna(row["日期"]) or pd.isna(row["时间段"]): continue
                gantt_data.append(dict(
                    Task=row["分配教室"],
                    Start=f"{row['日期']} {row['时间段'].split('-')[0]}",
                    Finish=f"{row['日期']} {row['时间段'].split('-')[1]}",
                    Resource=f"{row['科目']}({row['班级']})"
                ))
            if gantt_data:
                fig = ff.create_gantt(gantt_data, group_tasks=True, index_col='Resource', show_colorbar=True, title="教室安排甘特图")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无数据可展示甘特图")
        except Exception as e:
            st.warning(f"甘特图生成失败: {e}")
        # 监考老师工作量
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
        st.warning("请先上传全部4类基础表格！")
