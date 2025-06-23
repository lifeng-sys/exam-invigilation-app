import streamlit as st
import pandas as pd
import io
import plotly.figure_factory as ff
from collections import defaultdict

st.set_page_config(page_title="智能排考系统", layout="wide")
st.title("期末考试智能排考系统（支持单双号分考场与傻瓜式操作）")

# ========== 1. 数据上传与参数设置区 ==========
st.header("步骤1：上传四类模板表格")

colA, colB = st.columns(2)
with colA:
    max_per_day = st.number_input("每位老师每天最多监考场数", min_value=1, max_value=10, value=3)
with colB:
    st.info("如需分单双号请在‘分单双号’列填‘是’，大教室会自动合并考场。")

exam_file = st.file_uploader("考试安排表（.xlsx）", type=["xlsx"])
rooms_file = st.file_uploader("教室表（.xlsx）", type=["xlsx"])
teachers_file = st.file_uploader("教师表（.xlsx）", type=["xlsx"])
timeslots_file = st.file_uploader("考试时间段表（.xlsx）", type=["xlsx"])

# ========== 2. 数据读取函数 ==========
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

# ========== 3. 主排考逻辑（含单双号） ==========
st.header("步骤2：一键自动排考（含单双号分考场）")

def auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df, max_per_day):
    # 复制原始表
    schedule_rows = []
    teachers = teachers_df["姓名"].tolist()
    teacher_stats = defaultdict(lambda: defaultdict(int))  # 老师当天分配场次数
    teacher_total = defaultdict(int)                      # 老师总场次数
    used = set()           # (日期,时间段,教室)
    used_teacher = set()   # (日期,时间段,教师)
    used_class = set()     # (日期,时间段,班级)

    # 统考同组同步排
    tk_groups = exam_df[exam_df["是否统考"]=="是"].groupby(["科目","考试类型"])
    assigned_times = set()
    # 先排统考
    for (subject, etype), group in tk_groups:
        slot = None
        for _, ts in timeslots_df.iterrows():
            key = (ts["日期"], ts["时间段"])
            if key not in assigned_times:
                slot = key
                assigned_times.add(key)
                break
        if slot is None:
            for _, row in group.iterrows():
                schedule_rows.append({
                    **row, "日期": "", "时间段": "", "分配教室": "", "监考老师1": "", "监考老师2": "", "备注": "无可用时间"
                })
            continue
        avail_rooms = rooms_df
        for _, exam_row in group.iterrows():
            # 机考只用机房
            r = avail_rooms
            if etype == "机考":
                r = r[r["是否为机房"]=="是"]
            else:
                r = r[r["是否为机房"]=="否"]
            this_room, room_type = None, None
            for _, room_row in r.iterrows():
                room_key = (slot[0], slot[1], room_row["教室编号"])
                if room_key not in used:
                    this_room = room_row["教室编号"]
                    room_type = room_row["是否大教室"]
                    used.add(room_key)
                    break
            if this_room is None:
                schedule_rows.append({
                    **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": "", "监考老师1": "", "监考老师2": "", "备注": "无可用教室"
                })
                continue
            # 分配老师（均衡场次数）
            tcnt = 2 if room_type=="是" else 1
            assigned_teachers = []
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
                schedule_rows.append({
                    **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": this_room, "监考老师1": "", "监考老师2": "", "备注": "无可用老师"
                })
                continue
            schedule_rows.append({
                **exam_row, "日期": slot[0], "时间段": slot[1], "分配教室": this_room,
                "监考老师1": assigned_teachers[0],
                "监考老师2": assigned_teachers[1] if tcnt==2 else "",
                "备注": ""
            })
            used_class.add((slot[0], slot[1], exam_row["班级"]))

    # 非统考自动排考（含单双号逻辑）
    others = exam_df[exam_df["是否统考"]!="是"]
    for _, exam_row in others.iterrows():
        # 是否需要分单双号
        if str(exam_row.get("分单双号", "")).strip() == "是":
            # 先尝试大教室整体排
            assigned = False
            for _, ts in timeslots_df.iterrows():
                key = (ts["日期"], ts["时间段"])
                r = rooms_df[rooms_df["是否大教室"]=="是"]
                if exam_row["考试类型"] == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                for _, room_row in r.iterrows():
                    room_key = (key[0], key[1], room_row["教室编号"])
                    if room_key in used: continue
                    # 分配老师
                    tcnt = 2
                    assigned_teachers = []
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][key[0]], teacher_total[t])
                    )
                    for t in candidate_teachers:
                        if teacher_stats[t][key[0]] < max_per_day and (key[0], key[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers)==tcnt: break
                    if len(assigned_teachers)<tcnt: continue
                    # 分配
                    used.add(room_key)
                    for t in assigned_teachers:
                        used_teacher.add((key[0], key[1], t))
                        teacher_stats[t][key[0]] += 1
                        teacher_total[t] += 1
                    schedule_rows.append({
                        **exam_row, "日期": key[0], "时间段": key[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0], "监考老师2": assigned_teachers[1],
                        "备注": "大教室，整班不分单双号"
                    })
                    used_class.add((key[0], key[1], exam_row["班级"]))
                    assigned = True
                    break
                if assigned: break
            if assigned: continue
            # 如果没有大教室可排，拆为单双号
            assigned = False
            for _, ts in timeslots_df.iterrows():
                key = (ts["日期"], ts["时间段"])
                # 非大教室（且机考仅机房）
                r = rooms_df[rooms_df["是否大教室"]=="否"]
                if exam_row["考试类型"] == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                rooms_iter = list(r.iterrows())
                if len(rooms_iter)<2: continue  # 需要两间教室
                # 分配两个教室
                room_rows = []
                room_keys = []
                for _, room_row in rooms_iter:
                    room_key = (key[0], key[1], room_row["教室编号"])
                    if room_key not in used:
                        room_rows.append(room_row)
                        room_keys.append(room_key)
                    if len(room_rows)==2: break
                if len(room_rows)<2: continue
                # 分配老师（每考场1人）
                tgroup = []
                candidate_teachers = sorted(
                    teachers,
                    key=lambda t: (teacher_stats[t][key[0]], teacher_total[t])
                )
                for t in candidate_teachers:
                    if teacher_stats[t][key[0]] < max_per_day and (key[0], key[1], t) not in used_teacher:
                        tgroup.append(t)
                    if len(tgroup)==2: break
                if len(tgroup)<2: continue
                # 写入单双号考场
                names = [f"{exam_row['班级']}(单)", f"{exam_row['班级']}(双)"]
                for j in range(2):
                    used.add(room_keys[j])
                    used_teacher.add((key[0], key[1], tgroup[j]))
                    teacher_stats[tgroup[j]][key[0]] += 1
                    teacher_total[tgroup[j]] += 1
                    schedule_rows.append({
                        **exam_row, "班级": names[j], "日期": key[0], "时间段": key[1], "分配教室": room_rows[j]["教室编号"],
                        "监考老师1": tgroup[j], "监考老师2": "", "备注": "分单双号考场"
                    })
                    used_class.add((key[0], key[1], names[j]))
                assigned = True
                break
            if not assigned:
                schedule_rows.append({
                    **exam_row, "日期": "", "时间段": "", "分配教室": "", "监考老师1": "", "监考老师2": "",
                    "备注": "无可用大教室或教室资源拆单双号"
                })
        else:
            # 不分单双号，普通排考
            assigned = False
            for _, ts in timeslots_df.iterrows():
                key = (ts["日期"], ts["时间段"])
                r = rooms_df
                if exam_row["考试类型"] == "机考":
                    r = r[r["是否为机房"]=="是"]
                else:
                    r = r[r["是否为机房"]=="否"]
                for _, room_row in r.iterrows():
                    room_key = (key[0], key[1], room_row["教室编号"])
                    if room_key in used: continue
                    tcnt = 2 if room_row["是否大教室"]=="是" else 1
                    candidate_teachers = sorted(
                        teachers,
                        key=lambda t: (teacher_stats[t][key[0]], teacher_total[t])
                    )
                    assigned_teachers = []
                    for t in candidate_teachers:
                        if teacher_stats[t][key[0]] < max_per_day and (key[0], key[1], t) not in used_teacher:
                            assigned_teachers.append(t)
                            if len(assigned_teachers)==tcnt: break
                    if len(assigned_teachers)<tcnt: continue
                    # 分配
                    used.add(room_key)
                    for t in assigned_teachers:
                        used_teacher.add((key[0], key[1], t))
                        teacher_stats[t][key[0]] += 1
                        teacher_total[t] += 1
                    schedule_rows.append({
                        **exam_row, "日期": key[0], "时间段": key[1], "分配教室": room_row["教室编号"],
                        "监考老师1": assigned_teachers[0],
                        "监考老师2": assigned_teachers[1] if tcnt==2 else "",
                        "备注": ""
                    })
                    used_class.add((key[0], key[1], exam_row["班级"]))
                    assigned = True
                    break
                if assigned: break
            if not assigned:
                schedule_rows.append({
                    **exam_row, "日期": "", "时间段": "", "分配教室": "", "监考老师1": "", "监考老师2": "",
                    "备注": "无可用时间/教室/老师"
                })
    # 输出DataFrame
    schedule_df = pd.DataFrame(schedule_rows)
    return schedule_df, teacher_total

# ========== 4. 自动排考、结果展示与美化 ==========
if st.button("一键自动排考"):
    if exam_df is not None and rooms_df is not None and teachers_df is not None and timeslots_df is not None:
        sched, teacher_total = auto_schedule(exam_df, rooms_df, teachers_df, timeslots_df, max_per_day)
        st.success("排考完成！下方可调整、筛选、导出、可视化")
        # 冲突/单双号行高亮
        def highlight_row(r):
            if "单双号" in str(r["备注"]): return ['background-color: #ffd;']*len(r)
            if "大教室" in str(r["备注"]): return ['background-color: #e6f7ff']*len(r)
            if "无可用" in str(r["备注"]): return ['background-color: #faa']*len(r)
            return ['']*len(r)
        st.subheader("全排考表（可编辑微调后导出）")
        st.dataframe(sched.style.apply(highlight_row, axis=1), use_container_width=True)
        st.download_button("导出完整排考Excel", sched.to_excel(index=False), file_name="exam_schedule_all.xlsx")
        # 可视化与分视图
        st.subheader("筛选视图与导出")
        tab1, tab2, tab3 = st.tabs(["按教师", "按教室", "按班级"])
        with tab1:
            tsel = st.selectbox("教师筛选", ["全部"]+teachers)
            dfv = sched if tsel=="全部" else sched[(sched["监考老师1"]==tsel)|(sched["监考老师2"]==tsel)]
            st.dataframe(dfv)
            st.download_button("导出教师排表", dfv.to_excel(index=False), file_name="exam_by_teacher.xlsx")
        with tab2:
            rsel = st.selectbox("教室筛选", ["全部"]+list(rooms_df["教室编号"]))
            dfv = sched if rsel=="全部" else sched[sched["分配教室"]==rsel]
            st.dataframe(dfv)
            st.download_button("导出教室排表", dfv.to_excel(index=False), file_name="exam_by_room.xlsx")
        with tab3:
            csel = st.selectbox("班级筛选", ["全部"]+list(sched["班级"].unique()))
            dfv = sched if csel=="全部" else sched[sched["班级"]==csel]
            st.dataframe(dfv)
            st.download_button("导出班级排表", dfv.to_excel(index=False), file_name="exam_by_class.xlsx")
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
        st.download_button("导出监考老师工作量表", teacher_stat.to_excel(index=False), file_name="teacher_stat.xlsx")
    else:
        st.warning("请先上传全部4类基础表格！")
import io

# 例：完整排考Excel导出
output = io.BytesIO()
sched.to_excel(output, index=False)
output.seek(0)
st.download_button("导出完整排考Excel", output, file_name="exam_schedule_all.xlsx")

# 教师排表导出
output2 = io.BytesIO()
dfv.to_excel(output2, index=False)
output2.seek(0)
st.download_button("导出教师排表", output2, file_name="exam_by_teacher.xlsx")

# 教室排表导出
output3 = io.BytesIO()
dfv.to_excel(output3, index=False)
output3.seek(0)
st.download_button("导出教室排表", output3, file_name="exam_by_room.xlsx")

# 班级排表导出
output4 = io.BytesIO()
dfv.to_excel(output4, index=False)
output4.seek(0)
st.download_button("导出班级排表", output4, file_name="exam_by_class.xlsx")

# 监考老师工作量导出
output5 = io.BytesIO()
teacher_stat.to_excel(output5, index=False)
output5.seek(0)
st.download_button("导出监考老师工作量表", output5, file_name="teacher_stat.xlsx")
