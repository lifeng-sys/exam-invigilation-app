import pandas as pd
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
    slot_dt = datetime.strptime(start_time, "%Y-%m-%d %H:%M")
    exam_slots = []
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
    schedule['考试日期'] = schedule['考试时间'].apply(lambda x: str(x)[:10])
    data = []
    for t in set(sum([x.split('、') for x in schedule['监考老师'] if pd.notnull(x)], [])):
        df_t = schedule[schedule['监考老师'].str.contains(t)]
        cnt_per_day = df_t.groupby('考试日期').size()
        for d, c in cnt_per_day.items():
            if c > 1:
                data.append({'老师': t, '日期': d, '当天监考次数': c, '备注': "一天多场"})
    return pd.DataFrame(data)
