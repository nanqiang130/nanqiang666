import os
import io
import tempfile
import zipfile
from collections import defaultdict

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import subprocess
import matplotlib.font_manager as fm

def _setup_cjk_font():
    cjk_keywords = ['CJK', 'WenQuanYi', 'SimHei', 'Noto Sans SC']
    existing = [f.name for f in fm.fontManager.ttflist
                if any(k in f.name for k in cjk_keywords)]
    if existing:
        return existing[0]
    font_url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/SimplifiedChinese/NotoSansSC-Regular.otf"
    font_dir = os.path.join(os.path.expanduser("~"), ".fonts")
    os.makedirs(font_dir, exist_ok=True)
    font_path = os.path.join(font_dir, "NotoSansSC-Regular.otf")
    if not os.path.exists(font_path):
        try:
            import urllib.request
            urllib.request.urlretrieve(font_url, font_path)
        except:
            return None
    fm.fontManager.addfont(font_path)
    return "Noto Sans SC"

_cjk_font = _setup_cjk_font()
if _cjk_font:
    plt.rcParams['font.sans-serif'] = [_cjk_font, 'DejaVu Sans']
else:
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
SUBJECT_ORDER = ["数学", "语文", "英语", "物理", "化学", "生物", "地理", "历史", "政治", "日语", "俄语"]
SUBJECT_RANK = {s: i for i, s in enumerate(SUBJECT_ORDER)}

# 条件格式颜色
GREEN_FILL = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
RED_FILL = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF")

# 标色规则: (关键词, 绿色阈值, 红色阈值, 是否白字, 是否百分比)
COLOR_RULES = [
    ("日均习题任务", 0.9, 0.9, True, False),
    ("完成率",  0.85, 0.80, True, True),
    ("阅卷率",  0.80, 0.80, True, True),
    ("覆盖率",  0.70, None, True, True),
    ("得分率",  0.80, None, True, True),
]

# 图表色板
COLORS = ['#d97757', '#6a9bcc', '#788c5d', '#c4a35a', '#8b7cb6', '#c97b84', '#5d8c8c']
def parse_rate(val):
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return float(val) / 100 if float(val) > 1 else float(val)
    s = str(val).strip().replace("%", "")
    try:
        f = float(s)
        return f / 100 if f > 1 else f
    except:
        return None


def get_color_rule(col_name):
    if not col_name:
        return None
    name = col_name.strip()
    if "以上" in name or "以下" in name or "标绿" in name or "标红" in name:
        return None
    for keyword, green_above, red_below, white_font, is_percent in COLOR_RULES:
        if keyword in name:
            return (green_above, red_below, white_font, is_percent)
    return None


def detect_cols(ws, header_row=3):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v:
            m[v.strip()] = c
    return m


def sort_rows(data, col_names, subj_col="学科", sort_col="习题任务"):
    si = col_names.index(subj_col) if subj_col in col_names else 0
    ki = col_names.index(sort_col) if sort_col in col_names else -1

    def key(row):
        subj = row[si] or ""
        rank = SUBJECT_RANK.get(subj, 99)
        try:
            val = float(row[ki]) if ki >= 0 and row[ki] not in (None, "") else 0
        except:
            val = 0
        return (rank, -val)

    return sorted(data, key=key)


def write_sheet(wb, name, headers, data, auto_color=True):
    ws = wb.create_sheet(name)
    for ci, h in enumerate(headers, 1):
        ws.cell(1, ci, h)
    for ri, row in enumerate(data, 2):
        for ci, val in enumerate(row, 1):
            ws.cell(ri, ci, val)
    if auto_color:
        col_rules = {}
        for ci, h in enumerate(headers, 1):
            rule = get_color_rule(h)
            if rule:
                col_rules[ci] = rule
        for ri, row in enumerate(data, 2):
            for ci, val in enumerate(row, 1):
                if ci in col_rules:
                    green_above, red_below, use_white, is_percent = col_rules[ci]
                    rate = parse_rate(val) if is_percent else (float(val) if val not in (None, "") else None)
                    cell = ws.cell(ri, ci)
                    if rate is not None:
                        if green_above is not None and rate >= green_above:
                            cell.fill = GREEN_FILL
                            if use_white:
                                cell.font = WHITE_FONT
                        elif red_below is not None and rate < red_below:
                            cell.fill = RED_FILL
                            if use_white:
                                cell.font = WHITE_FONT


def generate_excel(input_bytes):
    """从上传的Excel字节流生成统计表，返回 wb_out 对象"""
    wb_in = openpyxl.load_workbook(io.BytesIO(input_bytes))
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    ws_t = wb_in["教师教学详情"]
    t_cols = detect_cols(ws_t)
    teacher_rows = []
    for r in range(4, ws_t.max_row + 1):
        row = {name: ws_t.cell(r, c).value for name, c in t_cols.items()}
        if row.get("姓名"):
            teacher_rows.append(row)

    s1_cols = ["姓名", "学科", "任务量", "习题任务", "日均习题任务", "生均题量", "完成率", "阅卷率"]
    s1_data = sort_rows([[r.get(c, "") for c in s1_cols] for r in teacher_rows], s1_cols)
    write_sheet(wb_out, "教师基础数据", s1_cols, s1_data)
    wb_out["教师基础数据"].cell(1, 15, "完成率：85%以上标绿，80%以下标红")
    wb_out["教师基础数据"].cell(2, 15, "阅卷率：80%以上标绿，80%以下标红")

    s2_cols = ["姓名", "学科", "习题任务", "重练任务数", "重练覆盖率", "重练完成率", "重练得分率"]
    s2_data = sort_rows([[r.get(c, "") for c in s2_cols] for r in teacher_rows], s2_cols)
    write_sheet(wb_out, "错题重练数据", s2_cols, s2_data)
    wb_out["错题重练数据"].cell(1, 14, "重练覆盖率：70%以上标绿")
    wb_out["错题重练数据"].cell(2, 14, "完成率、得分率：80%以上标绿")

    s3_cols = ["姓名", "学科", "靶向任务", "靶向习题", "靶向完成率", "靶向班", "靶向学生", "打标签", "音频讲解", "视频讲解"]
    s3_data = sort_rows([[r.get(c, "") for c in s3_cols] for r in teacher_rows], s3_cols)
    write_sheet(wb_out, "靶向+资源建设", s3_cols, s3_data, auto_color=False)

    ws_c = wb_in["班级学习行为"]
    c_cols = detect_cols(ws_c)
    s4_cols = ["班级", "习题任务", "师均习题任务", "生均题量", "完成率", "得分率",
               "阅卷率", "重练任务数", "重练覆盖率", "重练完成率", "重练得分率",
               "自学人次", "自学题量", "观看人次", "观看时长"]
    s4_data = []
    for r in range(4, ws_c.max_row + 1):
        subj = ws_c.cell(r, c_cols.get("学科", 3)).value
        if subj and str(subj).strip() == "总计":
            s4_data.append([ws_c.cell(r, c_cols.get(cn, 0)).value if cn in c_cols else "" for cn in s4_cols])
    write_sheet(wb_out, "各班级数据", s4_cols, s4_data, auto_color=False)

    return wb_out


def wb_to_bytes(wb):
    """将workbook转为字节流"""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, dpi=150, bbox_inches='tight', facecolor='white', format='png')
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def chart_teacher_basic(wb):
    """返回 [(标题, 图片bytes, 文件名), ...]"""
    results = []
    ws = wb['教师基础数据']
    data = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name is None:
            break
        data.append({
            '姓名': name,
            '学科': ws.cell(r, 2).value or '未知',
            '日均习题任务': ws.cell(r, 5).value or 0,
            '完成率': ws.cell(r, 7).value or 0,
            '阅卷率': ws.cell(r, 8).value or 0,
        })

    groups = defaultdict(list)
    for r in data:
        groups[r['学科']].append(r)

    for subj, items in groups.items():
        names = [r['姓名'] for r in items]
        fig, ax = plt.subplots(figsize=(max(8, len(names)*0.8), 5))
        x = np.arange(len(names))
        w = 0.35
        ax.bar(x - w/2, [r['完成率'] for r in items], w, label='完成率', color=COLORS[0])
        ax.bar(x + w/2, [r['阅卷率'] for r in items], w, label='阅卷率', color=COLORS[1])
        ax.set_title(f'{subj} - 教师完成率与阅卷率', fontsize=14)
        ax.set_ylabel('百分比 (%)')
        ax.set_xticks(x)
        ax.set_xticklabels(names, rotation=45, ha='right', fontsize=9)
        ax.legend()
        ax.set_ylim(0, 105)
        ax.axhline(y=80, color='#999', linestyle='--', linewidth=0.8)
        ax.grid(axis='y', alpha=0.3)
        results.append((f'{subj} - 完成率与阅卷率', fig_to_bytes(fig), f'教师基础数据_{subj}_柱状图.png'))

    # 日均习题任务
    fig, ax = plt.subplots(figsize=(max(12, len(data)*0.5), 6))
    names = [r['姓名'] for r in data]
    daily = [r['日均习题任务'] for r in data]
    colors_bar = [COLORS[2] if v >= 0.9 else COLORS[0] for v in daily]
    ax.bar(names, daily, color=colors_bar)
    ax.set_title('教师日均习题任务', fontsize=14)
    ax.set_ylabel('日均习题任务')
    ax.axhline(y=0.9, color='#999', linestyle='--', linewidth=0.8)
    plt.xticks(rotation=60, ha='right', fontsize=8)
    ax.grid(axis='y', alpha=0.3)
    results.append(('教师日均习题任务', fig_to_bytes(fig), '教师基础数据_日均习题任务_柱状图.png'))

    return results


def chart_error_retrain(wb):
    results = []
    ws = wb['错题重练数据']
    data = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name is None:
            break
        data.append({
            '姓名': name,
            '学科': ws.cell(r, 2).value or '未知',
            '重练覆盖率': ws.cell(r, 5).value or 0,
            '重练完成率': ws.cell(r, 6).value or 0,
            '重练得分率': ws.cell(r, 7).value or 0,
        })

    groups = defaultdict(list)
    for r in data:
        groups[r['学科']].append(r)

    for subj, items in groups.items():
        names = [r['姓名'] for r in items]
        fig, ax = plt.subplots(figsize=(max(8, len(names)*1.0), 5))
        x = np.arange(len(names))
        w = 0.25
        ax.bar(x - w, [r['重练覆盖率'] for r in items], w, label='重练覆盖率', color=COLORS[0])
        ax.bar(x, [r['重练完成率'] for r in items], w, label='重练完成率', color=COLORS[1])
        ax.bar(x + w, [r['重练得分率'] for r in items], w, label='重练得分率', color=COLORS[2])
        ax.set_title(f'{subj} - 错题重练数据', fontsize=14)
        ax.set_ylabel('百分比 (%)')
        ax.set_xticks(x)
        ax.set_xticklabels(names, rotation=45, ha='right', fontsize=9)
        ax.legend()
        ax.set_ylim(0, 110)
        ax.axhline(y=80, color='#999', linestyle='--', linewidth=0.8)
        ax.grid(axis='y', alpha=0.3)
        results.append((f'{subj} - 错题重练数据', fig_to_bytes(fig), f'错题重练数据_{subj}_柱状图.png'))

    return results


def chart_targeting(wb):
    results = []
    ws = wb['靶向+资源建设']
    data_all = []
    data_with_target = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name is None:
            break
        row = {
            '姓名': name,
            '靶向任务': ws.cell(r, 3).value,
            '靶向习题': ws.cell(r, 4).value,
            '靶向完成率': ws.cell(r, 5).value,
        }
        data_all.append(row)
        if row['靶向任务'] is not None:
            data_with_target.append(row)

    if not data_with_target:
        return results

    names = [r['姓名'] for r in data_with_target]
    fig, ax = plt.subplots(figsize=(max(8, len(names)*0.8), 5))
    x = np.arange(len(names))
    w = 0.35
    ax.bar(x - w/2, [r['靶向任务'] or 0 for r in data_with_target], w, label='靶向任务', color=COLORS[0])
    ax.bar(x + w/2, [r['靶向习题'] or 0 for r in data_with_target], w, label='靶向习题', color=COLORS[1])
    ax.set_title('靶向+资源建设 - 靶向任务与习题', fontsize=14)
    ax.set_ylabel('数量')
    ax.set_xticks(x)
    ax.set_xticklabels(names, rotation=45, ha='right', fontsize=9)
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    results.append(('靶向任务与习题', fig_to_bytes(fig), '靶向资源建设_柱状图.png'))

    data_with_rate = [r for r in data_all if r['靶向完成率'] is not None]
    if data_with_rate:
        fig, ax = plt.subplots(figsize=(max(8, len(data_with_rate)*0.8), 5))
        ax.bar([r['姓名'] for r in data_with_rate],
               [r['靶向完成率'] or 0 for r in data_with_rate], color=COLORS[2])
        ax.set_title('靶向完成率', fontsize=14)
        ax.set_ylabel('百分比 (%)')
        plt.xticks(rotation=45, ha='right', fontsize=9)
        ax.axhline(y=80, color='#999', linestyle='--', linewidth=0.8)
        ax.grid(axis='y', alpha=0.3)
        results.append(('靶向完成率', fig_to_bytes(fig), '靶向资源建设_完成率_柱状图.png'))

    return results


def chart_class_data(wb):
    results = []
    ws = wb['各班级数据']
    data = []
    for r in range(2, ws.max_row + 1):
        cls = ws.cell(r, 1).value
        if cls is None:
            break
        data.append({
            '班级': cls,
            '完成率': ws.cell(r, 5).value or 0,
            '得分率': ws.cell(r, 6).value or 0,
            '阅卷率': ws.cell(r, 7).value or 0,
            '重练覆盖率': ws.cell(r, 9).value or 0,
            '重练完成率': ws.cell(r, 10).value or 0,
            '重练得分率': ws.cell(r, 11).value or 0,
        })

    classes = [r['班级'] for r in data]
    x = np.arange(len(classes))
    w = 0.25

    fig, ax = plt.subplots(figsize=(max(10, len(classes)*0.8), 5.5))
    ax.bar(x - w, [r['完成率'] for r in data], w, label='完成率', color=COLORS[0])
    ax.bar(x, [r['得分率'] for r in data], w, label='得分率', color=COLORS[1])
    ax.bar(x + w, [r['阅卷率'] for r in data], w, label='阅卷率', color=COLORS[2])
    ax.set_title('各班级 - 完成率/得分率/阅卷率', fontsize=14)
    ax.set_ylabel('百分比 (%)')
    ax.set_xticks(x)
    ax.set_xticklabels(classes, rotation=45, ha='right', fontsize=9)
    ax.legend()
    ax.set_ylim(0, 110)
    ax.axhline(y=80, color='#999', linestyle='--', linewidth=0.8)
    ax.grid(axis='y', alpha=0.3)
    results.append(('各班级 - 完成率/得分率/阅卷率', fig_to_bytes(fig), '各班级数据_率指标_柱状图.png'))

    fig, ax = plt.subplots(figsize=(max(10, len(classes)*0.8), 5.5))
    ax.bar(x - w, [r['重练覆盖率'] for r in data], w, label='重练覆盖率', color=COLORS[0])
    ax.bar(x, [r['重练完成率'] for r in data], w, label='重练完成率', color=COLORS[1])
    ax.bar(x + w, [r['重练得分率'] for r in data], w, label='重练得分率', color=COLORS[2])
    ax.set_title('各班级 - 重练数据', fontsize=14)
    ax.set_ylabel('百分比 (%)')
    ax.set_xticks(x)
    ax.set_xticklabels(classes, rotation=45, ha='right', fontsize=9)
    ax.legend()
    ax.set_ylim(0, 110)
    ax.axhline(y=80, color='#999', linestyle='--', linewidth=0.8)
    ax.grid(axis='y', alpha=0.3)
    results.append(('各班级 - 重练数据', fig_to_bytes(fig), '各班级数据_重练_柱状图.png'))

    return results


def generate_all_charts(wb):
    """生成全部图表，返回 [(标题, 图片bytes, 文件名), ...]"""
    all_charts = []
    all_charts.extend(chart_teacher_basic(wb))
    all_charts.extend(chart_error_retrain(wb))
    all_charts.extend(chart_targeting(wb))
    all_charts.extend(chart_class_data(wb))
    return all_charts


def charts_to_zip(charts):
    """将所有图表打包为zip字节流"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for title, img_bytes, filename in charts:
            zf.writestr(filename, img_bytes)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════
# Streamlit 界面
# ═══════════════════════════════════════════

st.set_page_config(page_title="教学报告生成工具", page_icon="📊", layout="wide")

st.title("教学报告 → 教师数据统计表 + 柱状图")
st.markdown("上传教学报告Excel，一键生成带条件格式的统计表和各维度柱状图")

uploaded_file = st.file_uploader("上传教学报告 Excel 文件", type=['xlsx', 'xls'])

if uploaded_file is not None:
    with st.spinner("正在处理..."):
        try:
            input_bytes = uploaded_file.read()

            # 生成Excel
            wb = generate_excel(input_bytes)
            excel_bytes = wb_to_bytes(wb)

            # 生成图表
            charts = generate_all_charts(wb)
            zip_bytes = charts_to_zip(charts)

            st.success(f"处理完成！共生成 {len(charts)} 张图表")

            # 下载区
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="下载统计表 Excel",
                    data=excel_bytes,
                    file_name="教师数据统计表_生成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="下载全部图表 (ZIP)",
                    data=zip_bytes,
                    file_name="教师数据统计图表.zip",
                    mime="application/zip"
                )

            # 图表预览
            st.markdown("---")
            st.subheader("图表预览")

            # 按分组展示
            sections = {
                "教师基础数据": [],
                "错题重练数据": [],
                "靶向+资源建设": [],
                "各班级数据": [],
            }
            for title, img_bytes, filename in charts:
                if "教师基础数据" in filename:
                    sections["教师基础数据"].append((title, img_bytes, filename))
                elif "错题重练" in filename:
                    sections["错题重练数据"].append((title, img_bytes, filename))
                elif "靶向" in filename:
                    sections["靶向+资源建设"].append((title, img_bytes, filename))
                elif "各班级" in filename:
                    sections["各班级数据"].append((title, img_bytes, filename))

            for section_name, items in sections.items():
                if not items:
                    continue
                st.markdown(f"**{section_name}**")
                for title, img_bytes, filename in items:
                    st.image(img_bytes, caption=title, use_container_width=True)
                    st.download_button(
                        label=f"下载此图",
                        data=img_bytes,
                        file_name=filename,
                        mime="image/png",
                        key=filename
                    )
                st.markdown("")

        except Exception as e:
            st.error(f"处理失败：{e}")
            st.info("请确认上传的是包含「教师教学详情」和「班级学习行为」Sheet的教学报告文件。")
else:
    st.info("请上传教学报告Excel文件开始使用")
    st.markdown("""
    **使用说明：**
    1. 点击上方区域上传教学报告 Excel 文件
    2. 系统自动生成4个Sheet的统计表（带条件格式标色）
    3. 同时生成各维度的柱状图，可在线预览和下载
    4. 支持一键下载Excel统计表和全部图表（ZIP打包）
    """)
