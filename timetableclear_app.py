"""
================================================================================
小学校時間割自動生成アプリ  timetable_app.py
================================================================================

【実行方法】
1. ライブラリをインストール:
   pip install streamlit openpyxl ortools pandas

2. アプリを起動:
   streamlit run timetable_app.py

3. ブラウザで http://localhost:8501 を開く

【動作環境】
  Python 3.10 以上 / streamlit >= 1.28 / openpyxl >= 3.1
  ortools >= 9.0 / pandas >= 1.5

================================================================================
"""

# ============================================================
# インポート
# ============================================================
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import io
import pandas as pd
from datetime import datetime
import time
import copy

# ============================================================
# ① 固定定数（変更禁止）
# ============================================================

SUBJECTS = [
    "国語", "算数", "理科", "社会", "外国語",
    "体育", "図工", "音楽", "家庭", "道徳",
    "総合", "学活", "生活", "その他1", "その他2", "その他3"
]

DAYS   = ["月", "火", "水", "木", "金"]
GRADES = [1, 2, 3, 4, 5, 6]

COMMON_SUBJECTS = ["体育", "図工", "音楽", "家庭", "外国語"]

COLOR_HEADER = "D9D9D9"
COLOR_MANUAL = "CCE5FF"
COLOR_AUTO   = "CCFFCC"
COLOR_SG     = "FFE5CC"
COLOR_ABSENT = "C0C0C0"

STEP_LABELS = {
    1: "STEP 1: クラス数設定",
    2: "STEP 2: 教員登録",
    3: "STEP 3: コマ数設定",
    4: "STEP 4: 週コマ数設定",
    5: "STEP 5: 担当教員割り当て",
    6: "STEP 6: 不在コマ設定",
    7: "STEP 7: 特別教室設定",
    8: "STEP 8: 手動時間割入力",
    9: "STEP 9: 少人数学級設定",
    0: "🎲 時間割生成・出力",
}

# ============================================================
# ② session_state 全キー初期化
# ============================================================

def init_session_state():
    defaults = {
        "grade_classes":          {g: 1 for g in GRADES},
        "teachers":               [],
        "periods_per_day":        {d: 6 for d in DAYS},
        "periods_per_day_grade":  {g: {d: 6 for d in DAYS} for g in GRADES},
        "weekly_periods":         {},
        "assignments":            {},
        "unavailable":            {},
        "special_rooms":          [],
        "manual_timetable":       {},
        "manual_timetable_sg":    {},
        "small_group_classes":    {},
        "cross_grade_sync":        [],
        "soft_grade_grouping":    True,
        "soft_priority_subjects": [],
        "soft_first_subjects":    [],
        "generated_timetable":    {},
        "generated_timetable_sg": {},
        "timetable_history":      [],   # 過去の生成結果を保持するリスト
        "current_step":           1,
        "timetable_pattern_no":   0,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

# ============================================================
# ③ ヘルパー関数
# ============================================================

def get_all_classes() -> list[str]:
    result = []
    for g in GRADES:
        count = st.session_state["grade_classes"].get(g, 0)
        for c in range(1, count + 1):
            result.append(f"{g}年{c}組")
    return result

def get_grade(class_name: str) -> int:
    return int(class_name[0])

def get_class_index(class_name: str) -> int:
    classes = get_all_classes()
    return classes.index(class_name) if class_name in classes else 999

def get_periods_for_class(cls: str, day: str) -> int:
    """クラスの学年に応じた1日のコマ数を返す。"""
    grade = get_grade(cls)
    return st.session_state["periods_per_day_grade"].get(grade, {}).get(
        day, st.session_state["periods_per_day"].get(day, 6)
    )

def get_total_periods_per_week() -> int:
    return sum(st.session_state["periods_per_day"].values())

def get_max_periods() -> int:
    return max(st.session_state["periods_per_day"].values())

def get_teachers_for_slot(class_name: str, subject: str) -> list[str]:
    raw = st.session_state["assignments"].get(class_name, {}).get(subject, [])
    # 万一 ["山八,金場"] のようにカンマ区切りが1要素リストに入っていた場合も正しく展開する
    result = []
    for item in raw:
        for t in str(item).split(","):
            t = t.strip()
            if t:
                result.append(t)
    return result

def is_special_room_subject(subject: str, cls: str = None) -> tuple[bool, int]:
    grade = get_grade(cls) if cls else None
    for room in st.session_state["special_rooms"]:
        gs = room.get("grade_subjects", [])
        for g, s in gs:
            if s == subject and (grade is None or g == grade):
                return True, room["capacity"]
    return False, 999

def ensure_class_keys():
    cls_list = get_all_classes()
    wp   = st.session_state["weekly_periods"]
    asgn = st.session_state["assignments"]
    for cls in cls_list:
        wp.setdefault(cls, {s: 0 for s in SUBJECTS})
        asgn.setdefault(cls, {s: [] for s in SUBJECTS})
    for cls in list(wp.keys()):
        if cls not in cls_list:
            del wp[cls]
    for cls in list(asgn.keys()):
        if cls not in cls_list:
            del asgn[cls]

# ============================================================
# ④ Excel 共通書式ヘルパー
# ============================================================

def _xl_fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def _xl_border() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _xl_align(wrap: bool = True) -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def _xl_write(ws, row: int, col: int, value,
              color: str | None = None, bold: bool = False,
              width: int | None = None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.border    = _xl_border()
    cell.alignment = _xl_align()
    if color:
        cell.fill = _xl_fill(color)
    if bold:
        cell.font = Font(bold=True)
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return cell

# ============================================================
# Part 2: Excel I/O 補助スタイル関数
# ============================================================

def _apply_header_style(ws, row_num: int = 1):
    """指定行にヘッダースタイル（太字・薄グレー・罫線）を適用する。"""
    fill   = PatternFill("solid", fgColor=COLOR_HEADER)
    font   = Font(bold=True)
    thin   = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[row_num]:
        cell.fill      = fill
        cell.font      = font
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _set_col_widths(ws, width: int = 14):
    """全列の幅を設定する。"""
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = width


def _apply_cell_border(ws):
    """全データセルに罫線を適用する。"""
    thin   = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border


# ============================================================
# Part 2: Excel関数① テンプレート生成
# ============================================================

@st.cache_data
def generate_template_excel() -> bytes:
    """
    全10シートのテンプレートExcelをBytesIOで返す。
    サイドバーのダウンロードボタンに渡す。
    内容は固定なので st.cache_data でキャッシュする。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── シート1: クラス設定 ──────────────────────────────────
    ws = wb.create_sheet("クラス設定")
    ws.append(["学年", "クラス数"])
    for g in GRADES:
        ws.append([f"{g}年", 1])
    _apply_header_style(ws)
    _set_col_widths(ws, 12)

    # ── シート2: 先生リスト (ヘッダーなし: A=先生名) ─────────────
    ws = wb.create_sheet("先生リスト")
    ws.append(["（例）古澤"])
    _set_col_widths(ws, 16)

    # ── シート2b: 教科リスト (ヘッダーなし: A=教科名, B=週コマ数, C=必修フラグ) ──
    ws = wb.create_sheet("教科リスト")
    for subj in SUBJECTS:
        ws.append([subj, 0, 1])
    _set_col_widths(ws, 14)

    # ── シート3: コマ数設定 ──────────────────────────────────
    ws = wb.create_sheet("コマ数設定")
    ws.append(["曜日", "1日のコマ数"])
    for d in DAYS:
        ws.append([d, 6])
    _apply_header_style(ws)
    _set_col_widths(ws, 14)

    # ── シート4: 週コマ数 ────────────────────────────────────
    ws = wb.create_sheet("週コマ数")
    ws.append(["学年クラス"] + SUBJECTS)
    ws.append(["（例）1年1組"] + [0] * len(SUBJECTS))
    _apply_header_style(ws)
    _set_col_widths(ws, 10)
    ws.column_dimensions["A"].width = 14

    # ── シート5: 担当割り当て ────────────────────────────────
    ws = wb.create_sheet("担当割り当て")
    ws.append(["学年クラス"] + SUBJECTS)
    ws.append(["（例）1年1組"] + [""] * len(SUBJECTS))
    ws.cell(row=3, column=1).value = "※複数教員はカンマ区切り（例: 古澤,田中）"
    _apply_header_style(ws)
    _set_col_widths(ws, 10)
    ws.column_dimensions["A"].width = 14

    # ── シート6: 不在コマ ────────────────────────────────────
    ws = wb.create_sheet("不在コマ")
    ws.append(["教員名", "曜日", "時限"])
    ws.append(["（例）古澤", "月", 4])
    ws.append(["（例）古澤", "水", 5])
    _apply_header_style(ws)
    _set_col_widths(ws, 14)

    # ── シート7: 特別教室 ────────────────────────────────────
    ws = wb.create_sheet("特別教室")
    ws.append(["教室名", "対応教科", "同時使用可能クラス数"])
    ws.append(["理科室", "理科", 1])
    ws.append(["体育館", "保体", 1])
    ws.append(["音楽室", "音楽", 1])
    _apply_header_style(ws)
    _set_col_widths(ws, 18)

    # ── シート8: 手動時間割 ──────────────────────────────────
    ws = wb.create_sheet("手動時間割")
    ws.append(["学年クラス", "曜日", "時限", "教科"])
    ws.append(["（例）1年1組", "月", 1, "国語"])
    _apply_header_style(ws)
    _set_col_widths(ws, 14)

    # ── シート9: 少人数学級設定 ──────────────────────────────
    ws = wb.create_sheet("少人数学級設定")
    ws.append([
        "少人数学級名",
        "所属クラス（カンマ区切り）",
        "担当教員（カンマ区切り）",
        "同期グループ番号",
        "グループ内クラス（カンマ区切り）"
    ])
    ws.append(["少人数A", "1年1組,2年3組,3年1組", "田中", 1, "1年1組,3年1組"])
    ws.append(["少人数A", "", "", 2, "2年3組"])
    ws.cell(row=4, column=1).value = "※同期グループ未設定なら列D・Eは空欄"
    _apply_header_style(ws)
    _set_col_widths(ws, 22)

    # ── シート10: 少人数手動時間割 ───────────────────────────
    ws = wb.create_sheet("少人数手動時間割")
    ws.append(["少人数学級名", "曜日", "時限", "教科"])
    ws.append(["（例）少人数A", "月", 2, "音楽"])
    _apply_header_style(ws)
    _set_col_widths(ws, 16)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# Part 1b: 入力正規化ヘルパー
# ============================================================
import unicodedata as _ud

def normalize_value(val) -> str:
    """全角→半角 + strip"""
    if val is None:
        return ""
    return _ud.normalize("NFKC", str(val)).strip()

def normalize_str(val) -> str:
    """normalize_value の別名（互換性）"""
    return normalize_value(val)

def safe_int(val, default: int = 0) -> int:
    """正規化して整数に変換。失敗時は default を返す"""
    try:
        return int(normalize_value(val))
    except (ValueError, TypeError):
        return default

# ============================================================
# Part 2: Excel関数② 設定ファイル読み込み
# ============================================================

def load_settings_from_excel(uploaded_file):
    """
    Excelファイルを読み込み session_state を全て上書きする。
    シートが存在しない場合はそのSTEPをスキップし警告を表示する。
    """
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)

        # ── シート1: クラス設定 ──────────────────────────────
        if "クラス設定" in wb.sheetnames:
            ws = wb["クラス設定"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    try:
                        g = int(str(row[0]).replace("年", ""))
                        val = int(row[1])
                        st.session_state["grade_classes"][g] = val
                        # ウィジェットキャッシュも更新
                        st.session_state[f"grade_{g}"] = val
                    except (ValueError, TypeError):
                        pass

        # ── シート2: 先生リスト (ヘッダーなし, row1から) / 旧: 教員リスト 両対応 ──
        _teacher_sheet = None
        if "先生リスト" in wb.sheetnames:
            _teacher_sheet = wb["先生リスト"]
        elif "教員リスト" in wb.sheetnames:
            _teacher_sheet = wb["教員リスト"]
        if _teacher_sheet is not None:
            teachers = []
            # 先生リストはヘッダーなし (row1 = データ先頭)
            # 教員リストはヘッダーあり (row1 = ヘッダー) → どちらも全行チェック
            for row in _teacher_sheet.iter_rows(min_row=1, values_only=True):
                val = row[0] if row else None
                if (val and str(val).strip()
                        and str(val).strip() not in ("教員名", "先生名")
                        and not str(val).startswith("（例）")):
                    teachers.append(normalize_str(str(val)))
            st.session_state["teachers"] = teachers

        # ── シート2b: 教科リスト (ヘッダーなし: A=教科名, B=週コマ数, C=必修フラグ) ──
        if "教科リスト" in wb.sheetnames:
            ws = wb["教科リスト"]
            subject_defaults = {}
            for row in ws.iter_rows(min_row=1, values_only=True):
                if not row or not row[0]:
                    continue
                subj = normalize_str(str(row[0]))
                if subj not in SUBJECTS:
                    continue
                weekly_cnt = safe_int(row[1]) if len(row) > 1 else 0
                # required_flag = row[2] if len(row) > 2 else 1  # 将来利用可
                subject_defaults[subj] = weekly_cnt
            # 教科リストの週コマ数を全クラスのデフォルトとして適用
            if subject_defaults:
                classes = [f"{g}年{c}組"
                           for g in GRADES
                           for c in range(1, st.session_state["grade_classes"].get(g, 0) + 1)]
                if classes:
                    wp = st.session_state.get("weekly_periods", {})
                    for cls in classes:
                        if cls not in wp:
                            wp[cls] = {}
                        for subj, cnt in subject_defaults.items():
                            if wp[cls].get(subj, 0) == 0:
                                wp[cls][subj] = cnt
                    st.session_state["weekly_periods"] = wp

        # ── シート3: コマ数設定 ──────────────────────────────
        if "コマ数設定" in wb.sheetnames:
            ws = wb["コマ数設定"]
            headers = [c.value for c in ws[1]]
            ppd_grade = st.session_state.setdefault(
                "periods_per_day_grade", {g: {d: 6 for d in DAYS} for g in GRADES}
            )
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                cell0 = str(row[0]).strip()
                # 学年別フォーマット（新）: "1年", "2年"...
                if "年" in cell0:
                    try:
                        g = int(cell0.replace("年", ""))
                    except ValueError:
                        continue
                    ppd_grade.setdefault(g, {})
                    for col_idx, d in enumerate(DAYS, start=1):
                        if col_idx < len(row) and row[col_idx] is not None:
                            try:
                                val = int(row[col_idx])
                                ppd_grade[g][d] = val
                                st.session_state[f"ppd_{g}_{d}"] = val
                            except (ValueError, TypeError):
                                pass
                # 旧フォーマット（曜日別）: "月", "火"...
                elif cell0 in DAYS and len(row) > 1 and row[1] is not None:
                    try:
                        val = int(row[1])
                        for g in GRADES:
                            ppd_grade.setdefault(g, {})[cell0] = val
                            st.session_state[f"ppd_{g}_{cell0}"] = val
                    except (ValueError, TypeError):
                        pass
            # periods_per_day（最大値）を更新
            for d in DAYS:
                st.session_state["periods_per_day"][d] = max(
                    ppd_grade.get(g, {}).get(d, 6) for g in GRADES
                )

        # ── シート4: 週コマ数 ────────────────────────────────
        if "週コマ数" in wb.sheetnames:
            ws = wb["週コマ数"]
            headers = [c.value for c in ws[1]][1:]
            weekly = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                cls = row[0]
                if not cls or str(cls).startswith("（例）"):
                    continue
                cls = str(cls).strip()
                weekly[cls] = {}
                for i, subj in enumerate(headers):
                    if subj and i + 1 < len(row):
                        weekly[cls][subj] = int(row[i + 1] or 0)
            st.session_state["weekly_periods"] = weekly

        # ── シート5: 担当割り当て ────────────────────────────
        if "担当割り当て" in wb.sheetnames:
            ws = wb["担当割り当て"]
            headers = [c.value for c in ws[1]][1:]
            assignments = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                cls = row[0]
                if not cls or str(cls).startswith("（例）") \
                        or str(cls).startswith("※"):
                    continue
                cls = str(cls).strip()
                assignments[cls] = {}
                for i, subj in enumerate(headers):
                    if subj and i + 1 < len(row) and row[i + 1]:
                        teachers = [t.strip()
                                    for t in str(row[i + 1]).split(",")
                                    if t.strip()]
                        assignments[cls][subj] = teachers
                    else:
                        assignments[cls][subj] = []
            st.session_state["assignments"] = assignments

        # ── シート6: 不在コマ ────────────────────────────────
        if "不在コマ" in wb.sheetnames:
            ws = wb["不在コマ"]
            unavail = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                teacher, day, period = row[0], row[1], row[2]
                if not teacher or str(teacher).startswith("（例）"):
                    continue
                teacher = str(teacher).strip()
                if day in DAYS and period:
                    unavail.setdefault(teacher, {}).setdefault(
                        str(day), []).append(int(period))
            st.session_state["unavailable"] = unavail

        # ── シート7: 特別教室 ────────────────────────────────
        if "特別教室" in wb.sheetnames:
            ws = wb["特別教室"]
            rooms = []
            room_map = {}  # name -> room dict
            for row in ws.iter_rows(min_row=2, values_only=True):
                name, grade_val, subj, cap = row[0], row[1], row[2], row[3]
                if not name or str(name).startswith("（例）"):
                    continue
                if not name or not cap:
                    continue
                name = str(name).strip()
                try:
                    cap_int = int(cap)
                except (ValueError, TypeError):
                    cap_int = 1
                if name not in room_map:
                    room_map[name] = {"name": name, "capacity": cap_int, "grade_subjects": []}
                if grade_val and subj:
                    try:
                        grade_int = int(grade_val)
                    except (ValueError, TypeError):
                        grade_int = GRADES[0]
                    room_map[name]["grade_subjects"].append([grade_int, str(subj).strip()])
            rooms = list(room_map.values())
            st.session_state["special_rooms"] = rooms

        # ── シート8: 手動時間割 ──────────────────────────────
        if "手動時間割" in wb.sheetnames:
            ws = wb["手動時間割"]
            manual = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                cls, day, period, subj = row[0], row[1], row[2], row[3]
                if not cls or str(cls).startswith("（例）"):
                    continue
                if cls and day in DAYS and period and subj:
                    cls = str(cls).strip()
                    manual.setdefault(cls, {}).setdefault(
                        str(day), {})[int(period)] = str(subj).strip()
            st.session_state["manual_timetable"] = manual

        # ── シート9: 少人数学級設定 ──────────────────────────
        if "少人数学級設定" in wb.sheetnames:
            ws = wb["少人数学級設定"]
            sg = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                sg_name = row[0]
                if sg_name and not str(sg_name).startswith("（例）") \
                        and not str(sg_name).startswith("※"):
                    sg_name = str(sg_name).strip()
                    if sg_name not in sg:
                        sg[sg_name] = {
                            "classes":        [normalize_str(c)
                                               for c in str(row[1]).split(",")
                                               if c.strip()] if row[1] else [],
                            "teachers":       [t.strip()
                                               for t in str(row[2]).split(",")
                                               if t.strip()] if row[2] else [],
                            "weekly_periods": {},
                            "sync_groups":    []
                        }
                    if row[3] and row[4]:
                        group_classes = [c.strip()
                                         for c in str(row[4]).split(",")
                                         if c.strip()]
                        group_no = int(row[3])
                        while len(sg[sg_name]["sync_groups"]) < group_no:
                            sg[sg_name]["sync_groups"].append([])
                        sg[sg_name]["sync_groups"][group_no - 1] = group_classes
                elif not sg_name and sg:
                    last_sg = list(sg.keys())[-1]
                    if row[3] and row[4]:
                        group_classes = [c.strip()
                                         for c in str(row[4]).split(",")
                                         if c.strip()]
                        group_no = int(row[3])
                        while len(sg[last_sg]["sync_groups"]) < group_no:
                            sg[last_sg]["sync_groups"].append([])
                        sg[last_sg]["sync_groups"][group_no - 1] = group_classes
            st.session_state["small_group_classes"] = sg
            # ウィジェットキーを強制上書き（Excelロード後の古いキャッシュを排除）
            all_valid_classes = get_all_classes()
            for sg_name_, sg_data_ in sg.items():
                st.session_state[f"sg_cls_{sg_name_}"] = [
                    c for c in sg_data_.get("classes", [])
                    if c in all_valid_classes
                ]
                st.session_state[f"sg_teacher_{sg_name_}"] = sg_data_.get("teachers", [])
                st.session_state[f"sg_js_{sg_name_}"] = sg_data_.get("joint_subjects", [])
                for subj_ in sg_data_.get("joint_subjects", []):
                    st.session_state[f"sg_wp_{sg_name_}_{subj_}"] = \
                        sg_data_.get("weekly_periods", {}).get(subj_, 0)
                for gi_, grp_ in enumerate(sg_data_.get("sync_groups", [])):
                    st.session_state[f"sg_group_{sg_name_}_{gi_}"] = grp_

        # ── シート10: 少人数手動時間割 ───────────────────────
        if "少人数手動時間割" in wb.sheetnames:
            ws = wb["少人数手動時間割"]
            manual_sg = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                sg_name, day, period, subj = row[0], row[1], row[2], row[3]
                if not sg_name or str(sg_name).startswith("（例）"):
                    continue
                if sg_name and day in DAYS and period and subj:
                    sg_name = str(sg_name).strip()
                    manual_sg.setdefault(sg_name, {}).setdefault(
                        str(day), {})[int(period)] = str(subj).strip()
            st.session_state["manual_timetable_sg"] = manual_sg

        # ── ウィジェットキーをクリアして各STEPの入力欄を再初期化 ──
        # Streamlit はウィジェットキーが session_state に残っていると
        # value= パラメータより優先するため、Excel ロード後に古いキーを
        # 削除しておかないと読み込んだ値が画面に反映されない。
        # ※ "grade_" は "grade_classes" を誤って削除しないよう、
        #   数字サフィックスのもの（grade_1, grade_2, grade_3）のみ対象にする。
        for _k in list(st.session_state.keys()):
            if _k.startswith("grade_") and _k[6:].isdigit():
                del st.session_state[_k]
            elif _k.startswith(("ppd_", "wp_", "assign_", "unavail_")):
                del st.session_state[_k]

        st.success("✅ Excelファイルから設定を読み込みました（全STEP上書き）")

    except Exception as e:
        st.error(f"❌ 読み込みエラー: {e}")
        st.info("💡 テンプレートExcelを使用して入力してください")


# ============================================================
# Part 2: Excel関数③ 設定ファイル書き出し
# ============================================================

def save_settings_to_excel() -> bytes:
    """
    現在のsession_stateをExcelに書き出してBytesIOで返す。
    generate_template_excel()と同じシート構成・書式で出力する。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # シート1: クラス設定
    ws = wb.create_sheet("クラス設定")
    ws.append(["学年", "クラス数"])
    for g in GRADES:
        ws.append([f"{g}年", st.session_state["grade_classes"].get(g, 0)])
    _apply_header_style(ws); _set_col_widths(ws, 12)

    # シート2: 先生リスト (ヘッダーなし)
    ws = wb.create_sheet("先生リスト")
    for t in st.session_state["teachers"]:
        ws.append([t])
    _set_col_widths(ws, 16)

    # シート2b: 教科リスト (ヘッダーなし)
    ws = wb.create_sheet("教科リスト")
    weekly_p = st.session_state.get("weekly_periods", {})
    classes = [f"{g}年{c}組"
               for g in GRADES
               for c in range(1, st.session_state["grade_classes"].get(g, 0) + 1)]
    first_cls = classes[0] if classes else None
    for subj in SUBJECTS:
        cnt = weekly_p.get(first_cls, {}).get(subj, 0) if first_cls else 0
        ws.append([subj, cnt, 1])
    _set_col_widths(ws, 14)

    # シート3: コマ数設定（学年別）
    ws = wb.create_sheet("コマ数設定")
    ws.append(["学年"] + DAYS)
    ppd_grade = st.session_state.get("periods_per_day_grade", {})
    for g in GRADES:
        row = [f"{g}年"] + [ppd_grade.get(g, {}).get(d, 6) for d in DAYS]
        ws.append(row)
    _apply_header_style(ws); _set_col_widths(ws, 10)

    # シート4: 週コマ数
    ws = wb.create_sheet("週コマ数")
    ws.append(["学年クラス"] + SUBJECTS)
    for cls in get_all_classes():
        row = [cls]
        for subj in SUBJECTS:
            row.append(
                st.session_state["weekly_periods"].get(cls, {}).get(subj, 0))
        ws.append(row)
    _apply_header_style(ws); _set_col_widths(ws, 10)
    ws.column_dimensions["A"].width = 14

    # シート5: 担当割り当て
    ws = wb.create_sheet("担当割り当て")
    ws.append(["学年クラス"] + SUBJECTS)
    for cls in get_all_classes():
        row = [cls]
        for subj in SUBJECTS:
            teachers = st.session_state["assignments"].get(
                cls, {}).get(subj, [])
            row.append(",".join(teachers))
        ws.append(row)
    _apply_header_style(ws); _set_col_widths(ws, 10)
    ws.column_dimensions["A"].width = 14

    # シート6: 不在コマ
    ws = wb.create_sheet("不在コマ")
    ws.append(["教員名", "曜日", "時限"])
    for teacher, day_map in st.session_state["unavailable"].items():
        for day, periods in day_map.items():
            for p in periods:
                ws.append([teacher, day, p])
    _apply_header_style(ws); _set_col_widths(ws, 14)

    # シート7: 特別教室
    ws = wb.create_sheet("特別教室")
    ws.append(["教室名", "対象学年", "対応教科", "同時使用可能クラス数"])
    for room in st.session_state["special_rooms"]:
        gs = room.get("grade_subjects", [])
        if gs:
            for g, s in gs:
                ws.append([room["name"], g, s, room["capacity"]])
        else:
            ws.append([room["name"], "", "", room["capacity"]])
    _apply_header_style(ws); _set_col_widths(ws, 18)

    # シート8: 手動時間割
    ws = wb.create_sheet("手動時間割")
    ws.append(["学年クラス", "曜日", "時限", "教科"])
    for cls, day_map in st.session_state["manual_timetable"].items():
        for day, period_map in day_map.items():
            for period, subj in period_map.items():
                ws.append([cls, day, period, subj])
    _apply_header_style(ws); _set_col_widths(ws, 14)

    # シート9: 少人数学級設定
    ws = wb.create_sheet("少人数学級設定")
    ws.append(["少人数学級名", "所属クラス（カンマ区切り）",
               "担当教員（カンマ区切り）", "同期グループ番号",
               "グループ内クラス（カンマ区切り）"])
    for sg_name, sg_data in st.session_state["small_group_classes"].items():
        sync_groups = sg_data.get("sync_groups", [])
        if sync_groups:
            for gi, group in enumerate(sync_groups):
                if gi == 0:
                    ws.append([
                        sg_name,
                        ",".join(sg_data.get("classes", [])),
                        ",".join(sg_data.get("teachers", [])),
                        gi + 1,
                        ",".join(group)
                    ])
                else:
                    ws.append(["", "", "", gi + 1, ",".join(group)])
        else:
            ws.append([
                sg_name,
                ",".join(sg_data.get("classes", [])),
                ",".join(sg_data.get("teachers", [])),
                "", ""
            ])
    _apply_header_style(ws); _set_col_widths(ws, 22)

    # シート10: 少人数手動時間割
    ws = wb.create_sheet("少人数手動時間割")
    ws.append(["少人数学級名", "曜日", "時限", "教科"])
    for sg_name, day_map in st.session_state["manual_timetable_sg"].items():
        for day, period_map in day_map.items():
            for period, subj in period_map.items():
                ws.append([sg_name, day, period, subj])
    _apply_header_style(ws); _set_col_widths(ws, 16)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# Part 2: サイドバー UI
# ============================================================

def render_sidebar():
    """
    サイドバーにExcelボタン群・STEPナビゲーションを表示する。
    """
    st.sidebar.title("📅 小学校時間割生成")
    st.sidebar.markdown("---")

    # Excelテンプレートダウンロード
    st.sidebar.download_button(
        label="📥 テンプレートDL",
        data=generate_template_excel(),
        file_name="時間割テンプレート.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="入力用テンプレートをダウンロードします"
    )

    # 設定ファイルアップロード
    uploaded = st.sidebar.file_uploader(
        "📤 設定ファイルをアップロード",
        type=["xlsx"],
        help="保存済みの設定ファイルをアップロードします（全STEP上書き）"
    )
    if uploaded:
        # ファイルの同一性を名前+サイズで判定し、新規アップロード時のみ読み込む
        # （画面遷移のたびに再読み込みされて手動入力が消えるのを防ぐ）
        file_key = f"{uploaded.name}_{uploaded.size}"
        if st.session_state.get("_last_uploaded_excel_key") != file_key:
            load_settings_from_excel(uploaded)
            st.session_state["_last_uploaded_excel_key"] = file_key

    # 現在設定を保存
    st.sidebar.download_button(
        label="💾 現在の設定を保存",
        data=save_settings_to_excel(),
        file_name=f"時間割設定_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="現在の設定をExcelに保存します"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📋 設定STEP")

    # STEPナビゲーションボタン
    step_labels = {
        1: "① クラス数設定",
        2: "② 教員登録",
        3: "③ コマ数設定",
        4: "④ 週コマ数",
        5: "⑤ 担当教員割り当て",
        6: "⑥ 不在コマ設定",
        7: "⑦ 特別教室設定",
        8: "⑧ 手動時間割入力",
        9: "⑨ 少人数学級設定",
        10: "⑩ 学年間教科同期設定",
    }
    for step_num, label in step_labels.items():
        if st.sidebar.button(label, key=f"nav_{step_num}",
                             use_container_width=True):
            st.session_state["current_step"] = step_num
            st.rerun()

    st.sidebar.markdown("---")
    if st.sidebar.button("🚀 時間割を生成・確認",
                         type="primary", use_container_width=True):
        st.session_state["current_step"] = 0
        st.rerun()


# ============================================================
# Part 2: STEP 1〜5 UI
# ============================================================

def render_step1():
    st.header("STEP 1｜学年・クラス数設定")
    st.caption("各学年のクラス数を設定してください（1〜6クラス）")

    cols = st.columns(len(GRADES))
    for i, g in enumerate(GRADES):
        with cols[i]:
            st.session_state["grade_classes"][g] = st.number_input(
                f"{g}年生のクラス数",
                min_value=1, max_value=6,
                value=st.session_state["grade_classes"].get(g, 1),
                step=1,
                key=f"grade_{g}"
            )

    classes = get_all_classes()
    st.markdown("---")
    st.info(f"📋 設定クラス一覧（全 **{len(classes)}** クラス）\n\n"
            + "　".join(classes))


def render_step2():
    st.header("STEP 2｜教員登録")
    st.caption("時間割に関わる教員を全員登録してください")

    col1, col2 = st.columns([3, 1])
    with col1:
        new_name = st.text_input(
            "教員名を入力",
            placeholder="例: 古澤",
            key="new_teacher_input",
            label_visibility="collapsed"
        )
    with col2:
        if st.button("➕ 追加", use_container_width=True):
            name = new_name.strip()
            if name and name not in st.session_state["teachers"]:
                st.session_state["teachers"].append(name)
                st.rerun()
            elif name in st.session_state["teachers"]:
                st.warning("同じ名前の教員がすでに登録されています")

    st.markdown("---")
    st.markdown(f"#### 登録済み教員（{len(st.session_state['teachers'])}名）")

    if not st.session_state["teachers"]:
        st.info("教員が登録されていません")
    else:
        for i, teacher in enumerate(st.session_state["teachers"]):
            col1, col2 = st.columns([5, 1])
            col1.write(f"👤 {teacher}")
            if col2.button("🗑️ 削除", key=f"del_teacher_{i}",
                           use_container_width=True):
                st.session_state["teachers"].pop(i)
                st.rerun()


def render_step3():
    st.header("STEP 3｜曜日別1日コマ数設定（学年別）")
    st.caption("各学年・各曜日の1日のコマ数を設定してください（1〜8コマ）")

    ppd_grade = st.session_state["periods_per_day_grade"]
    # 学年がGRADESに合わせて初期化されているか確認
    for g in GRADES:
        ppd_grade.setdefault(g, {d: 6 for d in DAYS})

    header_cols = st.columns([2] + [1] * len(DAYS))
    with header_cols[0]:
        st.markdown("**学年**")
    for i, d in enumerate(DAYS):
        with header_cols[i + 1]:
            st.markdown(f"**{d}曜**")

    for g in GRADES:
        row_cols = st.columns([2] + [1] * len(DAYS))
        with row_cols[0]:
            st.markdown(f"**{g}年生**")
        for i, d in enumerate(DAYS):
            with row_cols[i + 1]:
                val = st.number_input(
                    f"{g}年{d}曜",
                    min_value=1, max_value=8,
                    value=ppd_grade[g].get(d, 6),
                    step=1,
                    key=f"ppd_{g}_{d}",
                    label_visibility="collapsed"
                )
                ppd_grade[g][d] = val

    # periods_per_day（最大値）を自動計算して更新
    for d in DAYS:
        st.session_state["periods_per_day"][d] = max(
            ppd_grade[g].get(d, 6) for g in GRADES
        )

    st.markdown("---")
    rows = []
    for g in GRADES:
        total_g = sum(ppd_grade[g].get(d, 6) for d in DAYS)
        rows.append(f"{g}年: 週{total_g}コマ")
    st.info("📊 学年別 週合計コマ数\n\n" + "　".join(rows))


def render_step4():
    st.header("STEP 4｜週コマ数設定")
    st.caption("各クラス・各教科の週あたりのコマ数を入力してください")

    classes = get_all_classes()
    total_per_week = get_total_periods_per_week()

    if not classes:
        st.warning("先にSTEP 1でクラス数を設定してください")
        return

    for cls in classes:
        with st.expander(f"📘 {cls}", expanded=False):
            if cls not in st.session_state["weekly_periods"]:
                st.session_state["weekly_periods"][cls] = {s: 0 for s in SUBJECTS}

            total = 0
            cols_per_row = 5
            for row_start in range(0, len(SUBJECTS), cols_per_row):
                row_subjects = SUBJECTS[row_start:row_start + cols_per_row]
                cols = st.columns(cols_per_row)
                for j, subj in enumerate(row_subjects):
                    with cols[j]:
                        val = st.number_input(
                            subj,
                            min_value=0,
                            max_value=total_per_week,
                            value=st.session_state["weekly_periods"][cls].get(subj, 0),
                            step=1,
                            key=f"wp_{cls}_{subj}"
                        )
                        st.session_state["weekly_periods"][cls][subj] = val
                        total += val

            if total > total_per_week:
                st.error(
                    f"⚠️ 合計 {total} コマ → 週上限 {total_per_week} コマを超過！")
            elif total == total_per_week:
                st.success(f"✅ 合計 {total} コマ（週上限ちょうど）")
            else:
                st.warning(
                    f"📝 合計 {total} コマ（週上限まで残り {total_per_week - total} コマ）")


def render_step5():
    st.header("STEP 5｜教科担当教員割り当て")
    st.caption("各クラス・各教科の担当教員を選択してください（複数選択可）")

    classes  = get_all_classes()
    teachers = st.session_state["teachers"]

    if not teachers:
        st.warning("先にSTEP 2で教員を登録してください")
        return
    if not classes:
        st.warning("先にSTEP 1でクラス数を設定してください")
        return

    for cls in classes:
        with st.expander(f"📘 {cls}", expanded=False):
            if cls not in st.session_state["assignments"]:
                st.session_state["assignments"][cls] = {s: [] for s in SUBJECTS}

            cols_per_row = 5
            for row_start in range(0, len(SUBJECTS), cols_per_row):
                row_subjects = SUBJECTS[row_start:row_start + cols_per_row]
                cols = st.columns(cols_per_row)
                for j, subj in enumerate(row_subjects):
                    with cols[j]:
                        current = st.session_state["assignments"][cls].get(subj, [])
                        valid_current = [t for t in current if t in teachers]
                        selected = st.multiselect(
                            subj,
                            options=teachers,
                            default=valid_current,
                            key=f"assign_{cls}_{subj}"
                        )
                        st.session_state["assignments"][cls][subj] = selected


# ============================================================
# Part 3: STEP 6〜9 UI
# ============================================================

def render_step6():
    st.header("STEP 6｜教員不在コマ設定")
    st.caption("各教員が授業に入れない時限をチェックしてください")

    teachers = st.session_state["teachers"]
    if not teachers:
        st.warning("先にSTEP 2で教員を登録してください")
        return

    for teacher in teachers:
        with st.expander(f"👤 {teacher}", expanded=False):
            if teacher not in st.session_state["unavailable"]:
                st.session_state["unavailable"][teacher] = {}

            for day in DAYS:
                periods = st.session_state["periods_per_day"].get(day, 6)
                current_unavail = st.session_state["unavailable"][teacher].get(day, [])

                st.markdown(f"**{day}曜日**")
                cols = st.columns(periods)
                selected = []
                for p in range(1, periods + 1):
                    with cols[p - 1]:
                        checked = st.checkbox(
                            f"{p}限",
                            value=(p in current_unavail),
                            key=f"unavail_{teacher}_{day}_{p}"
                        )
                        if checked:
                            selected.append(p)
                st.session_state["unavailable"][teacher][day] = selected

            total_unavail = sum(
                len(v) for v in st.session_state["unavailable"][teacher].values()
            )
            if total_unavail > 0:
                st.caption(f"合計不在コマ数: {total_unavail} コマ/週")


def render_step7():
    st.header("STEP 7｜特別教室設定")
    st.caption(
        "特別教室を登録し、使用する学年×教科の組み合わせと同時使用可能クラス数を設定してください。\n\n"
        "例：「第一体育館」に 1年体育・2年体育・3年体育 を登録 → "
        "同じ時限にこれらのクラスが合計で設定した上限数まで使用できます。"
    )

    if st.button("➕ 特別教室を追加", key="add_room"):
        st.session_state["special_rooms"].append(
            {"name": "", "capacity": 1, "grade_subjects": []}
        )
        st.rerun()

    if not st.session_state["special_rooms"]:
        st.info("特別教室が登録されていません。上のボタンで追加してください。")
        return

    for i, room in enumerate(st.session_state["special_rooms"]):
        # 旧フォーマット（grade/subject）を新フォーマットに自動変換
        if "grade" in room and "subject" in room and "grade_subjects" not in room:
            room["grade_subjects"] = [(room.pop("grade"), room.pop("subject"))]

        with st.expander(f"🏫 {room.get('name') or f'教室{i+1}'}", expanded=True):
            col1, col2, col3 = st.columns([4, 2, 1])
            with col1:
                room["name"] = st.text_input(
                    "教室名",
                    value=room.get("name", ""),
                    placeholder="例: 第一体育館",
                    key=f"room_name_{i}"
                )
            with col2:
                room["capacity"] = st.number_input(
                    "同時使用可能クラス数",
                    min_value=1, max_value=10,
                    value=room.get("capacity", 1),
                    step=1,
                    key=f"room_cap_{i}"
                )
            with col3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️ 削除", key=f"del_room_{i}", use_container_width=True):
                    st.session_state["special_rooms"].pop(i)
                    st.rerun()

            # 学年×教科ペアの追加
            st.markdown("**使用する学年×教科**")
            col_g, col_s, col_add = st.columns([2, 3, 2])
            with col_g:
                add_grade = st.selectbox(
                    "学年", GRADES, key=f"room_add_g_{i}",
                    format_func=lambda g: f"{g}年"
                )
            with col_s:
                add_subj = st.selectbox(
                    "教科", SUBJECTS, key=f"room_add_s_{i}"
                )
            with col_add:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ 追加", key=f"room_add_btn_{i}", use_container_width=True):
                    gs = room.setdefault("grade_subjects", [])
                    pair = [add_grade, add_subj]
                    if pair not in gs:
                        gs.append(pair)
                        st.rerun()

            # 登録済みペア一覧
            gs_list = room.get("grade_subjects", [])
            if gs_list:
                for j, (g, s) in enumerate(gs_list):
                    c1, c2 = st.columns([6, 1])
                    with c1:
                        st.markdown(f"・{g}年　{s}")
                    with c2:
                        if st.button("✕", key=f"room_del_gs_{i}_{j}"):
                            gs_list.pop(j)
                            st.rerun()
            else:
                st.info("学年×教科の組み合わせを追加してください")


def render_step8():
    st.header("STEP 8｜手動時間割入力")
    st.caption(
        "自動生成前に手動で配置したいコマを入力してください。"
        "手動入力したコマは自動生成で上書きされません（水色で表示）。"
    )

    tab_normal, tab_sg = st.tabs(["📘 通常学級", "📗 少人数学級"])

    # ── 通常学級タブ ─────────────────────────────────────────
    with tab_normal:
        classes = get_all_classes()
        if not classes:
            st.warning("先にSTEP 1でクラス数を設定してください")
        else:
            cls = st.selectbox("クラスを選択", classes, key="manual_cls_select")

            if cls not in st.session_state["manual_timetable"]:
                st.session_state["manual_timetable"][cls] = {}

            options = ["（未入力）"] + SUBJECTS

            for day in DAYS:
                periods = st.session_state["periods_per_day"].get(day, 6)
                st.markdown(f"**{day}曜日**")
                cols = st.columns(periods)
                day_data = st.session_state["manual_timetable"][cls].setdefault(day, {})

                for p in range(1, periods + 1):
                    with cols[p - 1]:
                        current_val = day_data.get(p, "（未入力）")
                        idx = options.index(current_val) \
                            if current_val in options else 0
                        chosen = st.selectbox(
                            f"{p}限",
                            options=options,
                            index=idx,
                            key=f"manual_{cls}_{day}_{p}"
                        )
                        if chosen == "（未入力）":
                            day_data.pop(p, None)
                        else:
                            day_data[p] = chosen

            total_manual = sum(
                len(v) for v in
                st.session_state["manual_timetable"].get(cls, {}).values()
            )
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            col1.info(f"📌 {cls} の手動入力コマ数: {total_manual} コマ")
            if col2.button(f"🗑️ {cls} をすべてクリア",
                           key=f"clear_manual_{cls}"):
                st.session_state["manual_timetable"][cls] = {}
                st.rerun()

    # ── 少人数学級タブ ───────────────────────────────────────
    with tab_sg:
        sg_list = list(st.session_state["small_group_classes"].keys())
        if not sg_list:
            st.info("先にSTEP 9で少人数学級を登録してください")
        else:
            sg_name = st.selectbox(
                "少人数学級を選択", sg_list, key="manual_sg_select"
            )

            if sg_name not in st.session_state["manual_timetable_sg"]:
                st.session_state["manual_timetable_sg"][sg_name] = {}

            options = ["（未入力）"] + SUBJECTS

            for day in DAYS:
                periods = st.session_state["periods_per_day"].get(day, 6)
                st.markdown(f"**{day}曜日**")
                cols = st.columns(periods)
                day_data = st.session_state["manual_timetable_sg"][sg_name].setdefault(day, {})

                for p in range(1, periods + 1):
                    with cols[p - 1]:
                        current_val = day_data.get(p, "（未入力）")
                        idx = options.index(current_val) \
                            if current_val in options else 0
                        chosen = st.selectbox(
                            f"{p}限",
                            options=options,
                            index=idx,
                            key=f"manual_sg_{sg_name}_{day}_{p}"
                        )
                        if chosen == "（未入力）":
                            day_data.pop(p, None)
                        else:
                            day_data[p] = chosen

            total_manual_sg = sum(
                len(v) for v in
                st.session_state["manual_timetable_sg"].get(sg_name, {}).values()
            )
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            col1.info(f"📌 {sg_name} の手動入力コマ数: {total_manual_sg} コマ")
            if col2.button(f"🗑️ {sg_name} をすべてクリア",
                           key=f"clear_manual_sg_{sg_name}"):
                st.session_state["manual_timetable_sg"][sg_name] = {}
                st.rerun()


def render_step9():
    st.header("STEP 9｜少人数学級設定")
    st.caption(
        "少人数学級を登録し、所属クラス（原籍学級）・担当教員・"
        "同期グループ（R6制約）を設定してください"
    )

    classes  = get_all_classes()
    teachers = st.session_state["teachers"]

    col1, col2 = st.columns([3, 1])
    with col1:
        new_sg = st.text_input(
            "少人数学級名を入力",
            placeholder="例: 少人数A、少人数理科1",
            key="new_sg_name_input",
            label_visibility="collapsed"
        )
    with col2:
        if st.button("➕ 追加", key="add_sg", use_container_width=True):
            name = new_sg.strip()
            if name and name not in st.session_state["small_group_classes"]:
                st.session_state["small_group_classes"][name] = {
                    "classes":        [],
                    "teachers":       [],
                    "joint_subjects": [],
                    "weekly_periods": {},
                    "sync_groups":    []
                }
                st.rerun()
            elif name in st.session_state["small_group_classes"]:
                st.warning("同じ名前の少人数学級がすでに存在します")

    if not st.session_state["small_group_classes"]:
        st.info("少人数学級が登録されていません")
        return

    for sg_name in list(st.session_state["small_group_classes"].keys()):
        sg_data = st.session_state["small_group_classes"][sg_name]
        with st.expander(f"📗 {sg_name}", expanded=True):

            # ── 所属クラス設定 ───────────────────────────────
            # セッションステートキーを初期化（初回のみ）
            key_cls = f"sg_cls_{sg_name}"
            if key_cls not in st.session_state:
                st.session_state[key_cls] = [
                    c for c in sg_data.get("classes", []) if c in classes
                ]
            st.multiselect(
                "所属クラス（原籍学級）",
                options=classes,
                key=key_cls,
                help="この少人数学級に生徒が所属する原籍学級を全て選択"
            )
            st.session_state["small_group_classes"][sg_name]["classes"] = \
                st.session_state[key_cls]

            # ── 担当教員設定 ─────────────────────────────────
            key_teacher = f"sg_teacher_{sg_name}"
            if key_teacher not in st.session_state:
                st.session_state[key_teacher] = [
                    t for t in sg_data.get("teachers", []) if t in teachers
                ]
            st.multiselect(
                "担当教員",
                options=teachers,
                key=key_teacher,
                help="R1（教員重複禁止）の対象となります"
            )
            st.session_state["small_group_classes"][sg_name]["teachers"] = \
                st.session_state[key_teacher]

            # ── 合同教科選択 ─────────────────────────────────
            key_js = f"sg_js_{sg_name}"
            if key_js not in st.session_state:
                st.session_state[key_js] = sg_data.get("joint_subjects", [])
            st.multiselect(
                "合同教科（すべての教科から選択）",
                options=SUBJECTS,
                key=key_js,
                help="このグループで合同授業を行う教科を選択してください"
            )
            st.session_state["small_group_classes"][sg_name]["joint_subjects"] = \
                st.session_state[key_js]
            joint_subjs = st.session_state[key_js]

            # ── 週コマ数（選択した合同教科のみ） ─────────────
            st.markdown("**週コマ数（合同教科）**")
            if "weekly_periods" not in sg_data:
                st.session_state["small_group_classes"][sg_name]["weekly_periods"] = {}
            if joint_subjs:
                cols = st.columns(min(len(joint_subjs), 6))
                for ci, subj in enumerate(joint_subjs):
                    with cols[ci % 6]:
                        key_wp = f"sg_wp_{sg_name}_{subj}"
                        if key_wp not in st.session_state:
                            st.session_state[key_wp] = \
                                sg_data.get("weekly_periods", {}).get(subj, 0)
                        st.number_input(
                            subj,
                            min_value=0, max_value=10,
                            step=1,
                            key=key_wp
                        )
                        st.session_state["small_group_classes"][sg_name][
                            "weekly_periods"][subj] = st.session_state[key_wp]
            else:
                st.info("合同教科を選択してください")

            # ── 同期グループ設定（R6） ────────────────────────
            st.markdown("---")
            st.markdown("**同期グループ設定（R6: 少人数学級同期制約）**")
            st.caption(
                "同じグループ内のクラスは、同一時限に"
                "「全員が共通教科」または「全員が非共通教科」になるよう制約されます。"
                "グループ未設定の場合は所属クラス全員に一括適用されます。"
            )

            current_classes = st.session_state["small_group_classes"][sg_name]["classes"]

            if st.button("➕ グループを追加", key=f"add_group_{sg_name}"):
                st.session_state["small_group_classes"][sg_name]["sync_groups"].append([])
                st.rerun()

            current_groups = st.session_state["small_group_classes"][sg_name]["sync_groups"]

            if not current_groups:
                st.info(
                    "💡 グループ未設定: 所属クラス全員（"
                    + ", ".join(current_classes) + "）を一括で同期します"
                )
            else:
                for gi in range(len(current_groups)):
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        valid_group = [c for c in current_groups[gi]
                                       if c in current_classes]
                        key_grp = f"sg_group_{sg_name}_{gi}"
                        if key_grp not in st.session_state:
                            st.session_state[key_grp] = valid_group
                        st.multiselect(
                            f"グループ {gi + 1}（2クラス以上選択）",
                            options=current_classes,
                            key=key_grp,
                            help="このグループ内で共通教科の時間帯を揃えます"
                        )
                        st.session_state["small_group_classes"][sg_name][
                            "sync_groups"][gi] = st.session_state[key_grp]
                    with col2:
                        st.markdown("　")
                        if st.button("🗑️", key=f"del_group_{sg_name}_{gi}",
                                     use_container_width=True):
                            st.session_state["small_group_classes"][sg_name][
                                "sync_groups"].pop(gi)
                            # グループキーをクリアして再初期化させる
                            for k in list(st.session_state.keys()):
                                if k.startswith(f"sg_group_{sg_name}_"):
                                    del st.session_state[k]
                            st.rerun()

                for gi, group in enumerate(current_groups):
                    if len(group) < 2:
                        st.warning(
                            f"⚠️ グループ {gi + 1}: "
                            "同期グループは2クラス以上必要です"
                        )

            # ── 少人数学級の削除 ─────────────────────────────
            st.markdown("---")
            if st.button(
                f"🗑️ 「{sg_name}」を削除",
                key=f"del_sg_{sg_name}",
                type="secondary"
            ):
                del st.session_state["small_group_classes"][sg_name]
                st.session_state["manual_timetable_sg"].pop(sg_name, None)
                # 関連するウィジェットキーをクリア
                for k in list(st.session_state.keys()):
                    if k.startswith(f"sg_cls_{sg_name}") or \
                       k.startswith(f"sg_teacher_{sg_name}") or \
                       k.startswith(f"sg_wp_{sg_name}_") or \
                       k.startswith(f"sg_group_{sg_name}_"):
                        del st.session_state[k]
                st.rerun()


# ============================================================
# STEP 10: 学年間教科同期設定
# ============================================================

def render_step10():
    st.header("STEP 10｜学年間教科同期設定")
    st.caption(
        "異なる学年の教科を同じ曜日・同じ時限に揃えるハード制約を設定します。\n\n"
        "例：「5年 家庭」と「6年 音楽」を同期 → 5年の各クラスが家庭の時限に、"
        "6年のいずれかのクラスが音楽を担当する時限が一致します。"
    )

    pairs = st.session_state["cross_grade_sync"]

    # ── 新規ペア追加 ───────────────────────────────────────
    st.markdown("#### 同期ペアを追加")
    col1, col2, col3, col4, col5 = st.columns([2, 3, 2, 3, 2])
    with col1:
        new_g1 = st.selectbox("学年①", GRADES, key="cgs_g1")
    with col2:
        new_s1 = st.selectbox("教科①", SUBJECTS, key="cgs_s1")
    with col3:
        new_g2 = st.selectbox("学年②", GRADES, key="cgs_g2")
    with col4:
        new_s2 = st.selectbox("教科②", SUBJECTS, key="cgs_s2")
    with col5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ 追加", key="cgs_add", use_container_width=True):
            if new_g1 == new_g2 and new_s1 == new_s2:
                st.warning("同じ学年・同じ教科のペアは追加できません")
            else:
                new_pair = {"grade1": new_g1, "subject1": new_s1,
                            "grade2": new_g2, "subject2": new_s2}
                # 重複チェック（順序問わず）
                reverse = {"grade1": new_g2, "subject1": new_s2,
                           "grade2": new_g1, "subject2": new_s1}
                if new_pair in pairs or reverse in pairs:
                    st.warning("同じ組み合わせがすでに登録されています")
                else:
                    pairs.append(new_pair)
                    st.rerun()

    st.markdown("---")

    # ── 登録済みペア一覧 ───────────────────────────────────
    if not pairs:
        st.info("同期ペアが登録されていません")
        return

    st.markdown("#### 登録済み同期ペア")
    for i, pair in enumerate(pairs):
        col_a, col_b = st.columns([6, 1])
        with col_a:
            st.markdown(
                f"**{i+1}.** {pair['grade1']}年「{pair['subject1']}」　↔　"
                f"{pair['grade2']}年「{pair['subject2']}」"
            )
        with col_b:
            if st.button("🗑️", key=f"cgs_del_{i}", help="削除"):
                pairs.pop(i)
                st.rerun()


# ============================================================
# Part 3: ソフト制約設定UI
# ============================================================

def render_soft_constraints():
    """
    ソフト制約（S1学年集約・S2優先教科）の設定UIを表示する。
    render_generate()内の生成ボタン上部に配置する。
    """
    st.markdown("### 🎛️ ソフト制約設定")

    st.session_state["soft_grade_grouping"] = st.checkbox(
        "S1: 学年集約を有効にする",
        value=st.session_state.get("soft_grade_grouping", True),
        help=(
            "同一教員が同じ曜日に複数クラスを担当する場合、"
            "できるだけ同じ学年の授業が連続するように調整します"
        )
    )

    if st.session_state["soft_grade_grouping"]:
        st.session_state["soft_priority_subjects"] = st.multiselect(
            "S2: 優先教科（学年集約を特に強く適用する教科）",
            options=SUBJECTS,
            default=st.session_state.get("soft_priority_subjects", []),
            help=(
                "選択した教科は学年集約のペナルティが10倍になります。"
                "例: 理科・音楽・保体など準備が必要な教科を選択してください"
            )
        )

    st.markdown("---")
    st.markdown("#### 📚 教科別段階生成")
    st.session_state["soft_first_subjects"] = st.multiselect(
        "先に配置する教科（Phase 1）",
        options=SUBJECTS,
        default=st.session_state.get("soft_first_subjects", []),
        help=(
            "選択した教科を Phase 1 として先に全クラスに配置し、"
            "その後に残りの教科（Phase 2）を配置します。\n"
            "「📚 教科別段階生成」ボタンで実行してください。"
        )
    )


# ============================================================
# Part 3: バリデーション関数
# ============================================================

def validate_all_settings() -> list[str]:
    """
    生成前に全設定の整合性チェックを行い、警告メッセージのリストを返す。
    リストが空なら問題なし。
    """
    warnings_list = []
    classes        = get_all_classes()
    total_per_week = get_total_periods_per_week()

    # ① 教員未登録チェック
    if not st.session_state["teachers"]:
        warnings_list.append("教員が1人も登録されていません（STEP 2）")

    # ② 週コマ数の合計チェック
    for cls in classes:
        total = sum(st.session_state["weekly_periods"].get(cls, {}).values())
        if total != total_per_week:
            warnings_list.append(
                f"{cls}: 週コマ数の合計が {total} コマ "
                f"（週上限 {total_per_week} コマと不一致）"
            )

    # ③ 担当教員未設定チェック（週コマ数>0なのに担当なし）
    for cls in classes:
        for subj in SUBJECTS:
            weekly = st.session_state["weekly_periods"].get(cls, {}).get(subj, 0)
            if weekly > 0:
                teachers = st.session_state["assignments"].get(
                    cls, {}).get(subj, [])
                if not teachers:
                    warnings_list.append(
                        f"{cls} / {subj}: 週{weekly}コマ設定だが担当教員が未設定"
                    )

    # ④ 特別教室の教科名整合性チェック
    for room in st.session_state["special_rooms"]:
        if room.get("subject") not in SUBJECTS:
            warnings_list.append(
                f"特別教室「{room.get('name', '?')}」の対応教科が不正"
            )

    # ⑤ 少人数学級の同期グループ整合性チェック
    for sg_name, sg_data in st.session_state["small_group_classes"].items():
        for gi, group in enumerate(sg_data.get("sync_groups", [])):
            if len(group) < 2:
                warnings_list.append(
                    f"少人数学級「{sg_name}」のグループ{gi+1}: "
                    "2クラス以上必要です"
                )
            for cls in group:
                if cls not in sg_data.get("classes", []):
                    warnings_list.append(
                        f"少人数学級「{sg_name}」のグループ{gi+1}: "
                        f"「{cls}」が所属クラスに含まれていません"
                    )

    # ⑥ R7事前チェック（教員の週総授業コマ数が上限内か）
    for teacher in st.session_state["teachers"]:
        weekly_load = sum(
            st.session_state["weekly_periods"].get(cls, {}).get(subj, 0)
            for cls in classes
            for subj, tlist in st.session_state["assignments"].get(cls, {}).items()
            if teacher in tlist
        )
        if weekly_load == 0:
            continue
        max_weekly = sum(
            max(0, st.session_state["periods_per_day"].get(day, 6)
                - len(st.session_state["unavailable"].get(teacher, {}).get(day, []))
                - 1)
            for day in DAYS
        )
        if weekly_load > max_weekly:
            warnings_list.append(
                f"教員「{teacher}」: 週{weekly_load}コマ必要ですが"
                f"R7制約後の上限は{max_weekly}コマです（担当クラスを減らすか不在設定を見直してください）"
            )

    # ⑦ 手動コマのR1・R5チェック
    manual_tt = st.session_state.get("manual_timetable", {})
    if manual_tt:
        for day in DAYS:
            total_p = st.session_state["periods_per_day"].get(day, 6)
            for p in range(1, total_p + 1):
                seen_t: dict = {}
                for cls in classes:
                    subj = manual_tt.get(cls, {}).get(day, {}).get(p)
                    if not subj:
                        continue
                    # R5: 同日同教科
                    day_map = manual_tt.get(cls, {}).get(day, {})
                    same = [s for pk, s in day_map.items() if pk != p and s == subj]
                    if same:
                        warnings_list.append(
                            f"手動コマR5違反: {cls} {day}曜に「{subj}」が複数コマ設定されています")
                    # R1: 教員重複
                    for t in st.session_state["assignments"].get(cls, {}).get(subj, []):
                        if t in seen_t:
                            warnings_list.append(
                                f"手動コマR1違反: {t}先生が{day}曜{p}限に"
                                f"{seen_t[t]}と{cls}の両方に設定されています")
                        seen_t[t] = cls
                        # ⑧ 手動コマと不在コマの競合チェック（R4）
                        if p in st.session_state["unavailable"].get(t, {}).get(day, []):
                            warnings_list.append(
                                f"手動コマと不在コマが競合: {t}先生は{day}曜{p}限が不在設定ですが、"
                                f"{cls}の「{subj}」が手動配置されています（時間割を生成できません）"
                            )

    return warnings_list


# ============================================================
# Part 4: 制約チェック関数
# ============================================================

def is_valid_placement(
    timetable: dict,
    timetable_sg: dict,
    cls: str,
    day: str,
    period: int,
    subject: str,
    assignments: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    small_group_classes: dict,
    weekly_remaining: dict
) -> bool:
    """R1〜R8を全チェックする。すべてパスしたらTrue。"""

    # ── 学年別コマ数チェック ─────────────────────────────────
    if period > get_periods_for_class(cls, day):
        return False

    teachers = assignments.get(cls, {}).get(subject, [])

    # ── R1: 教員重複禁止 ─────────────────────────────────────
    for t in teachers:
        for other_cls, schedule in timetable.items():
            if other_cls == cls:
                continue
            other_subj = schedule.get(day, {}).get(period)
            if other_subj:
                if t in assignments.get(other_cls, {}).get(other_subj, []):
                    return False
        for sg_name, sg_schedule in timetable_sg.items():
            sg_subj = sg_schedule.get(day, {}).get(period)
            if sg_subj:
                sg_teachers = small_group_classes.get(sg_name, {}).get("teachers", [])
                if t in sg_teachers:
                    return False

    # ── R3: 特別教室同時使用制限 ─────────────────────────────
    cls_grade = get_grade(cls)
    for room in special_rooms:
        gs = room.get("grade_subjects", [])
        if [cls_grade, subject] not in gs and (cls_grade, subject) not in gs:
            continue
        count = sum(
            1 for other_cls, schedule in timetable.items()
            if other_cls != cls
            and ([get_grade(other_cls), schedule.get(day, {}).get(period)] in gs
                 or (get_grade(other_cls), schedule.get(day, {}).get(period)) in gs)
        )
        if count >= room["capacity"]:
            return False

    # ── R4: 教員不在コマ尊重 ─────────────────────────────────
    for t in teachers:
        if period in unavailable.get(t, {}).get(day, []):
            return False

    # ── R5: 同一クラス・同一曜日に同一教科は1コマまで ───────
    for p, s in timetable.get(cls, {}).get(day, {}).items():
        if p != period and s == subject:
            return False

    # ── R6: 少人数学級同期制約 ───────────────────────────────
    for sg_name, sg_data in small_group_classes.items():
        groups = sg_data.get("sync_groups") or [sg_data.get("classes", [])]
        for group in groups:
            if cls not in group:
                continue
            other_subjects = [
                timetable.get(oc, {}).get(day, {}).get(period)
                for oc in group if oc != cls
            ]
            other_subjects = [s for s in other_subjects if s]
            if not other_subjects:
                continue
            joint_subjs = sg_data.get("joint_subjects", [])
            this_is_common = subject in joint_subjs
            for os_ in other_subjects:
                if (os_ in joint_subjs) != this_is_common:
                    return False

    # ── R8: 学年間教科同期制約（一方向・少ない方が基準） ────────
    # コマ数の少ない側のすべてのコマは、多い側のクラスと必ず同期する。
    # 多い側の余りコマは、少ない側が存在しない時限に自由配置。
    cross_grade_sync = st.session_state.get("cross_grade_sync", [])
    cls_grade = get_grade(cls)
    all_classes = get_all_classes()
    for pair in cross_grade_sync:
        g1, s1, g2, s2 = pair["grade1"], pair["subject1"], pair["grade2"], pair["subject2"]
        grade1_classes = [c for c in all_classes if get_grade(c) == g1]
        grade2_classes = [c for c in all_classes if get_grade(c) == g2]

        # 残コマ数で少ない方（minority）を動的に判定
        total_g1 = sum(weekly_remaining.get(c, {}).get(s1, 0) for c in grade1_classes)
        total_g2 = sum(weekly_remaining.get(c, {}).get(s2, 0) for c in grade2_classes)

        if total_g1 <= total_g2:
            # g1(s1)がminority: s1を配置するなら同時限にg2のs2が必要
            minority_grade, minority_subj = g1, s1
            majority_grade, majority_subj = g2, s2
            minority_classes, majority_classes = grade1_classes, grade2_classes
        else:
            # g2(s2)がminority: s2を配置するなら同時限にg1のs1が必要
            minority_grade, minority_subj = g2, s2
            majority_grade, majority_subj = g1, s1
            minority_classes, majority_classes = grade2_classes, grade1_classes

        # minority側を配置するとき: majority側が全て別教科で埋まっていたらNG
        if cls_grade == minority_grade and subject == minority_subj:
            filled_maj = [
                timetable.get(c, {}).get(day, {}).get(period)
                for c in majority_classes
            ]
            filled_maj_vals = [s for s in filled_maj if s is not None]
            if len(filled_maj_vals) == len(majority_classes) and \
                    not any(s == majority_subj for s in filled_maj_vals):
                return False

        # majority側がminority_subjを持っているとき:
        # minority側を別教科で埋めてしまい、minority側が全て埋まったらNG
        if cls_grade == minority_grade and subject != minority_subj:
            has_maj_subj = any(
                timetable.get(c, {}).get(day, {}).get(period) == majority_subj
                for c in majority_classes
            )
            if has_maj_subj:
                other_min = [c for c in minority_classes if c != cls]
                filled_other = [
                    timetable.get(c, {}).get(day, {}).get(period)
                    for c in other_min
                ]
                if all(s is not None for s in filled_other) and \
                        not any(s == minority_subj for s in filled_other):
                    return False

    # ── R7: 各教員の当日最低1コマ空き ───────────────────────
    total_periods_today = periods_per_day.get(day, 6)
    for t in teachers:
        assigned_today = sum(
            1 for oc, sc in timetable.items()
            for p_ in range(1, total_periods_today + 1)
            if sc.get(day, {}).get(p_) and
               t in assignments.get(oc, {}).get(sc[day].get(p_, ""), [])
        )
        assigned_today += sum(
            1 for sg_name, sg_sc in timetable_sg.items()
            for p_ in range(1, total_periods_today + 1)
            if sg_sc.get(day, {}).get(p_) and
               t in small_group_classes.get(sg_name, {}).get("teachers", [])
        )
        unavail_today = len(unavailable.get(t, {}).get(day, []))
        max_assignable = total_periods_today - unavail_today - 1
        if assigned_today >= max_assignable:
            return False

    return True


# ============================================================
# Part 4: MRV タイブレークスコア
# ============================================================

def tiebreak_score(
    cls: str,
    day: str,
    period: int,
    timetable: dict,
    timetable_sg: dict,
    small_group_classes: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    weekly_remaining: dict,
    priority_subjects: list,
    assignments: dict,
    teacher_day_cnt: dict = None
) -> float:
    """8段階のタイブレークスコアを計算して返す。高いほど優先。"""
    score = 0.0

    # 同期グループ判定
    in_sync = False
    sync_group_size = 0
    for sg_data in small_group_classes.values():
        groups = sg_data.get("sync_groups") or [sg_data.get("classes", [])]
        for group in groups:
            if cls in group:
                in_sync = True
                sync_group_size = max(sync_group_size, len(group))

    # 特別教室が必要か判定
    needs_room = False
    min_room_capacity = 999
    remaining = weekly_remaining.get(cls, {})
    cls_grade = get_grade(cls)
    for room in special_rooms:
        gs = room.get("grade_subjects", [])
        room_subjects_for_cls = {s for g, s in gs if g == cls_grade}
        if any(s in room_subjects_for_cls and rem > 0 for s, rem in remaining.items()):
            needs_room = True
            min_room_capacity = min(min_room_capacity, room["capacity"])

    # 第1位: R6×R3複合
    if in_sync and needs_room:
        score += 10000.0 + sync_group_size * 100.0
        if min_room_capacity < 999:
            score += 1000.0 / min_room_capacity
    # 第2位: R6のみ
    elif in_sync:
        score += 5000.0 + sync_group_size * 100.0
    # 第3位: R3のみ
    elif needs_room:
        score += 3000.0
        if min_room_capacity < 999:
            score += 1000.0 / min_room_capacity

    # 第4位: R7リスク（当日残り空き=1）
    # teacher_day_cntキャッシュがあればO(1)、なければ従来のO(n)フォールバック
    teachers_of_cls = list(set(
        t for s, rem in remaining.items() if rem > 0
        for t in assignments.get(cls, {}).get(s, [])
    ))
    total_today = periods_per_day.get(day, 6)
    for t in teachers_of_cls:
        if teacher_day_cnt is not None:
            assigned_today = teacher_day_cnt.get(t, {}).get(day, 0)
        else:
            assigned_today = sum(
                1 for oc, sc in timetable.items()
                for p_ in range(1, total_today + 1)
                if sc.get(day, {}).get(p_) and
                   t in assignments.get(oc, {}).get(sc[day].get(p_, ""), [])
            )
        unavail_today = len(unavailable.get(t, {}).get(day, []))
        if total_today - unavail_today - assigned_today - 1 <= 1:
            score += 2000.0
            break

    # 第5位: 不在コマが多い教員
    max_unavail = max(
        (sum(len(v) for v in unavailable.get(t, {}).values())
         for t in teachers_of_cls), default=0
    )
    score += max_unavail * 10.0

    # 第6位: 週残りコマ数の最大値
    score += max(remaining.values(), default=0) * 5.0

    # 第7位: 優先教科あり
    if any(s in priority_subjects for s, r in remaining.items() if r > 0):
        score += 2.0

    # 第8位: クラス番号（再現性）
    score -= get_class_index(cls) * 0.001

    return score


# ============================================================
# Part 4: ソフト制約スコア
# ============================================================

def soft_score(
    subject: str,
    cls: str,
    day: str,
    timetable: dict,
    assignments: dict,
    priority_subjects: list,
    period: int = 0
) -> float:
    """ソフト制約スコア。高いほど優先して配置する。
    S1: 学年集約（同学年の授業を同じ教員が連続して担当）
    S2: R6ソフト（同期グループ内で共通教科の種別が揃う）
    """
    score = 0.0
    small_group_classes = st.session_state.get("small_group_classes", {})

    # ── S1: 学年集約スコア ────────────────────────────────────
    if st.session_state.get("soft_grade_grouping", True):
        PRIORITY_WEIGHT = 10.0
        NORMAL_WEIGHT   = 1.0
        weight = PRIORITY_WEIGHT if subject in priority_subjects else NORMAL_WEIGHT
        this_grade = get_grade(cls)
        teachers = assignments.get(cls, {}).get(subject, [])
        for t in teachers:
            for other_cls, schedule in timetable.items():
                if other_cls == cls:
                    continue
                for p_, s_ in schedule.get(day, {}).items():
                    if not s_:
                        continue
                    if t not in assignments.get(other_cls, {}).get(s_, []):
                        continue
                    score += weight if get_grade(other_cls) == this_grade else -weight

        # S1b: 同学年・同教科・同曜日集約スコア
        # 同じ学年の他クラスがこの教科をすでに配置済みの場合、
        # 同じ曜日なら加点、異なる曜日なら減点
        DAY_ALIGN = 3.0
        w_day = DAY_ALIGN * weight
        for other_cls, schedule in timetable.items():
            if other_cls == cls or get_grade(other_cls) != this_grade:
                continue
            same_day = subject in schedule.get(day, {}).values()
            diff_day = any(
                subject in periods.values()
                for d_, periods in schedule.items()
                if d_ != day
            )
            if same_day:
                score += w_day
            elif diff_day:
                score -= w_day

        # S3: 学年交互防止（1〜4限で同一教員が学年を行き来しない）
        # 配置しようとしている時限(period)が1〜4限の場合のみ評価
        # 優先教科は weight が 10 倍になるため、自動的に強く効く
        GRADE_FLOW = 5.0
        w_s3 = GRADE_FLOW * weight   # 優先教科なら 5×10=50、通常なら 5×1=5
        S3_TARGET = {1, 2, 3, 4}
        if period in S3_TARGET:
            teachers_here = assignments.get(cls, {}).get(subject, [])
            for t in teachers_here:
                for adj_p in (period - 1, period + 1):
                    if adj_p not in S3_TARGET:
                        continue
                    # 隣接時限でTが担当しているクラスを探す
                    for other_cls, schedule in timetable.items():
                        if other_cls == cls:
                            continue
                        adj_subj = schedule.get(day, {}).get(adj_p)
                        if not adj_subj:
                            continue
                        if t not in assignments.get(other_cls, {}).get(adj_subj, []):
                            continue
                        # 同学年なら加点、異学年なら減点（交互防止）
                        if get_grade(other_cls) == this_grade:
                            score += w_s3
                        else:
                            score -= w_s3

    # ── S2: R6ソフト制約スコア ───────────────────────────────
    # 同期グループ内で既に配置済みの他クラスと「共通教科かどうか」が
    # 一致すれば加点、不一致なら減点
    if period > 0 and small_group_classes:
        R6_MATCH    = 500.0
        R6_MISMATCH = 500.0
        for sg_data in small_group_classes.values():
            joint_subjs = sg_data.get("joint_subjects", [])
            this_common = subject in joint_subjs
            groups = sg_data.get("sync_groups") or [sg_data.get("classes", [])]
            for group in groups:
                if cls not in group:
                    continue
                for oc in group:
                    if oc == cls:
                        continue
                    other_subj = timetable.get(oc, {}).get(day, {}).get(period)
                    if not other_subj:
                        continue
                    other_common = other_subj in joint_subjs
                    if this_common == other_common:
                        score += R6_MATCH
                    else:
                        score -= R6_MISMATCH

    return score


# ============================================================
# Part 4: 前処理関数
# ============================================================

def preprocess(
    timetable: dict,
    timetable_sg: dict,
    manual_timetable: dict,
    manual_timetable_sg: dict,
    weekly_periods: dict,
    weekly_periods_sg: dict
) -> tuple[dict, dict]:
    """
    手動入力コマをロックし、残り週コマ数を計算して返す。
    戻り値: (weekly_remaining, weekly_remaining_sg)
    """
    # 通常学級の残り週コマ数を初期化
    weekly_remaining = {
        cls: dict(weekly_periods.get(cls, {}))
        for cls in timetable
    }

    # 手動入力コマをロック
    for cls, day_map in manual_timetable.items():
        if cls not in timetable:
            continue
        for day, period_map in day_map.items():
            for period, subj in period_map.items():
                timetable[cls].setdefault(day, {})[period] = subj
                if subj in weekly_remaining.get(cls, {}):
                    weekly_remaining[cls][subj] = max(
                        0, weekly_remaining[cls][subj] - 1
                    )

    # 少人数学級の残り週コマ数を初期化
    weekly_remaining_sg = {
        sg: dict(weekly_periods_sg.get(sg, {}))
        for sg in timetable_sg
    }

    # 少人数手動入力コマをロック
    for sg_name, day_map in manual_timetable_sg.items():
        if sg_name not in timetable_sg:
            continue
        for day, period_map in day_map.items():
            for period, subj in period_map.items():
                timetable_sg[sg_name].setdefault(day, {})[period] = subj
                if subj in weekly_remaining_sg.get(sg_name, {}):
                    weekly_remaining_sg[sg_name][subj] = max(
                        0, weekly_remaining_sg[sg_name][subj] - 1
                    )

    return weekly_remaining, weekly_remaining_sg


# ============================================================
# Part 4: MRV バックトラッキングソルバー
# ============================================================

def solve_backtracking(
    timetable: dict,
    timetable_sg: dict,
    weekly_remaining: dict,
    assignments: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    small_group_classes: dict,
    priority_subjects: list,
    deadline: float = 30.0,
    forbidden_slots: frozenset = frozenset()
) -> str:
    """
    MRV + バックトラッキング（逆引きキャッシュ高速化版）。
    戻り値: "solved" / "timeout" / "failed"

    forbidden_slots: frozenset of (cls, day, period, subject) tuples.
        指定したコマ・教科の組み合わせを禁止することで、
        「別パターン生成」時に異なる解へ誘導する。

    【高速化】
    teacher_busy[t][day]  : set of periods where teacher t is already assigned
    teacher_day_cnt[t][day]: int  count of assignments for R7 check
    room_usage[subj][day][p]: int count of classes using special room at (day,p)
    すべて O(1) アクセス。is_valid_placement の O(n_classes) スキャンを排除。
    """
    start_time = time.time()

    # ── 逆引きキャッシュ初期化 ───────────────────────────────
    all_teachers = list(set(
        t for cls_asgn in assignments.values()
        for subj_teachers in cls_asgn.values()
        for t in subj_teachers
    ))

    # teacher_busy[t][day] = {period, ...}  既に担当しているコマのset
    teacher_busy: dict = {t: {d: set() for d in DAYS} for t in all_teachers}
    # teacher_day_cnt[t][day] = 当日担当コマ数（R7用）
    teacher_day_cnt: dict = {t: {d: 0 for d in DAYS} for t in all_teachers}
    # room_usage[subj][day][p] = 特別教室を使用中のクラス数
    # room_usage[room_idx][day][p] = 使用中クラス数
    max_p_global = max(periods_per_day.values(), default=8)
    room_usage: dict = {
        idx: {d: {p: 0 for p in range(1, max_p_global + 1)} for d in DAYS}
        for idx in range(len(special_rooms))
    }

    # ── キャッシュに手動ロック済みコマを反映 ────────────────
    def _register(cls_: str, day_: str, period_: int, subj_: str, sign: int):
        """sign=+1: 登録, sign=-1: 取消"""
        for t in assignments.get(cls_, {}).get(subj_, []):
            if t in teacher_busy:
                if sign == 1:
                    teacher_busy[t][day_].add(period_)
                else:
                    teacher_busy[t][day_].discard(period_)
                teacher_day_cnt[t][day_] += sign
        cls_g = get_grade(cls_)
        for ridx, room_ in enumerate(special_rooms):
            gs_ = room_.get("grade_subjects", [])
            if [cls_g, subj_] in gs_ or (cls_g, subj_) in gs_:
                room_usage[ridx][day_][period_] = max(
                    0, room_usage[ridx][day_].get(period_, 0) + sign
                )

    for cls_, day_map in timetable.items():
        for day_, period_map in day_map.items():
            for period_, subj_ in period_map.items():
                if subj_:
                    _register(cls_, day_, period_, subj_, 1)
    # 少人数学級の手動コマもteacher_busy に反映
    for sg_name, day_map in timetable_sg.items():
        sg_teachers = small_group_classes.get(sg_name, {}).get("teachers", [])
        for day_, period_map in day_map.items():
            for period_, subj_ in period_map.items():
                if subj_:
                    for t in sg_teachers:
                        if t in teacher_busy:
                            teacher_busy[t][day_].add(period_)
                            teacher_day_cnt[t][day_] += 1

    # ── 高速版 is_valid ──────────────────────────────────────
    def fast_valid(cls: str, day: str, period: int, subject: str) -> bool:
        # 禁止配置チェック（別パターン生成時のみ有効）
        if forbidden_slots and (cls, day, period, subject) in forbidden_slots:
            return False

        teachers = assignments.get(cls, {}).get(subject, [])

        # R1: O(1) – キャッシュ参照
        for t in teachers:
            if period in teacher_busy.get(t, {}).get(day, set()):
                return False

        # R3: O(1) – room_usage 参照
        cls_g = get_grade(cls)
        for ridx, room_ in enumerate(special_rooms):
            gs_ = room_.get("grade_subjects", [])
            if [cls_g, subject] not in gs_ and (cls_g, subject) not in gs_:
                continue
            if room_usage[ridx][day].get(period, 0) >= room_["capacity"]:
                return False

        # R4: 教員不在
        for t in teachers:
            if period in unavailable.get(t, {}).get(day, []):
                return False

        # R5: 同日同教科
        if subject in timetable.get(cls, {}).get(day, {}).values():
            # period 自体はまだ未配置なので values に入っていれば重複
            day_slots = timetable[cls][day]
            for p_, s_ in day_slots.items():
                if p_ != period and s_ == subject:
                    return False

        # R6: 少人数同期 → ソフト制約化のためハードチェックを削除
        # （soft_score でR6適合度を加点する方式に変更）

        # R7: O(1) – teacher_day_cnt 参照
        # 「出勤可能コマ数 - 1」を上限とする。出勤可能コマが0以下なら制約不要。
        total_today = periods_per_day.get(day, 6)
        for t in teachers:
            unavail_cnt   = len(unavailable.get(t, {}).get(day, []))
            available     = total_today - unavail_cnt  # 出勤可能コマ数
            max_ok        = available - 1              # 最低1コマ空き
            if max_ok <= 0:
                continue  # 出勤不可の曜日はR7対象外
            if teacher_day_cnt.get(t, {}).get(day, 0) >= max_ok:
                return False

        return True

    # ── 未配置スロット取得 ───────────────────────────────────
    def get_unassigned_slots():
        """残コマのあるクラスの未配置スロットのみ返す（修正2）"""
        return [
            (cls, day, p)
            for cls, schedule in timetable.items()
            if any(v > 0 for v in weekly_remaining.get(cls, {}).values())
            for day in DAYS
            for p in range(1, periods_per_day.get(day, 6) + 1)
            if schedule.get(day, {}).get(p) is None
        ]

    def get_candidates(cls: str, day: str, period: int) -> list:
        return [
            s for s in SUBJECTS
            if weekly_remaining.get(cls, {}).get(s, 0) > 0
            and fast_valid(cls, day, period, s)
        ]

    # ── バックトラッキング本体 ───────────────────────────────
    def backtrack() -> str:
        if time.time() - start_time > deadline:
            return "timeout"

        # 【修正1】残コマが全部0なら完成（空きスロットが残っていてもOK）
        if all(v == 0 for rem in weekly_remaining.values()
               for v in rem.values()):
            return "solved"

        slots = get_unassigned_slots()
        if not slots:
            return "failed"  # スロット埋めたが残コマあり

        # MRV + タイブレーク
        # slot_priority計算時の候補を保持して再利用（2重計算を防ぐ）
        slot_cands: dict = {}
        def slot_priority(slot):
            c, d, p = slot
            cands = get_candidates(c, d, p)
            slot_cands[slot] = cands
            tb = tiebreak_score(
                c, d, p, timetable, timetable_sg,
                small_group_classes, special_rooms,
                unavailable, periods_per_day,
                weekly_remaining, priority_subjects, assignments,
                teacher_day_cnt  # キャッシュを渡してR7リスクをO(1)化
            )
            return (len(cands), -tb)

        best_slot = min(slots, key=slot_priority)
        cls, day, period = best_slot
        candidates = slot_cands.get(best_slot) or get_candidates(cls, day, period)

        if not candidates:
            return "failed"

        candidates.sort(
            key=lambda s: soft_score(
                s, cls, day, timetable, assignments, priority_subjects, period
            ),
            reverse=True
        )

        for subj in candidates:
            # 配置 + キャッシュ更新
            timetable[cls][day][period] = subj
            weekly_remaining[cls][subj] -= 1
            _register(cls, day, period, subj, 1)

            result = backtrack()
            if result in ("solved", "timeout"):
                return result

            # 取消 + キャッシュ巻き戻し
            timetable[cls][day][period] = None
            weekly_remaining[cls][subj] += 1
            _register(cls, day, period, subj, -1)

        return "failed"

    return backtrack()


# ============================================================
# Part 4: OR-Tools フォールバックソルバー
# ============================================================

def solve_ortools(
    timetable: dict,
    timetable_sg: dict,
    weekly_remaining: dict,
    assignments: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    small_group_classes: dict,
    priority_subjects: list,
    random_seed: int = 0,
    time_limit: float = 60.0
) -> str:
    """
    Google OR-Tools CP-SAT で時間割を生成する。
    time_limit: ソルバーの最大実行秒数（デフォルト60秒）
    戻り値: "solved" / "failed"
    """
    try:
        from ortools.sat.python import cp_model

        model   = cp_model.CpModel()
        classes = list(timetable.keys())
        x       = {}  # x[cls][day][p][subj] ∈ {0,1}

        # ── 変数定義 ─────────────────────────────────────────
        for cls in classes:
            x[cls] = {}
            for day in DAYS:
                x[cls][day] = {}
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    x[cls][day][p] = {
                        subj: model.NewBoolVar(f"x_{cls}_{day}_{p}_{subj}")
                        for subj in SUBJECTS
                    }

        # ── 手動ロック済みコマを固定 ─────────────────────────
        for cls in classes:
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    locked = timetable[cls].get(day, {}).get(p)
                    if locked:
                        model.Add(x[cls][day][p][locked] == 1)
                        for subj in SUBJECTS:
                            if subj != locked:
                                model.Add(x[cls][day][p][subj] == 0)

        # ── 基本制約: 各コマに高々1教科 ─────────────────────
        for cls in classes:
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    model.AddAtMostOne(x[cls][day][p][s] for s in SUBJECTS)
                    # 学年別コマ数を超えるスロットは全て0固定
                    if p > get_periods_for_class(cls, day):
                        for s in SUBJECTS:
                            model.Add(x[cls][day][p][s] == 0)

        # ── R1: 教員重複禁止 ─────────────────────────────────
        for t in st.session_state["teachers"]:
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    tvars = [
                        x[cls][day][p][subj]
                        for cls in classes
                        for subj in SUBJECTS
                        if t in assignments.get(cls, {}).get(subj, [])
                    ]
                    if tvars:
                        model.AddAtMostOne(tvars)

        # ── R2: 週コマ数完全消化 ─────────────────────────────
        for cls in classes:
            for subj in SUBJECTS:
                target = weekly_remaining.get(cls, {}).get(subj, 0)
                pvars = [
                    x[cls][day][p][subj]
                    for day in DAYS
                    for p in range(1, periods_per_day.get(day, 6) + 1)
                    if not timetable[cls].get(day, {}).get(p)
                ]
                if pvars:
                    model.Add(sum(pvars) == target)

        # ── R3: 特別教室同時使用制限 ─────────────────────────
        for room in special_rooms:
            cap = room["capacity"]
            gs  = room.get("grade_subjects", [])
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    room_vars = [
                        x[c][day][p][s]
                        for c in classes
                        for g, s in gs
                        if get_grade(c) == g and c in x
                        and not timetable[c].get(day, {}).get(p)
                    ]
                    # 手動配置済みも合算
                    manual_count = sum(
                        1 for c in classes
                        for g, s in gs
                        if get_grade(c) == g and timetable[c].get(day, {}).get(p) == s
                    )
                    if room_vars:
                        model.Add(sum(room_vars) + manual_count <= cap)

        # ── R4: 教員不在コマ ─────────────────────────────────
        for t in st.session_state["teachers"]:
            for day in DAYS:
                for p in unavailable.get(t, {}).get(day, []):
                    if p > periods_per_day.get(day, 6):
                        continue
                    for cls in classes:
                        for subj in SUBJECTS:
                            if t in assignments.get(cls, {}).get(subj, []):
                                model.Add(x[cls][day][p][subj] == 0)

        # ── R5: 同曜日・同教科1コマまで ─────────────────────
        for cls in classes:
            for day in DAYS:
                for subj in SUBJECTS:
                    model.Add(sum(
                        x[cls][day][p][subj]
                        for p in range(1, periods_per_day.get(day, 6) + 1)
                    ) <= 1)

        # ── R8: 学年間教科同期制約（一方向・minority→majority） ──
        # コマ数の少ない側(minority)がある時限には必ず多い側(majority)も存在する。
        # 多い側の余りコマは少ない側が存在しない時限に自由配置。
        cross_grade_sync = st.session_state.get("cross_grade_sync", [])
        all_classes_list = get_all_classes()
        for pair in cross_grade_sync:
            g1, s1, g2, s2 = pair["grade1"], pair["subject1"], pair["grade2"], pair["subject2"]
            grade1_cls = [c for c in all_classes_list if get_grade(c) == g1 and c in classes]
            grade2_cls = [c for c in all_classes_list if get_grade(c) == g2 and c in classes]
            if not grade1_cls or not grade2_cls:
                continue

            # 残コマ合計でminority/majorityを決定
            total_g1 = sum(weekly_remaining.get(c, {}).get(s1, 0) for c in grade1_cls)
            total_g2 = sum(weekly_remaining.get(c, {}).get(s2, 0) for c in grade2_cls)
            if total_g1 <= total_g2:
                min_cls, min_subj = grade1_cls, s1
                maj_cls, maj_subj = grade2_cls, s2
            else:
                min_cls, min_subj = grade2_cls, s2
                maj_cls, maj_subj = grade1_cls, s1

            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    # has_min: minority側のいずれかがmin_subjを持つ
                    has_min = model.NewBoolVar(
                        f"r8_min_{g1}_{s1}_{g2}_{s2}_{day}_{p}")
                    sum_min = sum(x[c][day][p][min_subj] for c in min_cls
                                  if not timetable[c].get(day, {}).get(p))
                    manual_min = sum(
                        1 for c in min_cls
                        if timetable[c].get(day, {}).get(p) == min_subj
                    )
                    model.Add(sum_min + manual_min >= 1).OnlyEnforceIf(has_min)
                    model.Add(sum_min + manual_min == 0).OnlyEnforceIf(has_min.Not())

                    # has_maj: majority側のいずれかがmaj_subjを持つ
                    has_maj = model.NewBoolVar(
                        f"r8_maj_{g1}_{s1}_{g2}_{s2}_{day}_{p}")
                    sum_maj = sum(x[c][day][p][maj_subj] for c in maj_cls
                                  if not timetable[c].get(day, {}).get(p))
                    manual_maj = sum(
                        1 for c in maj_cls
                        if timetable[c].get(day, {}).get(p) == maj_subj
                    )
                    model.Add(sum_maj + manual_maj >= 1).OnlyEnforceIf(has_maj)
                    model.Add(sum_maj + manual_maj == 0).OnlyEnforceIf(has_maj.Not())

                    # has_min → has_maj（一方向: minorityがあればmajorityも必要）
                    model.Add(has_maj >= has_min)

        # ── 目的関数ペナルティ項（R6ソフト＋S1学年集約） ────────
        penalty_terms = []

        # ── R6: 少人数学級同期（ソフト制約・ペナルティ） ────────
        # ハード制約から外し、違反1件あたり大きなペナルティを課す。
        # これにより「できるだけ揃える」方向に最適化しつつ、
        # 揃えられない場合も時間割全体は生成できる。
        R6_PENALTY = 1000
        for sg_name, sg_data in small_group_classes.items():
            groups = sg_data.get("sync_groups") or [sg_data.get("classes", [])]
            for group in groups:
                valid_group = [c for c in group if c in classes]
                if len(valid_group) < 2:
                    continue
                for day in DAYS:
                    for p in range(1, periods_per_day.get(day, 6) + 1):
                        # is_common_vars[c] = 1 ならそのクラスはこの時限に合同教科
                        joint_subjs = sg_data.get("joint_subjects", [])
                        is_common_vars = {}
                        for c in valid_group:
                            iv = model.NewBoolVar(f"ic_{sg_name}_{c}_{day}_{p}")
                            if joint_subjs:
                                cs = sum(x[c][day][p][s] for s in joint_subjs)
                            else:
                                cs = model.NewConstant(0)
                            model.Add(cs >= 1).OnlyEnforceIf(iv)
                            model.Add(cs == 0).OnlyEnforceIf(iv.Not())
                            is_common_vars[c] = iv
                        # 隣接するクラスペアで不一致のときペナルティ変数を立てる
                        # ── R6ペナルティ変数: mismatch = ref XOR c ───────
                        # AddBoolXOr([a,b,c]) は a XOR b XOR c = 1 を意味する
                        # mismatch = ref XOR c を表すには
                        # ref XOR c XOR mismatch.Not() = 1 とする
                        ref = valid_group[0]
                        for c in valid_group[1:]:
                            mismatch = model.NewBoolVar(
                                f"r6mismatch_{sg_name}_{ref}_{c}_{day}_{p}")
                            model.AddBoolXOr(
                                [is_common_vars[ref],
                                 is_common_vars[c],
                                 mismatch.Not()]
                            )
                            penalty_terms.append(R6_PENALTY * mismatch)

        # ── R7: 教員1日最低1コマ空き ─────────────────────────
        for t in st.session_state["teachers"]:
            for day in DAYS:
                total_today = periods_per_day.get(day, 6)
                unavail_cnt = len(unavailable.get(t, {}).get(day, []))
                available   = total_today - unavail_cnt   # 出勤可能コマ数
                # 少人数学級の手動ロック済みコマ数を考慮して上限を減算
                sg_cnt = sum(
                    1 for sg_name, sg_sched in timetable_sg.items()
                    for p in range(1, total_today + 1)
                    if sg_sched.get(day, {}).get(p)
                    and t in small_group_classes.get(sg_name, {}).get("teachers", [])
                    and p not in unavailable.get(t, {}).get(day, [])
                )
                max_ok      = available - 1 - sg_cnt     # 最低1コマ空き（少人数分を減算）
                if max_ok <= 0:
                    continue  # 出勤不可またはすでに上限の曜日はR7対象外
                tvars = [
                    x[cls][day][p][subj]
                    for cls in classes
                    for p in range(1, total_today + 1)
                    for subj in SUBJECTS
                    if t in assignments.get(cls, {}).get(subj, [])
                    and p not in unavailable.get(t, {}).get(day, [])
                ]
                if tvars:
                    model.Add(sum(tvars) <= max_ok)

        # ── S1: 学年集約ペナルティ（時限集約 + 曜日集約） ──────────
        if st.session_state.get("soft_grade_grouping", True):
            PW, NW = 10, 1

            # S1a: 同教員・異学年が同日の異なる時限に配置されたらペナルティ
            for t in st.session_state["teachers"]:
                for day in DAYS:
                    dp = periods_per_day.get(day, 6)
                    for p1 in range(1, dp + 1):
                        for p2 in range(p1 + 1, dp + 1):
                            for c1 in classes:
                                for c2 in classes:
                                    if c1 == c2 or get_grade(c1) == get_grade(c2):
                                        continue
                                    for s1 in SUBJECTS:
                                        if t not in assignments.get(c1, {}).get(s1, []):
                                            continue
                                        for s2 in SUBJECTS:
                                            if t not in assignments.get(c2, {}).get(s2, []):
                                                continue
                                            w = PW if (s1 in priority_subjects
                                                       or s2 in priority_subjects) else NW
                                            pv = model.NewBoolVar(
                                                f"p_{t}_{day}_{p1}_{p2}_{c1}_{c2}")
                                            model.AddBoolAnd([
                                                x[c1][day][p1][s1],
                                                x[c2][day][p2][s2]
                                            ]).OnlyEnforceIf(pv)
                                            model.AddBoolOr([
                                                x[c1][day][p1][s1].Not(),
                                                x[c2][day][p2][s2].Not()
                                            ]).OnlyEnforceIf(pv.Not())
                                            penalty_terms.append(w * pv)

            # S1b: 同学年・同教科は同じ曜日に集約するペナルティ
            # has_on_day[c][s][d] = クラスcが教科sを曜日dに持つか
            has_on_day = {}
            for c in classes:
                has_on_day[c] = {}
                for s in SUBJECTS:
                    has_on_day[c][s] = {}
                    for d in DAYS:
                        dp = periods_per_day.get(d, 6)
                        xvars = [x[c][d][p][s] for p in range(1, dp + 1)]
                        hv = model.NewBoolVar(f"hod_{c}_{s}_{d}")
                        model.AddBoolOr(xvars).OnlyEnforceIf(hv)
                        model.AddBoolAnd([xv.Not() for xv in xvars]).OnlyEnforceIf(hv.Not())
                        has_on_day[c][s][d] = hv

            # 同学年クラスのペアで、同教科が異なる曜日に配置されたらペナルティ
            grade_cls_map: dict[int, list] = {}
            for c in classes:
                grade_cls_map.setdefault(get_grade(c), []).append(c)

            for gcls in grade_cls_map.values():
                for i, c1 in enumerate(gcls):
                    for c2 in gcls[i + 1:]:
                        for s in SUBJECTS:
                            w = PW if s in priority_subjects else NW
                            for d in DAYS:
                                # c1とc2でsの曜日配置が食い違うときペナルティ
                                mismatch = model.NewBoolVar(
                                    f"daymm_{c1}_{c2}_{s}_{d}")
                                model.AddBoolXOr([
                                    has_on_day[c1][s][d],
                                    has_on_day[c2][s][d],
                                    mismatch.Not()
                                ])
                                penalty_terms.append(w * mismatch)

            # S3: 学年交互防止（1〜4限で同一教員が連続して異学年を担当しない）
            # 優先教科が絡む遷移は S3_PW、それ以外は S3_NW でペナルティを課す
            S3_PW, S3_NW = 10, 1
            S3_TARGET = [p for p in range(1, 5)
                         if p <= min(periods_per_day.get(d, 6) for d in DAYS)]
            grades_in_school = sorted({get_grade(c) for c in classes})

            if len(grades_in_school) >= 2 and len(S3_TARGET) >= 2:
                # teaches_grade_var[t][d][p][g]: 全教科版
                teaches_grade_var = {}
                # teaches_prio_grade[t][d][p][g]: 優先教科のみ版
                teaches_prio_grade = {}
                for t in st.session_state["teachers"]:
                    teaches_grade_var[t] = {}
                    teaches_prio_grade[t] = {}
                    for d in DAYS:
                        teaches_grade_var[t][d] = {}
                        teaches_prio_grade[t][d] = {}
                        for p in S3_TARGET:
                            teaches_grade_var[t][d][p] = {}
                            teaches_prio_grade[t][d][p] = {}
                            for g in grades_in_school:
                                # 全教科
                                tvars_g = [
                                    x[c][d][p][s]
                                    for c in classes if get_grade(c) == g
                                    for s in SUBJECTS
                                    if t in assignments.get(c, {}).get(s, [])
                                    and p <= periods_per_day.get(d, 6)
                                ]
                                tg = model.NewBoolVar(f"tg_{t}_{d}_{p}_{g}")
                                if tvars_g:
                                    model.AddBoolOr(tvars_g).OnlyEnforceIf(tg)
                                    model.AddBoolAnd(
                                        [v.Not() for v in tvars_g]
                                    ).OnlyEnforceIf(tg.Not())
                                else:
                                    model.Add(tg == 0)
                                teaches_grade_var[t][d][p][g] = tg

                                # 優先教科のみ
                                pvars_g = [
                                    x[c][d][p][s]
                                    for c in classes if get_grade(c) == g
                                    for s in priority_subjects
                                    if t in assignments.get(c, {}).get(s, [])
                                    and p <= periods_per_day.get(d, 6)
                                ]
                                tpg = model.NewBoolVar(f"tpg_{t}_{d}_{p}_{g}")
                                if pvars_g:
                                    model.AddBoolOr(pvars_g).OnlyEnforceIf(tpg)
                                    model.AddBoolAnd(
                                        [v.Not() for v in pvars_g]
                                    ).OnlyEnforceIf(tpg.Not())
                                else:
                                    model.Add(tpg == 0)
                                teaches_prio_grade[t][d][p][g] = tpg

                # 連続する時限 (p, p+1) で学年が切り替わるとペナルティ
                # 優先教科が絡む遷移は S3_PW、それ以外は S3_NW
                for t in st.session_state["teachers"]:
                    for d in DAYS:
                        for p in S3_TARGET[:-1]:   # (1,2),(2,3),(3,4)
                            p2 = p + 1
                            if p2 not in S3_TARGET:
                                continue
                            for g1 in grades_in_school:
                                for g2 in grades_in_school:
                                    if g1 == g2:
                                        continue
                                    tg1 = teaches_grade_var[t][d][p][g1]
                                    tg2 = teaches_grade_var[t][d][p2][g2]
                                    # 学年遷移変数
                                    pv = model.NewBoolVar(
                                        f"s3_{t}_{d}_{p}_{g1}_{g2}")
                                    model.AddBoolAnd([tg1, tg2]).OnlyEnforceIf(pv)
                                    model.AddBoolOr(
                                        [tg1.Not(), tg2.Not()]
                                    ).OnlyEnforceIf(pv.Not())
                                    # 基本ペナルティ（全遷移に S3_NW）
                                    penalty_terms.append(S3_NW * pv)
                                    # 優先教科が絡む場合は追加ペナルティ
                                    if priority_subjects:
                                        tpg1 = teaches_prio_grade[t][d][p][g1]
                                        pv_p = model.NewBoolVar(
                                            f"s3p_{t}_{d}_{p}_{g1}_{g2}")
                                        model.AddBoolAnd(
                                            [pv, tpg1]
                                        ).OnlyEnforceIf(pv_p)
                                        model.AddBoolOr(
                                            [pv.Not(), tpg1.Not()]
                                        ).OnlyEnforceIf(pv_p.Not())
                                        # 追加分: PW - NW = 9
                                        penalty_terms.append(
                                            (S3_PW - S3_NW) * pv_p)

        if penalty_terms:
            model.Minimize(sum(penalty_terms))

        # ── ソルバー実行 ─────────────────────────────────────
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds  = float(time_limit)
        solver.parameters.num_search_workers   = 4
        if random_seed != 0:
            solver.parameters.random_seed = random_seed % (2**31 - 1)
        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            for cls in classes:
                for day in DAYS:
                    for p in range(1, periods_per_day.get(day, 6) + 1):
                        if timetable[cls].get(day, {}).get(p):
                            continue
                        for subj in SUBJECTS:
                            if solver.Value(x[cls][day][p][subj]) == 1:
                                timetable[cls].setdefault(day, {})[p] = subj
                                # weekly_remaining を減算して整合性を保つ
                                if subj in weekly_remaining.get(cls, {}):
                                    weekly_remaining[cls][subj] = max(
                                        0, weekly_remaining[cls][subj] - 1)
                                break
            return "solved"
        return "failed"

    except Exception as e:
        st.error(f"OR-Toolsエラー: {e}")
        return "failed"



# ============================================================
# Part 4: 学年別段階生成
# ============================================================

def solve_grade_by_grade(
    timetable: dict,
    timetable_sg: dict,
    weekly_remaining: dict,
    assignments: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    small_group_classes: dict,
    priority_subjects: list,
    mrv_deadline: float = 20.0,
    ort_time_limit: float = 60.0
) -> str:
    """
    学年ごとに順番に解く段階的生成。

    クロス学年の制約処理:
    ・R1（教員重複）: 前学年の使用済みコマを extra_unavail として次学年に引き継ぐ
    ・R3（教室容量）: 前学年の per-slot 使用数を特別教室の容量から引いて渡す
    ・少人数同期のある学年を先に解く（制約が厳しいため）
    戻り値: "solved" / "failed_grade{学年番号}"
    """
    from collections import defaultdict

    # ── 学年別クラスグループ ─────────────────────────────────
    grade_groups: dict[int, list] = defaultdict(list)
    for cls in timetable:
        grade_groups[get_grade(cls)].append(cls)

    # ── 少人数同期グループを学年別にフィルタリング ─────────────
    def sg_for_classes(cls_list):
        result = {}
        for sg_name, sg_data in small_group_classes.items():
            in_list = [c for c in sg_data.get("classes", []) if c in cls_list]
            if not in_list:
                continue
            filtered_groups = [
                [c for c in g if c in cls_list]
                for g in (sg_data.get("sync_groups") or [sg_data.get("classes", [])])
                if sum(1 for c in g if c in cls_list) >= 2
            ]
            if filtered_groups:
                result[sg_name] = {**sg_data,
                                   "classes": in_list,
                                   "sync_groups": filtered_groups}
        return result

    # ── 処理順: 少人数同期のある学年を優先 ──────────────────
    sorted_grades = sorted(grade_groups.keys(),
                           key=lambda g: -len(sg_for_classes(grade_groups[g])))

    # ── 追加不在（前学年で使用した教員スロット）───────────────
    # teacher -> day -> [period, ...]
    extra_unavail: dict[str, dict[str, list]] = defaultdict(
        lambda: defaultdict(list))
    # 教室使用カウント: subject -> day -> period -> count
    room_slot_used: dict[str, dict[str, dict[int, int]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(int)))

    def make_merged_unavail():
        result = {}
        all_teachers = set(unavailable) | set(extra_unavail)
        for t in all_teachers:
            result[t] = {
                d: sorted(set(
                    list(unavailable.get(t, {}).get(d, [])) +
                    list(extra_unavail[t].get(d, []))
                ))
                for d in DAYS
            }
        return result

    def make_adjusted_rooms(grade):
        """
        前学年の per-slot 使用数を引いた特別教室リストを返す。
        時限ごとに残容量が変わるが、OR-Tools/MRV の仕様上
        最も制約の厳しい時限（最小残容量）を各教室の有効容量とする。
        さらに詳細な時限制約は MRV キャッシュ経由で自動適用される。
        """
        adjusted = []
        for room in special_rooms:
            subj = room["subject"]
            cap  = room["capacity"]
            # 全時限のうち他学年が最も多く使う時限の使用数
            max_used = 0
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    used = room_slot_used[subj][day].get(p, 0)
                    max_used = max(max_used, used)
            new_cap = max(0, cap - max_used)
            adjusted.append({**room, "capacity": new_cap})
        return adjusted

    def make_combined_timetable(grade_classes):
        """
        この学年のみ空スロット、他学年は解済みスロット、を含む
        combined timetable を作成する（MRV の R3 キャッシュ用）。
        """
        combined = {}
        for cls, sched in timetable.items():
            combined[cls] = {d: dict(sched[d]) for d in DAYS}
        # grade_classes の slots を空にする（今から解く）
        for cls in grade_classes:
            combined[cls] = {d: dict(timetable[cls][d]) for d in DAYS}
        return combined

    # ── 学年ごとに生成 ────────────────────────────────────────
    for grade in sorted_grades:
        grade_classes = grade_groups[grade]
        sg_g = sg_for_classes(grade_classes)
        mv   = make_merged_unavail()

        # この学年の timetable/wr（手動コマはすでに preprocess 済み）
        tt_g = {cls: {d: dict(timetable[cls][d]) for d in DAYS}
                for cls in grade_classes}
        wr_g = {cls: dict(weekly_remaining[cls]) for cls in grade_classes}

        # MRV: combined_timetable を渡して R3 クロス学年を正確に反映
        combined_tt = make_combined_timetable(grade_classes)
        # combined_tt の grade_classes 部分は tt_g と同一なので上書き
        for cls in grade_classes:
            combined_tt[cls] = tt_g[cls]

        result = solve_backtracking(
            combined_tt, {}, wr_g, assignments,
            special_rooms, mv, periods_per_day, sg_g,
            priority_subjects, deadline=mrv_deadline
        )

        if result == "solved":
            # combined_tt から grade_classes の結果を tt_g に同期
            for cls in grade_classes:
                tt_g[cls] = combined_tt[cls]

        if result != "solved":
            # OR-Tools フォールバック（元の容量のまま渡す）
            # クロス学年の R3 は extra_unavail による R1 回避が間接的に軽減する
            tt_g2 = {cls: {d: dict(timetable[cls][d]) for d in DAYS}
                     for cls in grade_classes}
            wr_g2 = {cls: dict(weekly_remaining[cls]) for cls in grade_classes}
            result = solve_ortools(
                tt_g2, {}, wr_g2, assignments,
                special_rooms, mv, periods_per_day, sg_g,
                priority_subjects, time_limit=ort_time_limit
            )
            if result == "solved":
                tt_g = tt_g2
                wr_g = wr_g2

        if result != "solved":
            return f"failed_grade{grade}"

        # ── 解けた学年をメインの timetable/wr に書き戻す ──────
        for cls in grade_classes:
            timetable[cls] = tt_g[cls]
            weekly_remaining[cls] = wr_g[cls]

        # ── 教員使用スロットと教室使用スロットを次学年へ引き継ぐ ──
        for cls in grade_classes:
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    subj = timetable[cls][day].get(p)
                    if not subj:
                        continue
                    # R1: 教員スロット蓄積
                    for t in assignments.get(cls, {}).get(subj, []):
                        if p not in extra_unavail[t][day]:
                            extra_unavail[t][day].append(p)
                    # R3: 教室使用カウント
                    for room in special_rooms:
                        if room["subject"] == subj:
                            room_slot_used[subj][day][p] += 1

    return "solved"


# ============================================================
# Part 4: 教科別段階生成
# ============================================================

def solve_subject_by_subject(
    timetable: dict,
    timetable_sg: dict,
    weekly_remaining: dict,
    assignments: dict,
    special_rooms: list,
    unavailable: dict,
    periods_per_day: dict,
    small_group_classes: dict,
    priority_subjects: list,
    first_subjects: list,
    mrv_deadline: float = 20.0,
    ort_time_limit: float = 60.0
) -> str:
    """
    選択した教科を先に配置し、残りの教科を後から配置する段階的生成。

    Phase 1: first_subjects を全クラスに配置
    Phase 2: 残りの教科を配置（Phase 1 の教員スロットを引き継ぐ）

    戻り値: "solved" / "failed_phase1" / "failed_phase2"
    """
    remaining_subjects = [s for s in SUBJECTS if s not in first_subjects]

    for phase_idx, phase_subjects in enumerate([first_subjects, remaining_subjects], 1):
        if not phase_subjects:
            continue

        # この段階で配置する教科のみに絞った weekly_remaining
        wr_phase = {
            cls: {s: cnt for s, cnt in weekly_remaining[cls].items()
                  if s in phase_subjects}
            for cls in weekly_remaining
        }

        # 全クラスで残コマが 0 なら skip
        if all(v == 0 for rem in wr_phase.values() for v in rem.values()):
            continue

        # MRV 前のクリーンな timetable を保存（ORT フォールバック用）
        tt_clean = {cls: {d: dict(timetable[cls][d]) for d in DAYS}
                    for cls in timetable}

        result = solve_backtracking(
            timetable, timetable_sg,
            wr_phase, assignments,
            special_rooms, unavailable,
            periods_per_day, small_group_classes,
            priority_subjects, deadline=mrv_deadline
        )

        if result != "solved":
            # MRV のタイムアウト・失敗で timetable が汚染されている可能性があるため
            # クリーンな状態に戻してから OR-Tools に渡す
            for cls in timetable:
                for d in DAYS:
                    timetable[cls][d] = dict(tt_clean[cls][d])
            wr_phase2 = {
                cls: {s: cnt for s, cnt in weekly_remaining[cls].items()
                      if s in phase_subjects}
                for cls in weekly_remaining
            }
            result = solve_ortools(
                timetable, timetable_sg,
                wr_phase2, assignments,
                special_rooms, unavailable,
                periods_per_day, small_group_classes,
                priority_subjects, time_limit=ort_time_limit
            )

        if result != "solved":
            return f"failed_phase{phase_idx}"

        # timetable は in-place 更新済みなので次フェーズでそのまま使う

    return "solved"


# ============================================================
# Part 4: エラーレポート
# ============================================================

def report_r6_violations(timetable: dict):
    """
    生成済み時間割のR6ソフト制約の達成状況を集計して表示する。
    違反 = 同期グループ内で、ある時限に共通教科と非共通教科が混在しているコマ。
    """
    small_group_classes = st.session_state.get("small_group_classes", {})
    if not small_group_classes:
        return

    total_slots   = 0  # チェック対象スロット（ペア×時限）数
    violations    = []  # (sg_name, day, period, {cls: subj}) のリスト

    for sg_name, sg_data in small_group_classes.items():
        groups = sg_data.get("sync_groups") or [sg_data.get("classes", [])]
        for group in groups:
            valid_group = [c for c in group if c in timetable]
            if len(valid_group) < 2:
                continue
            periods_per_day = st.session_state["periods_per_day"]
            for day in DAYS:
                for p in range(1, periods_per_day.get(day, 6) + 1):
                    subjects = {
                        c: timetable[c][day].get(p)
                        for c in valid_group
                    }
                    filled = {c: s for c, s in subjects.items() if s}
                    if len(filled) < 2:
                        continue
                    total_slots += 1
                    joint_subjs = sg_data.get("joint_subjects", [])
                    flags = [s in joint_subjs for s in filled.values()]
                    if len(set(flags)) > 1:  # 混在
                        violations.append((sg_name, day, p, subjects))

    if total_slots == 0:
        return

    ok_count = total_slots - len(violations)
    rate = ok_count / total_slots * 100

    if not violations:
        st.success(
            f"✅ R6（少人数同期）: 全 {total_slots} スロット一致 "
            f"（達成率 100%）"
        )
    else:
        st.warning(
            f"⚠️ R6（少人数同期）: {ok_count}/{total_slots} スロット一致 "
            f"（達成率 {rate:.0f}%）　違反 {len(violations)} 件"
        )
        with st.expander("R6違反の詳細を見る", expanded=False):
            for sg_name, day, p, subjects in violations:
                joint_subjs = small_group_classes.get(sg_name, {}).get("joint_subjects", [])
                detail = "　".join(
                    f"{c}={'★' if s in joint_subjs else ''}{s}"
                    for c, s in subjects.items() if s
                )
                st.markdown(f"- **{sg_name}** {day}{p}限：{detail}")


def report_conflict(timetable: dict, weekly_remaining: dict):
    """解が見つからなかった場合に未配置コマと改善策を表示する。"""
    st.error("❌ 時間割を完成させることができませんでした")

    unplaced = [
        f"・{cls} ／ {subj}: あと **{rem}** コマ未配置"
        for cls, subj_map in weekly_remaining.items()
        for subj, rem in subj_map.items()
        if rem > 0
    ]
    if unplaced:
        st.markdown("**未配置コマ一覧:**")
        for item in unplaced:
            st.markdown(item)

    st.info(
        "💡 **改善策の例**\n"
        "- 週コマ数が多い教科を減らす（STEP 4）\n"
        "- 教員の不在コマを見直す（STEP 6）\n"
        "- 特別教室の同時使用可能クラス数を増やす（STEP 7）\n"
        "- 少人数学級の同期グループを再設定する（STEP 9）\n"
        "- 手動入力コマを減らす（STEP 8）"
    )


# ============================================================
# Part 4: 時間割表示
# ============================================================

def display_class_timetable(
    cls: str,
    timetable: dict,
    manual_timetable: dict,
    is_small_group: bool = False
):
    """クラスの時間割をグリッド表示する。手動=水色・自動=薄緑。"""
    max_p = get_max_periods()
    data  = {}

    for day in DAYS:
        day_periods = st.session_state["periods_per_day"].get(day, 6)
        col_data = []
        for p in range(1, max_p + 1):
            if p > day_periods:
                col_data.append("─")
                continue
            subj = timetable.get(cls, {}).get(day, {}).get(p, "")
            if not subj:
                col_data.append("")
                continue
            if is_small_group:
                teachers = st.session_state["small_group_classes"] \
                    .get(cls, {}).get("teachers", [])
            else:
                teachers = get_teachers_for_slot(cls, subj)
            cell = f"{subj}\n{'・'.join(teachers)}" if teachers else subj
            col_data.append(cell)
        data[day] = col_data

    df = pd.DataFrame(data, index=[f"{p}限" for p in range(1, max_p + 1)])
    st.dataframe(df, use_container_width=True)
    st.caption("🔵 水色=手動入力　🟢 薄緑=自動生成")


def display_teacher_timetable(
    teacher: str,
    timetable: dict,
    timetable_sg: dict
):
    """教員の時間割をグリッド表示する。空き・不在も色分け。"""
    max_p = get_max_periods()
    data  = {}

    for day in DAYS:
        day_periods = st.session_state["periods_per_day"].get(day, 6)
        col_data = []
        for p in range(1, max_p + 1):
            if p > day_periods:
                col_data.append("─")
                continue
            if p in st.session_state["unavailable"].get(teacher, {}).get(day, []):
                col_data.append("🚫不在")
                continue
            cell = "空き"
            for cls, schedule in timetable.items():
                subj = schedule.get(day, {}).get(p)
                if subj and teacher in get_teachers_for_slot(cls, subj):
                    cell = f"{cls}\n{subj}"
                    break
            if cell == "空き":
                for sg_name, sg_schedule in timetable_sg.items():
                    subj = sg_schedule.get(day, {}).get(p)
                    if subj and teacher in st.session_state["small_group_classes"] \
                            .get(sg_name, {}).get("teachers", []):
                        cell = f"{sg_name}\n{subj}"
                        break
            col_data.append(cell)
        data[day] = col_data

    df = pd.DataFrame(data, index=[f"{p}限" for p in range(1, max_p + 1)])
    st.dataframe(df, use_container_width=True)
    st.caption("🚫 不在　空欄=空きコマ")


def display_all_teachers_timetable(timetable: dict, timetable_sg: dict):
    """全教員の時間割を一覧表示する。縦=教員、横=曜日×時限。"""
    teachers      = st.session_state["teachers"]
    periods_per_day = st.session_state["periods_per_day"]
    unavailable   = st.session_state["unavailable"]
    small_group_classes = st.session_state["small_group_classes"]

    # 列ラベルを生成（例: 月1, 月2, ..., 金6）
    columns = [
        f"{day}{p}"
        for day in DAYS
        for p in range(1, periods_per_day.get(day, 6) + 1)
    ]

    rows = {}
    for teacher in teachers:
        row = []
        for day in DAYS:
            day_periods = periods_per_day.get(day, 6)
            for p in range(1, day_periods + 1):
                if p in unavailable.get(teacher, {}).get(day, []):
                    row.append("🚫不在")
                    continue
                cell = ""
                for cls, schedule in timetable.items():
                    subj = schedule.get(day, {}).get(p)
                    if subj and teacher in get_teachers_for_slot(cls, subj):
                        cell = f"{cls} {subj}"
                        break
                if not cell:
                    for sg_name, sg_schedule in timetable_sg.items():
                        subj = sg_schedule.get(day, {}).get(p)
                        if subj and teacher in small_group_classes.get(
                                sg_name, {}).get("teachers", []):
                            cell = f"{sg_name} {subj}"
                            break
                row.append(cell)
        rows[teacher] = row

    df = pd.DataFrame(rows, index=columns).T
    st.dataframe(df, use_container_width=True)
    st.caption("🚫 不在　空欄=空きコマ")


# ============================================================
# Part 4: Excel 出力
# ============================================================

def export_timetable_to_excel(
    timetable: dict,
    timetable_sg: dict,
    manual_timetable: dict
) -> bytes:
    """
    生成済み時間割をExcelに出力する。
    ①通常学級シート ②少人数学級シート ③教員別シート
    """
    wb  = openpyxl.Workbook()
    wb.remove(wb.active)
    thin        = Side(style="thin")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_align  = Alignment(wrap_text=True, horizontal="center", vertical="center")
    max_p       = get_max_periods()

    def make_fill(color):
        return PatternFill("solid", fgColor=color)

    def write_timetable_sheet(ws, cls, schedule, manual, is_sg=False):
        """共通の時間割シート書き込み処理。"""
        # ヘッダー行（曜日）
        ws.append(["時限"] + DAYS)
        _apply_header_style(ws)

        for p in range(1, max_p + 1):
            row_cells = []
            for day in DAYS:
                day_periods = st.session_state["periods_per_day"].get(day, 6)
                if p > day_periods:
                    row_cells.append("─")
                    continue
                subj = schedule.get(day, {}).get(p, "")
                if not subj:
                    row_cells.append("")
                    continue
                if is_sg:
                    teachers = st.session_state["small_group_classes"] \
                        .get(cls, {}).get("teachers", [])
                else:
                    teachers = get_teachers_for_slot(cls, subj)
                cell_val = f"{subj}\n{'・'.join(teachers)}" if teachers else subj
                row_cells.append(cell_val)

            ws.append([f"{p}限"] + row_cells)
            row_idx = p + 1
            ws.row_dimensions[row_idx].height = 35

            for col_idx, day in enumerate(DAYS, start=2):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border    = border
                cell.alignment = wrap_align
                day_periods = st.session_state["periods_per_day"].get(day, 6)
                if p > day_periods:
                    continue
                subj = schedule.get(day, {}).get(p, "")
                if not subj:
                    continue
                is_manual = manual.get(cls, {}).get(day, {}).get(p) is not None
                if is_sg:
                    cell.fill = make_fill(COLOR_SG)
                elif is_manual:
                    cell.fill = make_fill(COLOR_MANUAL)
                else:
                    cell.fill = make_fill(COLOR_AUTO)

            # 時限列のスタイル
            lc = ws.cell(row=row_idx, column=1)
            lc.border    = border
            lc.alignment = wrap_align
            lc.fill      = make_fill(COLOR_HEADER)
            lc.font      = Font(bold=True)

        # 列幅設定
        ws.column_dimensions["A"].width = 8
        for col_idx in range(2, len(DAYS) + 2):
            ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # ── ①全学年一覧シート（縦軸:学年クラス、横軸:曜日×時限） ──
    ws = wb.create_sheet("全学年一覧")
    ppd_grade = st.session_state.get("periods_per_day_grade", {})
    all_cls_list = get_all_classes()

    # 横軸: 曜日×時限の列を構築
    col_headers = []  # (day, period) のリスト
    for day in DAYS:
        day_max = max(
            (ppd_grade.get(g, {}).get(day, 6) for g in GRADES),
            default=max_p
        )
        for p in range(1, day_max + 1):
            col_headers.append((day, p))

    # ヘッダー行1: 曜日（各曜日の先頭列にmerge）
    ws.cell(row=1, column=1, value="学年クラス")
    ws.cell(row=1, column=1).fill      = make_fill(COLOR_HEADER)
    ws.cell(row=1, column=1).font      = Font(bold=True)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=1, column=1).border    = border
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    day_col_ranges = {}  # day -> (start_col, end_col)
    for col_i, (day, p) in enumerate(col_headers):
        col = col_i + 2
        if day not in day_col_ranges:
            day_col_ranges[day] = [col, col]
        else:
            day_col_ranges[day][1] = col

    for day, (sc, ec) in day_col_ranges.items():
        if sc == ec:
            cell = ws.cell(row=1, column=sc, value=day)
        else:
            ws.merge_cells(start_row=1, start_column=sc, end_row=1, end_column=ec)
            cell = ws.cell(row=1, column=sc, value=day)
        cell.fill      = make_fill(COLOR_HEADER)
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border

    # ヘッダー行2: 時限
    for col_i, (day, p) in enumerate(col_headers):
        col = col_i + 2
        cell = ws.cell(row=2, column=col, value=f"{p}限")
        cell.fill      = make_fill(COLOR_HEADER)
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border

    # データ行: 各クラス
    for row_i, cls in enumerate(all_cls_list):
        row = row_i + 3
        g   = get_grade(cls)

        # クラス名ラベル
        lc = ws.cell(row=row, column=1, value=cls)
        lc.fill      = make_fill(COLOR_HEADER)
        lc.font      = Font(bold=True)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border    = border
        ws.row_dimensions[row].height = 20

        schedule = timetable.get(cls, {})
        for col_i, (day, p) in enumerate(col_headers):
            col  = col_i + 2
            cell = ws.cell(row=row, column=col)
            cell.border    = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # この学年のこの曜日のコマ数上限を超える列はグレー
            day_max = ppd_grade.get(g, {}).get(day, 6)
            if p > day_max:
                cell.value = ""
                cell.fill  = make_fill(COLOR_ABSENT)
                continue

            subj = schedule.get(day, {}).get(p, "")
            cell.value = subj if subj else ""
            if subj:
                is_manual = manual_timetable.get(cls, {}).get(day, {}).get(p) is not None
                cell.fill = make_fill(COLOR_MANUAL if is_manual else COLOR_AUTO)

    # 列幅・行高
    ws.column_dimensions["A"].width = 10
    for col_i in range(len(col_headers)):
        ws.column_dimensions[get_column_letter(col_i + 2)].width = 6
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 16

    # ── ②通常学級シート（個別） ──────────────────────────────
    for cls in get_all_classes():
        ws = wb.create_sheet(cls)
        write_timetable_sheet(
            ws, cls,
            timetable.get(cls, {}),
            manual_timetable, is_sg=False
        )

    # ── ②少人数学級シート ────────────────────────────────────
    for sg_name in st.session_state["small_group_classes"]:
        ws = wb.create_sheet(sg_name)
        write_timetable_sheet(
            ws, sg_name,
            timetable_sg.get(sg_name, {}),
            {}, is_sg=True
        )

    # ── ③教員別シート ────────────────────────────────────────
    for teacher in st.session_state["teachers"]:
        ws = wb.create_sheet(f"教員_{teacher}")
        ws.append(["時限"] + DAYS)
        _apply_header_style(ws)

        for p in range(1, max_p + 1):
            row_cells = []
            for day in DAYS:
                day_periods = st.session_state["periods_per_day"].get(day, 6)
                if p > day_periods:
                    row_cells.append("─")
                    continue
                if p in st.session_state["unavailable"].get(teacher, {}).get(day, []):
                    row_cells.append("不在")
                    continue
                cell_val = "空き"
                for cls, schedule in timetable.items():
                    subj = schedule.get(day, {}).get(p)
                    if subj and teacher in get_teachers_for_slot(cls, subj):
                        cell_val = f"{cls}\n{subj}"
                        break
                if cell_val == "空き":
                    for sg_name, sg_schedule in timetable_sg.items():
                        subj = sg_schedule.get(day, {}).get(p)
                        if subj and teacher in st.session_state[
                                "small_group_classes"].get(sg_name, {}).get("teachers", []):
                            cell_val = f"{sg_name}\n{subj}"
                            break
                row_cells.append(cell_val)

            ws.append([f"{p}限"] + row_cells)
            row_idx = p + 1
            ws.row_dimensions[row_idx].height = 35

            for col_idx, day in enumerate(DAYS, start=2):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border    = border
                cell.alignment = wrap_align
                day_periods = st.session_state["periods_per_day"].get(day, 6)
                if p > day_periods:
                    continue
                val = row_cells[col_idx - 2]
                if val == "不在":
                    cell.fill = make_fill(COLOR_ABSENT)
                elif val == "空き":
                    pass
                else:
                    cell.fill = make_fill(COLOR_AUTO)

            lc = ws.cell(row=row_idx, column=1)
            lc.border = border; lc.alignment = wrap_align
            lc.fill   = make_fill(COLOR_HEADER)
            lc.font   = Font(bold=True)

        ws.column_dimensions["A"].width = 8
        for col_idx in range(2, len(DAYS) + 2):
            ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # ── ④全教員一覧シート ─────────────────────────────────────
    ws = wb.create_sheet("全教員一覧")
    periods_per_day = st.session_state["periods_per_day"]
    unavailable     = st.session_state["unavailable"]
    sg_classes      = st.session_state["small_group_classes"]

    # ヘッダー行（教員名 | 月1 月2 ... 金6）
    col_labels = [
        f"{day}{p}"
        for day in DAYS
        for p in range(1, periods_per_day.get(day, 6) + 1)
    ]
    ws.append(["教員名"] + col_labels)
    _apply_header_style(ws)

    for teacher in st.session_state["teachers"]:
        row_vals = [teacher]
        for day in DAYS:
            day_periods = periods_per_day.get(day, 6)
            for p in range(1, day_periods + 1):
                if p in unavailable.get(teacher, {}).get(day, []):
                    row_vals.append("不在")
                    continue
                cell_val = ""
                for cls, schedule in timetable.items():
                    subj = schedule.get(day, {}).get(p)
                    if subj and teacher in get_teachers_for_slot(cls, subj):
                        cell_val = f"{cls} {subj}"
                        break
                if not cell_val:
                    for sg_name, sg_schedule in timetable_sg.items():
                        subj = sg_schedule.get(day, {}).get(p)
                        if subj and teacher in sg_classes.get(
                                sg_name, {}).get("teachers", []):
                            cell_val = f"{sg_name} {subj}"
                            break
                row_vals.append(cell_val)
        ws.append(row_vals)

    # スタイル適用
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border    = border
            cell.alignment = wrap_align
            if cell.column == 1:
                cell.fill = make_fill(COLOR_HEADER)
                cell.font = Font(bold=True)
            elif cell.value == "不在":
                cell.fill = make_fill(COLOR_ABSENT)
            elif cell.value:
                cell.fill = make_fill(COLOR_AUTO)
    ws.column_dimensions["A"].width = 14
    for col_idx in range(2, len(col_labels) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 12

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# Part 4: 生成画面 render_generate()
# ============================================================

def render_generate():
    st.header("🎲 時間割生成・確認・出力")

    classes = get_all_classes()
    if not classes:
        st.warning("先にSTEP 1〜9を設定してください")
        return

    # ── バリデーション ────────────────────────────────────────
    warnings_list = validate_all_settings()
    if warnings_list:
        with st.expander(
            f"⚠️ 設定に {len(warnings_list)} 件の警告があります（クリックで展開）",
            expanded=False
        ):
            for w in warnings_list:
                st.warning(w)

    # ── ソフト制約設定 ────────────────────────────────────────
    render_soft_constraints()
    st.markdown("---")

    # ── 生成ボタン ────────────────────────────────────────────
    col1, col2, col3 = st.columns([3, 2, 1])
    with col1:
        generate_btn = st.button(
            "🚀 時間割を自動生成する",
            type="primary",
            use_container_width=True
        )
    with col2:
        has_prev = bool(st.session_state.get("generated_timetable"))
        alt_btn = st.button(
            "🔀 別パターンを生成する",
            use_container_width=True,
            disabled=not has_prev,
            help="現在と異なる配置パターンを探します（生成後に有効になります）"
        )
    with col3:
        time_limit = st.number_input(
            "制限（秒）",
            min_value=10, max_value=120, value=30, step=5,
            help="この秒数を超えるとOR-Toolsに切り替えます"
        )

    first_subjects = st.session_state.get("soft_first_subjects", [])
    subject_btn = st.button(
        "📚 教科別段階生成",
        use_container_width=True,
        disabled=not first_subjects,
        help="先に配置する教科を選択してから実行してください（ソフト制約設定で選択）"
    )

    # パターン番号表示
    pattern_no = st.session_state.get("timetable_pattern_no", 0)
    if has_prev:
        label = "標準パターン" if pattern_no == 1 else f"パターン #{pattern_no}"
        st.caption(f"🎲 現在表示中: {label}　　別パターンを生成するたびに番号が増えます")

    # ── 生成処理の共通ロジック ───────────────────────────────
    def _run_generation(forbidden_slots: frozenset = frozenset(), ort_seed: int = 0):
        """
        forbidden_slots: 別パターン生成時に禁止する (cls,day,period,subject) の集合。
                         通常生成時は空集合。
        ort_seed:        OR-Toolsフォールバック時のランダムシード。
        """
        with st.spinner("時間割を生成中..."):
            timetable = {
                cls: {day: {} for day in DAYS}
                for cls in classes
            }
            timetable_sg = {
                sg: {day: {} for day in DAYS}
                for sg in st.session_state["small_group_classes"]
            }
            weekly_periods_sg = {
                sg: sg_data.get("weekly_periods", {})
                for sg, sg_data in st.session_state["small_group_classes"].items()
            }
            weekly_remaining, weekly_remaining_sg = preprocess(
                timetable, timetable_sg,
                st.session_state["manual_timetable"],
                st.session_state["manual_timetable_sg"],
                st.session_state["weekly_periods"],
                weekly_periods_sg
            )
            priority_subjects = st.session_state.get("soft_priority_subjects", [])
            assignments        = st.session_state["assignments"]
            special_rooms      = st.session_state["special_rooms"]
            unavailable        = st.session_state["unavailable"]
            periods_per_day    = st.session_state["periods_per_day"]
            sg_classes         = st.session_state["small_group_classes"]

            result = solve_backtracking(
                timetable, timetable_sg,
                weekly_remaining, assignments,
                special_rooms, unavailable,
                periods_per_day, sg_classes,
                priority_subjects,
                deadline=float(time_limit),
                forbidden_slots=forbidden_slots
            )

            used_ortools = False
            used_grade_split = False

            if result in ("timeout", "failed"):
                if result == "timeout":
                    st.warning(
                        f"⏱ {time_limit}秒以内に解が見つかりませんでした。"
                        " OR-Tools (CP-SAT) にフォールバックします…"
                    )
                else:
                    st.info(
                        "🔄 MRVバックトラッキングで解が見つかりませんでした。"
                        " OR-Tools (CP-SAT) で再挑戦します…"
                    )
                timetable_for_ort = {
                    cls: {day: {} for day in DAYS} for cls in timetable
                }
                timetable_sg_for_ort = {
                    sg: {day: {} for day in DAYS} for sg in timetable_sg
                }
                weekly_remaining_ort, _ = preprocess(
                    timetable_for_ort, timetable_sg_for_ort,
                    st.session_state["manual_timetable"],
                    st.session_state["manual_timetable_sg"],
                    st.session_state["weekly_periods"],
                    weekly_periods_sg
                )
                result = solve_ortools(
                    timetable_for_ort, timetable_sg_for_ort,
                    weekly_remaining_ort, assignments,
                    special_rooms, unavailable,
                    periods_per_day, sg_classes,
                    priority_subjects,
                    random_seed=ort_seed,
                    time_limit=float(time_limit)
                )
                if result == "solved":
                    timetable    = timetable_for_ort
                    timetable_sg = timetable_sg_for_ort
                    weekly_remaining = weekly_remaining_ort
                used_ortools = True

            # ── 学年別段階生成フォールバック ──────────────────
            # MRV・OR-Toolsとも失敗し、複数学年がある場合のみ試みる
            n_grades = len(set(get_grade(c) for c in classes))
            if result != "solved" and n_grades >= 2:
                st.info(
                    "🔀 学年別段階生成を試みます…"
                    "（各学年を順番に生成し、教員の衝突を調整します）"
                )
                # 新しいtimetableを用意してpreprocess
                timetable_gs = {cls: {day: {} for day in DAYS} for cls in classes}
                timetable_sg_gs = {
                    sg: {day: {} for day in DAYS} for sg in timetable_sg
                }
                wr_gs, _ = preprocess(
                    timetable_gs, timetable_sg_gs,
                    st.session_state["manual_timetable"],
                    st.session_state["manual_timetable_sg"],
                    st.session_state["weekly_periods"],
                    weekly_periods_sg
                )
                result = solve_grade_by_grade(
                    timetable_gs, timetable_sg_gs,
                    wr_gs, assignments,
                    special_rooms, unavailable,
                    periods_per_day, sg_classes,
                    priority_subjects,
                    mrv_deadline=float(time_limit),
                    ort_time_limit=float(time_limit) * 2
                )
                if result == "solved":
                    timetable    = timetable_gs
                    timetable_sg = timetable_sg_gs
                    weekly_remaining = wr_gs
                used_grade_split = True

            if result == "solved":
                st.session_state["generated_timetable"]    = timetable
                st.session_state["generated_timetable_sg"] = timetable_sg
                if used_grade_split:
                    method = "学年別段階生成"
                elif used_ortools:
                    method = "OR-Tools CP-SAT"
                else:
                    method = "MRVバックトラッキング"
                pno = st.session_state.get("timetable_pattern_no", 0)
                lbl = "標準パターン" if pno == 1 else f"パターン #{pno}"
                from datetime import datetime as _dt
                history_entry = {
                    "timetable":    timetable,
                    "timetable_sg": timetable_sg,
                    "label":        f"{lbl}（{method}）",
                    "timestamp":    _dt.now().strftime("%H:%M:%S"),
                }
                st.session_state.setdefault("timetable_history", []).append(history_entry)
                st.success(f"✅ {lbl} の生成が完了しました！（{method}）")
                report_r6_violations(timetable)
            else:
                report_conflict(timetable, weekly_remaining)
        return result

    # ── 教科別段階生成 ──────────────────────────────────────────
    def _run_subject_generation():
        """選択した教科を先に配置し、残りを後から配置する段階的生成。"""
        fs = st.session_state.get("soft_first_subjects", [])
        with st.spinner(f"教科別段階生成中… Phase 1: {', '.join(fs)}"):
            timetable = {cls: {day: {} for day in DAYS} for cls in classes}
            timetable_sg = {
                sg: {day: {} for day in DAYS}
                for sg in st.session_state["small_group_classes"]
            }
            weekly_periods_sg = {
                sg: sg_data.get("weekly_periods", {})
                for sg, sg_data in st.session_state["small_group_classes"].items()
            }
            weekly_remaining, _ = preprocess(
                timetable, timetable_sg,
                st.session_state["manual_timetable"],
                st.session_state["manual_timetable_sg"],
                st.session_state["weekly_periods"],
                weekly_periods_sg
            )
            priority_subjects = st.session_state.get("soft_priority_subjects", [])
            assignments     = st.session_state["assignments"]
            special_rooms   = st.session_state["special_rooms"]
            unavailable     = st.session_state["unavailable"]
            periods_per_day = st.session_state["periods_per_day"]
            sg_classes      = st.session_state["small_group_classes"]

            result = solve_subject_by_subject(
                timetable, timetable_sg,
                weekly_remaining, assignments,
                special_rooms, unavailable,
                periods_per_day, sg_classes,
                priority_subjects,
                first_subjects=fs,
                mrv_deadline=float(time_limit),
                ort_time_limit=float(time_limit) * 2
            )

            if result == "solved":
                st.session_state["generated_timetable"]    = timetable
                st.session_state["generated_timetable_sg"] = timetable_sg
                pno = st.session_state.get("timetable_pattern_no", 0)
                lbl = "標準パターン" if pno == 1 else f"パターン #{pno}"
                from datetime import datetime as _dt
                history_entry = {
                    "timetable":    timetable,
                    "timetable_sg": timetable_sg,
                    "label":        f"{lbl}（教科別段階生成: {', '.join(fs)} → 残り）",
                    "timestamp":    _dt.now().strftime("%H:%M:%S"),
                }
                st.session_state.setdefault("timetable_history", []).append(history_entry)
                st.success(
                    f"✅ {lbl} の生成が完了しました！（教科別段階生成）\n"
                    f"Phase 1: {', '.join(fs)} → Phase 2: 残りの教科"
                )
                report_r6_violations(timetable)
            elif result == "failed_phase1":
                st.error(
                    f"❌ Phase 1（{', '.join(fs)}）の配置に失敗しました。\n"
                    "選択した教科の担当教員・週コマ数・教室設定を確認してください。"
                )
            else:
                st.error(
                    "❌ Phase 2（残りの教科）の配置に失敗しました。\n"
                    "Phase 1 の配置後に残るスロットが不足している可能性があります。"
                )
        return result

    # ── 通常生成 ─────────────────────────────────────────────
    if generate_btn:
        st.session_state["timetable_pattern_no"] = 1
        result = _run_generation()
        if result != "solved":
            return

    # ── 別パターン生成 ────────────────────────────────────────
    if alt_btn:
        import random as _rnd
        prev_tt = st.session_state.get("generated_timetable", {})
        # 前の解から全配置を収集し、約25%をランダムに禁止する
        all_placements = [
            (cls, day, p, subj)
            for cls, day_map in prev_tt.items()
            for day, period_map in day_map.items()
            for p, subj in period_map.items()
            if subj
        ]
        if all_placements:
            n_forbid = max(1, len(all_placements) // 4)
            forbidden = frozenset(_rnd.sample(all_placements, n_forbid))
        else:
            forbidden = frozenset()
        pno = st.session_state.get("timetable_pattern_no", 1) + 1
        st.session_state["timetable_pattern_no"] = pno
        result = _run_generation(
            forbidden_slots=forbidden,
            ort_seed=pno
        )
        if result != "solved":
            # 別パターンが見つからなかった場合はカウンタを戻す
            st.session_state["timetable_pattern_no"] = pno - 1
            return

    # ── 教科別段階生成ボタン ──────────────────────────────────
    if subject_btn:
        st.session_state["timetable_pattern_no"] = 1
        result = _run_subject_generation()
        if result != "solved":
            return

    # ── 生成済み時間割の表示 ──────────────────────────────────
    history = st.session_state.get("timetable_history", [])

    if not history:
        st.info("上の「時間割を自動生成する」ボタンを押してください")
        return

    st.markdown("---")

    # 履歴が複数ある場合は選択UIを表示
    if len(history) > 1:
        history_labels = [
            f"{h['timestamp']}  {h['label']}"
            for h in reversed(history)
        ]
        selected_label = st.selectbox(
            "📋 閲覧する時間割を選択",
            options=history_labels,
            index=0,
            help="過去に生成した時間割と現在の時間割を切り替えて閲覧できます",
        )
        selected_pos = history_labels.index(selected_label)
        selected_entry = list(reversed(history))[selected_pos]
        timetable    = selected_entry["timetable"]
        timetable_sg = selected_entry["timetable_sg"]
        if selected_pos == 0:
            st.caption("✅ 最新の時間割を表示しています")
        else:
            st.caption(f"🕐 過去の時間割を表示中: {selected_entry['label']}")
    else:
        timetable    = history[-1]["timetable"]
        timetable_sg = history[-1]["timetable_sg"]

    tab_cls, tab_sg, tab_teacher, tab_all, tab_dl = st.tabs(
        ["📘 クラス別", "📗 少人数学級", "👤 教員別", "👥 全教員一覧", "📥 ダウンロード"]
    )

    with tab_cls:
        selected_cls = st.selectbox(
            "クラスを選択",
            options=get_all_classes(),
            key="disp_cls_select"
        )
        if selected_cls:
            st.subheader(f"📘 {selected_cls} の時間割")
            display_class_timetable(
                selected_cls, timetable,
                st.session_state["manual_timetable"]
            )

    with tab_sg:
        sg_list = list(st.session_state["small_group_classes"].keys())
        if not sg_list:
            st.info("少人数学級が登録されていません")
        else:
            selected_sg = st.selectbox(
                "少人数学級を選択",
                options=sg_list,
                key="disp_sg_select"
            )
            if selected_sg:
                st.subheader(f"📗 {selected_sg} の時間割")
                display_class_timetable(
                    selected_sg, timetable_sg,
                    st.session_state["manual_timetable_sg"],
                    is_small_group=True
                )

    with tab_teacher:
        teachers = st.session_state["teachers"]
        if not teachers:
            st.info("教員が登録されていません")
        else:
            selected_t = st.selectbox(
                "教員を選択",
                options=teachers,
                key="disp_teacher_select"
            )
            if selected_t:
                st.subheader(f"👤 {selected_t} の時間割")
                display_teacher_timetable(selected_t, timetable, timetable_sg)

    with tab_all:
        teachers = st.session_state["teachers"]
        if not teachers:
            st.info("教員が登録されていません")
        else:
            st.subheader("👥 全教員の時間割一覧")
            display_all_teachers_timetable(timetable, timetable_sg)

    with tab_dl:
        st.markdown("#### 📥 時間割をExcelでダウンロード")
        fname = f"時間割_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="📥 時間割Excelをダウンロード",
            data=export_timetable_to_excel(
                timetable, timetable_sg,
                st.session_state["manual_timetable"]
            ),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("クラス別・少人数学級・教員別シートが含まれます")


# ============================================================
# ⑤ main()（完成版）
# ============================================================

def main():
    st.set_page_config(
        page_title="小学校時間割自動生成",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    init_session_state()
    render_sidebar()

    step = st.session_state["current_step"]
    if   step == 1: render_step1()
    elif step == 2: render_step2()
    elif step == 3: render_step3()
    elif step == 4: render_step4()
    elif step == 5: render_step5()
    elif step == 6: render_step6()
    elif step == 7: render_step7()
    elif step == 8: render_step8()
    elif step == 9: render_step9()
    elif step == 10: render_step10()
    elif step == 0: render_generate()


if __name__ == "__main__":
    main()

