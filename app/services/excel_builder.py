"""Build the 10-sheet estimation Excel from DB session data."""
import json
import io
import aiosqlite
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

HEADER_ROWS = 7
SEP_ROWS = 7
DATA1_START = 9

SN1 = '文档规模及工作量估计'
SN2 = '代码规模及工作量估计'
SN3 = '汇总表(文档相关)'
SN4 = '汇总表(代码相关)'
SN5 = '配置管理工作量估计'
SN6 = '质量保证工作量估计'
SN7 = '其他管理类工作量估计'
SN8 = '评审工作量'
SN9 = '规模估计结果报告'
SN10 = '工作量估计结果报告'

_ASSUME_DOC_SCALE = (
    '文档规模估计基于：\n'
    '1）文档假设采用公司模板编写，页面设置为A4，字号为小四；\n'
    '2）出所测试文档假设按照配置项测试文档模板编写；\n'
    '3）文档规模包含封面在内的所有页；\n'
    '4）参考软件研制任务书的功能数量；\n'
    '5）参考XXXXXX项目；\n'
    '6）专家法估计偏差阈值按20%考虑。'
)
_ASSUME_DOC_WL = (
    '文档编写工作量和其他工程活动工作量基于：\n'
    '1）文档编写人员均有同类文档的编写经验，掌握438C军用软件文档编写要求；\n'
    '2）参考文档估计规模；\n'
    '3）参考XXXXXX项目；\n'
    '4）专家法估计偏差阈值按20%考虑。'
)
_ASSUME_CODE_SCALE = (
    '代码规模估计基于：\n'
    '1）软件源代码的度量按源代码行（LOC）计算，只考虑有效代码行，不考虑注释。\n'
    '2）软件采用B/S架构，国产化麒麟系统和达梦数据库，代码编写采用C++、Java和JS语言，'
    '使用微软VS Code和IDEA开发工具；\n'
    '3）参考XXXXXX的开发经验和规模经验；\n'
    '4）参考软件研制任务书中的功能点数量；\n'
    '5）只对可能发生变化或新增的软件工作产品进行估计。\n'
    '6）专家法估计偏差阈值按20%考虑。'
)
_ASSUME_CODE_WL = (
    '编码、单元测试及软件BUG修改的工作量估计基于：\n'
    '1）编码人员均有仿真类和评估类项目的开发经验，且熟悉B/S架构的开发模式；\n'
    '2）参考XXXXXX和XXXXXX中期评估项目的开发经验和规模经验；\n'
    '3）参考代码估计规模；\n'
    '4）专家法估计偏差阈值按20%考虑。'
)
_ASSUME_CM = (
    '配置管理活动工作量的估计，基于公司提供的单次活动工时经验数据、项目周期内预计开展的活动次数'
    '及配置活动内容计算所得。公司提供的单次活动工时经验值见《软件估计指南》'
)
_ASSUME_QA = (
    '质量保证活动工作量的估计，基于公司提供的单次活动工时经验数据、项目周期内预计开展的质量活动'
    '次数计算所得。公司提供的单次活动工时经验值见《软件估计指南》。'
)
_ASSUME_MGMT = (
    '工作例会、测量等管理类活动工作量的估计，基于公司提供的单次活动工时经验数据、项目周期内预计'
    '开展的活动次数及预计参与人数计算所得。公司提供的单次活动工时经验值见《软件估计指南》。'
)
_ASSUME_REVIEW = (
    '评审活动工作量的估计，基于公司提供的单次活动工时经验数据、评审对象内容、参与评审人数计算所得。'
    '公司提供的单次活动工时经验值见《软件估计指南》'
)
_NOTE_12 = (
    '说明：\n'
    '1）当估计类别选择"工作量"时，只需填写总值一栏；\n'
    '2）当估计类别选择"规模"时，总值=新增规模+重用规模；\n'
    '3）规模的度量单位为"行"（代码）或"页"（文档），工作量为人时；\n'
    '4）估计对象中，各软件模块应能对应到软件研制任务书中的功能点；\n'
    '5）"估计假设和约束条件"一栏应说明用于参考的软件研制任务书、类似项目信息、其他可能有用的相关信息；\n'
    '6）重用包括在原有基础上修改使用和直接使用'
)
_NOTE_9 = (
    '说明：\n'
    '1）总值=新增规模+重用规模；\n'
    '2）当估计方法为专家估计法时，总值一栏取自估计汇总表中的取整后估计值；\n'
    '3）规模的度量单位为"行"（代码）或"页"（文档）；\n'
    '4）估计对象中，各软件模块应能对应到软件研制任务书中的功能点；\n'
    '5）重用包括在原有基础上修改使用和直接使用。'
)
_NOTE_10 = (
    '说明：\n'
    '1）当估计方法为专家估计法时，估计值一栏取值附录B估计汇总表（专家估计法）中的取整后估计值；\n'
    '2）"估计假设和约束条件"一栏中应说明的内容包括但不限于：\n'
    '①估计方法为运算法时，所采用的各项系数及依据。\n'
    '②估计方法为专家估计法时，偏差阈值、用于参考的软件研制任务书、规模数据、类似项目信息、其他可能有用的相关信息。'
)

# ── 格式常量 ─────────────────────────────────────────────────────────────────

_THIN = Side(style='thin')
_THICK = Side(style='thick')
_FONT = Font(name='宋体', size=11)
_TITLE_FONT = Font(name='宋体', size=16)
_BORDER_A = Border(top=_THIN, bottom=_THIN, left=_THICK, right=_THIN)
_BORDER = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)
_ALIGN = Alignment(horizontal='left', vertical='center', wrap_text=True)
_TITLE_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Sheet 1/2 精确列宽 (A-Y, 25列，M为空白分隔列)
_CW_12 = {
    'A': 13.0, 'B': 13.0, 'C': 1.83203125, 'D': 13.0, 'E': 3.5,
    'F': 13.0, 'G': 8.83203125, 'H': 4.6640625, 'I': 5.33203125,
    'J': 8.1640625, 'K': 2.33203125, 'L': 12.6640625,
    'M': 2.0,
    'N': 13.0, 'O': 13.0, 'P': 3.1640625, 'Q': 13.0, 'R': 4.1640625,
    'S': 13.0, 'T': 13.0, 'U': 7.6640625, 'V': 2.0, 'W': 8.6640625,
    'X': 1.5, 'Y': 11.0,
}
# Sheet 1/2 标题行高和表头行高（标题行=row1, 表头=rows2-8）
_HH_12_TITLE = 41.0
_HH_12_HDR = {2: 29.0, 3: 33.0, 4: 21.5, 5: 109.5, 6: 84.0, 7: 15.5, 8: 18.0}
# Sheet 3/4 行高（含新增标题行 row1=41, 其余+1偏移）
_HH_34 = {1: 41.0, 2: 31.2, 3: 27.5, 4: 20.0, 5: 102.0, 6: 101.0, 8: 15.5}
# Sheet 3/4 列宽 (A-N, 14列)
_CW_34 = {
    'A': 6.0, 'B': 13.0, 'C': 13.0, 'D': 13.0, 'E': 13.0, 'F': 8.0,
    'G': 2.5, 'H': 11.0, 'I': 2.5, 'J': 11.0, 'K': 11.0, 'L': 10.0,
    'M': 10.0, 'N': 13.0,
}


def _apply_borders_font(ws, max_row: int, max_col: int, sep_cols: set = None):
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if sep_cols and cell.column in sep_cols:
                continue
            cell.border = _BORDER_A if cell.column == 1 else _BORDER
            cell.font = _FONT
            cell.alignment = _ALIGN


def _set_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def _set_page_a4(ws):
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = 'portrait'


def _fmt_sheet12(ws, block2_title: int, block2_start: int, block2_title_h: float = 49.0):
    """Column widths, row heights, borders, font for Sheet 1 or 2."""
    _set_widths(ws, _CW_12)
    _set_page_a4(ws)
    ws.row_dimensions[1].height = _HH_12_TITLE
    ws.row_dimensions[block2_title].height = block2_title_h
    for r, h in _HH_12_HDR.items():
        ws.row_dimensions[r].height = h
        ws.row_dimensions[block2_start + (r - 2)].height = h
    _apply_borders_font(ws, ws.max_row, 25, sep_cols={13})


def _fmt_sheet34(ws, max_col: int = 14):
    """Column widths, row heights, borders, font for Sheet 3 or 4."""
    _set_widths(ws, _CW_34)
    _set_page_a4(ws)
    for r, h in _HH_34.items():
        ws.row_dimensions[r].height = h
    _apply_borders_font(ws, ws.max_row, max_col)


def _fmt_activity(ws, max_col: int):
    """Borders and font for activity workload sheets (5-8)."""
    _set_page_a4(ws)
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18.0
    _apply_borders_font(ws, ws.max_row, max_col)


def _fmt_report(ws, max_col: int):
    """Borders and font for result report sheets (9-10)."""
    _set_page_a4(ws)
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16.0
    _apply_borders_font(ws, ws.max_row, max_col)


def _q(name: str) -> str:
    return f"'{name}'" if any(c in name for c in ' ()[]') else name


def _apply_title_style(ws, r: int, c: int = 1):
    """Apply center + size-16 font to a title cell (call after _apply_borders_font)."""
    cell = ws.cell(row=r, column=c)
    cell.font = _TITLE_FONT
    cell.alignment = _TITLE_ALIGN


def _mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _w(ws, r, c, v):
    ws.cell(row=r, column=c).value = v


def _wmc(ws, r1, c1, r2, c2, v):
    _mc(ws, r1, c1, r2, c2)
    ws.cell(row=r1, column=c1).value = v


async def _load_data(db: aiosqlite.Connection, session_id: int) -> dict:
    async with db.execute("SELECT * FROM estimate_sessions WHERE id=?", (session_id,)) as cur:
        session = dict(await cur.fetchone())
    project_id = session['project_id']

    async with db.execute("SELECT * FROM projects WHERE id=?", (project_id,)) as cur:
        project = dict(await cur.fetchone())

    async with db.execute(
        "SELECT * FROM sheet1_estimates WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        s1_all = [dict(r) for r in await cur.fetchall()]

    doc_rows = [r for r in s1_all if r['config_item_id'] is not None]  # 规模行
    all_wl1 = [r for r in s1_all if r['config_item_id'] is None]       # 工作量行（WBS 工程任务）

    N_doc = len(doc_rows)
    N_wl1 = len(all_wl1)
    N_item1 = N_doc + N_wl1
    OFFSET1 = N_item1 + 11

    async with db.execute(
        "SELECT * FROM code_scale_estimates WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        s2_all = [dict(r) for r in await cur.fetchall()]

    N_code = len(s2_all)
    N_item2 = 2 * N_code
    OFFSET2 = N_item2 + 11

    async with db.execute(
        "SELECT * FROM activity_workload_snapshots WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        act_all = [dict(r) for r in await cur.fetchall()]

    cm_rows = [r for r in act_all if r['activity_type'] == 'cm']
    qa_rows = [r for r in act_all if r['activity_type'] == 'qa']
    mgmt_rows = [r for r in act_all if r['activity_type'] == 'mgmt']

    async with db.execute(
        "SELECT * FROM review_workload_snapshots WHERE session_id=? ORDER BY id", (session_id,)
    ) as cur:
        rev_rows = [dict(r) for r in await cur.fetchall()]

    async with db.execute(
        "SELECT id, resources, duration_days, workload FROM wbs_tasks WHERE project_id=?",
        (project_id,)
    ) as cur:
        task_map = {r['id']: dict(r) for r in await cur.fetchall()}

    try:
        estimators = json.loads(project.get('estimators') or '{}')
    except Exception:
        estimators = {}

    def _pad(lst, n=4):
        return (list(lst) + [''] * n)[:n]

    doc_est = _pad(estimators.get('doc', []))
    code_est = _pad(estimators.get('code', []))

    return {
        'project': project,
        'session_id': session_id,
        'doc_rows': doc_rows,
        'all_wl1': all_wl1,
        'N_doc': N_doc,
        'N_wl1': N_wl1,
        'N_item1': N_item1,
        'OFFSET1': OFFSET1,
        's2_all': s2_all,
        'N_code': N_code,
        'N_item2': N_item2,
        'OFFSET2': OFFSET2,
        'cm_rows': cm_rows,
        'qa_rows': qa_rows,
        'mgmt_rows': mgmt_rows,
        'rev_rows': rev_rows,
        'task_map': task_map,
        'doc_est': doc_est,
        'code_est': code_est,
        'cm_est': estimators.get('cm', ''),
        'qa_est': estimators.get('qa', ''),
        'mgmt_est': estimators.get('mgmt', ''),
        'review_est': estimators.get('review', ''),
    }


def _write_block_header(ws, block_row: int, est_left: str, est_right: str, project: dict,
                        assumption_scale: str, assumption_wl: str):
    """Write 7-row block header (Sheet 1/2 style) starting at block_row."""
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    round_num = project.get('round', 1)

    r = block_row
    _wmc(ws, r, 1, r, 2, '项目名称')
    _wmc(ws, r, 3, r, 6, proj_name)
    _wmc(ws, r, 7, r, 8, '项目编号')
    _wmc(ws, r, 9, r, 12, proj_code)
    _wmc(ws, r, 14, r, 15, '项目名称')
    _wmc(ws, r, 16, r, 19, proj_name)
    _wmc(ws, r, 20, r, 21, '项目编号')
    _wmc(ws, r, 22, r, 25, proj_code)

    r = block_row + 1
    _w(ws, r, 1, '估计人')
    _wmc(ws, r, 2, r, 3, est_left)
    _wmc(ws, r, 4, r, 5, '估计类别')
    _wmc(ws, r, 6, r, 8, '■规模  ■工作量')
    _wmc(ws, r, 9, r, 10, '估计轮次')
    _wmc(ws, r, 11, r, 12, round_num)
    _w(ws, r, 14, '估计人')
    _wmc(ws, r, 15, r, 16, est_right)
    _wmc(ws, r, 17, r, 18, '估计类别')
    _wmc(ws, r, 19, r, 21, '■规模  ■工作量')
    _wmc(ws, r, 22, r, 23, '估计轮次')
    _wmc(ws, r, 24, r, 25, round_num)

    r = block_row + 2
    _wmc(ws, r, 1, r, 3, '估计方法')
    _wmc(ws, r, 4, r, 12, '专家估计法')
    _wmc(ws, r, 14, r, 16, '估计方法')
    _wmc(ws, r, 17, r, 25, '专家估计法')

    r = block_row + 3
    _wmc(ws, r, 1, r + 1, 3, '估计假设和约束条件')
    _wmc(ws, r, 4, r, 12, assumption_scale)      # 行4：规模假设
    _wmc(ws, r + 1, 4, r + 1, 12, assumption_wl) # 行5：工作量假设
    _wmc(ws, r, 14, r + 1, 16, '估计假设和约束条件')
    _wmc(ws, r, 17, r, 25, assumption_scale)
    _wmc(ws, r + 1, 17, r + 1, 25, assumption_wl)

    r = block_row + 5
    _wmc(ws, r, 1, r + 1, 1, '序号')
    _wmc(ws, r, 2, r + 1, 5, '估计对象')
    _wmc(ws, r, 6, r + 1, 6, '度量\n单位')
    _wmc(ws, r, 7, r, 11, '估计值')
    _w(ws, r, 12, '备注')
    _wmc(ws, r, 14, r + 1, 14, '序号')
    _wmc(ws, r, 15, r + 1, 18, '估计对象')
    _wmc(ws, r, 19, r + 1, 19, '度量\n单位')
    _wmc(ws, r, 20, r, 24, '估计值')
    _w(ws, r, 25, '备注')

    r = block_row + 6
    _w(ws, r, 7, '总值')
    _wmc(ws, r, 8, r, 9, '新增规模')
    _wmc(ws, r, 10, r, 11, '重用规模')
    _w(ws, r, 20, '总值')
    _wmc(ws, r, 21, r, 22, '新增规模')
    _wmc(ws, r, 23, r, 24, '重用规模')


def _write_scale_row(ws, r, seq, name, v_left, v_right, unit):
    _w(ws, r, 1, seq)
    _wmc(ws, r, 2, r, 5, name)
    _w(ws, r, 6, unit)
    _w(ws, r, 7, f'=H{r}+J{r}')
    _wmc(ws, r, 8, r, 9, v_left)
    _wmc(ws, r, 10, r, 11, 0)
    _w(ws, r, 14, seq)
    _wmc(ws, r, 15, r, 18, name)
    _w(ws, r, 19, unit)
    _w(ws, r, 20, f'=U{r}+W{r}')
    _wmc(ws, r, 21, r, 22, v_right)
    _wmc(ws, r, 23, r, 24, 0)


def _write_workload_row(ws, r, seq, name, wl_left, wl_right):
    _w(ws, r, 1, seq)
    _wmc(ws, r, 2, r, 5, name)
    _w(ws, r, 6, '人时')
    _w(ws, r, 7, wl_left)
    _wmc(ws, r, 8, r, 9, '—')
    _wmc(ws, r, 10, r, 11, '—')
    _w(ws, r, 14, seq)
    _wmc(ws, r, 15, r, 18, name)
    _w(ws, r, 19, '人时')
    _w(ws, r, 20, wl_right)
    _wmc(ws, r, 21, r, 22, '—')
    _wmc(ws, r, 23, r, 24, '—')


def _build_sheet1(ws, d: dict):
    N_doc = d['N_doc']
    N_item1 = d['N_item1']
    OFFSET1 = d['OFFSET1']
    doc_est = d['doc_est']
    project = d['project']

    BLOCK2_TITLE = N_item1 + 12
    BLOCK2_START = N_item1 + 13
    DATA2_START = BLOCK2_START + HEADER_ROWS

    # 标题行
    _wmc(ws, 1, 1, 1, 12, '文档规模及其他工程活动工作量估计表')
    _wmc(ws, 1, 14, 1, 25, '文档规模及工程活动工作量估计表')
    _wmc(ws, BLOCK2_TITLE, 1, BLOCK2_TITLE, 12, '文档规模及工程活动工作量估计表')
    _wmc(ws, BLOCK2_TITLE, 14, BLOCK2_TITLE, 25, '文档规模及工程活动工作量估计表')

    _write_block_header(ws, 2, doc_est[0], doc_est[1], project, _ASSUME_DOC_SCALE, _ASSUME_DOC_WL)
    _write_block_header(ws, BLOCK2_START, doc_est[2], doc_est[3], project, _ASSUME_DOC_SCALE, _ASSUME_DOC_WL)

    for i, row in enumerate(d['doc_rows']):
        r1 = DATA1_START + i
        r2 = DATA2_START + i
        _write_scale_row(ws, r1, i + 1, row['item_name'], row['v1'], row['v2'], '页')
        _write_scale_row(ws, r2, i + 1, row['item_name'], row['v3'], row['v4'], '页')

    for i, row in enumerate(d['all_wl1']):
        r1 = DATA1_START + N_doc + 1 + i   # +1 空行隔开规模区和工作量区
        r2 = DATA2_START + N_doc + 1 + i
        _write_workload_row(ws, r1, i + 1, row['item_name'], row['wl1'], row['wl2'])
        _write_workload_row(ws, r2, i + 1, row['item_name'], row['wl3'], row['wl4'])

    # 说明区块：合并 A:L 和 M:X，数据结束后空 1 行
    note_r1 = DATA1_START + N_item1 + 2
    note_r2 = DATA2_START + N_item1 + 2
    _wmc(ws, note_r1, 1,  note_r1, 12, _NOTE_12)
    _wmc(ws, note_r1, 14, note_r1, 25, _NOTE_12)
    _wmc(ws, note_r2, 1,  note_r2, 12, _NOTE_12)
    _wmc(ws, note_r2, 14, note_r2, 25, _NOTE_12)
    ws.row_dimensions[note_r1].height = 120
    ws.row_dimensions[note_r2].height = 120

    _fmt_sheet12(ws, BLOCK2_TITLE, BLOCK2_START, block2_title_h=49.0)
    _apply_title_style(ws, 1, 1)
    _apply_title_style(ws, 1, 14)
    _apply_title_style(ws, BLOCK2_TITLE, 1)
    _apply_title_style(ws, BLOCK2_TITLE, 14)


def _build_sheet2(ws, d: dict):
    N_code = d['N_code']
    N_item2 = d['N_item2']
    code_est = d['code_est']
    project = d['project']
    N_doc = d['N_doc']

    BLOCK2_TITLE = N_item2 + 12
    BLOCK2_START = N_item2 + 13
    DATA2_START = BLOCK2_START + HEADER_ROWS

    # 标题行
    _wmc(ws, 1, 1, 1, 12, '代码规模及实现工作量估计表')
    _wmc(ws, 1, 14, 1, 25, '代码规模及实现工作量估计表')
    _wmc(ws, BLOCK2_TITLE, 1, BLOCK2_TITLE, 12, '代码规模及实现工作量估计表')
    _wmc(ws, BLOCK2_TITLE, 14, BLOCK2_TITLE, 25, '代码规模及实现工作量估计表')

    _write_block_header(ws, 2, code_est[0], code_est[1], project, _ASSUME_CODE_SCALE, _ASSUME_CODE_WL)
    _write_block_header(ws, BLOCK2_START, code_est[2], code_est[3], project, _ASSUME_CODE_SCALE, _ASSUME_CODE_WL)

    for i, row in enumerate(d['s2_all']):
        seq = N_doc + 1 + i
        r1 = DATA1_START + i
        r2 = DATA2_START + i
        _write_scale_row(ws, r1, seq, row['module_name'], row['v1'], row['v2'], '行')
        _write_scale_row(ws, r2, seq, row['module_name'], row['v3'], row['v4'], '行')

    for i, row in enumerate(d['s2_all']):
        r1 = DATA1_START + N_code + 1 + i   # +1 空行隔开
        r2 = DATA2_START + N_code + 1 + i
        _write_workload_row(ws, r1, i + 1, row['module_name'], row['wl1'], row['wl2'])
        _write_workload_row(ws, r2, i + 1, row['module_name'], row['wl3'], row['wl4'])

    note_r1 = DATA1_START + N_item2 + 2
    note_r2 = DATA2_START + N_item2 + 2
    _wmc(ws, note_r1, 1,  note_r1, 12, _NOTE_12)
    _wmc(ws, note_r1, 14, note_r1, 25, _NOTE_12)
    _wmc(ws, note_r2, 1,  note_r2, 12, _NOTE_12)
    _wmc(ws, note_r2, 14, note_r2, 25, _NOTE_12)
    ws.row_dimensions[note_r1].height = 120
    ws.row_dimensions[note_r2].height = 120

    _fmt_sheet12(ws, BLOCK2_TITLE, BLOCK2_START, block2_title_h=47.0)
    _apply_title_style(ws, 1, 1)
    _apply_title_style(ws, 1, 14)
    _apply_title_style(ws, BLOCK2_TITLE, 1)
    _apply_title_style(ws, BLOCK2_TITLE, 14)


def _write_summary_header(ws, project: dict, estimators_str: str,
                          assumption_scale: str, assumption_wl: str, title: str = ''):
    """Sheet 3/4 header rows 1-8 with merged cells matching sample layout."""
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    round_num = project.get('round', 1)

    # Row 1: 标题行
    _wmc(ws, 1, 1, 1, 14, title)

    # Row 2: 项目名称
    _w(ws, 2, 1, '项目名称')
    _wmc(ws, 2, 2, 2, 8, proj_name)
    _wmc(ws, 2, 9, 2, 10, '项目编号')
    _wmc(ws, 2, 11, 2, 14, proj_code)

    # Row 3: 估计人
    _w(ws, 3, 1, '估计人')
    _wmc(ws, 3, 2, 3, 5, estimators_str)
    _w(ws, 3, 6, '估计类别')
    _wmc(ws, 3, 7, 3, 11, '■规模  ■工作量')
    _w(ws, 3, 12, '估计轮次')
    _wmc(ws, 3, 13, 3, 14, round_num)

    # Row 4: 估计方法
    _wmc(ws, 4, 1, 4, 3, '估计方法')
    _wmc(ws, 4, 4, 4, 14, '专家估计法')

    # Rows 5-6: assumptions (label spans A5:D6)
    _wmc(ws, 5, 1, 6, 4, '估计假设和约束条件')
    _wmc(ws, 5, 5, 5, 14, assumption_scale)
    _wmc(ws, 6, 5, 6, 14, assumption_wl)

    # Rows 7-8: column headers
    _wmc(ws, 7, 1, 8, 1, '序号')
    _wmc(ws, 7, 2, 8, 5, '估计对象')
    _wmc(ws, 7, 6, 8, 7, '度量单位')
    _wmc(ws, 7, 8, 7, 12, '估计值')
    _wmc(ws, 7, 13, 8, 14, '估计值\n取整')
    _wmc(ws, 8, 8, 8, 9, '平均')
    _w(ws, 8, 10, '最大')
    _w(ws, 8, 11, '最小')
    _w(ws, 8, 12, '偏差率')


def _build_sheet3(ws, d: dict):
    N_doc = d['N_doc']
    N_wl1 = d['N_wl1']
    N_item1 = d['N_item1']
    OFFSET1 = d['OFFSET1']
    doc_est = d['doc_est']
    project = d['project']
    task_map = d['task_map']
    all_wl1 = d['all_wl1']

    est_str = '、'.join(e for e in doc_est if e)
    _write_summary_header(ws, project, est_str, _ASSUME_DOC_SCALE, _ASSUME_DOC_WL,
                          title='估计汇总表（文档相关、工程活动相关）')

    # Scale section
    for i, row in enumerate(d['doc_rows']):
        r = DATA1_START + i
        r3 = r + OFFSET1
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 5, f"='{SN1}'!B{r}")
        _wmc(ws, r, 6, r, 7, '页')
        _wmc(ws, r, 8, r, 9, f"=('{SN1}'!G{r}+'{SN1}'!T{r}+'{SN1}'!G{r3}+'{SN1}'!T{r3})/4")
        _w(ws, r, 10, f"=MAX('{SN1}'!G{r},'{SN1}'!T{r},'{SN1}'!G{r3},'{SN1}'!T{r3})")
        _w(ws, r, 11, f"=MIN('{SN1}'!G{r},'{SN1}'!T{r},'{SN1}'!G{r3},'{SN1}'!T{r3})")
        _w(ws, r, 12, f'=MAX(ABS(J{r}-H{r}),ABS(K{r}-H{r}))/H{r}')
        ws.cell(row=r, column=12).number_format = '0.00%'
        _w(ws, r, 13, f'=ROUND(H{r},0)')

    # Workload section（与规模区之间有一行空行，行号 +1）
    for i, row in enumerate(all_wl1):
        r = DATA1_START + N_doc + 1 + i
        r3 = r + OFFSET1
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 5, f"='{SN1}'!B{r}")
        _wmc(ws, r, 6, r, 7, '人时')
        _wmc(ws, r, 8, r, 9, f"=('{SN1}'!G{r}+'{SN1}'!T{r}+'{SN1}'!G{r3}+'{SN1}'!T{r3})/4")
        _w(ws, r, 10, f"=MAX('{SN1}'!G{r},'{SN1}'!T{r},'{SN1}'!G{r3},'{SN1}'!T{r3})")
        _w(ws, r, 11, f"=MIN('{SN1}'!G{r},'{SN1}'!T{r},'{SN1}'!G{r3},'{SN1}'!T{r3})")
        _w(ws, r, 12, f'=MAX(ABS(J{r}-H{r}),ABS(K{r}-H{r}))/H{r}')
        ws.cell(row=r, column=12).number_format = '0.00%'
        _w(ws, r, 13, f'=ROUND(H{r},0)')

    # Schedule section
    DATA_END = DATA1_START + N_item1
    SCHED_TITLE = DATA_END + 8
    SCHED_HDR = SCHED_TITLE + 1
    SCHED_DATA = SCHED_TITLE + 2

    _wmc(ws, SCHED_TITLE, 1, SCHED_TITLE, 14, '工作量与进度安排表')
    ws.row_dimensions[SCHED_TITLE].height = 41.0
    ws.row_dimensions[SCHED_HDR].height = 30.0
    _w(ws, SCHED_HDR, 1, '序号')
    _wmc(ws, SCHED_HDR, 2, SCHED_HDR, 5, '估计对象')
    _wmc(ws, SCHED_HDR, 6, SCHED_HDR, 7, '工作量估计值')
    _wmc(ws, SCHED_HDR, 8, SCHED_HDR, 10, '计划分配资源')
    _wmc(ws, SCHED_HDR, 11, SCHED_HDR, 12, '每日可用\n工作时间')
    _w(ws, SCHED_HDR, 13, '任务所需工作日')

    for i, row in enumerate(all_wl1):
        r = SCHED_DATA + i
        wl_row = DATA1_START + N_doc + 1 + i  # 与上方工作量区对齐
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 5, f'=B{wl_row}')
        _wmc(ws, r, 6, r, 7, f'=M{wl_row}')

        tid = row.get('wbs_task_id')
        task = task_map.get(tid, {}) if tid else {}
        resources = task.get('resources') or ''
        dur = task.get('duration_days') or 1
        em = row.get('em_target') or 0
        k_val = max(1, round(em / dur)) if dur > 0 else 1

        _wmc(ws, r, 8, r, 10, resources)
        _wmc(ws, r, 11, r, 12, k_val)
        _w(ws, r, 13, f'=ROUNDUP(F{r}/K{r},0)')

    _fmt_sheet34(ws, max_col=14)
    _apply_title_style(ws, 1)


def _build_sheet4(ws, d: dict):
    N_code = d['N_code']
    N_item2 = d['N_item2']
    OFFSET2 = d['OFFSET2']
    code_est = d['code_est']
    project = d['project']
    task_map = d['task_map']
    N_doc = d['N_doc']

    est_str = '、'.join(e for e in code_est if e)
    _write_summary_header(ws, project, est_str, _ASSUME_CODE_SCALE, _ASSUME_CODE_WL,
                          title='汇总表（代码相关）')
    ws.row_dimensions[5].height = 121.0  # 代码规模假设文字较长

    # Scale section
    for i, row in enumerate(d['s2_all']):
        r = DATA1_START + i
        r3 = r + OFFSET2
        seq = N_doc + 1 + i
        _w(ws, r, 1, seq)
        _wmc(ws, r, 2, r, 5, f"='{SN2}'!B{r}")
        _wmc(ws, r, 6, r, 7, '行')
        _wmc(ws, r, 8, r, 9, f"=('{SN2}'!G{r}+'{SN2}'!T{r}+'{SN2}'!G{r3}+'{SN2}'!T{r3})/4")
        _w(ws, r, 10, f"=MAX('{SN2}'!G{r},'{SN2}'!T{r},'{SN2}'!G{r3},'{SN2}'!T{r3})")
        _w(ws, r, 11, f"=MIN('{SN2}'!G{r},'{SN2}'!T{r},'{SN2}'!G{r3},'{SN2}'!T{r3})")
        _w(ws, r, 12, f'=MAX(ABS(J{r}-H{r}),ABS(K{r}-H{r}))/H{r}')
        ws.cell(row=r, column=12).number_format = '0.00%'
        _w(ws, r, 13, f'=ROUND(H{r},0)')

    # Workload section（+1 空行隔开）
    for i, row in enumerate(d['s2_all']):
        r = DATA1_START + N_code + 1 + i
        r3 = r + OFFSET2
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 5, f"='{SN2}'!B{r}")
        _wmc(ws, r, 6, r, 7, '人时')
        _wmc(ws, r, 8, r, 9, f"=('{SN2}'!G{r}+'{SN2}'!T{r}+'{SN2}'!G{r3}+'{SN2}'!T{r3})/4")
        _w(ws, r, 10, f"=MAX('{SN2}'!G{r},'{SN2}'!T{r},'{SN2}'!G{r3},'{SN2}'!T{r3})")
        _w(ws, r, 11, f"=MIN('{SN2}'!G{r},'{SN2}'!T{r},'{SN2}'!G{r3},'{SN2}'!T{r3})")
        _w(ws, r, 12, f'=MAX(ABS(J{r}-H{r}),ABS(K{r}-H{r}))/H{r}')
        ws.cell(row=r, column=12).number_format = '0.00%'
        _w(ws, r, 13, f'=ROUND(H{r},0)')

    # Schedule section for Sheet 4
    DATA_END4 = DATA1_START + N_item2
    SCHED_TITLE = DATA_END4 + 8
    SCHED_HDR = SCHED_TITLE + 1
    SCHED_DATA = SCHED_TITLE + 2

    _wmc(ws, SCHED_TITLE, 1, SCHED_TITLE, 14, '代码工作量与进度安排')
    ws.row_dimensions[SCHED_TITLE].height = 41.0
    ws.row_dimensions[SCHED_HDR].height = 30.0
    _w(ws, SCHED_HDR, 1, '序号')
    _wmc(ws, SCHED_HDR, 2, SCHED_HDR, 5, '估计对象')
    _wmc(ws, SCHED_HDR, 6, SCHED_HDR, 7, '工作量估计值')
    _wmc(ws, SCHED_HDR, 8, SCHED_HDR, 10, '计划分配资源')
    _wmc(ws, SCHED_HDR, 11, SCHED_HDR, 12, '每日可用\n工作时间')
    _w(ws, SCHED_HDR, 13, '任务所需工作日')

    for i, row in enumerate(d['s2_all']):
        r = SCHED_DATA + i
        wl_row = DATA1_START + N_code + 1 + i
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 5, f'=B{wl_row}')
        _wmc(ws, r, 6, r, 7, f'=M{wl_row}')

        tid = row.get('wbs_task_id')
        task = task_map.get(tid, {}) if tid else {}
        resources = task.get('resources') or ''
        dur = task.get('duration_days') or 1
        em = row.get('em_target') or 0
        k_val = max(1, round(em / dur)) if dur > 0 else 1

        _wmc(ws, r, 8, r, 10, resources)
        _wmc(ws, r, 11, r, 12, k_val)
        _w(ws, r, 13, f'=ROUNDUP(F{r}/K{r},0)')

    _fmt_sheet34(ws, max_col=14)
    _apply_title_style(ws, 1)


def _build_sheet5(ws, d: dict):
    """CM workload sheet."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    cm_est = d['cm_est']

    _wmc(ws, 1, 1, 1, 6, '配置管理工作量估计')
    ws.row_dimensions[1].height = 41.0

    _w(ws, 2, 1, '项目名称')
    _wmc(ws, 2, 2, 2, 3, proj_name)
    _w(ws, 2, 4, '项目编号')
    _wmc(ws, 2, 5, 2, 6, proj_code)

    _w(ws, 3, 1, '估计方法')
    _wmc(ws, 3, 2, 3, 3, '经验法')
    _w(ws, 3, 4, '估计人员')
    _wmc(ws, 3, 5, 3, 6, cm_est)

    _w(ws, 4, 1, '估计假设与约束条件')
    _wmc(ws, 4, 2, 4, 6, _ASSUME_CM)
    ws.row_dimensions[4].height = 58.0

    _wmc(ws, 5, 1, 5, 6, '估计过程')
    ws.row_dimensions[5].height = 20.0

    _w(ws, 6, 2, '阶段时长（月）')
    _w(ws, 6, 3, '配置管理活动')
    _w(ws, 6, 4, '工作量估计\n（人时）')
    _w(ws, 6, 5, '单次活动工时\n（人时）')
    _w(ws, 6, 6, '活动次数')
    ws.row_dimensions[6].height = 30.0

    # group by phase and merge phase name cells
    phase_groups = []
    for row in d['cm_rows']:
        phase = row.get('phase_name') or ''
        if not phase_groups or phase_groups[-1][0] != phase:
            phase_groups.append((phase, []))
        phase_groups[-1][1].append(row)

    r = 7
    for phase, activities in phase_groups:
        phase_start = r
        for act in activities:
            _w(ws, r, 3, act['activity_name'])
            _w(ws, r, 5, act.get('workload') or 0)
            _w(ws, r, 6, 1)
            _w(ws, r, 4, f'=E{r}*F{r}')
            r += 1
        phase_end = r - 1
        if phase_start == phase_end:
            _w(ws, phase_start, 1, phase)
        else:
            _wmc(ws, phase_start, 1, phase_end, 1, phase)

    last_data = r - 1
    total_row = r
    _wmc(ws, total_row, 1, total_row, 3, '总计')
    _w(ws, total_row, 4, f'=SUM(D7:D{last_data})')
    d['_cm_total_row'] = total_row
    _fmt_activity(ws, max_col=6)
    _apply_title_style(ws, 1)


def _build_sheet6(ws, d: dict):
    """QA workload sheet."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    qa_est = d['qa_est']

    _wmc(ws, 1, 1, 1, 7, '质量保证工作量估计')
    ws.row_dimensions[1].height = 41.0

    _wmc(ws, 2, 1, 2, 2, '项目名称')
    _wmc(ws, 2, 3, 2, 4, proj_name)
    _w(ws, 2, 5, '项目编号')
    _wmc(ws, 2, 6, 2, 7, proj_code)

    _wmc(ws, 3, 1, 3, 2, '估计方法')
    _wmc(ws, 3, 3, 3, 4, '经验法')
    _w(ws, 3, 5, '估计人员')
    _wmc(ws, 3, 6, 3, 7, qa_est)

    _wmc(ws, 4, 1, 4, 2, '估计假设和约束条件')
    _wmc(ws, 4, 3, 4, 7, _ASSUME_QA)
    ws.row_dimensions[4].height = 44.0

    _wmc(ws, 5, 1, 5, 7, '估计过程')
    ws.row_dimensions[5].height = 20.0

    _w(ws, 6, 2, '阶段时长(月)')
    _w(ws, 6, 3, '质量活动')
    _w(ws, 6, 4, '工作量估计')
    _w(ws, 6, 5, '单次活动工时')
    _w(ws, 6, 7, '审核次数')
    ws.row_dimensions[6].height = 15.0

    # group by phase and merge phase name cells
    phase_groups = []
    for row in d['qa_rows']:
        phase = row.get('phase_name') or ''
        if not phase_groups or phase_groups[-1][0] != phase:
            phase_groups.append((phase, []))
        phase_groups[-1][1].append(row)

    r = 7
    for phase, activities in phase_groups:
        phase_start = r
        for act in activities:
            _w(ws, r, 3, act['activity_name'])
            _w(ws, r, 5, act.get('workload') or 0)
            _w(ws, r, 7, 1)
            _w(ws, r, 4, f'=E{r}*G{r}')
            r += 1
        phase_end = r - 1
        if phase_start == phase_end:
            _w(ws, phase_start, 1, phase)
        else:
            _wmc(ws, phase_start, 1, phase_end, 1, phase)

    last_data = r - 1
    total_row = r
    _wmc(ws, total_row, 1, total_row, 3, '总计')
    _w(ws, total_row, 4, f'=SUM(D7:D{last_data})')
    d['_qa_total_row'] = total_row
    _fmt_activity(ws, max_col=7)
    _apply_title_style(ws, 1)


def _build_sheet7(ws, d: dict):
    """Other management workload sheet."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    mgmt_est = d['mgmt_est']

    _wmc(ws, 1, 1, 1, 6, '其他管理类工作量估计')
    ws.row_dimensions[1].height = 41.0

    _w(ws, 2, 1, '项目名称')
    _wmc(ws, 2, 2, 2, 3, proj_name)
    _w(ws, 2, 4, '项目编号')
    _wmc(ws, 2, 5, 2, 6, proj_code)

    _w(ws, 3, 1, '估计人员')
    _wmc(ws, 3, 2, 3, 3, mgmt_est)
    _w(ws, 3, 4, '估计方法')
    _wmc(ws, 3, 5, 3, 6, '经验法')

    _w(ws, 4, 1, '估计假设和约束条件')
    _wmc(ws, 4, 2, 4, 6, _ASSUME_MGMT)
    ws.row_dimensions[4].height = 41.0

    _wmc(ws, 5, 1, 5, 6, '估计过程')
    ws.row_dimensions[5].height = 20.0

    _w(ws, 6, 2, '管理活动')
    _w(ws, 6, 3, '工作量（人时）')
    _w(ws, 6, 4, '活动次数')
    _w(ws, 6, 5, '单次活动工时经验值（人时）')
    _w(ws, 6, 6, '参加人数')
    ws.row_dimensions[6].height = 30.0

    # group by phase and merge phase name cells
    phase_groups = []
    for row in d['mgmt_rows']:
        phase = row.get('phase_name') or ''
        if not phase_groups or phase_groups[-1][0] != phase:
            phase_groups.append((phase, []))
        phase_groups[-1][1].append(row)

    r = 7
    for phase, activities in phase_groups:
        phase_start = r
        for act in activities:
            _w(ws, r, 2, act['activity_name'])
            _w(ws, r, 5, act.get('workload') or 0)
            _w(ws, r, 4, 1)
            _w(ws, r, 6, 1)
            _w(ws, r, 3, f'=D{r}*E{r}*F{r}')
            r += 1
        phase_end = r - 1
        if phase_start == phase_end:
            _w(ws, phase_start, 1, phase)
        else:
            _wmc(ws, phase_start, 1, phase_end, 1, phase)

    last_data = r - 1
    total_row = r
    _w(ws, total_row, 1, '总计')
    _w(ws, total_row, 3, f'=SUM(C7:C{last_data})')
    d['_mgmt_total_row'] = total_row
    _fmt_activity(ws, max_col=6)
    _apply_title_style(ws, 1)


def _build_sheet8(ws, d: dict):
    """Review workload sheet."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    review_est = d['review_est']

    _wmc(ws, 1, 1, 1, 5, '评审工作量估计')
    ws.row_dimensions[1].height = 41.0

    _wmc(ws, 2, 1, 2, 2, '项目名称')
    _w(ws, 2, 3, proj_name)
    _w(ws, 2, 4, '项目编号')
    _w(ws, 2, 5, proj_code)

    _wmc(ws, 3, 1, 3, 2, '估计人员')
    _w(ws, 3, 3, review_est)
    _w(ws, 3, 4, '估计方法')
    _w(ws, 3, 5, '经验法')

    _wmc(ws, 4, 1, 4, 2, '估计假设与约束条件')
    _wmc(ws, 4, 3, 4, 5, _ASSUME_REVIEW)
    ws.row_dimensions[4].height = 45.0

    _wmc(ws, 5, 1, 5, 5, '估计过程')
    ws.row_dimensions[5].height = 20.0

    _w(ws, 6, 1, '序号')
    _w(ws, 6, 2, '评审对象')
    _w(ws, 6, 3, '评审参加人数')
    _w(ws, 6, 4, '评审时长（小时）')
    _w(ws, 6, 5, '评审总工时（人时）')

    for i, row in enumerate(d['rev_rows']):
        r = 7 + i
        _w(ws, r, 1, i + 1)
        _w(ws, r, 2, row['activity_name'])
        participants = row.get('participants') or 1
        total_hours = row.get('total_hours') or 0
        duration = (total_hours / participants) if participants > 0 else total_hours
        _w(ws, r, 3, participants)
        _w(ws, r, 4, duration)
        _w(ws, r, 5, f'=C{r}*D{r}')

    last_data = 6 + len(d['rev_rows'])
    total_row = last_data + 1
    _wmc(ws, total_row, 1, total_row, 2, '总计')
    _w(ws, total_row, 5, f'=SUM(E7:E{last_data})')
    d['_review_total_row'] = total_row
    _fmt_activity(ws, max_col=5)
    _apply_title_style(ws, 1)


def _build_sheet9(ws, d: dict):
    """Scale results report."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    N_doc = d['N_doc']
    N_code = d['N_code']

    # Row 1: 标题行
    _wmc(ws, 1, 1, 1, 9, '规模估计结果报告')
    ws.row_dimensions[1].height = 42.0

    # Row 2: 项目名称
    _wmc(ws, 2, 1, 2, 2, '项目名称')
    _wmc(ws, 2, 3, 2, 6, proj_name)
    _w(ws, 2, 7, '项目编号')
    _wmc(ws, 2, 8, 2, 9, proj_code)

    # Row 3: 文档规模估计 section header
    _wmc(ws, 3, 1, 3, 9, '文档规模估计')
    ws.row_dimensions[3].height = 20.0

    # Row 4: doc estimators
    _wmc(ws, 4, 1, 4, 2, '估计人员')
    _wmc(ws, 4, 3, 4, 4, '、'.join(e for e in d['doc_est'] if e))
    _wmc(ws, 4, 5, 4, 6, '估计方法')
    _wmc(ws, 4, 7, 4, 9, '专家估计法')

    # Row 5: doc assumption
    _wmc(ws, 5, 1, 5, 2, '估计假设和约束条件')
    _wmc(ws, 5, 3, 5, 9, _ASSUME_DOC_SCALE)
    ws.row_dimensions[5].height = 107.0

    # Row 6: 代码规模估计 section header
    _wmc(ws, 6, 1, 6, 9, '代码规模估计')
    ws.row_dimensions[6].height = 20.0

    # Row 7: code estimators
    _wmc(ws, 7, 1, 7, 2, '估计人员')
    _wmc(ws, 7, 3, 7, 4, '、'.join(e for e in d['code_est'] if e))
    _wmc(ws, 7, 5, 7, 6, '估计方法')
    _wmc(ws, 7, 7, 7, 9, '专家估计法')

    # Row 8: code assumption
    _wmc(ws, 8, 1, 8, 2, '估计假设和约束条件')
    _wmc(ws, 8, 3, 8, 9, _ASSUME_CODE_SCALE)
    ws.row_dimensions[8].height = 124.0

    # Row 9: 估计值列表 section header
    _wmc(ws, 9, 1, 9, 9, '估计值列表')
    ws.row_dimensions[9].height = 20.0

    # Rows 10-11: column headers
    _wmc(ws, 10, 1, 11, 1, '序号')
    _wmc(ws, 10, 2, 11, 4, '估计对象')
    _wmc(ws, 10, 5, 11, 5, '度量单位')
    _wmc(ws, 10, 6, 10, 8, '估计值')
    _wmc(ws, 10, 9, 11, 9, '备注')
    _w(ws, 11, 6, '总值')
    _w(ws, 11, 7, '新增规模')
    _w(ws, 11, 8, '重用规模')

    DOC_DATA_START = 12
    # Doc scale rows
    for i, row in enumerate(d['doc_rows']):
        r = DOC_DATA_START + i
        sn3_r = DATA1_START + i
        _w(ws, r, 1, i + 1)
        _wmc(ws, r, 2, r, 4, f"='{SN3}'!B{sn3_r}")
        _w(ws, r, 5, f"='{SN3}'!F{sn3_r}")
        _w(ws, r, 6, f"='{SN3}'!M{sn3_r}")
        sn1_r = sn3_r
        sn1_r3 = sn1_r + d['OFFSET1']
        _w(ws, r, 7, f"=ROUND(('{SN1}'!H{sn1_r}+'{SN1}'!U{sn1_r}+'{SN1}'!H{sn1_r3}+'{SN1}'!U{sn1_r3})/4,0)")
        _w(ws, r, 8, f"=ROUND(('{SN1}'!J{sn1_r}+'{SN1}'!W{sn1_r}+'{SN1}'!J{sn1_r3}+'{SN1}'!W{sn1_r3})/4,0)")

    doc_end = DOC_DATA_START - 1 + N_doc

    # Code scale rows
    for i, row in enumerate(d['s2_all']):
        r = doc_end + 2 + i
        sn4_r = DATA1_START + i
        _w(ws, r, 1, N_doc + 1 + i)
        _wmc(ws, r, 2, r, 4, f"='{SN4}'!B{sn4_r}")
        _w(ws, r, 5, '行')
        _w(ws, r, 6, f"='{SN4}'!M{sn4_r}")
        sn2_r = sn4_r
        sn2_r3 = sn2_r + d['OFFSET2']
        _w(ws, r, 7, f"=ROUND(('{SN2}'!H{sn2_r}+'{SN2}'!U{sn2_r}+'{SN2}'!H{sn2_r3}+'{SN2}'!U{sn2_r3})/4,0)")
        _w(ws, r, 8, f'=F{r}-G{r}')

    code_end = doc_end + 1 + N_code

    total_row = code_end + 2
    _wmc(ws, total_row, 1, total_row, 4, '文档规模')
    ws.row_dimensions[total_row].height = 21.0
    _w(ws, total_row, 5, '页')
    _w(ws, total_row, 6, f'=SUM(F{DOC_DATA_START}:F{doc_end})')
    _w(ws, total_row, 7, f'=SUM(G{DOC_DATA_START}:G{doc_end})')
    _w(ws, total_row, 8, f'=SUM(H{DOC_DATA_START}:H{doc_end})')

    total_row2 = total_row + 1
    _wmc(ws, total_row2, 1, total_row2, 4, '代码规模')
    _w(ws, total_row2, 5, '行')
    _w(ws, total_row2, 6, f'=SUM(F{doc_end+2}:F{code_end})')
    _w(ws, total_row2, 7, f'=SUM(G{doc_end+2}:G{code_end})')
    _w(ws, total_row2, 8, f'=SUM(H{doc_end+2}:H{code_end})')

    # 说明行
    note_row = total_row2 + 1
    _wmc(ws, note_row, 1, note_row, 9, _NOTE_9)
    ws.row_dimensions[note_row].height = 101.0

    _fmt_report(ws, max_col=9)
    _apply_title_style(ws, 1)


def _build_sheet10(ws, d: dict):
    """Workload results report."""
    project = d['project']
    proj_name = project.get('name', '')
    proj_code = f"ZDPX{project.get('code', '')}"
    N_doc = d['N_doc']
    N_wl1 = d['N_wl1']
    N_code = d['N_code']

    # Row 1: 标题行
    _wmc(ws, 1, 1, 1, 8, '工作量估计结果报告')
    ws.row_dimensions[1].height = 41.0

    # Row 2: 项目名称
    _w(ws, 2, 1, '项目名称')
    _wmc(ws, 2, 2, 2, 4, proj_name)
    _wmc(ws, 2, 5, 2, 6, '项目编号')
    _wmc(ws, 2, 7, 2, 8, proj_code)

    # Rows 3-8: assumption label spans A3:B8, content spans C:H
    _wmc(ws, 3, 1, 8, 2, '估计假设和约束条件')
    _wmc(ws, 3, 3, 3, 8, _ASSUME_DOC_WL)
    _wmc(ws, 4, 3, 4, 8, _ASSUME_CODE_WL)
    _wmc(ws, 5, 3, 5, 8, _ASSUME_CM)
    _wmc(ws, 6, 3, 6, 8, _ASSUME_QA)
    _wmc(ws, 7, 3, 7, 8, _ASSUME_MGMT)
    _wmc(ws, 8, 3, 8, 8, _ASSUME_REVIEW)
    ws.row_dimensions[3].height = 95.0
    ws.row_dimensions[4].height = 87.0
    ws.row_dimensions[5].height = 61.0
    ws.row_dimensions[6].height = 57.0
    ws.row_dimensions[7].height = 59.0
    ws.row_dimensions[8].height = 44.0

    # Row 9: column headers
    _w(ws, 9, 1, '活动类别')
    _wmc(ws, 9, 2, 9, 4, '估计对象')
    _wmc(ws, 9, 5, 9, 6, '估计值')
    _w(ws, 9, 7, '估计人员')
    _w(ws, 9, 8, '估计方法')

    est_str_doc = '、'.join(e for e in d['doc_est'] if e)
    est_str_code = '、'.join(e for e in d['code_est'] if e)

    # Engineering workload rows (all_wl1: WBS engineering tasks for Sheet1)
    eng_data_start = 10
    for i, row in enumerate(d['all_wl1']):
        r = eng_data_start + i
        sn3_r = DATA1_START + N_doc + 1 + i  # workload row in Sheet3 (+1 blank separator)
        _wmc(ws, r, 2, r, 4, f"='{SN3}'!B{sn3_r}")
        _wmc(ws, r, 5, r, 6, f"='{SN3}'!M{sn3_r}")
        _w(ws, r, 7, est_str_doc)
        _w(ws, r, 8, '专家法')

    # Code workload rows
    code_eng_start = eng_data_start + N_wl1
    for i, row in enumerate(d['s2_all']):
        r = code_eng_start + i
        sn4_r = DATA1_START + N_code + 1 + i  # workload row in Sheet4 (+1 blank separator)
        _wmc(ws, r, 2, r, 4, f"='{SN4}'!B{sn4_r}")
        _wmc(ws, r, 5, r, 6, f"='{SN4}'!M{sn4_r}")
        _w(ws, r, 7, est_str_code)
        _w(ws, r, 8, '专家法')

    eng_end = code_eng_start + N_code - 1
    eng_total_row = eng_end + 1

    # Merge '工程类' for A column across all eng data rows (not including total)
    if eng_data_start <= eng_end:
        _wmc(ws, eng_data_start, 1, eng_end, 1, '工程类')
    elif eng_data_start == eng_total_row:
        _w(ws, eng_data_start, 1, '工程类')

    _wmc(ws, eng_total_row, 2, eng_total_row, 4, '合计')
    _wmc(ws, eng_total_row, 5, eng_total_row, 6, f'=SUM(E{eng_data_start}:E{eng_total_row-1})')

    # CM / QA / review / mgmt summary rows
    cm_row = eng_total_row + 1
    qa_row = cm_row + 1
    review_row = qa_row + 1
    mgmt_row = review_row + 1

    cm_total = d.get('_cm_total_row', 6 + len(d['cm_rows']) + 1)
    qa_total = d.get('_qa_total_row', 6 + len(d['qa_rows']) + 1)
    review_total = d.get('_review_total_row', 6 + len(d['rev_rows']) + 1)
    mgmt_total = d.get('_mgmt_total_row', 6 + len(d['mgmt_rows']) + 1)

    _wmc(ws, cm_row, 1, cm_row, 4, '配置管理类')
    _wmc(ws, cm_row, 5, cm_row, 6, f"='{SN5}'!D{cm_total}")
    _w(ws, cm_row, 7, d['cm_est'])
    _w(ws, cm_row, 8, '经验法')

    _wmc(ws, qa_row, 1, qa_row, 4, '质量保证类')
    _wmc(ws, qa_row, 5, qa_row, 6, f"='{SN6}'!D{qa_total}")
    _w(ws, qa_row, 7, d['qa_est'])
    _w(ws, qa_row, 8, '经验法')

    _wmc(ws, review_row, 1, review_row, 4, '评审类')
    _wmc(ws, review_row, 5, review_row, 6, f"='{SN8}'!E{review_total}")
    _w(ws, review_row, 7, d['review_est'])
    _w(ws, review_row, 8, '经验法')

    _wmc(ws, mgmt_row, 1, mgmt_row, 4, '其他管理类')
    _wmc(ws, mgmt_row, 5, mgmt_row, 6, f"='{SN7}'!C{mgmt_total}")
    _w(ws, mgmt_row, 7, d['mgmt_est'])
    _w(ws, mgmt_row, 8, '经验法')

    total_row = mgmt_row + 1
    _wmc(ws, total_row, 1, total_row, 8, '总计')
    ws.row_dimensions[total_row].height = 20.0
    total_row2 = total_row + 1
    _wmc(ws, total_row2, 1, total_row2, 2, '工作量')
    _wmc(ws, total_row2, 3, total_row2, 8, f'=E{eng_total_row}+E{cm_row}+E{qa_row}+E{review_row}+E{mgmt_row}')

    # 说明行
    note_row = total_row2 + 1
    _wmc(ws, note_row, 1, note_row, 8, _NOTE_10)
    ws.row_dimensions[note_row].height = 83.0

    _fmt_report(ws, max_col=8)
    _apply_title_style(ws, 1)


async def build_excel(db: aiosqlite.Connection, session_id: int) -> bytes:
    d = await _load_data(db, session_id)

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet(SN1)
    ws2 = wb.create_sheet(SN2)
    ws3 = wb.create_sheet(SN3)
    ws4 = wb.create_sheet(SN4)
    ws5 = wb.create_sheet(SN5)
    ws6 = wb.create_sheet(SN6)
    ws7 = wb.create_sheet(SN7)
    ws8 = wb.create_sheet(SN8)
    ws9 = wb.create_sheet(SN9)
    ws10 = wb.create_sheet(SN10)

    _build_sheet1(ws1, d)
    _build_sheet2(ws2, d)
    _build_sheet3(ws3, d)
    _build_sheet4(ws4, d)
    _build_sheet5(ws5, d)
    _build_sheet6(ws6, d)
    _build_sheet7(ws7, d)
    _build_sheet8(ws8, d)
    # _cm_total_row / _qa_total_row / _mgmt_total_row / _review_total_row
    # are now set inside the respective builders above; no recalculation needed.

    _build_sheet9(ws9, d)
    _build_sheet10(ws10, d)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
