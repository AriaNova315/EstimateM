"""生成 excel_layout.md：Excel builder 所需的完整列坐标参考"""
import openpyxl
from openpyxl.utils import get_column_letter

SAMPLE = "/Users/gaoyiman/Desktop/Code/5000B/EstimateM/规模及工作量估计表.xlsx"
OUT    = "/Users/gaoyiman/Desktop/Code/5000B/EstimateM/excel_layout.md"

wb = openpyxl.load_workbook(SAMPLE, data_only=False)

lines = []
def p(*args): lines.append(" ".join(str(a) for a in args))

def row_summary(ws, rn):
    vals = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(rn, c).value
        if v is not None:
            vals[get_column_letter(c)] = repr(v)[:60]
    return vals

p("# EstimateM Excel 布局参考（供 excel_builder.py 使用）")
p()
p("从样例文件 `规模及工作量估计表.xlsx` 提取，行列坐标均为 1-indexed。")
p()

# Sheet 名称
p("## Sheet 名称映射")
p()
for i, n in enumerate(wb.sheetnames, 1):
    p(f"- 样例 Sheet {i}: `{n}`")
p()
p("生成文件使用的 Sheet 名称（去掉 ok 后缀）：")
clean = [
    "文档规模及工作量估计",
    "代码规模及工作量估计",
    "汇总表(文档相关)",
    "汇总表(代码相关)",
    "配置管理工作量估计",
    "质量保证工作量估计",
    "其他管理类工作量估计",
    "评审工作量",
    "规模估计结果报告",
    "工作量估计结果报告",
]
for i, n in enumerate(clean, 1):
    p(f"- Sheet {i}: `{n}`")

p()
p("---")
p()
p("## Sheet 1 / Sheet 2：4 人估算表布局")
p()
p("### 整体结构（以 Sheet 1 文档为例，Sheet 2 代码相同）")
p()
ws1 = wb[wb.sheetnames[0]]
ws2 = wb[wb.sheetnames[1]]
p(f"样例 Sheet1: {ws1.max_row} 行 x {ws1.max_column} 列")
p(f"样例 Sheet2: {ws2.max_row} 行 x {ws2.max_column} 列")
p()
p("每张 Sheet 分为上下两块，每块横向并排 2 个估算人：")
p()
p("```")
p("上块（行 1 ~ 行 35 for N_item=28）")
p("  左半 A-L: 估算人 1")
p("  右半 M-X: 估算人 2")
p("分隔区（7 行固定，说明文字）")
p("下块（行 43 ~ 行 77 for N_item=28）")
p("  左半 A-L: 估算人 3")
p("  右半 M-X: 估算人 4")
p("尾部（7 行固定，说明文字）")
p("```")
p()
p("**行号计算公式（设 N_item = 总数据行数 = 文档数 + 工程活动数）：**")
p()
p("```python")
p("HEADER_ROWS = 7               # 每块头部行数")
p("SEP_ROWS    = 7               # 块间/块后说明行数")
p("DATA1_START = HEADER_ROWS + 1  # = 8：上块数据起始行")
p("DATA1_END   = DATA1_START + N_item - 1")
p("SEP1_START  = DATA1_END + 1")
p("SEP1_END    = SEP1_START + SEP_ROWS - 1")
p("BLOCK2_START= SEP1_END + 1   # 下块第一行（= 1 + HEADER_ROWS + N_item + SEP_ROWS）")
p("DATA2_START = BLOCK2_START + HEADER_ROWS  # 下块数据起始行")
p("DATA2_END   = DATA2_START + N_item - 1")
p("# offset 即 DATA2_START - DATA1_START = 1 + N_item + SEP_ROWS + HEADER_ROWS - 1")
p("BLOCK2_OFFSET = N_item + SEP_ROWS + HEADER_ROWS  # 样例 28+7+7=42")
p("```")
p()
p("### 列映射（左半 A-L；右半 M-X 结构相同，列号 +12）")
p()
p("| 字段 | 左半列 | 右半列 | 备注 |")
p("|------|--------|--------|------|")
p("| 序号 | A(1) | M(13) | 整数 |")
p("| 估计对象 | B-E(2-5)合并 | N-Q(14-17)合并 | 文档/模块名称 |")
p("| 度量单位 | F(6) | R(18) | 页/行/人时 |")
p("| 总值 | G(7) | S(19) | 规模=`=H+J`；工作量=直接值 |")
p("| 新增规模/值 | H-I(8-9)合并 | T-U(20-21)合并 | 规模行=数值；工作量行='—' |")
p("| 重用规模 | J-K(10-11)合并 | V-W(22-23)合并 | 规模行=0；工作量行='—' |")
p("| 备注 | L(12) | X(24) | 通常空 |")
p()
p("### 头部 7 行内容（相对于块起始行）")
p()
p("| 偏移 | 左半关键内容 | 右半关键内容 |")
p("|------|-------------|-------------|")
p("| +0 | A='项目名称'(A:B合并), C:F=项目名称值, G:H='项目编号', I:L=编号值 | M='项目名称'(M:N合并), O:R=值, S:T='项目编号', U:X=值 |")
p("| +1 | A='估计人', B:C=姓名, D:E='估计类别', F:H='■规模  ■工作量', I:J='估计轮次', K:L=轮次值 | M='估计人', N:O=姓名, P:Q='估计类别', R:T=同左, U:V='估计轮次', W:X=轮次值 |")
p("| +2 | A:C='估计方法', D:L='专家估计法' | M:O='估计方法', P:X='专家估计法' |")
p("| +3 | A:C='估计假设和约束条件'(合并+3/+4), D:L=文本 | M:O=同左, P:X=文本 |")
p("| +4 | (延续+3合并) | (延续) |")
p("| +5 | A:A='序号'(合并+5/+6), B:E='估计对象'(合并+5/+6), F:F='度量\\n单位'(合并+5/+6), G:K='估计值'(合并+5) | M=同A, N:Q=同B:E, R=同F, S:W='估计值' |")
p("| +6 | G='总值', H:I='新增规模', J:K='重用规模' | S='总值', T:U='新增规模', V:W='重用规模' |")
p()
p("### 数据行公式（Sheet 1 规模行）")
p()
p("```python")
p("# 对第 i 个文档（0-indexed），数据行号 r = DATA1_START + i")
p("# 左半（估算人 1）")
p("ws.cell(r, 1).value  = i + 1        # 序号")
p("merge_cells(ws, r, 2, r, 5)         # B:E 合并")
p("ws.cell(r, 2).value  = doc_name")
p("ws.cell(r, 6).value  = '页'")
p("ws.cell(r, 7).value  = f'=H{r}+J{r}'   # 总值")
p("merge_cells(ws, r, 8, r, 9)         # H:I 合并")
p("ws.cell(r, 8).value  = v1_pages     # 估算人1 的页数")
p("merge_cells(ws, r, 10, r, 11)       # J:K 合并")
p("ws.cell(r, 10).value = 0            # 重用规模")
p("# 右半（估算人 2）列 +12")
p("ws.cell(r, 13).value = i + 1")
p("merge_cells(ws, r, 14, r, 17)")
p("ws.cell(r, 14).value = doc_name")
p("ws.cell(r, 18).value = '页'")
p("ws.cell(r, 19).value = f'=T{r}+V{r}'")
p("merge_cells(ws, r, 20, r, 21)")
p("ws.cell(r, 20).value = v2_pages")
p("merge_cells(ws, r, 22, r, 23)")
p("ws.cell(r, 22).value = 0")
p("```")
p()
p("### 数据行公式（工作量行，规模字段填 '—'）")
p()
p("```python")
p("ws.cell(r, 7).value  = wl1_hours   # 直接值")
p("ws.cell(r, 8).value  = '—'         # H:I 合并")
p("ws.cell(r, 10).value = '—'         # J:K 合并")
p("ws.cell(r, 19).value = wl2_hours")
p("ws.cell(r, 20).value = '—'         # T:U 合并")
p("ws.cell(r, 22).value = '—'         # V:W 合并")
p("```")
p()
p("### Sheet 2 特殊说明")
p()
p("- 规模度量单位 = '行'（LOC）")
p("- 序号延续文档序号（样例中代码模块从 24 开始，因为有 23 个文档配置项）")
p("- 样例工作量行使用 `=ROUNDUP(G{scale_row}/divisor,0)` 公式；**生成时改用直接值**（工作量已由估算算法独立生成）")
p("- Sheet 2 样例数据行数 = 4 个代码模块 + 1 个 bug 修改 = 5 行（上块共 16 行）")
p()

p("---")
p()
p("## Sheet 3 / Sheet 4：汇总表")
p()
ws3 = wb[wb.sheetnames[2]]
p(f"样例 Sheet3: {ws3.max_row} 行 x {ws3.max_column} 列")
p()
p("### 头部（行 1-7，固定）")
p()
for rn in range(1, 8):
    sv = row_summary(ws3, rn)
    if sv:
        p(f"行{rn}: " + "  ".join(f"{k}={v}" for k, v in sv.items()))
p()
p("### 数据行公式模板")
p()
p("设 SN1='文档规模及工作量估计'，汇总表数据行 r（从 8 开始），")
p("对应 Sheet1 上块行 r1 = r，下块行 r3 = r + BLOCK2_OFFSET：")
p()
for rn in range(8, 10):
    sv = row_summary(ws3, rn)
    if sv:
        p(f"行{rn}样例: " + "  ".join(f"{k}={v}" for k, v in sv.items()))
p()
p("```python")
p("SN1 = '文档规模及工作量估计'")
p("# r   = 汇总表数据行（从 DATA_START=8 开始）")
p("# r1  = Sheet1 上块对应行 = DATA1_START + (r - 8)")
p("# r3  = r1 + BLOCK2_OFFSET（= r1 + N_item + 14）")
p()
p("# 列 B（估计对象名，引用 Sheet1）")
p("B{r} = f\"='{SN1}'!B{r1}\"")
p()
p("# 列 H:I（平均值）- 合并 H:I")
p("H{r} = f\"=('{SN1}'!G{r1}+'{SN1}'!S{r1}+'{SN1}'!G{r3}+'{SN1}'!S{r3})/4\"")
p()
p("# 列 J:K（最大值）- 合并 J:K")
p("J{r} = f\"=MAX('{SN1}'!G{r1},'{SN1}'!S{r1},'{SN1}'!G{r3},'{SN1}'!S{r3})\"")
p()
p("# 列 L:M（最小值）- 合并 L:M（样例中 K 列）")
p("K{r} = f\"=MIN('{SN1}'!G{r1},'{SN1}'!S{r1},'{SN1}'!G{r3},'{SN1}'!S{r3})\"")
p()
p("# 列 L（偏差率）")
p("L{r} = f\"=MAX(ABS(J{r}-H{r}),ABS(K{r}-H{r}))/H{r}\"  # 格式设为百分比")
p()
p("# 列 M（取整均值）")
p("M{r} = f\"=ROUND(H{r},0)\"")
p("```")
p()
p("### 汇总表下半部分：工作量与进度安排表（行 43+）")
p()
p("固定标题行：'工作量与进度安排表'")
p("列映射：A=序号, B:E=估计对象(引用上方M列对应行的B), F:G=工作量估计值(引用M列),")
p("        H:J=计划分配资源(人名), K:L=每日可用工作时间(人数×Tm), M=任务所需工作日")
p()
for rn in range(43, 60):
    sv = row_summary(ws3, rn)
    if sv:
        p(f"行{rn}: " + "  ".join(f"{k}={v}" for k, v in sv.items()))

p()
p("---")
p()
p("## Sheets 5-8：CM / QA / 管理 / 评审（单人实际值）")
p()
for si in range(4, 8):
    sn = wb.sheetnames[si]
    ws_x = wb[sn]
    p(f"### Sheet {si+1}：{sn}（{ws_x.max_row} 行 x {ws_x.max_column} 列）")
    p()
    for rn in range(1, min(ws_x.max_row + 1, 50)):
        sv = row_summary(ws_x, rn)
        if sv:
            p(f"  行{rn:3d}: " + "  ".join(f"{k}={v}" for k, v in sv.items()))
    p()

p("---")
p()
p("## Sheets 9-10：规模/工作量估计结果报告")
p()
for si in range(8, 10):
    sn = wb.sheetnames[si]
    ws_x = wb[sn]
    p(f"### Sheet {si+1}：{sn}（{ws_x.max_row} 行 x {ws_x.max_column} 列）")
    p()
    for rn in range(1, min(ws_x.max_row + 1, 80)):
        sv = row_summary(ws_x, rn)
        if sv:
            p(f"  行{rn:3d}: " + "  ".join(f"{k}={v}" for k, v in sv.items()))
    p()

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print(f"Done: {len(lines)} lines -> {OUT}")
