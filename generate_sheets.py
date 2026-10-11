import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import DoughnutChart, Reference

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color palettes
NAVY_HEADER = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
SUB_HEADER = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 연한 노랑 (입력 셀)
ACCENT_GREEN = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
GRAY_BG = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

WHITE_BOLD = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="맑은 고딕", size=14, bold=True, color="1F4E79")
SECTION_FONT = Font(name="맑은 고딕", size=11, bold=True, color="1F4E79")
BOLD_FONT = Font(name="맑은 고딕", size=10, bold=True)
REGULAR_FONT = Font(name="맑은 고딕", size=10)
SMALL_FONT = Font(name="맑은 고딕", size=9, color="595959")

THIN_BORDER_SIDE = Side(border_style="thin", color="D9D9D9")
BORDER_ALL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_TOP_BOTTOM = Border(top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)

# -------------------------------------------------------------------------
# 1. 대시보드
# -------------------------------------------------------------------------
ws_dash = wb.create_sheet("대시보드")
ws_dash.views.sheetView[0].showGridLines = True

ws_dash["A1"] = "개인 투자 포트폴리오 트래커 - 대시보드"
ws_dash["A1"].font = TITLE_FONT

ws_dash["A3"] = "기준 통화"
ws_dash["B3"] = "=설정!B2"
ws_dash["D3"] = "마지막 갱신 시각"
ws_dash["E3"] = "=설정!B8"
ws_dash["A3"].font = BOLD_FONT
ws_dash["D3"].font = BOLD_FONT

# Key metrics table
ws_dash["A5"] = "■ 핵심 투자 요약 지표"
ws_dash["A5"].font = SECTION_FONT

metrics = [
    ("누적 투자 원금", "=SUM(포트폴리오!G4:G100)", "₩#,##0"),
    ("현재 총자산", "=C8+C9", "₩#,##0"),
    ("투자 자산 평가액", "=SUM(포트폴리오!J4:J100)", "₩#,##0"),
    ("현금성 자산", "=SUM('계좌 및 현금'!F4:F50)", "₩#,##0"),
    ("평가 손익", "=C8-C6", "₩#,##0"),
    ("수익률", '=IF(C6>0, C10/C6, "-")', "0.00%"),
    ("1일 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=2, C7-INDEX(\'자산 기록\'!D4:D100, COUNTA(\'자산 기록\'!A4:A100)-1), "기록 부족")', "₩#,##0"),
    ("1주 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=5, C7-INDEX(\'자산 기록\'!D4:D100, MAX(1, COUNTA(\'자산 기록\'!A4:A100)-5)), "기록 부족")', "₩#,##0"),
    ("1개월 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=20, C7-INDEX(\'자산 기록\'!D4:D100, MAX(1, COUNTA(\'자산 기록\'!A4:A100)-20)), "기록 부족")', "₩#,##0"),
]

ws_dash["B5"] = "항목"
ws_dash["C5"] = "금액 / 수치"
ws_dash["B5"].font = WHITE_BOLD
ws_dash["C5"].font = WHITE_BOLD
ws_dash["B5"].fill = NAVY_HEADER
ws_dash["C5"].fill = NAVY_HEADER
ws_dash["B5"].alignment = Alignment(horizontal="center")
ws_dash["C5"].alignment = Alignment(horizontal="center")

for idx, (label, formula, num_format) in enumerate(metrics, start=6):
    ws_dash[f"B{idx}"] = label
    ws_dash[f"C{idx}"] = formula
    ws_dash[f"B{idx}"].font = BOLD_FONT
    ws_dash[f"C{idx}"].font = BOLD_FONT
    ws_dash[f"B{idx}"].border = BORDER_ALL
    ws_dash[f"C{idx}"].border = BORDER_ALL
    ws_dash[f"B{idx}"].fill = GRAY_BG
    ws_dash[f"C{idx}"].number_format = num_format
    ws_dash[f"C{idx}"].alignment = Alignment(horizontal="right")

# Account summary table
ws_dash["E5"] = "■ 계좌별 자산 요약"
ws_dash["E5"].font = SECTION_FONT

account_headers = ["계좌 구분", "투자 자산", "현금 잔액", "합계"]
for c_idx, h in enumerate(account_headers, start=5):
    cell = ws_dash.cell(row=6, column=c_idx, value=h)
    cell.font = WHITE_BOLD
    cell.fill = SUB_HEADER
    cell.alignment = Alignment(horizontal="center")
    cell.border = BORDER_ALL

accounts = ["ISA", "금현물계좌", "CMA", "해외계좌", "일반/현금"]
for r_idx, acc in enumerate(accounts, start=7):
    ws_dash.cell(row=r_idx, column=5, value=acc).font = BOLD_FONT
    ws_dash.cell(row=r_idx, column=5).border = BORDER_ALL
    ws_dash.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
    
    # 투자 자산 SUMIF
    cell_inv = ws_dash.cell(row=r_idx, column=6, value=f'=SUMIF(포트폴리오!$A$4:$A$100, "{acc}", 포트폴리오!$J$4:$J$100)')
    cell_inv.font = REGULAR_FONT
    cell_inv.number_format = "₩#,##0"
    cell_inv.border = BORDER_ALL
    cell_inv.alignment = Alignment(horizontal="right")
    
    # 현금 잔액 SUMIF
    cell_cash = ws_dash.cell(row=r_idx, column=7, value=f'=SUMIF(\'계좌 및 현금\'!$A$4:$A$50, "{acc}", \'계좌 및 현금\'!$F$4:$F$50)')
    cell_cash.font = REGULAR_FONT
    cell_cash.number_format = "₩#,##0"
    cell_cash.border = BORDER_ALL
    cell_cash.alignment = Alignment(horizontal="right")
    
    # 합계
    cell_tot = ws_dash.cell(row=r_idx, column=8, value=f'=F{r_idx}+G{r_idx}')
    cell_tot.font = BOLD_FONT
    cell_tot.number_format = "₩#,##0"
    cell_tot.border = BORDER_ALL
    cell_tot.alignment = Alignment(horizontal="right")

# Account summary total row
ws_dash.cell(row=12, column=5, value="합계").font = WHITE_BOLD
ws_dash.cell(row=12, column=5).fill = NAVY_HEADER
ws_dash.cell(row=12, column=5).alignment = Alignment(horizontal="center")
ws_dash.cell(row=12, column=5).border = BORDER_ALL

for c_idx, f_sum in [(6, "=SUM(F7:F11)"), (7, "=SUM(G7:G11)"), (8, "=SUM(H7:H11)")]:
    c = ws_dash.cell(row=12, column=c_idx, value=f_sum)
    c.font = WHITE_BOLD
    c.fill = NAVY_HEADER
    c.number_format = "₩#,##0"
    c.alignment = Alignment(horizontal="right")
    c.border = BORDER_ALL

# -------------------------------------------------------------------------
# 2. 포트폴리오
# -------------------------------------------------------------------------
ws_port = wb.create_sheet("포트폴리오")
ws_port.views.sheetView[0].showGridLines = True

ws_port["A1"] = "보유 포트폴리오 관리"
ws_port["A1"].font = TITLE_FONT

ws_port["A2"] = "목표 비중 합계:"
ws_port["B2"] = "=SUM(N4:N100)"
ws_port["B2"].number_format = "0.0%"
ws_port["B2"].font = BOLD_FONT

ws_port["C2"] = "미배분 비중:"
ws_port["D2"] = "=1-B2"
ws_port["D2"].number_format = "0.0%"
ws_port["D2"].font = BOLD_FONT

ws_port["E2"] = "상태 경고:"
ws_port["F2"] = '=IF(B2>1.0001, "⚠️ 목표 비중 100% 초과!", IF(B2<0.9999, "⚠️ 목표 비중 미달 (정규화 배분)", "✅ 정상 (100%)"))'
ws_port["F2"].font = BOLD_FONT

port_headers = [
    "계좌", "티커/코드", "종목명", "자산 유형", "통화", 
    "보유 수량", "누적 매입 금액", "현재가", "기준 환율", 
    "평가액(KRW)", "평가 손익", "수익률", "실제 비중", 
    "목표 비중", "비중 차이", "조회 상태/갱신"
]

for col_idx, h in enumerate(port_headers, start=1):
    c = ws_port.cell(row=3, column=col_idx, value=h)
    c.font = WHITE_BOLD
    c.fill = NAVY_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER_ALL

# Sample mock rows (ISA, 해외, 금현물 등)
sample_port_data = [
    ("ISA", "005930", "삼성전자", "주식", "KRW", 50, 3500000, 72000, 1, 0.20, "수동 입력/구글연동"),
    ("ISA", "379800", "TIGER 미국S&P500", "ETF", "KRW", 100, 1400000, 16500, 1, 0.25, "수동 입력/구글연동"),
    ("해외계좌", "QQQ", "Invesco QQQ Trust", "ETF", "USD", 10, 5200000, 480, 1380, 0.30, "수동 입력/구글연동"),
    ("해외계좌", "NVDA", "NVIDIA Corp", "주식", "USD", 15, 2500000, 130, 1380, 0.15, "수동 입력/구글연동"),
    ("금현물계좌", "금현물", "KRX 금 1g", "금", "KRW", 20, 1900000, 105000, 1, 0.10, "수동 입력"),
]

for row_idx, data in enumerate(sample_port_data, start=4):
    acc, ticker, name, atype, curr, qty, buy_amt, price, fx, target_w, status = data
    ws_port[f"A{row_idx}"] = acc
    ws_port[f"B{row_idx}"] = ticker
    ws_port[f"C{row_idx}"] = name
    ws_port[f"D{row_idx}"] = atype
    ws_port[f"E{row_idx}"] = curr
    ws_port[f"F{row_idx}"] = qty
    ws_port[f"G{row_idx}"] = buy_amt
    ws_port[f"H{row_idx}"] = price
    ws_port[f"I{row_idx}"] = fx
    
    # Formulas
    ws_port[f"J{row_idx}"] = f'=IF(B{row_idx}="","",F{row_idx}*H{row_idx}*I{row_idx})' # 평가액
    ws_port[f"K{row_idx}"] = f'=IF(B{row_idx}="","",J{row_idx}-G{row_idx})'          # 손익
    ws_port[f"L{row_idx}"] = f'=IF(G{row_idx}>0,K{row_idx}/G{row_idx},"-")'          # 수익률
    ws_port[f"M{row_idx}"] = f'=IF(SUM($J$4:$J$100)>0,J{row_idx}/SUM($J$4:$J$100),0)' # 실제 비중
    ws_port[f"N{row_idx}"] = target_w                                               # 목표 비중
    ws_port[f"O{row_idx}"] = f'=IF(N{row_idx}="","",M{row_idx}-N{row_idx})'          # 비중 차이
    ws_port[f"P{row_idx}"] = status

    # Styles
    for col_c in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "N"]:
        ws_port[f"{col_c}{row_idx}"].fill = INPUT_FILL
        ws_port[f"{col_c}{row_idx}"].border = BORDER_ALL
        ws_port[f"{col_c}{row_idx}"].font = REGULAR_FONT
    
    for col_c in ["J", "K", "L", "M", "O", "P"]:
        ws_port[f"{col_c}{row_idx}"].border = BORDER_ALL
        ws_port[f"{col_c}{row_idx}"].font = BOLD_FONT

    # Alignments & formats
    ws_port[f"A{row_idx}"].alignment = Alignment(horizontal="center")
    ws_port[f"B{row_idx}"].alignment = Alignment(horizontal="center")
    ws_port[f"D{row_idx}"].alignment = Alignment(horizontal="center")
    ws_port[f"E{row_idx}"].alignment = Alignment(horizontal="center")
    ws_port[f"F{row_idx}"].number_format = "#,##0"
    ws_port[f"G{row_idx}"].number_format = "₩#,##0"
    ws_port[f"H{row_idx}"].number_format = "#,##0.00"
    ws_port[f"I{row_idx}"].number_format = "#,##0.00"
    ws_port[f"J{row_idx}"].number_format = "₩#,##0"
    ws_port[f"K{row_idx}"].number_format = "₩#,##0"
    ws_port[f"L{row_idx}"].number_format = "0.00%"
    ws_port[f"M{row_idx}"].number_format = "0.0%"
    ws_port[f"N{row_idx}"].number_format = "0.0%"
    ws_port[f"O{row_idx}"].number_format = "+0.0%;-0.0%;0.0%"
    ws_port[f"P{row_idx}"].alignment = Alignment(horizontal="center")

# Prepare 15 blank styled rows for user inputs
for row_idx in range(9, 25):
    ws_port[f"A{row_idx}"] = ""
    ws_port[f"B{row_idx}"] = ""
    ws_port[f"J{row_idx}"] = f'=IF(B{row_idx}="","",F{row_idx}*H{row_idx}*I{row_idx})'
    ws_port[f"K{row_idx}"] = f'=IF(B{row_idx}="","",J{row_idx}-G{row_idx})'
    ws_port[f"L{row_idx}"] = f'=IF(G{row_idx}>0,K{row_idx}/G{row_idx},"-")'
    ws_port[f"M{row_idx}"] = f'=IF(SUM($J$4:$J$100)>0,J{row_idx}/SUM($J$4:$J$100),0)'
    ws_port[f"O{row_idx}"] = f'=IF(N{row_idx}="","",M{row_idx}-N{row_idx})'
    
    for c_letter in ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P"]:
        cell = ws_port[f"{c_letter}{row_idx}"]
        cell.border = BORDER_ALL
        cell.font = REGULAR_FONT
        if c_letter in ["A","B","C","D","E","F","G","H","I","N"]:
            cell.fill = INPUT_FILL
    
    ws_port[f"F{row_idx}"].number_format = "#,##0"
    ws_port[f"G{row_idx}"].number_format = "₩#,##0"
    ws_port[f"H{row_idx}"].number_format = "#,##0.00"
    ws_port[f"I{row_idx}"].number_format = "#,##0.00"
    ws_port[f"J{row_idx}"].number_format = "₩#,##0"
    ws_port[f"K{row_idx}"].number_format = "₩#,##0"
    ws_port[f"L{row_idx}"].number_format = "0.00%"
    ws_port[f"M{row_idx}"].number_format = "0.0%"
    ws_port[f"N{row_idx}"].number_format = "0.0%"
    ws_port[f"O{row_idx}"].number_format = "+0.0%;-0.0%;0.0%"

# -------------------------------------------------------------------------
# 3. 월 투자 배분
# -------------------------------------------------------------------------
ws_alloc = wb.create_sheet("월 투자 배분")
ws_alloc.views.sheetView[0].showGridLines = True

ws_alloc["A1"] = "월 적립식 투자 배분안 계산기"
ws_alloc["A1"].font = TITLE_FONT

# Parameters
alloc_inputs = [
    ("A3", "이번 달 투자 가능 금액", "C3", 500000, "₩#,##0"),
    ("A4", "현금 유보 금액", "C4", 0, "₩#,##0"),
    ("A5", "실제 배분 대상 금액", "C5", "=C3-C4", "₩#,##0"),
    ("A6", "배분 후 예상 총자산", "C6", "=SUM(포트폴리오!$J$4:$J$100)+C5", "₩#,##0"),
]

for label_pos, label, val_pos, val, num_fmt in alloc_inputs:
    ws_alloc[label_pos] = label
    ws_alloc[label_pos].font = BOLD_FONT
    ws_alloc[val_pos] = val
    ws_alloc[val_pos].font = BOLD_FONT
    ws_alloc[val_pos].number_format = num_fmt
    ws_alloc[val_pos].border = BORDER_ALL
    if label_pos in ["A3", "A4"]:
        ws_alloc[val_pos].fill = INPUT_FILL
    else:
        ws_alloc[val_pos].fill = ACCENT_GREEN

ws_alloc["A7"] = "※ 목표 비중 대비 현재 부족한 자산을 기준으로 월 투자금을 최적 배분합니다."
ws_alloc["A7"].font = SMALL_FONT

alloc_headers = [
    "계좌", "티커", "종목명", "현재 평가액", "현재 비중",
    "목표 비중", "정규화 비중", "목표 평가액(투자후)", "목표 대비 부족액",
    "권장 투자 금액", "예상 투자후 비중", "비고"
]

for c_idx, h in enumerate(alloc_headers, start=1):
    c = ws_alloc.cell(row=9, column=c_idx, value=h)
    c.font = WHITE_BOLD
    c.fill = NAVY_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER_ALL

# Allocation rows referencing 포트폴리오
for i in range(4, 9):
    r_target = 10 + (i - 4)
    ws_alloc[f"A{r_target}"] = f'=포트폴리오!A{i}'
    ws_alloc[f"B{r_target}"] = f'=포트폴리오!B{i}'
    ws_alloc[f"C{r_target}"] = f'=포트폴리오!C{i}'
    ws_alloc[f"D{r_target}"] = f'=포트폴리오!J{i}' # 현재 평가액
    ws_alloc[f"E{r_target}"] = f'=포트폴리오!M{i}' # 현재 비중
    ws_alloc[f"F{r_target}"] = f'=포트폴리오!N{i}' # 목표 비중
    
    # 정규화 비중 (목표 비중 합계로 나눔)
    ws_alloc[f"G{r_target}"] = f'=IF(SUM($F$10:$F$20)>0, F{r_target}/SUM($F$10:$F$20), 0)'
    # 목표 평가액(투자 후) = 정규화 비중 * 배분 후 예상 총자산 ($C$6)
    ws_alloc[f"H{r_target}"] = f'=G{r_target}*$C$6'
    # 목표 대비 부족액 = MAX(0, 목표 평가액 - 현재 평가액)
    ws_alloc[f"I{r_target}"] = f'=MAX(0, H{r_target}-D{r_target})'
    # 권장 투자 금액 = 부족액 비율로 실제 배분 대상 금액 ($C$5) 분배
    ws_alloc[f"J{r_target}"] = f'=IF(SUM($I$10:$I$20)>0, ROUND($C$5*(I{r_target}/SUM($I$10:$I$20)), -2), 0)'
    # 예상 투자후 비중 = (현재 평가액 + 권장 투자액) / 배분 후 총자산 ($C$6)
    ws_alloc[f"K{r_target}"] = f'=(D{r_target}+J{r_target})/$C$6'
    # 비고
    ws_alloc[f"L{r_target}"] = f'=IF(H{r_target}<=D{r_target}, "목표 초과 (매수 불필요)", IF(J{r_target}>0, "권장 매수", "-"))'

    for c_letter in ["A","B","C","D","E","F","G","H","I","J","K","L"]:
        cell = ws_alloc[f"{c_letter}{r_target}"]
        cell.border = BORDER_ALL
        cell.font = REGULAR_FONT
    
    ws_alloc[f"A{r_target}"].alignment = Alignment(horizontal="center")
    ws_alloc[f"B{r_target}"].alignment = Alignment(horizontal="center")
    ws_alloc[f"D{r_target}"].number_format = "₩#,##0"
    ws_alloc[f"E{r_target}"].number_format = "0.0%"
    ws_alloc[f"F{r_target}"].number_format = "0.0%"
    ws_alloc[f"G{r_target}"].number_format = "0.0%"
    ws_alloc[f"H{r_target}"].number_format = "₩#,##0"
    ws_alloc[f"I{r_target}"].number_format = "₩#,##0"
    ws_alloc[f"J{r_target}"].number_format = "₩#,##0"
    ws_alloc[f"J{r_target}"].font = BOLD_FONT
    ws_alloc[f"K{r_target}"].number_format = "0.0%"
    ws_alloc[f"L{r_target}"].alignment = Alignment(horizontal="center")

# Total row for allocation
tot_r = 15
ws_alloc[f"C{tot_r}"] = "합계"
ws_alloc[f"C{tot_r}"].font = WHITE_BOLD
ws_alloc[f"C{tot_r}"].fill = SUB_HEADER
ws_alloc[f"C{tot_r}"].alignment = Alignment(horizontal="center")
ws_alloc[f"C{tot_r}"].border = BORDER_ALL

ws_alloc[f"D{tot_r}"] = "=SUM(D10:D14)"
ws_alloc[f"D{tot_r}"].number_format = "₩#,##0"
ws_alloc[f"D{tot_r}"].font = WHITE_BOLD
ws_alloc[f"D{tot_r}"].fill = SUB_HEADER
ws_alloc[f"D{tot_r}"].border = BORDER_ALL

ws_alloc[f"J{tot_r}"] = "=SUM(J10:J14)"
ws_alloc[f"J{tot_r}"].number_format = "₩#,##0"
ws_alloc[f"J{tot_r}"].font = WHITE_BOLD
ws_alloc[f"J{tot_r}"].fill = SUB_HEADER
ws_alloc[f"J{tot_r}"].border = BORDER_ALL

ws_alloc[f"L{tot_r}"] = '=IF(J15<=C5, "✅ 배분 한도 준수", "⚠️ 투자금 초과")'
ws_alloc[f"L{tot_r}"].font = WHITE_BOLD
ws_alloc[f"L{tot_r}"].fill = SUB_HEADER
ws_alloc[f"L{tot_r}"].alignment = Alignment(horizontal="center")
ws_alloc[f"L{tot_r}"].border = BORDER_ALL

# -------------------------------------------------------------------------
# 4. 계좌 및 현금
# -------------------------------------------------------------------------
ws_cash = wb.create_sheet("계좌 및 현금")
ws_cash.views.sheetView[0].showGridLines = True

ws_cash["A1"] = "계좌별 현금 및 예수금 잔액 원장"
ws_cash["A1"].font = TITLE_FONT

cash_headers = ["계좌명", "자산 구분", "통화", "잔액", "기준 환율", "기준 통화 환산액(KRW)", "기준일/메모"]
for c_idx, h in enumerate(cash_headers, start=1):
    c = ws_cash.cell(row=3, column=c_idx, value=h)
    c.font = WHITE_BOLD
    c.fill = NAVY_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER_ALL

sample_cash_data = [
    ("ISA", "예수금", "KRW", 250000, 1, "배당금 및 대기자금"),
    ("CMA", "CMA 잔액", "KRW", 1500000, 1, "비상금 및 월 투자 유보금"),
    ("해외계좌", "달러 예수금", "USD", 350, 1380, "해외 매수 대기금"),
    ("일반/현금", "원화 현금", "KRW", 500000, 1, "파킹통장 잔액"),
]

for r_idx, (acc, gubun, curr, bal, fx, memo) in enumerate(sample_cash_data, start=4):
    ws_cash[f"A{r_idx}"] = acc
    ws_cash[f"B{r_idx}"] = gubun
    ws_cash[f"C{r_idx}"] = curr
    ws_cash[f"D{r_idx}"] = bal
    ws_cash[f"E{r_idx}"] = fx
    ws_cash[f"F{r_idx}"] = f'=D{r_idx}*E{r_idx}'
    ws_cash[f"G{r_idx}"] = memo

    for c_letter in ["A","B","C","D","E","G"]:
        ws_cash[f"{c_letter}{r_idx}"].fill = INPUT_FILL
        ws_cash[f"{c_letter}{r_idx}"].border = BORDER_ALL
        ws_cash[f"{c_letter}{r_idx}"].font = REGULAR_FONT
    
    ws_cash[f"F{r_idx}"].border = BORDER_ALL
    ws_cash[f"F{r_idx}"].font = BOLD_FONT

    ws_cash[f"A{r_idx}"].alignment = Alignment(horizontal="center")
    ws_cash[f"B{r_idx}"].alignment = Alignment(horizontal="center")
    ws_cash[f"C{r_idx}"].alignment = Alignment(horizontal="center")
    ws_cash[f"D{r_idx}"].number_format = "#,##0.00"
    ws_cash[f"E{r_idx}"].number_format = "#,##0.00"
    ws_cash[f"F{r_idx}"].number_format = "₩#,##0"

# Empty input rows
for r_idx in range(8, 20):
    ws_cash[f"F{r_idx}"] = f'=IF(D{r_idx}="","",D{r_idx}*IF(E{r_idx}="",1,E{r_idx}))'
    for c_letter in ["A","B","C","D","E","F","G"]:
        cell = ws_cash[f"{c_letter}{r_idx}"]
        cell.border = BORDER_ALL
        cell.font = REGULAR_FONT
        if c_letter != "F":
            cell.fill = INPUT_FILL
    ws_cash[f"F{r_idx}"].number_format = "₩#,##0"

# -------------------------------------------------------------------------
# 5. 자산 기록
# -------------------------------------------------------------------------
ws_hist = wb.create_sheet("자산 기록")
ws_hist.views.sheetView[0].showGridLines = True

ws_hist["A1"] = "총자산 일별/기간별 스냅샷 기록"
ws_hist["A1"].font = TITLE_FONT

hist_headers = ["기준일시", "투자 자산 평가액", "현금성 자산", "총자산", "데이터 상태"]
for c_idx, h in enumerate(hist_headers, start=1):
    c = ws_hist.cell(row=3, column=c_idx, value=h)
    c.font = WHITE_BOLD
    c.fill = NAVY_HEADER
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = BORDER_ALL

sample_hist_data = [
    ("2026-09-01", 16800000, 2600000, 19400000, "정상"),
    ("2026-09-15", 17200000, 2550000, 19750000, "정상"),
    ("2026-10-01", 18100000, 2700000, 20800000, "정상"),
    ("2026-10-08", 19355000, 2733000, 22088000, "정상"),
]

for r_idx, (dt, inv, cash, tot, st) in enumerate(sample_hist_data, start=4):
    ws_hist[f"A{r_idx}"] = dt
    ws_hist[f"B{r_idx}"] = inv
    ws_hist[f"C{r_idx}"] = cash
    ws_hist[f"D{r_idx}"] = tot
    ws_hist[f"E{r_idx}"] = st

    for c_letter in ["A","B","C","D","E"]:
        cell = ws_hist[f"{c_letter}{r_idx}"]
        cell.border = BORDER_ALL
        cell.font = REGULAR_FONT
        cell.fill = INPUT_FILL
    
    ws_hist[f"A{r_idx}"].alignment = Alignment(horizontal="center")
    ws_hist[f"B{r_idx}"].number_format = "₩#,##0"
    ws_hist[f"C{r_idx}"].number_format = "₩#,##0"
    ws_hist[f"D{r_idx}"].number_format = "₩#,##0"
    ws_hist[f"D{r_idx}"].font = BOLD_FONT
    ws_hist[f"E{r_idx}"].alignment = Alignment(horizontal="center")

# -------------------------------------------------------------------------
# 6. 설정
# -------------------------------------------------------------------------
ws_set = wb.create_sheet("설정")
ws_set.views.sheetView[0].showGridLines = True

ws_set["A1"] = "포트폴리오 환경 설정"
ws_set["A1"].font = TITLE_FONT

settings = [
    ("기준 통화", "KRW"),
    ("목표 비중 분모 기준", "투자 자산만"),
    ("시세 데이터 공급원", "Google 스프레드시트 (GOOGLEFINANCE / Apps Script)"),
    ("지원 계좌 목록", "ISA, 금현물계좌, CMA, 해외계좌, 일반/현금"),
    ("지원 자산 유형", "주식, ETF, 금, 현금, 기타"),
    ("월 배분 주문 단위", "100원 단위 반올림 / 정수주"),
    ("마지막 갱신 시각", "2026-10-08 23:15:00"),
]

for r_idx, (k, v) in enumerate(settings, start=2):
    ws_set[f"A{r_idx}"] = k
    ws_set[f"B{r_idx}"] = v
    ws_set[f"A{r_idx}"].font = BOLD_FONT
    ws_set[f"B{r_idx}"].font = REGULAR_FONT
    ws_set[f"A{r_idx}"].border = BORDER_ALL
    ws_set[f"B{r_idx}"].border = BORDER_ALL
    ws_set[f"A{r_idx}"].fill = GRAY_BG
    ws_set[f"B{r_idx}"].fill = INPUT_FILL

# -------------------------------------------------------------------------
# Adjust Column Widths for all sheets
# -------------------------------------------------------------------------
for ws in wb.worksheets:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.number_format and '₩' in cell.number_format:
                max_len = max(max_len, 15)
            else:
                max_len = max(max_len, len(val_str.encode('utf-8', errors='ignore')) // 2 + len(val_str) // 2)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

wb.save("주식_포트폴리오_관리.xlsx")
print("Workbook created successfully!")

