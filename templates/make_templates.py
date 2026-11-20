"""근태 표준 양식 2종 생성: python3 templates/make_templates.py
① attendance-std-daily.xlsx  일별 근태표 (직원마다 근태/근로시간/연장/야간 4줄)
② attendance-std-clock.xlsx  출퇴근 기록 (직원·날짜마다 한 줄)
예시 데이터는 모두 가상의 인물입니다.
"""
import calendar
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).parent
MARK = "근태표준양식 v1 — 이 칸은 지우지 마세요 (웹 점검 도구가 표준 양식을 알아보는 표시)"
Y, M = 2026, 10
LAST = calendar.monthrange(Y, M)[1]
CODES = ["근무", "휴근", "반차", "휴무", "주휴", "대휴", "연차", "병가", "경조", "공가", "휴가", "휴직", "결근", "입사전", "퇴사"]

INK = "14110F"
HEAD = PatternFill("solid", fgColor="2B2522")
CODE_ROW = PatternFill("solid", fgColor="FFF6E0")
EXAMPLE = PatternFill("solid", fgColor="EEF1F7")
WEEKEND = PatternFill("solid", fgColor="F6F3F1")
thin = Side(style="thin", color="D9D3CE")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WHITE_B = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="top", wrap_text=True)


def head(ws, row, labels, widths=None):
    for c, lab in enumerate(labels, 1):
        cell = ws.cell(row, c, lab)
        cell.fill, cell.font, cell.alignment, cell.border = HEAD, WHITE_B, CENTER, BOX
        if widths and c <= len(widths) and widths[c - 1]:
            ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]
    ws.row_dimensions[row].height = 30


def yn(ws, rng):
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True, showErrorMessage=True,
                        errorTitle="Y 또는 N", error="Y(해당) 또는 N(해당 없음)만 넣어 주세요.")
    ws.add_data_validation(dv)
    dv.add(rng)


def settings_sheet(wb):
    ws = wb.create_sheet("설정")
    ws["A1"] = MARK
    ws["A1"].font = Font(bold=True, color="E63C12")
    ws["A2"] = "사업장마다 한 줄. 한 파일에 여러 사업장을 넣어도 되고, 사업장마다 파일을 따로 만들어도 됩니다."
    head(ws, 3, ["사업장", "대상월", "상시5인이상", "탄력적근로시간제", "주휴일", "토요일약정휴일", "비고"],
         [22, 12, 13, 17, 16, 15, 40])
    rows = [["예시사업장", f"{Y}-{M:02d}", "Y", "N", "일요일", "N", "← 예시입니다. 실제 사업장으로 바꿔 쓰세요"]]
    for r, vals in enumerate(rows, 4):
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BOX
            cell.fill = EXAMPLE
    for r in range(4, 54):
        ws.cell(r, 2).number_format = "@"
        for c in range(1, 8):
            ws.cell(r, c).border = BOX
    yn(ws, "C4:D53")
    yn(ws, "F4:F53")
    dv = DataValidation(type="list", formula1='"일요일,토요일,교대(지정)"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E4:E53")
    notes = {
        "B3": "YYYY-MM 형식 (예: 2026-10). 텍스트로 입력.",
        "C3": "상시 근로자 5인 이상이면 Y. N이면 가산수당·52시간 판정 없이 시간만 집계.",
        "D3": "탄력적 근로시간제를 시행하면 Y.",
        "E3": "일요일 / 토요일 / 교대(지정). 교대·순환 근무로 주휴일이 사람마다 다르면 '교대(지정)'을 고르고, 근태표에 주휴일을 '주휴'로 표시.",
        "F3": "토요일을 취업규칙상 유급휴일로 정했으면 Y.",
    }
    for ref, t in notes.items():
        ws[ref].comment = Comment(t, "근태표준양식")
    ws.freeze_panes = "A4"
    return ws


def guide_sheet(wb, kind):
    ws = wb.active
    ws.title = "안내"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 70
    ws["A1"] = "근태 표준 양식 " + {"daily": "① 일별 근태표", "clock": "② 출퇴근 기록", "status": "③ 근무현황"}[kind]
    ws["A1"].font = Font(bold=True, size=16)
    lines = [
        "작성 순서",
        "1. [설정] 시트에 사업장 이름·대상월·사업장 조건을 적습니다.",
        {"daily": "2. [근태표] 시트에 직원마다 4줄(근태/근로시간/연장/야간)을 적습니다. 직원이 많으면 4줄 묶음을 복사해 붙여 쓰세요.",
         "clock": "2. [출퇴근] 시트에 직원·날짜마다 한 줄씩 적습니다. 출퇴근 기록기에서 내려받은 자료를 붙여넣어도 됩니다.",
         "status": "2. [근무현황] 시트에 직원마다 한 줄, 날짜 칸에 ○(근무)·/(휴무)·연차 등을 적습니다. 연장근로가 있으면 [시간외] 시트에 같은 사람의 날짜별 시간을 적습니다."}[kind],
        {"daily": "3. 시트 이름(설정·근태표)과 머리글 행은 바꾸지 마세요. 시트를 나눌 때는 '근태표_A동'처럼 이름 앞부분을 유지하세요.",
         "clock": "3. 시트 이름(설정·출퇴근)과 머리글 행은 바꾸지 마세요. 시트를 나눌 때는 '출퇴근_A동'처럼 이름 앞부분을 유지하세요.",
         "status": "3. 시트 이름(설정·근무현황·시간외)과 머리글 행은 바꾸지 마세요. 주간·심야를 나눌 때는 '근무현황(주간)', '근무현황(심야)', '시간외(심야)'처럼 이름 앞부분을 유지하세요."}[kind],
        "4. 웹 점검 도구에 올리면 칸 단위로 검증하고, 틀린 칸은 '근태표 K12'처럼 위치를 알려 줍니다.",
        "",
        "시간 입력 규칙",
        "· 근로시간 = 그날 실제 일한 시간(휴게시간 제외). 예: 09:00~18:00, 휴게 1시간 → 8",
        "· 야간 = 22시~다음날 06시 사이에 일한 시간. 근로시간 안에 포함된 시간입니다.",
    ]
    if kind == "daily":
        lines += [
            "· 근로시간을 비우면 근무 8시간·반차 4시간·휴근 8시간으로 보고, '연장' 줄의 시간을 더합니다.",
            "  (예: 근무 + 연장 2 → 10시간). 정확한 시간이 있으면 근로시간 줄에 직접 적으세요(그때 연장 줄은 무시).",
            "· 격일제·24시간 교대처럼 하루 근로시간이 8시간이 아니면 근로시간 줄을 반드시 적으세요.",
        ]
    elif kind == "status":
        lines += [
            "· ○ 하루 근로시간 = 그 사람의 '1일근로' 칸(비우면 8시간). 야간조라면 '1일야간'에 22~06시 근로시간을 적으세요(예: 22:00~07:00, 휴게 1시간 → 1일근로 8, 1일야간 7).",
            "· 그날만 시간이 다르면 날짜 칸에 ○ 대신 숫자(실근로시간)를 적어도 됩니다. 예: 12",
            "· [시간외] 시트의 시간은 그날 근로시간에 더합니다. 휴무(/)인 날 일했다면 근무현황에 ○ 또는 휴근으로 적어 주세요.",
            "· 입사일 전·퇴사일 후 날짜는 비워 두면 재직 기간이 아닌 것으로 봅니다. 재직 중 빈 칸은 휴무로 봅니다.",
            "· 머리글 위·아래의 '전체 소계' 같은 합계 행은 있어도 됩니다(성명이 없는 행은 읽지 않음).",
        ]
    else:
        lines += [
            "· 출근·퇴근 시각을 적으면 근로시간·야간시간을 자동 계산합니다. 퇴근이 출근보다 이르면 다음날 퇴근으로 봅니다.",
            "· 휴게(분)를 비우면 법정 최소 휴게(4시간 이상 30분, 8시간 이상 60분)를 빼고 계산합니다. 실제 휴게를 적는 것이 정확합니다.",
            "· 근로시간·야간시간 칸에 숫자를 적으면 계산값 대신 그 값을 씁니다. 휴게를 밤(22~06시)에 쉬면 야간시간을 직접 적는 것이 정확합니다.",
            "· 출근하지 않은 날도 한 줄 적고 근태(휴무·주휴·연차 등)를 넣어 주세요. 빠진 날은 휴무로 봅니다.",
        ]
    lines += [
        "",
        "직원 표시 (해당하면 Y)",
        "· 감단: 노동부 승인을 받은 감시·단속적 근로자(경비 등)",
        "· 교대: 교대제 근무자 (감단+교대면 노동절 근무 시 휴일근로수당 산정)",
        "· 초단기: 주 15시간 미만 근로자. 비워 두면 자료로 자동 판단합니다.",
    ]
    r = 3
    for t in lines:
        ws.cell(r, 1, t).font = BOLD if t and not t.startswith(("·", " ", "1", "2", "3", "4")) else Font()
        r += 1
    r += 1
    ws.cell(r, 1, "근태 코드").font = Font(bold=True, size=13)
    r += 1
    for c, lab in enumerate(["코드", "판정에서의 의미", "쓰는 경우"], 1):
        cell = ws.cell(r, c, lab)
        cell.fill, cell.font, cell.border = HEAD, WHITE_B, BOX
    table = ([("○", "근무 (1일근로 시간)", "근무현황 시트의 기본 근무 표시. O·◯·●도 같은 뜻"),
              ("/", "휴무", "근무현황 시트의 쉬는 날 표시"),
              ("숫자 (예: 12)", "그날 실근로시간", "그날만 근로시간이 다를 때")] if kind == "status" else []) + [
        ("근무", "근무 (기본 8시간)", "평일·휴무일 근무. 공휴일·주휴일에 일하면 자동으로 휴일근로로 판정"),
        ("휴근", "휴일 근무 (기본 8시간)", "교대 근무자가 지정 주휴일에 출근한 경우 등 휴일근로임을 표시할 때"),
        ("반차", "4시간 근무", "오전·오후 반차"),
        ("휴무", "쉬는 날 (무급휴무일 포함)", "소정근로일이 아닌 날"),
        ("주휴", "그 주의 주휴일", "교대·순환 근무로 주휴일이 요일과 다를 때 꼭 표시"),
        ("대휴", "휴일대체로 쉰 날", "휴일근로를 다른 날로 대체한 경우. 사전 지정·서면합의 요건 확인 표시"),
        ("연차", "유급 휴가", "연차·반차가 아닌 하루 연차"),
        ("병가·경조·공가·휴가", "휴가류", "회사 규정상 휴가"),
        ("휴직", "휴직", "육아·병 휴직 등"),
        ("결근", "결근", "무단·무급 결근 (주휴수당 미발생 가능 표시)"),
        ("입사전·퇴사", "재직 기간 아님", "월중 입사·퇴사자의 재직 전후 날짜 (판정에서 제외)"),
    ]
    for row in table:
        r += 1
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border, cell.alignment = BOX, WRAP
    r += 2
    ws.cell(r, 1, "파일 안의 연한 파란 칸은 예시입니다. 지우거나 실제 자료로 덮어 쓰세요.").font = Font(italic=True, color="6B6460")
    return ws


def daily():
    wb = Workbook()
    guide_sheet(wb, "daily")
    settings_sheet(wb)
    ws = wb.create_sheet("근태표")
    ws["A1"] = "일별 근태표 — 직원마다 4줄 (근태 / 근로시간 / 연장 / 야간). 날짜 1~31 열은 대상월과 상관없이 그대로 두세요."
    ws["A1"].font = BOLD
    fixed = ["사업장", "성명", "사번", "직무", "감단", "교대", "초단기", "구분"]
    head(ws, 3, fixed + [str(d) for d in range(1, 32)], [16, 10, 11, 10, 6, 6, 7, 9] + [5.2] * 31)
    for d in range(1, 32):
        c = 8 + d
        ws.cell(3, c).value = d
        if d <= LAST:
            wd = dt.date(Y, M, d).weekday()
            cell = ws.cell(2, c, "월화수목금토일"[wd])
            cell.alignment = CENTER
            cell.font = Font(color="E63C12" if wd >= 5 else "6B6460", size=9)
    ws.cell(2, 8, f"{Y}-{M:02d} 요일 (예시)").font = Font(size=9, color="6B6460")

    def weekday(d):
        return dt.date(Y, M, d).weekday()

    hol = {3, 5, 9}  # 2026-10 개천절(토)·대체공휴일·한글날 → 예시용
    ex = []
    # 1) 사무직: 평일 근무, 주말 휴무, 연장 일부, 공휴일 휴무, 10/9 출근(휴일근로)
    code, hrs, ot, ngt = {}, {}, {}, {}
    for d in range(1, LAST + 1):
        w = weekday(d)
        code[d] = "주휴" if w == 6 else "휴무" if w == 5 or d in hol else "근무"
    code[9] = "근무"
    code[20] = "연차"
    for d in (6, 7, 8, 14, 15, 21, 22):
        ot[d] = 2
    ex.append(("예시사업장", "김예시", "E001", "사무", "", "", "", code, hrs, ot, ngt))
    # 2) 경비(감단) 격일제 24시간: 근무일 근로 16h(휴게 8h 제외), 야간 6h
    code, hrs, ot, ngt = {}, {}, {}, {}
    for d in range(1, LAST + 1):
        if d % 2 == 1:
            code[d], hrs[d], ngt[d] = "근무", 16, 6
        else:
            code[d] = "휴무"
    ex.append(("예시사업장", "이샘플", "E002", "경비", "Y", "Y", "", code, hrs, ot, ngt))
    # 3) 교대 근무자: 4일 근무 2일 휴무, 주휴 표시, 10/15 입사
    code, hrs, ot, ngt = {}, {}, {}, {}
    for d in range(1, LAST + 1):
        if d < 15:
            code[d] = "입사전"
            continue
        k = (d - 15) % 6
        code[d] = "근무" if k < 4 else ("주휴" if k == 5 else "휴무")
        if code[d] == "근무" and k >= 2:
            hrs[d], ngt[d] = 8, 7  # 야간조 22:00~07:00, 휴게 1시간
    ex.append(("예시사업장", "박표준", "E003", "생산", "", "Y", "", code, hrs, ot, ngt))

    dv_code = DataValidation(type="list", formula1='"' + ",".join(CODES) + '"', allow_blank=True, showErrorMessage=True,
                             errorTitle="표준 근태코드", error="목록에 있는 근태코드만 넣어 주세요: " + "·".join(CODES))
    dv_num = DataValidation(type="decimal", operator="between", formula1="0", formula2="24", allow_blank=True,
                            showErrorMessage=True, errorTitle="시간", error="0~24 사이 숫자(시간)를 넣어 주세요. 30분은 0.5")
    ws.add_data_validation(dv_code)
    ws.add_data_validation(dv_num)
    labels = ["근태", "근로시간", "연장", "야간"]
    r = 4
    blocks = ex + [None] * 47  # 예시 3명 + 빈 묶음 47개 = 50명분
    for b in blocks:
        for k, lab in enumerate(labels):
            row = r + k
            ws.cell(row, 8, lab).font = BOLD if k == 0 else Font(color="6B6460")
            for c in range(1, 40):
                cell = ws.cell(row, c)
                cell.border = BOX
                if c >= 9:
                    cell.alignment = CENTER
                if k == 0:
                    cell.fill = CODE_ROW
                if c >= 9 and c - 8 <= LAST and weekday(c - 8) >= 5 and k > 0:
                    cell.fill = WEEKEND
            rng = f"{get_column_letter(9)}{row}:{get_column_letter(39)}{row}"
            (dv_code if k == 0 else dv_num).add(rng)
        if b:
            site, name, eid, job, g, s, u, code, hrs, ot, ngt = b
            for c, v in enumerate([site, name, eid, job, g, s, u], 1):
                cell = ws.cell(r, c, v or None)
                cell.fill = EXAMPLE
            for d in range(1, LAST + 1):
                c = 8 + d
                ws.cell(r, c, code.get(d)).fill = EXAMPLE
                for k, src in ((1, hrs), (2, ot), (3, ngt)):
                    if d in src:
                        cell = ws.cell(r + k, c, src[d])
                        cell.fill = EXAMPLE
        r += 4
    yn(ws, f"E4:G{r}")
    ws.freeze_panes = "I4"
    wb.move_sheet("근태표", offset=0)
    wb.active = wb.sheetnames.index("근태표")
    path = OUT / "attendance-std-daily.xlsx"
    wb.save(path)
    return path


def clock():
    wb = Workbook()
    guide_sheet(wb, "clock")
    settings_sheet(wb)
    ws = wb.create_sheet("출퇴근")
    ws["A1"] = "출퇴근 기록 — 직원·날짜마다 한 줄. 출근하지 않은 날도 근태(휴무·주휴·연차 등)를 넣어 한 줄 적어 주세요."
    ws["A1"].font = BOLD
    cols = ["사업장", "성명", "사번", "직무", "감단", "교대", "초단기", "일자", "근태", "출근", "퇴근", "휴게(분)", "근로시간", "야간시간", "비고"]
    head(ws, 3, cols, [16, 10, 11, 10, 6, 6, 7, 12, 8, 8, 8, 9, 10, 10, 30])
    tips = {"I3": "비우면 출퇴근 시각이 있을 때 '근무'로 봅니다. 휴일근무는 날짜로 자동 판정.",
            "J3": "09:00 형식. 퇴근이 출근보다 이르면 다음날 퇴근으로 봅니다.",
            "L3": "실제 휴게시간(분). 비우면 법정 최소(4h↑ 30분, 8h↑ 60분)로 계산.",
            "M3": "비우면 출퇴근-휴게로 자동 계산. 숫자를 넣으면 그 값을 씁니다.",
            "N3": "비우면 22~06시 겹친 시간으로 자동 계산."}
    for ref, t in tips.items():
        ws[ref].comment = Comment(t, "근태표준양식")
    rows = []
    hol = {3, 5, 9}
    for d in range(1, LAST + 1):
        day = dt.date(Y, M, d)
        w = day.weekday()
        if w == 6:
            rows.append(("김예시", "E001", "사무", "", day, "주휴", None, None, None, "", "", ""))
        elif w == 5 or (d in hol and d != 9):
            rows.append(("김예시", "E001", "사무", "", day, "휴무", None, None, None, "", "", ""))
        elif d == 20:
            rows.append(("김예시", "E001", "사무", "", day, "연차", None, None, None, "", "", ""))
        else:
            out = dt.time(20, 0) if d in (6, 7, 8) else dt.time(18, 0)
            rows.append(("김예시", "E001", "사무", "", day, "", dt.time(9, 0), out, 60, "", "", "한글날 출근" if d == 9 else ""))
    for d in range(1, LAST + 1):
        day = dt.date(Y, M, d)
        if d % 2 == 1:
            rows.append(("이샘플", "E002", "경비", "Y", day, "근무", dt.time(7, 0), dt.time(7, 0), 480, "", "", "격일 24시간, 휴게 8시간"))
        else:
            rows.append(("이샘플", "E002", "경비", "Y", day, "휴무", None, None, None, "", "", ""))
    for d in range(1, LAST + 1):
        day = dt.date(Y, M, d)
        k = (d - 1) % 6
        if k < 4:
            rows.append(("박표준", "E003", "생산", "", day, "", dt.time(22, 0), dt.time(7, 0), 60, "", "", "야간조"))
        else:
            rows.append(("박표준", "E003", "생산", "", day, "주휴" if k == 5 else "휴무", None, None, None, "", "", ""))
    r = 4
    for name, eid, job, g, day, code, a, b, brk, h, n, note in rows:
        vals = ["예시사업장", name, eid, job, g or None, "Y" if name in ("이샘플", "박표준") else None, None,
                day, code or None, a, b, brk, h or None, n or None, note or None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.fill, cell.border = EXAMPLE, BOX
        ws.cell(r, 8).number_format = "yyyy-mm-dd"
        ws.cell(r, 10).number_format = ws.cell(r, 11).number_format = "hh:mm"
        r += 1
    for rr in range(r, r + 1500):
        ws.cell(rr, 8).number_format = "yyyy-mm-dd"
        ws.cell(rr, 10).number_format = ws.cell(rr, 11).number_format = "hh:mm"
    dv_code = DataValidation(type="list", formula1='"' + ",".join(CODES) + '"', allow_blank=True, showErrorMessage=True,
                             errorTitle="표준 근태코드", error="목록에 있는 근태코드만 넣어 주세요.")
    ws.add_data_validation(dv_code)
    dv_code.add(f"I4:I{r + 1500}")
    yn(ws, f"E4:G{r + 1500}")
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:O{r - 1}"
    wb.move_sheet("출퇴근", offset=0)
    wb.active = wb.sheetnames.index("출퇴근")
    path = OUT / "attendance-std-clock.xlsx"
    wb.save(path)
    return path


def status():
    wb = Workbook()
    guide_sheet(wb, "status")
    settings_sheet(wb)
    ws = wb.create_sheet("근무현황")
    ws["A1"] = "근무현황"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "직원마다 한 줄 · ○ 근무 / 휴무 · 연차·반차·주휴·휴근·대휴 등 · 숫자 = 그날 실근로시간"
    ws["A2"].font = Font(size=9, color="6B6460")
    fixed = ["순번", "구분", "직무상세", "성명", "사번", "입사일", "퇴사일", "감단", "교대", "초단기", "1일근로", "1일야간"]
    nF = len(fixed)
    head(ws, 3, fixed + list(range(1, 32)) + ["근무일수"], [5, 8, 11, 9, 10, 11, 11, 5, 5, 6, 6, 6] + [4.2] * 31 + [7])
    for d in range(1, LAST + 1):
        wd = dt.date(Y, M, d).weekday()
        cell = ws.cell(4, nF + d, "월화수목금토일"[wd])
        cell.alignment, cell.border = CENTER, BOX
        cell.font = Font(color="E63C12" if wd >= 5 else "6B6460", size=9)
    ws.cell(4, 1, f"{Y}-{M:02d} 요일 (예시)").font = Font(size=9, color="6B6460")
    hol = {3, 5, 9}

    def wk(d):
        return dt.date(Y, M, d).weekday()

    people = []
    # 주간 상담: 평일 ○, 주말·공휴일 /, 일요일 주휴, 20일 연차
    days = {}
    for d in range(1, LAST + 1):
        days[d] = "주휴" if wk(d) == 6 else "/" if wk(d) == 5 or d in hol else "○"
    days[20] = "연차"
    days[9] = "○"
    people.append(("그룹장", "상담", "김예시", "E101", dt.date(2023, 3, 2), None, "", "", "", 8, None, days))
    # 심야조 격일: 22:00~07:00, 휴게 1h → 1일근로 8, 1일야간 7
    days = {d: ("○" if d % 2 else "/") for d in range(1, LAST + 1)}
    people.append(("상담원", "심야상담", "이샘플", "E102", dt.date(2024, 5, 13), None, "", "Y", "", 8, 7, days))
    # 월중 퇴사: 10/17 퇴사, 하루 12시간 근무한 날
    days = {d: ("주휴" if wk(d) == 6 else "/" if wk(d) == 5 or d in hol else "○") for d in range(1, 18)}
    days[14] = 12
    people.append(("상담원", "상담", "박표준", "E103", dt.date(2025, 1, 6), dt.date(Y, M, 17), "", "", "", 8, None, days))
    r = 5
    dv = DataValidation(type="list", formula1='"○,/,주휴,휴근,반차,대휴,연차,병가,경조,공가,휴가,휴직,결근"', allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    for n, (grp, job, name, eid, join, quit_, g, sft, u, base, night, days) in enumerate(people + [None] * 0, 1):
        vals = [n, grp, job, name, eid, join, quit_, g or None, sft or None, u or None, base, night]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.fill, cell.border = EXAMPLE, BOX
        ws.cell(r, 6).number_format = ws.cell(r, 7).number_format = "yyyy-mm-dd"
        for d in range(1, LAST + 1):
            cell = ws.cell(r, nF + d, days.get(d))
            cell.fill, cell.border, cell.alignment = EXAMPLE, BOX, CENTER
        col = get_column_letter(nF + LAST)
        ws.cell(r, nF + 32, f'=COUNTIF({get_column_letter(nF + 1)}{r}:{col}{r},"○")+COUNT({get_column_letter(nF + 1)}{r}:{col}{r})').border = BOX
        r += 1
    for rr in range(r, r + 300):
        for c in range(1, nF + 33):
            ws.cell(rr, c).border = BOX
        ws.cell(rr, 6).number_format = ws.cell(rr, 7).number_format = "yyyy-mm-dd"
    dv.add(f"{get_column_letter(nF + 1)}5:{get_column_letter(nF + 31)}{r + 300}")
    yn(ws, f"H5:J{r + 300}")
    ws.freeze_panes = ws.cell(5, nF + 1)

    ot = wb.create_sheet("시간외")
    ot["A1"] = "시간외 (연장근로 시간)"
    ot["A1"].font = Font(bold=True, size=14)
    ot["A2"] = "근무현황과 같은 사람(사번 또는 성명)의 날짜별 연장시간. 그날 근로시간에 더해집니다. 없는 사람은 안 적어도 됩니다."
    ot["A2"].font = Font(size=9, color="6B6460")
    head(ot, 3, ["순번", "성명", "사번"] + list(range(1, 32)) + ["합계"], [5, 9, 10] + [4.2] * 31 + [7])
    for d in range(1, LAST + 1):
        cell = ot.cell(4, 3 + d, "월화수목금토일"[wk(d)])
        cell.alignment, cell.font = CENTER, Font(color="E63C12" if wk(d) >= 5 else "6B6460", size=9)
    rows = [("김예시", "E101", {6: 2, 7: 2, 8: 1.5}), ("이샘플", "E102", {1: 1, 13: 1})]
    for n, (name, eid, h) in enumerate(rows, 1):
        rr = 4 + n
        for c, v in enumerate([n, name, eid], 1):
            cell = ot.cell(rr, c, v)
            cell.fill, cell.border = EXAMPLE, BOX
        for d in range(1, LAST + 1):
            cell = ot.cell(rr, 3 + d, h.get(d))
            cell.fill, cell.border, cell.alignment = EXAMPLE, BOX, CENTER
        ot.cell(rr, 35, f"=SUM(D{rr}:{get_column_letter(3 + LAST)}{rr})").border = BOX
    dvn = DataValidation(type="decimal", operator="between", formula1="0", formula2="24", allow_blank=True,
                         showErrorMessage=True, errorTitle="시간", error="0~24 사이 숫자(시간)를 넣어 주세요. 30분은 0.5")
    ot.add_data_validation(dvn)
    dvn.add(f"D5:AH305")
    ot.freeze_panes = "D5"
    wb.move_sheet("시간외", offset=-(len(wb.sheetnames) - 1))
    wb.move_sheet("근무현황", offset=-(len(wb.sheetnames) - 1))
    wb.active = 0
    path = OUT / "attendance-std-status.xlsx"
    wb.save(path)
    return path


if __name__ == "__main__":
    for p in (daily(), clock(), status()):
        print("wrote", p)
