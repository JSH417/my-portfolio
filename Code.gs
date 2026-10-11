/**
 * 개인 투자 포트폴리오 트래커 자동 생성 스크립트
 * 
 * [사용 방법]
 * 1. 구글 드라이브에서 빈 'Google 스프레드시트'를 하나 만듭니다. (또는 sheets.new 접속)
 * 2. 상단 메뉴: [확장 프로그램] -> [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 이 스크립트 전체를 붙여넣은 뒤 [저장] 아이콘을 누릅니다.
 * 4. 상단 함수 선택창에서 [setupEntireSpreadsheet] 를 선택하고 [실행] 버튼을 누릅니다.
 *    (권한 요청 창이 뜨면 고급 -> 안전하지 않음으로 이동 클릭 후 허용)
 * 5. 몇 초 후 6개 시트(대시보드, 포트폴리오, 월 투자 배분, 계좌 및 현금, 자산 기록, 설정)가
 *    모든 서식, 색상, 수식과 함께 자동으로 완성됩니다!
 */

function setupEntireSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 색상 정의
  var NAVY = "#1F4E79";
  var SUB_NAVY = "#2F5597";
  var INPUT_YELLOW = "#FFF2CC";
  var LIGHT_BLUE = "#D9E1F2";
  var LIGHT_GREEN = "#E2EFDA";
  var GRAY = "#F2F2F2";
  
  // 1. 설정 시트
  var setSheet = ss.getSheetByName("설정") || ss.insertSheet("설정");
  setSheet.clear();
  setSheet.getRange("A1").setValue("포트폴리오 환경 설정").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  var setVals = [
    ["기준 통화", "KRW"],
    ["목표 비중 분모 기준", "투자 자산만"],
    ["시세 데이터 공급원", "Google 스프레드시트 (GOOGLEFINANCE / Apps Script)"],
    ["지원 계좌 목록", "ISA, 금현물계좌, CMA, 해외계좌, 일반/현금"],
    ["지원 자산 유형", "주식, ETF, 금, 현금, 기타"],
    ["월 배분 주문 단위", "100원 단위 반올림 / 정수주"],
    ["마지막 갱신 시각", Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss")]
  ];
  setSheet.getRange(2, 1, setVals.length, 2).setValues(setVals);
  setSheet.getRange(2, 1, setVals.length, 1).setBackground(GRAY).setFontWeight("bold");
  setSheet.getRange(2, 2, setVals.length, 1).setBackground(INPUT_YELLOW);
  setSheet.getRange(2, 1, setVals.length, 2).setBorder(true, true, true, true, true, true);
  
  // 2. 자산 기록 시트
  var histSheet = ss.getSheetByName("자산 기록") || ss.insertSheet("자산 기록");
  histSheet.clear();
  histSheet.getRange("A1").setValue("총자산 일별/기간별 스냅샷 기록").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  var histHeaders = [["기준일시", "투자 자산 평가액", "현금성 자산", "총자산", "데이터 상태"]];
  histSheet.getRange(3, 1, 1, 5).setValues(histHeaders).setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  var sampleHist = [
    ["2026-09-01", 16800000, 2600000, 19400000, "정상"],
    ["2026-09-15", 17200000, 2550000, 19750000, "정상"],
    ["2026-10-01", 18100000, 2700000, 20800000, "정상"],
    ["2026-10-08", 19355000, 2733000, 22088000, "정상"]
  ];
  histSheet.getRange(4, 1, sampleHist.length, 5).setValues(sampleHist);
  histSheet.getRange(4, 2, sampleHist.length, 3).setNumberFormat("₩#,##0");
  histSheet.getRange(4, 1, sampleHist.length, 5).setBorder(true, true, true, true, true, true);
  
  // 3. 계좌 및 현금 시트
  var cashSheet = ss.getSheetByName("계좌 및 현금") || ss.insertSheet("계좌 및 현금");
  cashSheet.clear();
  cashSheet.getRange("A1").setValue("계좌별 현금 및 예수금 잔액 원장").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  var cashHeaders = [["계좌명", "자산 구분", "통화", "잔액", "기준 환율", "기준 통화 환산액(KRW)", "기준일/메모"]];
  cashSheet.getRange(3, 1, 1, 7).setValues(cashHeaders).setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  var sampleCash = [
    ["ISA", "예수금", "KRW", 250000, 1, "", "배당금 및 대기자금"],
    ["CMA", "CMA 잔액", "KRW", 1500000, 1, "", "비상금 및 월 투자 유보금"],
    ["해외계좌", "달러 예수금", "USD", 350, 1380, "", "해외 매수 대기금"],
    ["일반/현금", "원화 현금", "KRW", 500000, 1, "", "파킹통장 잔액"]
  ];
  cashSheet.getRange(4, 1, sampleCash.length, 7).setValues(sampleCash);
  for (var r = 4; r <= 7; r++) {
    cashSheet.getRange(r, 6).setFormula("=D" + r + "*E" + r);
  }
  cashSheet.getRange("D4:E7").setNumberFormat("#,##0.00");
  cashSheet.getRange("F4:F7").setNumberFormat("₩#,##0").setFontWeight("bold");
  cashSheet.getRange("A4:E7").setBackground(INPUT_YELLOW);
  cashSheet.getRange("G4:G7").setBackground(INPUT_YELLOW);
  cashSheet.getRange("A4:G7").setBorder(true, true, true, true, true, true);
  
  // 4. 포트폴리오 시트
  var portSheet = ss.getSheetByName("포트폴리오") || ss.insertSheet("포트폴리오");
  portSheet.clear();
  portSheet.getRange("A1").setValue("보유 포트폴리오 관리").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  portSheet.getRange("A2").setValue("목표 비중 합계:").setFontWeight("bold");
  portSheet.getRange("B2").setFormula("=SUM(N4:N100)").setNumberFormat("0.0%").setFontWeight("bold");
  portSheet.getRange("C2").setValue("미배분 비중:").setFontWeight("bold");
  portSheet.getRange("D2").setFormula("=1-B2").setNumberFormat("0.0%").setFontWeight("bold");
  portSheet.getRange("E2").setValue("상태 경고:").setFontWeight("bold");
  portSheet.getRange("F2").setFormula('=IF(B2>1.0001, "⚠️ 목표 비중 100% 초과!", IF(B2<0.9999, "⚠️ 목표 비중 미달 (정규화 배분)", "✅ 정상 (100%)"))').setFontWeight("bold");
  
  var portHeaders = [["계좌", "티커/코드", "종목명", "자산 유형", "통화", "보유 수량", "누적 매입 금액", "현재가", "기준 환율", "평가액(KRW)", "평가 손익", "수익률", "실제 비중", "목표 비중", "비중 차이", "조회 상태/갱신"]];
  portSheet.getRange(3, 1, 1, 16).setValues(portHeaders).setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  var samplePort = [
    ["ISA", "005930", "삼성전자", "주식", "KRW", 50, 3500000, 72000, 1, 0.20, "수동 입력/구글연동"],
    ["ISA", "379800", "TIGER 미국S&P500", "ETF", "KRW", 100, 1400000, 16500, 1, 0.25, "수동 입력/구글연동"],
    ["해외계좌", "QQQ", "Invesco QQQ Trust", "ETF", "USD", 10, 5200000, 480, 1380, 0.30, "수동 입력/구글연동"],
    ["해외계좌", "NVDA", "NVIDIA Corp", "주식", "USD", 15, 2500000, 130, 1380, 0.15, "수동 입력/구글연동"],
    ["금현물계좌", "금현물", "KRX 금 1g", "금", "KRW", 20, 1900000, 105000, 1, 0.10, "수동 입력"]
  ];
  
  for (var i = 0; i < samplePort.length; i++) {
    var row = 4 + i;
    var d = samplePort[i];
    portSheet.getRange(row, 1, 1, 9).setValues([[d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7], d[8]]]);
    portSheet.getRange(row, 10).setFormula('=IF(B' + row + '="","",F' + row + '*H' + row + '*I' + row + ')');
    portSheet.getRange(row, 11).setFormula('=IF(B' + row + '="","",J' + row + '-G' + row + ')');
    portSheet.getRange(row, 12).setFormula('=IF(G' + row + '>0,K' + row + '/G' + row + ',"-")');
    portSheet.getRange(row, 13).setFormula('=IF(SUM($J$4:$J$100)>0,J' + row + '/SUM($J$4:$J$100),0)');
    portSheet.getRange(row, 14).setValue(d[9]);
    portSheet.getRange(row, 15).setFormula('=IF(N' + row + '="","",M' + row + '-N' + row + ')');
    portSheet.getRange(row, 16).setValue(d[10]);
  }
  
  portSheet.getRange("F4:F8").setNumberFormat("#,##0");
  portSheet.getRange("G4:G8").setNumberFormat("₩#,##0");
  portSheet.getRange("H4:I8").setNumberFormat("#,##0.00");
  portSheet.getRange("J4:K8").setNumberFormat("₩#,##0").setFontWeight("bold");
  portSheet.getRange("L4:L8").setNumberFormat("0.00%").setFontWeight("bold");
  portSheet.getRange("M4:N8").setNumberFormat("0.0%");
  portSheet.getRange("O4:O8").setNumberFormat("+0.0%;-0.0%;0.0%").setFontWeight("bold");
  portSheet.getRange("A4:I8").setBackground(INPUT_YELLOW);
  portSheet.getRange("N4:N8").setBackground(INPUT_YELLOW);
  portSheet.getRange("A4:P8").setBorder(true, true, true, true, true, true);
  
  // 5. 월 투자 배분 시트
  var allocSheet = ss.getSheetByName("월 투자 배분") || ss.insertSheet("월 투자 배분");
  allocSheet.clear();
  allocSheet.getRange("A1").setValue("월 적립식 투자 배분안 계산기").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  allocSheet.getRange("A3").setValue("이번 달 투자 가능 금액").setFontWeight("bold");
  allocSheet.getRange("C3").setValue(500000).setNumberFormat("₩#,##0").setBackground(INPUT_YELLOW).setFontWeight("bold");
  allocSheet.getRange("A4").setValue("현금 유보 금액").setFontWeight("bold");
  allocSheet.getRange("C4").setValue(0).setNumberFormat("₩#,##0").setBackground(INPUT_YELLOW).setFontWeight("bold");
  allocSheet.getRange("A5").setValue("실제 배분 대상 금액").setFontWeight("bold");
  allocSheet.getRange("C5").setFormula("=C3-C4").setNumberFormat("₩#,##0").setBackground(LIGHT_GREEN).setFontWeight("bold");
  allocSheet.getRange("A6").setValue("배분 후 예상 총자산").setFontWeight("bold");
  allocSheet.getRange("C6").setFormula("=SUM(포트폴리오!$J$4:$J$100)+C5").setNumberFormat("₩#,##0").setBackground(LIGHT_GREEN).setFontWeight("bold");
  
  var allocHeaders = [["계좌", "티커", "종목명", "현재 평가액", "현재 비중", "목표 비중", "정규화 비중", "목표 평가액(투자후)", "목표 대비 부족액", "권장 투자 금액", "예상 투자후 비중", "비고"]];
  allocSheet.getRange(9, 1, 1, 12).setValues(allocHeaders).setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  for (var i = 0; i < 5; i++) {
    var pRow = 4 + i;
    var aRow = 10 + i;
    allocSheet.getRange(aRow, 1).setFormula("=포트폴리오!A" + pRow);
    allocSheet.getRange(aRow, 2).setFormula("=포트폴리오!B" + pRow);
    allocSheet.getRange(aRow, 3).setFormula("=포트폴리오!C" + pRow);
    allocSheet.getRange(aRow, 4).setFormula("=포트폴리오!J" + pRow).setNumberFormat("₩#,##0");
    allocSheet.getRange(aRow, 5).setFormula("=포트폴리오!M" + pRow).setNumberFormat("0.0%");
    allocSheet.getRange(aRow, 6).setFormula("=포트폴리오!N" + pRow).setNumberFormat("0.0%");
    allocSheet.getRange(aRow, 7).setFormula("=IF(SUM($F$10:$F$14)>0, F" + aRow + "/SUM($F$10:$F$14), 0)").setNumberFormat("0.0%");
    allocSheet.getRange(aRow, 8).setFormula("=G" + aRow + "*$C$6").setNumberFormat("₩#,##0");
    allocSheet.getRange(aRow, 9).setFormula("=MAX(0, H" + aRow + "-D" + aRow + ")").setNumberFormat("₩#,##0");
    allocSheet.getRange(aRow, 10).setFormula("=IF(SUM($I$10:$I$14)>0, ROUND($C$5*(I" + aRow + "/SUM($I$10:$I$14)), -2), 0)").setNumberFormat("₩#,##0").setFontWeight("bold");
    allocSheet.getRange(aRow, 11).setFormula("=(D" + aRow + "+J" + aRow + ")/$C$6").setNumberFormat("0.0%");
    allocSheet.getRange(aRow, 12).setFormula('=IF(H' + aRow + '<=D' + aRow + ', "목표 초과 (매수 불필요)", IF(J' + aRow + '>0, "권장 매수", "-"))');
  }
  allocSheet.getRange("A10:L14").setBorder(true, true, true, true, true, true);
  
  // 합계행
  allocSheet.getRange(15, 3).setValue("합계").setFontWeight("bold").setBackground(SUB_NAVY).setFontColor("white");
  allocSheet.getRange(15, 4).setFormula("=SUM(D10:D14)").setNumberFormat("₩#,##0").setFontWeight("bold").setBackground(SUB_NAVY).setFontColor("white");
  allocSheet.getRange(15, 10).setFormula("=SUM(J10:J14)").setNumberFormat("₩#,##0").setFontWeight("bold").setBackground(SUB_NAVY).setFontColor("white");
  allocSheet.getRange(15, 12).setFormula('=IF(J15<=C5, "✅ 배분 한도 준수", "⚠️ 투자금 초과")').setFontWeight("bold").setBackground(SUB_NAVY).setFontColor("white");
  
  // 6. 대시보드 시트
  var dashSheet = ss.getSheetByName("대시보드") || ss.insertSheet("대시보드");
  dashSheet.clear();
  dashSheet.getRange("A1").setValue("개인 투자 포트폴리오 트래커 - 대시보드").setFontSize(14).setFontWeight("bold").setFontColor(NAVY);
  dashSheet.getRange("A3").setValue("기준 통화").setFontWeight("bold");
  dashSheet.getRange("B3").setFormula("=설정!B2");
  dashSheet.getRange("D3").setValue("마지막 갱신 시각").setFontWeight("bold");
  dashSheet.getRange("E3").setFormula("=설정!B8");
  
  dashSheet.getRange("A5").setValue("■ 핵심 투자 요약 지표").setFontSize(11).setFontWeight("bold").setFontColor(NAVY);
  dashSheet.getRange(5, 2).setValue("항목").setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  dashSheet.getRange(5, 3).setValue("금액 / 수치").setBackground(NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  var dashMetrics = [
    ["누적 투자 원금", "=SUM(포트폴리오!G4:G100)", "₩#,##0"],
    ["현재 총자산", "=C8+C9", "₩#,##0"],
    ["투자 자산 평가액", "=SUM(포트폴리오!J4:J100)", "₩#,##0"],
    ["현금성 자산", "=SUM('계좌 및 현금'!F4:F50)", "₩#,##0"],
    ["평가 손익", "=C8-C6", "₩#,##0"],
    ["수익률", '=IF(C6>0, C10/C6, "-")', "0.00%"],
    ["1일 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=2, C7-INDEX(\'자산 기록\'!D4:D100, COUNTA(\'자산 기록\'!A4:A100)-1), "기록 부족")', "₩#,##0"],
    ["1주 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=5, C7-INDEX(\'자산 기록\'!D4:D100, MAX(1, COUNTA(\'자산 기록\'!A4:A100)-5)), "기록 부족")', "₩#,##0"],
    ["1개월 자산 변화", '=IF(COUNTA(\'자산 기록\'!A4:A100)>=20, C7-INDEX(\'자산 기록\'!D4:D100, MAX(1, COUNTA(\'자산 기록\'!A4:A100)-20)), "기록 부족")', "₩#,##0"]
  ];
  
  for (var m = 0; m < dashMetrics.length; m++) {
    var mr = 6 + m;
    dashSheet.getRange(mr, 2).setValue(dashMetrics[m][0]).setFontWeight("bold").setBackground(GRAY);
    dashSheet.getRange(mr, 3).setFormula(dashMetrics[m][1]).setFontWeight("bold").setNumberFormat(dashMetrics[m][2]);
  }
  dashSheet.getRange(6, 2, dashMetrics.length, 2).setBorder(true, true, true, true, true, true);
  
  dashSheet.getRange("E5").setValue("■ 계좌별 자산 요약").setFontSize(11).setFontWeight("bold").setFontColor(NAVY);
  var accHeaders = [["계좌 구분", "투자 자산", "현금 잔액", "합계"]];
  dashSheet.getRange(6, 5, 1, 4).setValues(accHeaders).setBackground(SUB_NAVY).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  
  var accList = ["ISA", "금현물계좌", "CMA", "해외계좌", "일반/현금"];
  for (var a = 0; a < accList.length; a++) {
    var ar = 7 + a;
    dashSheet.getRange(ar, 5).setValue(accList[a]).setFontWeight("bold").setHorizontalAlignment("center");
    dashSheet.getRange(ar, 6).setFormula('=SUMIF(포트폴리오!$A$4:$A$100, "' + accList[a] + '", 포트폴리오!$J$4:$J$100)').setNumberFormat("₩#,##0");
    dashSheet.getRange(ar, 7).setFormula('=SUMIF(\'계좌 및 현금\'!$A$4:$A$50, "' + accList[a] + '", \'계좌 및 현금\'!$F$4:$F$50)').setNumberFormat("₩#,##0");
    dashSheet.getRange(ar, 8).setFormula('=F' + ar + '+G' + ar).setNumberFormat("₩#,##0").setFontWeight("bold");
  }
  dashSheet.getRange(7, 5, accList.length, 4).setBorder(true, true, true, true, true, true);
  
  dashSheet.getRange(12, 5).setValue("합계").setFontWeight("bold").setBackground(NAVY).setFontColor("white").setHorizontalAlignment("center");
  dashSheet.getRange(12, 6).setFormula("=SUM(F7:F11)").setNumberFormat("₩#,##0").setFontWeight("bold").setBackground(NAVY).setFontColor("white");
  dashSheet.getRange(12, 7).setFormula("=SUM(G7:G11)").setNumberFormat("₩#,##0").setFontWeight("bold").setBackground(NAVY).setFontColor("white");
  dashSheet.getRange(12, 8).setFormula("=SUM(H7:H11)").setNumberFormat("₩#,##0").setFontWeight("bold").setBackground(NAVY).setFontColor("white");
  dashSheet.getRange(12, 5, 1, 4).setBorder(true, true, true, true, true, true);

  // 시트 순서 정렬
  ss.setActiveSheet(dashSheet);
  ss.moveChartToObject ? null : null; // dummy
}
/**
 * 개인 투자 포트폴리오 트래커 - Google Apps Script
 */

/**
 * 네이버 증권에서 국내 종목 현재가를 가져옵니다.
 * 사용법: =GET_NAVER_PRICE(B4)
 */
function GET_NAVER_PRICE(ticker) {
  if (!ticker) return "";
  var code = String(ticker).trim();
  // 숫자 앞 0이 빠졌을 경우 6자리 맞춤 (예: 5930 -> 005930)
  while (code.length < 6 && /^\d+$/.test(code)) {
    code = "0" + code;
  }
  
  try {
    var url = "https://finance.naver.com/item/main.naver?code=" + code;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var html = response.getContentText("EUC-KR");
    
    var match = html.match(/<p class="no_today">[\s\S]*?<span class="blind">([0-9,]+)<\/span>/);
    if (match && match[1]) {
      return parseInt(match[1].replace(/,/g, ""), 10);
    }
  } catch (e) {
    return "조회 실패";
  }
  return "조회 실패";
}

/**
 * 네이버 증권에서 국내 종목명을 가져옵니다.
 * 사용법: =GET_NAVER_NAME(B4)
 */
function GET_NAVER_NAME(ticker) {
  if (!ticker) return "";
  var code = String(ticker).trim();
  while (code.length < 6 && /^\d+$/.test(code)) {
    code = "0" + code;
  }
  
  try {
    var url = "https://finance.naver.com/item/main.naver?code=" + code;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var html = response.getContentText("EUC-KR");
    
    var match = html.match(/<h2><a href="[^"]*">([^<]+)<\/a><\/h2>/);
    if (match && match[1]) {
      return match[1].trim();
    }
  } catch (e) {
    return "조회 실패";
  }
  return "조회 실패";
}

/**
 * 국내 및 미국 주식 현재가 통합 자동 조회 함수
 * 사용법: =GET_PRICE(B4)
 */
function GET_PRICE(ticker) {
  if (!ticker) return "";
  var t = String(ticker).trim();
  if (/^\d+$/.test(t)) {
    return GET_NAVER_PRICE(t);
  }
  return ""; // 미국 주식은 시트에서 GOOGLEFINANCE 사용 권장
}

/**
 * 국내 및 미국 주식 종목명 통합 자동 조회 함수
 * 사용법: =GET_NAME(B4)
 */
function GET_NAME(ticker) {
  if (!ticker) return "";
  var t = String(ticker).trim();
  if (/^\d+$/.test(t)) {
    return GET_NAVER_NAME(t);
  }
  return "";
}

/**
 * 당일 총자산 스냅샷 자동 기록
 */
function recordDailySnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("대시보드");
  var historySheet = ss.getSheetByName("자산 기록");
  
  if (!dashSheet || !historySheet) return;
  
  var today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var investedAsset = dashSheet.getRange("C8").getValue();
  var cashAsset = dashSheet.getRange("C9").getValue();
  var totalAsset = dashSheet.getRange("C7").getValue();
  var status = "정상";
  
  var data = historySheet.getDataRange().getValues();
  var rowIndexToUpdate = -1;
  
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, "Asia/Seoul", "yyyy-MM-dd");
    }
    if (rowDate === today) {
      rowIndexToUpdate = i + 1;
      break;
    }
  }
  
  if (rowIndexToUpdate > 0) {
    historySheet.getRange(rowIndexToUpdate, 1, 1, 5).setValues([[
      today, investedAsset, cashAsset, totalAsset, status
    ]]);
  } else {
    historySheet.appendRow([today, investedAsset, cashAsset, totalAsset, status]);
  }
}
/**
 * KRX 금현물 1g 현재가를 실시간으로 가져옵니다.
 * 사용법: =GET_KRX_GOLD_PRICE()
 */
function GET_KRX_GOLD_PRICE() {
  try {
    var url = "https://m.stock.naver.com/marketindex/metals/M04020000";
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X)"
      }
    });
    var html = response.getContentText("UTF-8");
    
    // <strong class="DetailInfo_price... >177,480...</strong> 파싱
    var match = html.match(/DetailInfo_price[^>]*>([0-9,]+)/);
    if (match && match[1]) {
      return parseInt(match[1].replace(/,/g, ""), 10);
    }
  } catch (e) {
    return "조회 실패";
  }
  return "조회 실패";
}
