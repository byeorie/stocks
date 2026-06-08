// =====================================================
// 주식 대시보드 - Google Apps Script 웹앱 (v4)
// 거래내역 / 실현손익 / 현금 / 리밸런싱 기능 포함
// =====================================================
// 사용법:
// 1. 구글 시트에서 확장 프로그램 > Apps Script 열기
// 2. 이 코드 전체를 붙여넣기
// 3. setup() 함수를 한 번 실행 (시트 자동 생성)
// 4. 배포 > 새 배포 > 웹앱 > 액세스: 모든 사용자 > 배포
// 5. 웹앱 URL을 대시보드 HTML의 입력창에 붙여넣기
// =====================================================

const SHEET_PORTFOLIO = '포트폴리오';
const SHEET_WATCHLIST = '관심종목';
const SHEET_SAVING    = '적립식';
const SHEET_TRADES    = '거래내역';
const SHEET_MEMOS     = '투자메모';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, sheet: sheetName, row, data } = body;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 거래내역 추가 (별도 처리 — 포트폴리오 자동 업데이트 포함)
    if (action === 'addTrade') {
      return addTradeAndUpdatePortfolio(body);
    }

    // 거래내역 수정
    if (action === 'updateTrade') {
      const tradeSheet = ss.getSheetByName(SHEET_TRADES);
      if (!tradeSheet) throw new Error('거래내역 시트 없음');
      const rows = tradeSheet.getDataRange().getValues();
      const t = body.trade;
      const origDate = body.origDate;
      const origTicker = body.origTicker;
      for (let i = 1; i < rows.length; i++) {
        const rowDate = Utilities.formatDate(new Date(rows[i][0]), 'Asia/Seoul', 'yyyy-MM-dd');
        const rowTicker = String(rows[i][1]);
        if (rowDate === origDate && rowTicker === origTicker) {
          tradeSheet.getRange(i+1, 1).setValue(new Date(t.date));
          tradeSheet.getRange(i+1, 2).setNumberFormat('@');
          tradeSheet.getRange(i+1, 2).setValue("'" + t.ticker);
          tradeSheet.getRange(i+1, 3, 1, 9).setValues([[t.name, t.market, t.owner, t.account, t.type, t.qty, t.price, t.fee||0, t.memo||'']]);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    // 거래내역 삭제
    if (action === 'deleteTrade') {
      const tradeSheet = ss.getSheetByName(SHEET_TRADES);
      if (!tradeSheet) throw new Error('거래내역 시트 없음');
      const rows = tradeSheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        const rowDate = Utilities.formatDate(new Date(rows[i][0]), 'Asia/Seoul', 'yyyy-MM-dd');
        const rowTicker = String(rows[i][1]);
        if (rowDate === body.date && rowTicker === body.ticker && String(rows[i][6]) === String(body.type) && String(rows[i][7]) === String(body.qty)) {
          tradeSheet.deleteRow(i+1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    // 메모 추가
    if (action === 'addMemo') {
      const memoSheet = ss.getSheetByName(SHEET_MEMOS) || createMemosSheet(ss);
      const m = body.memo;
      memoSheet.appendRow([
        new Date(m.date || new Date()),
        m.ticker, m.name, m.market,
        m.type,   // 매수이유/손절계획/주가변동이유/매도이유
        m.content
      ]);
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    // 메모 삭제
    if (action === 'deleteMemo') {
      const memoSheet = ss.getSheetByName(SHEET_MEMOS);
      if (memoSheet) {
        const rows = memoSheet.getDataRange().getValues();
        for (let i = rows.length - 1; i >= 1; i--) {
          if (String(rows[i][1]) === String(body.ticker) &&
              Utilities.formatDate(new Date(rows[i][0]),'Asia/Seoul','yyyy-MM-dd') === body.date &&
              rows[i][4] === body.type) {
            memoSheet.deleteRow(i + 1); break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    // 현금 업데이트
    if (action === 'setCash') {
      const settingSheet = ss.getSheetByName('설정');
      settingSheet.getRange('B3').setValue(body.cash || 0);
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('시트 없음: ' + sheetName);

    if (action === 'add') {
      // A열(티커) 텍스트 형식 먼저 설정 후 행 추가 (0으로 시작하는 코드 보호)
      sheet.getRange('A:A').setNumberFormat('@');
      const addRow = sheet.getLastRow() + 1;
      sheet.getRange(addRow, 1).setNumberFormat('@');
      sheet.getRange(addRow, 1).setValue("'" + data[0]);
      if (data.length > 1) sheet.getRange(addRow, 2, 1, data.length - 1).setValues([data.slice(1)]);
    } else if (action === 'update') {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(row.ticker) &&
            String(rows[i][3]) === String(row.owner) &&
            String(rows[i][4]) === String(row.account || '')) {
          const numCols = data.length;
          sheet.getRange(i + 1, 1, 1, numCols).setValues([data]);
          // 현재가 수식 복원
          restoreFormula(sheet, sheetName, i + 1, data[0]);
          break;
        }
      }
    } else if (action === 'delete') {
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]) === String(row.ticker) &&
            String(rows[i][3]) === String(row.owner || '') &&
            String(rows[i][4]) === String(row.account || '')) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function restoreFormula(sheet, sheetName, rowNum, ticker) {
  const col = sheetName === '포트폴리오' ? 8 : sheetName === '적립식' ? 8 : 5;
  let formula = '';
  if (sheetName === '포트폴리오' || sheetName === '적립식') {
    // US 종목이면 티커 그대로, KR이면 KRX: 접두어
    const marketCol = sheetName === '포트폴리오' ? 3 : 0;
    const tickerRef = 'A' + rowNum;
    const marketVal = sheet.getRange(rowNum, marketCol === 3 ? 3 : 1).getValue();
    if (String(marketVal).toUpperCase() === 'KR') {
      formula = '=IFERROR(GOOGLEFINANCE("KRX:"&A'+rowNum+',"price"),0)';
    } else {
      formula = '=IFERROR(GOOGLEFINANCE(A'+rowNum+',"price"),0)';
    }
  } else if (sheetName === '관심종목') {
    const marketVal = sheet.getRange(rowNum, 3).getValue();
    if (String(marketVal).toUpperCase() === 'KR') {
      formula = '=IFERROR(GOOGLEFINANCE("KRX:"&A'+rowNum+',"price"),0)';
    } else {
      formula = '=IFERROR(GOOGLEFINANCE(A'+rowNum+',"price"),0)';
    }
  }
  if (formula) sheet.getRange(rowNum, col).setFormula(formula);
}

function doGet(e) {
  // 파라미터 없이 접근하면 대시보드 HTML 반환
  if (!e.parameter.action) {
    const scriptUrl = ScriptApp.getService().getUrl();
    const htmlContent = HtmlService.createHtmlOutputFromFile('index').getContent()
      .replace('REPLACE_SCRIPT_URL', scriptUrl);
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle('주식 대시보드')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const action = e.parameter.action || 'all';
  let data = {};

  // 티커 → 종목명 조회
  if (action === 'lookup') {
    const ticker = e.parameter.ticker || '';
    const market = e.parameter.market || 'US';
    data = lookupTicker(ticker, market);
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'all' || action === 'portfolio') data.portfolio = getPortfolio();
  if (action === 'all' || action === 'watchlist') data.watchlist = getWatchlist();
  if (action === 'all' || action === 'saving')    data.saving    = getSaving();
  if (action === 'all' || action === 'rate')      data.usdkrw    = getUSDKRW();
  if (action === 'all' || action === 'trades')    data.trades    = getTrades();
  if (action === 'all' || action === 'realized')  data.realized  = getRealized();
  if (action === 'all' || action === 'cash')      data.cash      = getCash();
  if (action === 'all' || action === 'memos')     data.memos     = getMemos();

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPortfolio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PORTFOLIO);
  if (!sheet) sheet = createPortfolioSheet(ss);

  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    result.push({
      ticker:  String(r[0]).padStart(r[2]==='KR'?6:0,'0'),   // A: 티커 (KR은 6자리 유지)
      name:    r[1],   // B: 종목명
      market:  r[2],   // C: 시장 (KR/US)
      owner:   r[3],   // D: 소유자
      account: r[4],   // E: 계좌
      buy:     parseFloat(r[5]) || 0,  // F: 매수가
      qty:     parseFloat(r[6]) || 0,  // G: 수량
      current: parseFloat(r[7]) || 0,  // H: 현재가 (GOOGLEFINANCE)
    });
  }
  return result;
}

function getWatchlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_WATCHLIST);
  if (!sheet) sheet = createWatchlistSheet(ss);

  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    result.push({
      ticker:  String(r[0]).padStart(r[2]==='KR'?6:0,'0'),
      name:    r[1],
      market:  r[2],
      ref:     parseFloat(r[3]) || 0,
      current: parseFloat(r[4]) || 0,
      memo:    r[5] || '',
    });
  }
  return result;
}

function lookupTicker(ticker, market) {
  if (!ticker) return { ok: false, name: '' };
  try {
    // GOOGLEFINANCE로 종목명 조회 (임시 셀 활용)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('설정');
    const lookupCell = sheet.getRange('D1');
    const nameCell   = sheet.getRange('E1');

    // 시장에 따라 티커 형식 조정
    const fullTicker = market === 'KR' ? 'KRX:' + ticker : ticker;
    lookupCell.setFormula('=IFERROR(GOOGLEFINANCE("' + fullTicker + '","name"),"")');

    // 수식 계산 대기
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const name = nameCell.getValue() || lookupCell.getValue();
    lookupCell.clearContent();

    if (name && name !== '' && name !== 0) {
      return { ok: true, name: String(name) };
    }
    return { ok: false, name: '' };
  } catch(e) {
    return { ok: false, name: '', error: e.message };
  }
}

function getSaving() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SAVING);
  if (!sheet) sheet = createSavingSheet(ss);

  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    result.push({
      ticker:       r[0],          // A: 티커
      name:         r[1],          // B: 종목명
      platform:     r[2],          // C: 플랫폼 (토스/카카오페이 등)
      owner:        r[3],          // D: 소유자
      monthlyKRW:   parseFloat(r[4]) || 0,  // E: 월 적립금액(원)
      avgBuyUSD:    parseFloat(r[5]) || 0,  // F: 평균매수가(USD) - 월 1회 수동 업데이트
      totalQty:     parseFloat(r[6]) || 0,  // G: 누적수량 - 월 1회 수동 업데이트
      current:      parseFloat(r[7]) || 0,  // H: 현재가(USD, GOOGLEFINANCE 자동)
      startDate:    r[8] ? Utilities.formatDate(new Date(r[8]), 'Asia/Seoul', 'yyyy-MM-dd') : '', // I: 시작일
      memo:         r[9] || '',     // J: 메모
    });
  }
  return result;
}

function getUSDKRW() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');
    if (sheet) return parseFloat(sheet.getRange('B2').getValue()) || 1547;
  } catch(e) {}
  return 1547;
}

function getMemos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MEMOS);
  if (!sheet) sheet = createMemosSheet(ss);
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1]) continue;
    result.push({
      date:    r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      ticker:  String(r[1]).padStart(r[3]==='KR'?6:0,'0'),
      name:    r[2],
      market:  r[3],
      type:    r[4],
      content: r[5],
    });
  }
  return result;
}

function getCash() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');
    if (sheet) return parseFloat(sheet.getRange('B3').getValue()) || 0;
  } catch(e) {}
  return 0;
}

function getTrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_TRADES);
  if (!sheet) sheet = createTradesSheet(ss);
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    result.push({
      date:    r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      ticker:  String(r[1]).padStart(r[3]==='KR'?6:0,'0'),
      name:    r[2],
      market:  r[3],
      owner:   r[4],
      account: r[5],
      type:    r[6],  // 매수/매도
      qty:     parseFloat(r[7]) || 0,
      price:   parseFloat(r[8]) || 0,
      fee:     parseFloat(r[9]) || 0,
      memo:    r[10] || '',
    });
  }
  return result;
}

// 실현손익 계산 (FIFO 방식)
function getRealized() {
  const trades = getTrades();
  const realized = {};  // key: ticker+owner

  // 종목별·소유자별 매수 큐 (FIFO)
  const queues = {};
  let totalRealized = 0;

  trades.forEach(t => {
    const key = t.ticker + '|' + t.owner;
    if (!queues[key]) queues[key] = [];

    if (t.type === '매수') {
      queues[key].push({ qty: t.qty, price: t.price });
    } else if (t.type === '매도') {
      let remainQty = t.qty;
      let costTotal = 0;
      while (remainQty > 0 && queues[key].length > 0) {
        const lot = queues[key][0];
        const useQty = Math.min(lot.qty, remainQty);
        costTotal += useQty * lot.price;
        lot.qty -= useQty;
        remainQty -= useQty;
        if (lot.qty <= 0) queues[key].shift();
      }
      const proceeds = t.qty * t.price - t.fee;
      const pnl = proceeds - costTotal;
      if (!realized[key]) realized[key] = { ticker: t.ticker, name: t.name, owner: t.owner, pnl: 0 };
      realized[key].pnl += pnl;
      totalRealized += pnl;
    }
  });

  return { items: Object.values(realized), total: totalRealized };
}

// 거래 추가 + 포트폴리오 평균매수가/수량 자동 업데이트
function addTradeAndUpdatePortfolio(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tradeSheet = ss.getSheetByName(SHEET_TRADES);
  if (!tradeSheet) tradeSheet = createTradesSheet(ss);

  const t = body.trade;
  const dateVal = t.date ? new Date(t.date) : new Date();
  // B열(티커) 텍스트 형식 먼저 설정 후 행 추가
  const tradeNewRow = tradeSheet.getLastRow() + 1;
  tradeSheet.getRange(tradeNewRow, 2).setNumberFormat('@');
  tradeSheet.getRange(tradeNewRow, 2).setValue("'" + t.ticker);
  tradeSheet.getRange(tradeNewRow, 1).setValue(dateVal);
  tradeSheet.getRange(tradeNewRow, 3, 1, 9).setValues([[t.name, t.market, t.owner, t.account, t.type, t.qty, t.price, t.fee||0, t.memo||'']]);

  // 포트폴리오 업데이트
  const pfSheet = ss.getSheetByName(SHEET_PORTFOLIO);
  if (pfSheet && (t.type === '매수' || t.type === '매도')) {
    const rows = pfSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const ticker = String(r[0]).padStart(r[2]==='KR'?6:0,'0');
      if (ticker === t.ticker && String(r[3]) === t.owner && String(r[4]) === (t.account||'')) {
        found = true;
        const curQty = parseFloat(r[6]) || 0;
        const curBuy = parseFloat(r[5]) || 0;
        let newQty, newBuy;
        if (t.type === '매수') {
          const curCost = curQty * curBuy;
          const addCost = t.qty * t.price;
          newQty = curQty + t.qty;
          newBuy = newQty > 0 ? (curCost + addCost) / newQty : 0;
        } else {
          newQty = Math.max(0, curQty - t.qty);
          newBuy = curBuy;
        }
        pfSheet.getRange(i+1, 6).setValue(Math.round(newBuy * 100) / 100);
        pfSheet.getRange(i+1, 7).setValue(Math.round(newQty * 10000) / 10000);
        // 현재가 수식 복원
        restoreFormula(pfSheet, SHEET_PORTFOLIO, i+1, t.ticker);
        break;
      }
    }

    // 포트폴리오에 없는 종목이면 새 행 추가 (매수일 때만)
    if (!found && t.type === '매수') {
      const formula = t.market === 'KR'
        ? '=IFERROR(GOOGLEFINANCE("KRX:"&A' + (pfSheet.getLastRow()+1) + ',"price"),0)'
        : '=IFERROR(GOOGLEFINANCE(A' + (pfSheet.getLastRow()+1) + ',"price"),0)';
      // A열(티커) 텍스트 형식 먼저 설정 후 행 추가 (0으로 시작하는 코드 보호)
      pfSheet.getRange('A:A').setNumberFormat('@');
      const newPfRow = pfSheet.getLastRow() + 1;
      pfSheet.getRange(newPfRow, 1).setNumberFormat('@');
      pfSheet.getRange(newPfRow, 1).setValue("'" + t.ticker);
      pfSheet.getRange(newPfRow, 2, 1, 7).setValues([[t.name, t.market, t.owner, t.account||'', t.price, t.qty, formula]]);
    }
  }

  // 거래 메모가 있으면 투자메모 시트에 자동 저장
  if (t.memo && t.memo.trim()) {
    const memoSheet = ss.getSheetByName(SHEET_MEMOS) || createMemosSheet(ss);
    memoSheet.appendRow([
      new Date(t.date || new Date()),
      t.ticker, t.name, t.market,
      t.type === '매수' ? '매수이유' : '매도이유',
      t.memo.trim()
    ]);
  }

  return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
}

// ── 시트 자동 생성 ────────────────────────────────────
function createMemosSheet(ss) {
  const sheet = ss.insertSheet(SHEET_MEMOS);
  const headers = ['날짜', '티커', '종목명', '시장', '메모유형', '내용'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#2d3748').setFontColor('white').setFontWeight('bold');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function createTradesSheet(ss) {
  const sheet = ss.insertSheet(SHEET_TRADES);
  const headers = ['날짜', '티커', '종목명', '시장', '소유자', '계좌', '구분(매수/매도)', '수량', '단가', '수수료', '메모'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a1a2e').setFontColor('white').setFontWeight('bold');
  sheet.getRange('B:B').setNumberFormat('@');  // 티커 텍스트 형식
  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd');  // 날짜 형식
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function createPortfolioSheet(ss) {
  const sheet = ss.insertSheet(SHEET_PORTFOLIO);
  const headers = ['티커', '종목명', '시장(KR/US)', '소유자', '계좌', '매수가', '수량', '현재가(자동)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a1a2e').setFontColor('white').setFontWeight('bold');

  // 예시 데이터
  const examples = [
    ['005930', '삼성전자', 'KR', '본인', '삼성증권', 70000, 10, '=IFERROR(GOOGLEFINANCE("KRX:"&A2,"price"),0)'],
    ['000660', 'SK하이닉스', 'KR', '본인', '삼성증권', 130000, 5, '=IFERROR(GOOGLEFINANCE("KRX:"&A3,"price"),0)'],
    ['AAPL', 'Apple', 'US', '배우자', '키움증권', 180, 3, '=IFERROR(GOOGLEFINANCE(A4,"price"),0)'],
    ['NVDA', 'NVIDIA', 'US', '배우자', '키움증권', 800, 2, '=IFERROR(GOOGLEFINANCE(A5,"price"),0)'],
  ];
  sheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
  // A열(티커)을 텍스트 형식으로 지정 → 앞자리 0 보존
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function createWatchlistSheet(ss) {
  const sheet = ss.insertSheet(SHEET_WATCHLIST);
  const headers = ['티커', '종목명', '시장(KR/US)', '기준가', '현재가(자동)', '메모'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a1a2e').setFontColor('white').setFontWeight('bold');

  const examples = [
    ['035720', '카카오', 'KR', 50000, '=IFERROR(GOOGLEFINANCE("KRX:"&A2,"price"),0)', '목표가 60,000'],
    ['TSLA', 'Tesla', 'US', 200, '=IFERROR(GOOGLEFINANCE(A3,"price"),0)', '변동성 주시'],
  ];
  sheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function createSettingsSheet(ss) {
  const sheet = ss.insertSheet('설정');
  sheet.getRange('A1').setValue('항목');
  sheet.getRa