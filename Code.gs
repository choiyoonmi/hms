/**
 * 해피트리 HMS V33 - 학부모 자동 발송 시스템
 * Google Apps Script용 Code.gs
 */

const CONFIG = {
  SPREADSHEET_ID: '1pfZn1ZLKqyguN7fwpjFqLn5L2osUcfc7LdWzyWCNRCw',
  APP_TITLE: '해피트리 HMS V33',
  DIRECTOR_EMAIL: 'white21040@kakao.com',
  ADMIN_PIN: '75356',
  LOGO_FILE_ID: '1vtxfk2Av6xnuk4ECuQSDjhCWNBkGOa0q',
  REPORT_ROOT_FOLDER: '🌳 해피트리 성장리포트',
};

const SHEETS = {
  STUDENTS: '학생목록',
  TEACHERS: '선생님목록',
  EVALUATIONS: '평가기록',
  ALERTS: '알림기록',
  PARENTS: '학부모정보',
  NOTICES: '공지사항',
  CONSULT_REQUESTS: '상담신청',
  SEND_HISTORY: '발송이력',
  PDF_HISTORY: 'PDF생성이력',
};

const ISSUE_MAP = {
  '0': '없음',
  '1': '숙제 부족',
  '2': '준비물 미지참',
  '3': '지각/피곤함',
  '4': '집중 부족',
  '5': '수업 참여 부족',
  '6': '이해 부족',
  '7': '오답 많음',
  '8': '계산 실수',
  '9': '해석 부족',
  '10': '서술형 약함',
  '11': '시험 불안',
  '12': '자신감 부족',
  '13': '친구 관계 상담 필요',
  '14': '학부모 상담 필요',
  '15': '칭찬할 점 있음',
  '16': '성장 모습 보임',
  '17': '노력 우수',
};

function doGet(e) {
  ensureSheets_();
  const page = e && e.parameter && e.parameter.page;
  if (page === 'parent') {
    return HtmlService.createTemplateFromFile('parent')
      .evaluate()
      .setTitle('해피트리 학부모 페이지')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialData() {
  const ss = ensureSheets_();
  return {
    today: todayKey_(),
    students: readObjects_(ss.getSheetByName(SHEETS.STUDENTS)).filter(s => text_(s['상태'] || '재원') !== '퇴원'),
    teachers: readObjects_(ss.getSheetByName(SHEETS.TEACHERS)).filter(t => text_(t['상태'] || '재직') !== '퇴사'),
    issues: ISSUE_MAP,
  };
}

function verifyAdminPin(pin) {
  return String(pin || '') === CONFIG.ADMIN_PIN;
}

function saveEvaluation(payload) {
  const ss = ensureSheets_();
  const sheet = ss.getSheetByName(SHEETS.EVALUATIONS);
  const alertSheet = ss.getSheetByName(SHEETS.ALERTS);
  const now = new Date();
  const timestamp = formatDateTime_(now);
  const dateKey = formatDate_(now);

  const habit = Number(payload.learningHabit || payload.homework || 0);
  const focus = Number(payload.focus || 0);
  const understanding = Number(payload.understanding || 0);
  const performance = Number(payload.performance || payload.attitude || 0);
  const total = habit + focus + understanding + performance;
  const avg = Math.round((total / 4) * 10) / 10;
  const issueCodes = Array.isArray(payload.issueCodes) && payload.issueCodes.length ? payload.issueCodes.map(String) : ['0'];
  const issueLabels = issueCodes.map(c => ISSUE_MAP[c] || c).join(', ');
  const status = decideStatus_(avg, issueCodes, focus, understanding);
  const comment = makeAutoComment_({
    studentName: payload.studentName,
    subject: payload.subject,
    habit,
    focus,
    understanding,
    performance,
    issueCodes,
    issueLabels,
    status,
  });

  const row = [
    timestamp,
    dateKey,
    payload.teacherName || '',
    payload.subject || '',
    payload.className || '',
    payload.studentName || '',
    habit,
    focus,
    understanding,
    performance,
    issueCodes.join(','),
    issueLabels,
    total,
    avg,
    status,
    payload.memo || '',
    comment,
  ];
  sheet.appendRow(row);

  if (status === '집중관리' || issueCodes.includes('14')) {
    const message = makeAlertMessage_(row);
    try {
      MailApp.sendEmail(CONFIG.DIRECTOR_EMAIL, '🚨 [해피트리 성장기록] 집중관리 학생 발생', message);
      alertSheet.appendRow([timestamp, payload.studentName || '', status, message, CONFIG.DIRECTOR_EMAIL, '발송완료']);
    } catch (e) {
      alertSheet.appendRow([timestamp, payload.studentName || '', status, message, CONFIG.DIRECTOR_EMAIL, '발송실패: ' + e.message]);
    }
  }
  return { ok: true, total, avg, status, comment };
}

function getDashboardData() {
  const ss = ensureSheets_();
  const evalRows = readObjects_(ss.getSheetByName(SHEETS.EVALUATIONS));
  const students = readObjects_(ss.getSheetByName(SHEETS.STUDENTS)).filter(s => text_(s['상태'] || '재원') !== '퇴원');
  const today = todayKey_();
  const todayRows = evalRows.filter(r => normalizeDateKey_(r['수업일자']) === today);

  const urgent = todayRows.filter(r => normalizeStatus_(r['상태']) === '집중관리');
  const watch = todayRows.filter(r => normalizeStatus_(r['상태']) === '관심관리');
  const normal = todayRows.filter(r => normalizeStatus_(r['상태']) === '정상');

  const classMap = {};
  students.forEach(s => {
    const cls = text_(s['반명'] || '미지정');
    if (!classMap[cls]) classMap[cls] = { className: cls, active: 0, recorded: 0, urgent: 0, watch: 0, normal: 0, percent: 0 };
    classMap[cls].active++;
  });
  const recordedSet = new Set();
  todayRows.forEach(r => {
    const name = text_(r['학생명']);
    const cls = text_(r['반명'] || '미지정');
    if (!name) return;
    recordedSet.add(name);
    if (!classMap[cls]) classMap[cls] = { className: cls, active: 0, recorded: 0, urgent: 0, watch: 0, normal: 0, percent: 0 };
    classMap[cls].recorded++;
    const st = normalizeStatus_(r['상태']);
    if (st === '집중관리') classMap[cls].urgent++;
    else if (st === '관심관리') classMap[cls].watch++;
    else classMap[cls].normal++;
  });
  Object.values(classMap).forEach(c => c.percent = c.active ? Math.round((c.recorded / c.active) * 100) : 0);

  const teacherMap = {};
  todayRows.forEach(r => {
    const t = text_(r['선생님명'] || '미지정');
    if (!teacherMap[t]) teacherMap[t] = { teacherName: t, count: 0, urgent: 0, watch: 0, normal: 0 };
    teacherMap[t].count++;
    const st = normalizeStatus_(r['상태']);
    if (st === '집중관리') teacherMap[t].urgent++;
    else if (st === '관심관리') teacherMap[t].watch++;
    else teacherMap[t].normal++;
  });

  const waitingStudents = students
    .filter(s => !recordedSet.has(text_(s['학생명'])))
    .map(s => ({ studentName: s['학생명'], className: s['반명'] }));

  const summary = makeStudentSummary_(evalRows);

  return serialize_({
    ok: true,
    today,
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    totalToday: todayRows.length,
    urgent,
    watch,
    normal,
    classStats: Object.values(classMap).sort((a, b) => a.className.localeCompare(b.className, 'ko')),
    teacherStats: Object.values(teacherMap).sort((a, b) => b.count - a.count),
    waitingStudents,
    summary,
    repeatedRisk: summary.filter(s => s.recentAvg <= 2.8 || s.lowFocus >= 2 || s.parentNeed >= 1),
    issueStats: getIssueStats_(evalRows, 30),
  });
}

function getMonthlyReport(studentName, monthKey) {
  const ss = ensureSheets_();
  const rows = readObjects_(ss.getSheetByName(SHEETS.EVALUATIONS))
    .filter(r => text_(r['학생명']) === text_(studentName) && normalizeDateKey_(r['수업일자']).slice(0, 7) === monthKey);
  const student = findStudent_(studentName);
  const count = rows.length;
  const avg = field => count ? round1_(rows.reduce((s, r) => s + Number(r[field] || 0), 0) / count) : 0;
  const avgHabit = avg('학습습관') || avg('숙제');
  const avgFocus = avg('집중도');
  const avgUnderstanding = avg('이해도');
  const avgPerformance = avg('수행도') || avg('태도');
  const avgTotal = count ? round1_((avgHabit + avgFocus + avgUnderstanding + avgPerformance) / 4) : 0;
  const temp = growthTemp_(avgTotal);
  const issueStats = getIssueStats_(rows, 999);
  const latest = rows.slice(-10).reverse();
  const comment = makeMonthlyComment_({ studentName, monthKey, count, avgHabit, avgFocus, avgUnderstanding, avgPerformance, avgTotal, temp, issueStats });
  const pdf = getLatestPdf_(studentName, monthKey);
  return serialize_({
    ok: true,
    studentName,
    className: student['반명'] || '',
    monthKey,
    count,
    avgHabit,
    avgFocus,
    avgUnderstanding,
    avgPerformance,
    avgTotal,
    temp,
    issueStats,
    latest,
    comment,
    pdfUrl: pdf ? pdf.fileUrl : '',
    pdfName: pdf ? pdf.fileName : '',
  });
}

function createMonthlyPdf(studentName, monthKey) {
  const report = getMonthlyReport(studentName, monthKey);
  const html = buildReportHtml_(report);
  const blob = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf');
  const fileName = `해피트리_성장리포트_${studentName}_${monthKey}.pdf`;
  blob.setName(fileName);
  const folder = getMonthFolder_(monthKey);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const ss = ensureSheets_();
  ss.getSheetByName(SHEETS.PDF_HISTORY).appendRow([formatDateTime_(new Date()), studentName, monthKey, fileName, file.getUrl(), folder.getUrl()]);
  return { ok: true, fileName, fileUrl: file.getUrl(), folderUrl: folder.getUrl() };
}

function makeParentMessage(studentName, monthKey) {
  const report = getMonthlyReport(studentName, monthKey);
  let pdfUrl = report.pdfUrl;
  if (!pdfUrl) {
    try { pdfUrl = createMonthlyPdf(studentName, monthKey).fileUrl; } catch (e) { pdfUrl = ''; }
  }
  const portal = getParentPortalUrl_(studentName);
  const text = `🌳 해피트리 성장리포트\n\n${studentName} 학생의 ${monthKey} 성장리포트가 준비되었습니다.\n\n성장온도: ${report.temp}\n평균 성장지수: ${report.avgTotal}\n평가 횟수: ${report.count}회\n\n${report.comment}\n\n📄 PDF 보기\n${pdfUrl || 'PDF 링크 준비 중'}\n\n📱 학부모 페이지\n${portal}\n\n- 해피트리학원 -`;
  return { ok: true, text, pdfUrl, portalUrl: portal, report };
}

function sendMonthlyReportEmail(studentName, monthKey) {
  const ss = ensureSheets_();
  const parent = getParentInfo_(studentName);
  if (!parent || !parent.email) throw new Error(`${studentName} 학생의 학부모 이메일이 없습니다. 학부모정보 시트를 확인하세요.`);
  const message = makeParentMessage(studentName, monthKey);
  const subject = `🌳 [해피트리] ${studentName} ${monthKey} 성장리포트`;
  MailApp.sendEmail(parent.email, subject, message.text);
  ss.getSheetByName(SHEETS.SEND_HISTORY).appendRow([formatDateTime_(new Date()), studentName, monthKey, parent.parentName || '', parent.email, '이메일', '발송완료', message.pdfUrl, message.portalUrl]);
  return { ok: true, studentName, email: parent.email, status: '발송완료' };
}

function bulkSendMonthlyReportEmails(monthKey, className) {
  const ss = ensureSheets_();
  const students = readObjects_(ss.getSheetByName(SHEETS.STUDENTS))
    .filter(s => text_(s['상태'] || '재원') !== '퇴원')
    .filter(s => !className || className === '전체' || text_(s['반명']) === text_(className));
  const results = [];
  students.forEach(s => {
    try {
      const report = getMonthlyReport(s['학생명'], monthKey);
      if (!report.count) {
        results.push({ studentName: s['학생명'], status: '기록없음' });
        return;
      }
      results.push(sendMonthlyReportEmail(s['학생명'], monthKey));
    } catch (e) {
      ss.getSheetByName(SHEETS.SEND_HISTORY).appendRow([formatDateTime_(new Date()), s['학생명'], monthKey, '', '', '이메일', '발송실패: ' + e.message, '', '']);
      results.push({ studentName: s['학생명'], status: '실패', message: e.message });
    }
  });
  return { ok: true, results };
}

function getSendCenterData(monthKey, className) {
  const ss = ensureSheets_();
  const students = readObjects_(ss.getSheetByName(SHEETS.STUDENTS))
    .filter(s => text_(s['상태'] || '재원') !== '퇴원')
    .filter(s => !className || className === '전체' || text_(s['반명']) === text_(className));
  const parents = readObjects_(ss.getSheetByName(SHEETS.PARENTS));
  const history = readObjects_(ss.getSheetByName(SHEETS.SEND_HISTORY));
  const rows = students.map(s => {
    const name = s['학생명'];
    const report = getMonthlyReport(name, monthKey);
    const parent = parents.find(p => text_(p['학생명']) === text_(name)) || {};
    const sent = history.filter(h => text_(h['학생명']) === text_(name) && text_(h['월']) === monthKey && text_(h['상태']).indexOf('완료') > -1).length;
    return { studentName: name, className: s['반명'], count: report.count, temp: report.temp, avgTotal: report.avgTotal, parentName: parent['보호자명'] || '', email: parent['이메일'] || '', sent };
  });
  return { ok: true, monthKey, className: className || '전체', rows };
}

function parentLogin(studentName, pin) {
  const parent = getParentInfo_(studentName);
  if (!parent) return { ok: false, message: '학부모정보가 없습니다.' };
  if (String(parent.pin || '') !== String(pin || '')) return { ok: false, message: '인증번호가 맞지 않습니다.' };
  return { ok: true, studentName: parent.studentName };
}

function getParentPortalData(studentName) {
  const monthKey = todayKey_().slice(0, 7);
  const report = getMonthlyReport(studentName, monthKey);
  const notices = readObjects_(ensureSheets_().getSheetByName(SHEETS.NOTICES)).slice(-5).reverse();
  return { ok: true, report, notices };
}

function submitParentConsultation(payload) {
  const ss = ensureSheets_();
  ss.getSheetByName(SHEETS.CONSULT_REQUESTS).appendRow([
    formatDateTime_(new Date()),
    payload.studentName || '',
    payload.parentName || '',
    payload.message || '',
    payload.preferredDate || '',
    payload.preferredTime || '',
    '대기',
  ]);
  return { ok: true };
}

function ensureSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  setupSheet_(ss, SHEETS.STUDENTS, ['학생명', '반명', '학년', '상태']);
  setupSheet_(ss, SHEETS.TEACHERS, ['선생님명', '담당과목', '상태']);
  setupSheet_(ss, SHEETS.EVALUATIONS, ['기록시간', '수업일자', '선생님명', '과목', '반명', '학생명', '학습습관', '집중도', '이해도', '수행도', '특이사항코드', '특이사항', '총점', '평균', '상태', '메모', '자동코멘트']);
  setupSheet_(ss, SHEETS.ALERTS, ['알림시간', '학생명', '상태', '알림내용', '수신메일', '발송상태']);
  setupSheet_(ss, SHEETS.PARENTS, ['학생명', '보호자명', '연락처', '비밀번호', '이메일', '상태']);
  setupSheet_(ss, SHEETS.NOTICES, ['날짜', '제목', '내용', '상태']);
  setupSheet_(ss, SHEETS.CONSULT_REQUESTS, ['신청시간', '학생명', '보호자명', '상담내용', '희망날짜', '희망시간', '상태']);
  setupSheet_(ss, SHEETS.SEND_HISTORY, ['발송시간', '학생명', '월', '보호자명', '이메일', '발송유형', '상태', 'PDF링크', '포털링크']);
  setupSheet_(ss, SHEETS.PDF_HISTORY, ['생성시간', '학생명', '월', '파일명', '파일링크', '폴더링크']);
  return ss;
}

function setupSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eef3ff');
    sheet.setFrozenRows(1);
  }
}

function readObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => text_(h));
  return values.slice(1).filter(r => r.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = serializeValue_(row[i]));
    return obj;
  });
}

function findStudent_(studentName) {
  return readObjects_(ensureSheets_().getSheetByName(SHEETS.STUDENTS)).find(s => text_(s['학생명']) === text_(studentName)) || {};
}

function getParentInfo_(studentName) {
  const p = readObjects_(ensureSheets_().getSheetByName(SHEETS.PARENTS)).find(x => text_(x['학생명']) === text_(studentName));
  if (!p) return null;
  return { studentName: p['학생명'], parentName: p['보호자명'], phone: p['연락처'], pin: p['비밀번호'], email: p['이메일'] };
}

function makeStudentSummary_(rows) {
  const map = {};
  rows.forEach(r => {
    const name = text_(r['학생명']);
    if (!name) return;
    if (!map[name]) map[name] = [];
    map[name].push(r);
  });
  return Object.keys(map).map(name => {
    const all = map[name];
    const latest = all.slice(-5);
    const avg = latest.reduce((s, r) => s + (Number(r['평균']) || Number(r['총점']) / 4 || 0), 0) / Math.max(latest.length, 1);
    return {
      studentName: name,
      className: latest[latest.length - 1]['반명'] || '',
      count: all.length,
      recentAvg: round1_(avg),
      lastStatus: normalizeStatus_(latest[latest.length - 1]['상태']),
      lowFocus: latest.filter(r => Number(r['집중도'] || 0) <= 2).length,
      parentNeed: latest.filter(r => hasIssue_(r['특이사항코드'], '14')).length,
      temp: growthTemp_(avg),
    };
  }).sort((a, b) => a.recentAvg - b.recentAvg);
}

function getIssueStats_(rows, days) {
  const since = new Date();
  since.setDate(since.getDate() - days);
  const count = {};
  rows.forEach(r => {
    const d = new Date(normalizeDateKey_(r['수업일자']));
    if (days < 900 && d < since) return;
    String(r['특이사항코드'] || '').split(',').map(x => x.trim()).filter(x => x && x !== '0').forEach(code => {
      const label = ISSUE_MAP[code] || code;
      count[label] = (count[label] || 0) + 1;
    });
  });
  return Object.keys(count).map(label => ({ label, count: count[label] })).sort((a, b) => b.count - a.count);
}

function getLatestPdf_(studentName, monthKey) {
  const rows = readObjects_(ensureSheets_().getSheetByName(SHEETS.PDF_HISTORY))
    .filter(r => text_(r['학생명']) === text_(studentName) && text_(r['월']) === monthKey);
  if (!rows.length) return null;
  const last = rows[rows.length - 1];
  return { fileName: last['파일명'], fileUrl: last['파일링크'] };
}

function getParentPortalUrl_(studentName) {
  const base = ScriptApp.getService().getUrl();
  return `${base}?page=parent&student=${encodeURIComponent(studentName)}`;
}

function getMonthFolder_(monthKey) {
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.REPORT_ROOT_FOLDER);
  return getOrCreateFolder_(root, monthKey);
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function buildReportHtml_(r) {
  const logo = getLogoDataUri_();
  const latestRows = (r.latest || []).map(x => `<tr><td>${x['수업일자'] || ''}</td><td>${x['과목'] || ''}</td><td>${x['학습습관'] || x['숙제'] || ''}</td><td>${x['집중도'] || ''}</td><td>${x['이해도'] || ''}</td><td>${x['수행도'] || x['태도'] || ''}</td><td>${x['상태'] || ''}</td></tr>`).join('');
  const issueRows = (r.issueStats || []).slice(0, 6).map(i => `<li>${i.label} ${i.count}회</li>`).join('') || '<li>특이사항 없음</li>';
  return `<!doctype html><html><head><meta charset="UTF-8"><style>
  @page{size:A4 landscape;margin:16mm}body{font-family:Arial,'Noto Sans KR',sans-serif;color:#11284a}.box{border:3px solid #12345b;border-radius:18px;padding:22px}.head{display:flex;gap:18px;align-items:center;border-bottom:2px solid #12345b;padding-bottom:16px}.logo{width:72px;height:72px;object-fit:contain}.title{font-size:34px;font-weight:900}.sub{color:#6b7280}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:16px 0}.card{border:1px solid #d8e0ef;border-radius:12px;padding:14px}.num{font-size:28px;font-weight:900}.cols{display:grid;grid-template-columns:1.2fr 1fr;gap:14px}.panel{border:1px solid #d8e0ef;border-radius:12px;padding:14px}table{width:100%;border-collapse:collapse;font-size:12px}td,th{border:1px solid #e5e7eb;padding:6px}h3{margin:0 0 8px}.foot{text-align:center;margin-top:14px;font-size:12px;color:#6b7280}</style></head><body><div class="box"><div class="head">${logo ? `<img class="logo" src="${logo}">` : ''}<div><div class="sub">HAPPYTREE GROWTH RECORD · V33</div><div class="title">해피트리 성장리포트</div><div class="sub">오늘의 작은 기록이 아이의 큰 성장을 만듭니다.</div></div><div style="margin-left:auto;text-align:right"><b>학생 ${r.studentName}</b><br>기간 ${r.monthKey}<br>${r.temp}</div></div><div class="grid"><div class="card">학습습관<div class="num">${r.avgHabit}</div></div><div class="card">집중도<div class="num">${r.avgFocus}</div></div><div class="card">이해도<div class="num">${r.avgUnderstanding}</div></div><div class="card">수행도<div class="num">${r.avgPerformance}</div></div></div><div class="cols"><div class="panel"><h3>학부모 공유용 성장 코멘트</h3><p>${r.comment}</p><h3>최근 성장기록</h3><table><tr><th>날짜</th><th>과목</th><th>습관</th><th>집중</th><th>이해</th><th>수행</th><th>상태</th></tr>${latestRows}</table></div><div><div class="panel"><h3>특이사항 통계</h3><ul>${issueRows}</ul></div><div class="panel" style="margin-top:14px"><h3>월간 요약</h3><p>평균 성장지수 ${r.avgTotal} · 평가 ${r.count}회</p></div></div></div><div class="foot">해피트리학원 · 해피트리는 매일의 작은 변화를 기록하며 학생의 성장을 함께 만들어가겠습니다. 🌳</div></div></body></html>`;
}

function getLogoDataUri_() {
  if (!CONFIG.LOGO_FILE_ID) return '';
  try {
    const blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (e) {
    return '';
  }
}

function decideStatus_(avg, codes, focus, understanding) {
  if (codes.includes('14') || codes.includes('13') || (codes.includes('4') && codes.includes('6')) || avg <= 2.5 || focus <= 1 || understanding <= 1) return '집중관리';
  if (avg <= 3.5 || codes.includes('1') || codes.includes('4') || codes.includes('6')) return '관심관리';
  return '정상';
}

function makeAutoComment_(r) {
  const positives = r.issueCodes.filter(c => ['15', '16', '17'].includes(c)).map(c => ISSUE_MAP[c]);
  const needs = r.issueCodes.filter(c => !['0', '15', '16', '17'].includes(c)).map(c => ISSUE_MAP[c]);
  let subjectText = '';
  if (r.subject === '국어') subjectText = '독해와 서술 표현을 중심으로';
  else if (r.subject === '영어') subjectText = '해석과 문장 적용을 중심으로';
  else if (r.subject === '수학') subjectText = '개념 이해와 문제 해결 과정을 중심으로';
  else subjectText = '오늘 수업에서';
  if (r.status === '정상') return `${subjectText} 안정적인 학습 흐름을 보였습니다. ${positives.length ? positives.join(', ') + '이 돋보입니다.' : '현재의 학습 습관을 잘 유지하면 좋겠습니다.'}`;
  if (r.status === '관심관리') return `${subjectText} 조금 더 관리가 필요한 모습이 있었습니다. ${needs.length ? needs.join(', ') + ' 부분을 함께 점검하겠습니다.' : '학습 흐름을 다시 잡아가겠습니다.'}`;
  return `${subjectText} 집중적인 관리가 필요한 기록이 확인되었습니다. ${needs.length ? needs.join(', ') + ' 부분을 우선적으로 살피겠습니다.' : '가정과의 협조가 필요할 수 있습니다.'}`;
}

function makeMonthlyComment_(r) {
  if (!r.count) return `${r.studentName} 학생은 ${r.monthKey} 성장기록이 아직 충분히 쌓이지 않았습니다.`;
  const weak = [];
  if (r.avgHabit < 3.2) weak.push('학습습관');
  if (r.avgFocus < 3.2) weak.push('집중도');
  if (r.avgUnderstanding < 3.2) weak.push('이해도');
  if (r.avgPerformance < 3.2) weak.push('수행도');
  if (!weak.length) return `${r.studentName} 학생은 이번 달 전반적으로 안정적인 성장 흐름을 보였습니다. 학습습관, 집중도, 이해도, 수행도가 고르게 유지되고 있어 긍정적인 성장이 기대됩니다.`;
  return `${r.studentName} 학생은 이번 달 ${weak.join(', ')} 부분에서 조금 더 세심한 관리가 필요합니다. 해피트리는 매 수업의 작은 변화를 기록하며 안정적인 성장 흐름을 만들어가겠습니다.`;
}

function makeAlertMessage_(row) {
  return `🚨 해피트리 집중관리 학생 발생\n\n학생: ${row[5]}\n반: ${row[4]}\n과목: ${row[3]}\n선생님: ${row[2]}\n학습습관: ${row[6]} / 집중도: ${row[7]} / 이해도: ${row[8]} / 수행도: ${row[9]}\n특이사항: ${row[11]}\n총점: ${row[12]} / 평균: ${row[13]}\n상태: ${row[14]}\n메모: ${row[15] || '-'}`;
}

function hasIssue_(codes, target) {
  return String(codes || '').split(',').map(x => x.trim()).includes(String(target));
}

function normalizeStatus_(s) {
  const v = text_(s);
  if (v === '긴급관리') return '집중관리';
  return v;
}

function growthTemp_(avg) {
  const n = Number(avg || 0);
  if (n >= 3.8) return '🟢 안정 성장';
  if (n >= 2.8) return '🟡 관심 성장';
  return '🔴 집중 관리';
}

function todayKey_() { return formatDate_(new Date()); }
function formatDate_(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'yyyy-MM-dd'); }
function formatDateTime_(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function normalizeDateKey_(v) { if (!v) return ''; try { return formatDate_(v); } catch(e) { return String(v).slice(0,10); } }
function text_(v) { return String(v == null ? '' : v).trim(); }
function round1_(n) { return Math.round(Number(n || 0) * 10) / 10; }
function serializeValue_(v) { return v instanceof Date ? formatDate_(v) : v; }
function serialize_(obj) { return JSON.parse(JSON.stringify(obj)); }
