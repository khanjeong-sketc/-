// =============================================
// AppState - 전역 상태
// =============================================
const AppState = {
  employees: [],
  leaderReviews: [],
  competencyReviews: [],
  peerFeedbacks: [],
  upwardFeedbacks: [],
  historyData: [],
  extraSheets: [],   // 추가 시트 데이터 [{sheetName, headers, rows}]
  divisionName: '',
  currentIndex: 0,
};

const chartInstances = {};

// =============================================
// 페이지 전환
// =============================================
function showPage(pageName) {
  document.getElementById('upload-page').style.display = 'none';
  document.getElementById('summary-page').style.display = 'none';
  document.getElementById('detail-page').style.display = 'none';
  document.getElementById('header-nav').style.display = 'none';
  document.getElementById('header-breadcrumb').textContent = '';

  if (pageName === 'upload') {
    document.getElementById('upload-page').style.display = 'flex';
  } else if (pageName === 'summary') {
    document.getElementById('summary-page').style.display = 'block';
    document.getElementById('header-breadcrumb').textContent = '부문 요약';
    renderSummaryPage();
  } else if (pageName === 'detail') {
    document.getElementById('detail-page').style.display = 'block';
    document.getElementById('header-nav').style.display = 'flex';
    renderDetailPage(AppState.currentIndex);
  }
}

function navigateDetail(index) {
  AppState.currentIndex = index;
  showPage('detail');
}

function navigatePrev() {
  if (AppState.currentIndex > 0) {
    AppState.currentIndex--;
    renderDetailPage(AppState.currentIndex);
  }
}

function navigateNext() {
  if (AppState.currentIndex < AppState.employees.length - 1) {
    AppState.currentIndex++;
    renderDetailPage(AppState.currentIndex);
  }
}

function updateNavButtons() {
  const idx = AppState.currentIndex;
  const total = AppState.employees.length;
  document.getElementById('btn-prev').disabled = idx === 0;
  document.getElementById('btn-next').disabled = idx === total - 1;
  document.getElementById('nav-counter').textContent = `${idx + 1} / ${total}`;
  const emp = AppState.employees[idx];
  document.getElementById('header-breadcrumb').textContent = emp ? emp.name : '';
}

// =============================================
// 파일 업로드
// =============================================
const uploadedFiles = { main: null, history: null };

function initUploader() {
  const inputMain = document.getElementById('file-input-main');
  const inputHistory = document.getElementById('file-input-history');

  inputMain.addEventListener('change', (e) => {
    if (e.target.files[0]) assignFile('main', e.target.files[0]);
  });
  inputHistory.addEventListener('change', (e) => {
    if (e.target.files[0]) assignFile('history', e.target.files[0]);
  });
}

function isHistoryFileName(name) {
  return /이력|대시보드|history|hist/i.test(name);
}

function assignFile(type, file) {
  uploadedFiles[type] = file;
  const label = type === 'main' ? 'file-name-main' : 'file-name-history';
  const el = document.getElementById(label);
  el.textContent = file.name;
  el.style.color = '#166534';
  // 메인 파일이 있으면 버튼 활성화
  document.getElementById('btn-start').disabled = !uploadedFiles.main;
  hideUploadError();
}

function processFiles() {
  if (!uploadedFiles.main) { showUploadError('올해 성과 평가 파일을 선택해주세요.'); return; }
  hideUploadError();
  showLoading(true);

  const readFile = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

  const tasks = [readFile(uploadedFiles.main)];
  if (uploadedFiles.history) tasks.push(readFile(uploadedFiles.history));

  Promise.all(tasks).then(([mainData, histData]) => {
    try {
      const mainWb = XLSX.read(mainData, { type: 'array', cellDates: true });

      // 파일명에서 부문명 추출: '2025년_통합 업무 성과 리뷰_미디어 사업' → '미디어 사업'
      const fileName = uploadedFiles.main.name.replace(/\.xlsx?$/i, '');
      const divMatch = fileName.match(/_([^_]+)$/);
      AppState.divisionName = divMatch ? divMatch[1].trim() : fileName;
      const result = parseWorkbook(mainWb);
      if (result.error) { showUploadError(result.error); showLoading(false); return; }

      if (histData) {
        const histWb = XLSX.read(histData, { type: 'array', cellDates: true });
        parseHistoryWorkbook(histWb);
      }

      showLoading(false);
      showPage('summary');
    } catch (err) {
      showUploadError('파일을 읽는 중 오류가 발생했습니다: ' + err.message);
      showLoading(false);
    }
  }).catch(err => {
    showUploadError('파일 읽기 실패: ' + err.message);
    showLoading(false);
  });
}

function showUploadError(msg) {
  const el = document.getElementById('upload-error');
  el.textContent = msg;
  el.style.display = 'block';
}
function hideUploadError() { document.getElementById('upload-error').style.display = 'none'; }
function showLoading(show) { document.getElementById('upload-loading').style.display = show ? 'block' : 'none'; }

// =============================================
// 엑셀 파싱
// =============================================
const SHEET_PATTERNS = {
  mainReview:    ['통합 리뷰', 'HQ_통합 리뷰', '미디어_통합 리뷰', 'AP개발_통합 리뷰', 'PM_통합 리뷰',
                  'Hwan_통합 리뷰', '세일즈 조직_통합 리뷰', '전사 마케팅, R&D 본부_통합 리뷰'],
  leaderSelf:    ['리더 셀프 리뷰', 'PM 셀프 리뷰'],
  competency:    ['역량 리뷰'],
  peerFeedback:  ['성장 피드백 (동료)', '성장 피드백(동료)', '동료 피드백'],
  upwardFeedback:['상향 피드백 (리더)', '상향 피드백(리더)', '상향 피드백 (Hwan님)', '상향 피드백'],
  // 알려진 시트 패턴 (추가 시트 감지 제외용)
  knownSheets:   ['통합 리뷰', '리더 셀프 리뷰', 'PM 셀프 리뷰', '역량 리뷰',
                  '성장 피드백', '동료 피드백', '상향 피드백', '리더 피드백',
                  '핵심인재보상', '3개년 성과'],
};

function findSheet(workbook, patterns) {
  for (const pattern of patterns) {
    if (workbook.SheetNames.includes(pattern)) return workbook.Sheets[pattern];
  }
  // 부분 매칭: 패턴이 시트명에 포함되거나, 시트명이 패턴으로 시작하는 경우
  for (const name of workbook.SheetNames) {
    for (const pattern of patterns) {
      if (name.includes(pattern) || pattern.includes(name)) return workbook.Sheets[name];
    }
  }
  // 앞 4글자 매칭 (fallback)
  for (const name of workbook.SheetNames) {
    for (const pattern of patterns) {
      const key = pattern.replace(/\s/g, '').substring(0, 4);
      if (key && name.replace(/\s/g, '').includes(key)) return workbook.Sheets[name];
    }
  }
  return null;
}

// 병합 셀을 언머지하여 첫 번째 셀 값을 나머지에 복사
// - 세로 병합(같은 컬럼, 여러 행): 항상 덮어씀 (부서 등 세로 병합 처리)
// - 가로 병합(같은 행, 여러 컬럼): 빈 셀만 채움 (헤더 병합 - 오른쪽은 서술형 칸)
function unmergeSheet(sheet) {
  const merges = sheet['!merges'];
  if (!merges) return sheet;
  merges.forEach(merge => {
    const firstCell = sheet[XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c })];
    if (!firstCell) return;
    const isVertical = merge.e.r > merge.s.r; // 세로 병합 여부
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        if (r === merge.s.r && c === merge.s.c) continue;
        const addr = XLSX.utils.encode_cell({ r, c });
        if (isVertical) {
          // 세로 병합: 항상 덮어씀
          sheet[addr] = { ...firstCell };
        } else {
          // 가로 병합: 빈 셀만 채움 (오른쪽 칸은 서술형 데이터 보존)
          if (!sheet[addr]) sheet[addr] = { ...firstCell };
        }
      }
    }
  });
  return sheet;
}

function sheetToRowsFromHeader(sheet, headerRowIndex) {
  unmergeSheet(sheet);
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const headers = [];
  const headerCount = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: headerRowIndex, c })];
    let name = cell ? String(cell.v).trim() : '';
    if (name) {
      headerCount[name] = (headerCount[name] || 0) + 1;
      if (headerCount[name] > 1) name = `${name}_${headerCount[name] - 1}`;
    }
    headers.push(name);
  }
  const rows = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
    const row = {};
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      const key = headers[c - range.s.c];
      if (key) {
        row[key] = cell ? cell.v : null;
        if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') hasData = true;
      }
    }
    if (hasData) rows.push(row);
  }
  return { headers, rows };
}

function parseWorkbook(workbook) {
  const mainSheet = findSheet(workbook, SHEET_PATTERNS.mainReview);
  if (!mainSheet) return { error: "'통합 리뷰' 시트를 찾을 수 없습니다." };

  // 사람 이름 판별: 한글 2~5자 또는 영문 이름 형태만 허용
  function isPersonName(val) {
    if (!val) return false;
    const s = String(val).trim();
    if (/^[가-힣]{2,5}$/.test(s)) return true;           // 한글 이름
    if (/^[가-힣]{2,5}\s*\([\w\s]+\)$/.test(s)) return true; // 홍길동(영문) 형태
    if (/^[A-Za-z][\w\s\-\.]{1,20}$/.test(s) && s.split(' ').length <= 4) return true; // 영문 이름
    return false;
  }

  const { rows: mainRows, headers: mainHeaders } = sheetToRowsFromHeader(mainSheet, 12);

  // 행12(0-indexed: 11)에서 그룹 헤더 읽기 → 리뷰 섹션 이름 동적 결정
  // 행12: ['', '전체인원...', '', '', '', '', '', '1차 의견', '', '1차 종합', ...]
  // 행13: ['', '소속', '성명', ..., '달성도', '리뷰', '평가 등급', '리뷰', ...]
  const groupHeaders = [];
  {
    const range = XLSX.utils.decode_range(mainSheet['!ref'] || 'A1:A1');
    let lastGroup = '';
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = mainSheet[XLSX.utils.encode_cell({ r: 11, c })]; // 행12 (0-indexed: 11)
      const val = cell ? String(cell.v).trim() : '';
      if (val && val !== lastGroup) lastGroup = val;
      groupHeaders.push(lastGroup);
    }
  }

  // ── 동적 섹션 파싱 ──────────────────────────────────────────
  // 행12(그룹헤더)와 행13(컬럼헤더)를 조합해서 섹션 목록 구성
  // 섹션 = { label, achieveKey, gradeKey, reviewKey }
  // 3차종합/최종등급/2차종합 섹션은 기본 reviewSections에서 제외 (별도 파싱)
  const skipPattern = /3차|최종|2차\s*종합/;

  // 그룹헤더에서 섹션 경계 찾기: 비어있지 않은 그룹헤더가 바뀌는 지점
  // 달성도 컬럼이 있는 인덱스부터 시작
  const achievStartIdx = mainHeaders.findIndex(h => h && h.includes('달성도'));
  const reviewSections = [];
  const reviewSectionLabels = []; // 요약 테이블 헤더용

  if (achievStartIdx >= 0) {
    // 그룹헤더 기준으로 섹션 분리
    const sectionBoundaries = [];
    let lastLabel = '';
    for (let i = achievStartIdx; i < mainHeaders.length; i++) {
      const g = groupHeaders[i] || '';
      if (g && g !== lastLabel) {
        sectionBoundaries.push({ label: g, startIdx: i });
        lastLabel = g;
      }
    }

    sectionBoundaries.forEach(({ label, startIdx }) => {
      if (skipPattern.test(label)) return;

      // 이 섹션 범위: startIdx ~ 다음 섹션 시작 전
      const nextBoundary = sectionBoundaries.find(b => b.startIdx > startIdx);
      const endIdx = nextBoundary ? nextBoundary.startIdx : mainHeaders.length;

      // 섹션 내에서 달성도, 등급, 리뷰 키 찾기
      let achieveKey = null, gradeKey = null, reviewKey = null;
      for (let i = startIdx; i < endIdx; i++) {
        const h = mainHeaders[i];
        if (!h) continue;
        if (h.includes('달성도') && !achieveKey) achieveKey = h;
        else if ((h.includes('등급') || h === '평가 등급') && !gradeKey) gradeKey = h;
        else if (!reviewKey) {
          // 리뷰 컬럼: '리뷰', '리뷰_N', '업무 리뷰' 모두 포함
          if (h === '리뷰' || /^리뷰[_\.]/.test(h) || h === '업무 리뷰') reviewKey = h;
        }
      }

      // 달성도나 리뷰 중 하나라도 있으면 섹션 추가
      if (achieveKey || reviewKey) {
        const cleanLabel = label.replace(/전체 인원 \d+명/, '').trim();
        reviewSections.push({ label: cleanLabel, achieveKey, gradeKey, reviewKey });
        reviewSectionLabels.push(cleanLabel);
      }
    });
  }

  // ── 3차 종합 / 최종 등급 섹션 별도 파싱 ──
  let finalGradeSection = null; // { label, gradeKey, reviewKey }
  if (achievStartIdx >= 0) {
    const sectionBoundaries2 = [];
    let lastLabel2 = '';
    for (let i = achievStartIdx; i < mainHeaders.length; i++) {
      const g = groupHeaders[i] || '';
      if (g && g !== lastLabel2) {
        sectionBoundaries2.push({ label: g, startIdx: i });
        lastLabel2 = g;
      }
    }
    // 3차 종합 또는 최종 등급 섹션 찾기
    const finalBoundary = sectionBoundaries2.find(b => /3차|최종/.test(b.label));
    if (finalBoundary) {
      const nextB = sectionBoundaries2.find(b => b.startIdx > finalBoundary.startIdx);
      const endIdx = nextB ? nextB.startIdx : mainHeaders.length;
      let gradeKey = null, reviewKey = null;
      for (let i = finalBoundary.startIdx; i < endIdx; i++) {
        const h = mainHeaders[i];
        if (!h) continue;
        if ((h.includes('등급') || h === '평가 등급') && !gradeKey) gradeKey = h;
        else if (!reviewKey) {
          // 리뷰/코멘트/의견 또는 등급이 아닌 나머지 텍스트 컬럼
          if (h === '리뷰' || /^리뷰[_\.]/.test(h) || h.includes('코멘트') || h.includes('의견')) reviewKey = h;
        }
      }
      // fallback: 등급 키 다음의 아무 컬럼이라도 리뷰로 사용
      if (gradeKey && !reviewKey) {
        for (let i = finalBoundary.startIdx; i < endIdx; i++) {
          const h = mainHeaders[i];
          if (h && h !== gradeKey) { reviewKey = h; break; }
        }
      }
      if (gradeKey || reviewKey) {
        const cleanLabel = finalBoundary.label.replace(/전체 인원 \d+명/, '').trim();
        finalGradeSection = { label: cleanLabel, gradeKey, reviewKey };
      }
    }
  }

  // reviewSections가 비어있으면 기존 방식으로 fallback
  const useDynamicSections = reviewSections.length > 0;
  AppState.reviewSectionLabels = reviewSectionLabels; // 요약 테이블 헤더용
  AppState.hasFinalGrade = !!finalGradeSection; // 최종 등급 존재 여부

  // fallback용 변수 (동적 섹션 파싱 실패 시)
  const hasWorkReviewCol = mainHeaders.includes('업무 리뷰');
  const reviewKeys = mainHeaders.filter(h => h === '리뷰' || /^리뷰[_\.]/.test(h));
  const review1Key = mainHeaders.find(h => h === '리뷰_1' || h === '리뷰.1') || reviewKeys[1] || reviewKeys[0] || null;

  AppState.employees = mainRows
    .filter(r => r['소속'] && r['성명'] && isPersonName(r['성명']))
    .map(r => ({
      department: r['소속'] || '',
      name: r['성명'] || '',
      position: r['직책'] || '',
      joinDate: formatDate(r['입사일']),
      totalCareer: r['총 경력'] || r['총경력'] || '',
      tenure: r['25년 근속'] || r['24년 근속'] || r['근속'] || '',
      // fallback용 (동적 섹션 없을 때)
      achievement: r['업무 달성도 (1-5)'] || r['업무달성도'] || null,
      workReview: (hasWorkReviewCol ? r['업무 리뷰'] : (reviewKeys[0] ? r[reviewKeys[0]] : '')) || '',
      grade1: r['평가 등급'] || r['1차종합등급'] || r['1차 종합 등급'] || '',
      review1: (review1Key ? r[review1Key] : '') || '',
      // 동적 섹션 데이터
      reviewSections: useDynamicSections ? reviewSections.map(s => ({
        label: s.label,
        achievement: s.achieveKey ? (r[s.achieveKey] ?? null) : null,
        grade: s.gradeKey ? (r[s.gradeKey] || '') : '',
        review: s.reviewKey ? (r[s.reviewKey] || '') : '',
      })) : [],
      prevYear1: r['작년등급'] || '',
      prevYear2: r['재작년등급'] || '',
      // 3차 종합 / 최종 등급
      finalGrade: finalGradeSection && finalGradeSection.gradeKey ? (r[finalGradeSection.gradeKey] || '') : '',
      finalComment: finalGradeSection && finalGradeSection.reviewKey ? (r[finalGradeSection.reviewKey] || '') : '',
      finalGradeLabel: finalGradeSection ? finalGradeSection.label : '',
    }));

  const leaderSheet = findSheet(workbook, SHEET_PATTERNS.leaderSelf);
  if (leaderSheet) {
    const { rows } = sheetToRowsFromHeader(leaderSheet, 3);
    AppState.leaderReviews = rows
      .filter(r => r['리뷰 대상자'] || r['리뷰대상자'])
      .map(r => ({
        reviewee: r['리뷰 대상자'] || r['리뷰대상자'] || '',
        org: r['조직'] || '',
        task: r['주 업무 내용'] || r['주업무내용'] || '',
        weight: r['가중치 (100 기준)'] || r['가중치'] || null,
        achievement: r['달성도 (1-5)'] || r['달성도'] || null,
        comment: r['코멘트 내용'] || r['코멘트'] || '',
      }));
  }

  const compSheet = findSheet(workbook, SHEET_PATTERNS.competency);
  if (compSheet) {
    const { rows } = sheetToRowsFromHeader(compSheet, 3);
    AppState.competencyReviews = rows
      .filter(r => r['팀원'])
      .map(r => ({
        org: r['조직'] || '',
        member: r['팀원'] || '',
        leader: r['작성 리더'] || r['작성리더'] || '',
        communication: r['커뮤니케이션'] || null,
        challenge: r['도전성'] || null,
        responsibility: r['책임감'] || null,
        teamwork: r['팀워크'] || null,
        expertise: r['업무전문성'] || null,
        growthLevel: r['역량/개발 성장 수준에 대한 평가'] || r['성장수준평가'] || '',
        advice: r['팀원의 업무와 성장에 도움이 되는 건설적인 조언'] || r['건설적조언'] || '',
      }));
  }

  const peerSheet = findSheet(workbook, SHEET_PATTERNS.peerFeedback);
  if (peerSheet) {
    const { rows } = sheetToRowsFromHeader(peerSheet, 3);
    AppState.peerFeedbacks = rows
      .filter(r => r['리뷰 대상자'] || r['리뷰대상자'])
      .map(r => {
        const findCol = (keywords) => {
          const key = Object.keys(r).find(k => keywords.some(kw => k.includes(kw)));
          return key ? r[key] : null;
        };
        // 보완 역량 선택값: '보완' 또는 '성장이 필요' 포함 키 중 _1 suffix 없는 것
        const improvAreaKey = Object.keys(r).find(k =>
          (k.includes('보완') || k.includes('성장이 필요')) && !k.endsWith('_1')
        );
        const improvArea = improvAreaKey ? r[improvAreaKey] : null;
        // 선택 이유: 같은 헤더명 + '_1' suffix (병합 셀 언머지 결과)
        const improvReasonKey = improvAreaKey ? (improvAreaKey + '_1') : null;
        const improvReason = (improvReasonKey && r[improvReasonKey]) || findCol(['선택한 이유', '이유']) || null;
        return {
          org: r['조직'] || '',
          reviewee: r['리뷰 대상자'] || r['리뷰대상자'] || '',
          average: r['점수'] || r['종합 평균'] || r['평균'] || null,
          score: r['점수'] || r['평균'] || null,
          improvementArea: improvArea ? String(improvArea) : '',
          improvementReason: improvReason ? String(improvReason) : '',
          positiveImpact: findCol(['긍정적인 영향', '긍정적 영향', '협업 관련', '긍정적']) || '',
        };
      });
  }

  const upwardSheet = findSheet(workbook, SHEET_PATTERNS.upwardFeedback);
  if (upwardSheet) {
    AppState.upwardFeedbacks = [];
    unmergeSheet(upwardSheet); // 세로 병합(부서 등) 먼저 처리

    const upRange = XLSX.utils.decode_range(upwardSheet['!ref'] || 'A1:A1');

    // 헤더 행 인덱스 찾기
    const headerRowIndices = [];
    for (let r = 0; r <= upRange.e.r; r++) {
      for (let c = upRange.s.c; c <= upRange.e.c; c++) {
        const cell = upwardSheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && ['부서', '닉네임', '리뷰 대상자'].includes(String(cell.v).trim())) {
          headerRowIndices.push(r);
          break;
        }
      }
    }

    // 디버그: 헤더 행 및 각 섹션 이름 행 출력
    console.group('🔍 상향 피드백 시트 디버그');
    console.log('헤더 행 인덱스:', headerRowIndices);
    headerRowIndices.forEach((hIdx) => {
      // 헤더 행 컬럼값 출력
      const hVals = [];
      for (let c = upRange.s.c; c <= Math.min(upRange.e.c, upRange.s.c + 15); c++) {
        const cell = upwardSheet[XLSX.utils.encode_cell({ r: hIdx, c })];
        if (cell && cell.v) hVals.push(String(cell.v).trim());
      }
      console.log(`  행${hIdx} 헤더:`, hVals);
      // 이름 행 후보 (헤더 위 5행)
      for (let i = hIdx - 1; i >= Math.max(0, hIdx - 5); i--) {
        const rVals = [];
        for (let c = upRange.s.c; c <= upRange.e.c; c++) {
          const cell = upwardSheet[XLSX.utils.encode_cell({ r: i, c })];
          if (cell && cell.v != null && String(cell.v).trim()) rVals.push(String(cell.v).trim());
        }
        if (rVals.length > 0) console.log(`    위 행${i}:`, rVals);
      }
    });
    console.groupEnd();

    headerRowIndices.forEach((hIdx, sectionIdx) => {
      const nextHIdx = headerRowIndices[sectionIdx + 1] != null
        ? headerRowIndices[sectionIdx + 1]
        : upRange.e.r + 1;

      // 헤더 추출 (이미 unmerge됐으므로 그냥 읽기, 중복 헤더는 _1 suffix)
      const colHeaders = [];
      const headerCount = {};
      for (let c = upRange.s.c; c <= upRange.e.c; c++) {
        const cell = upwardSheet[XLSX.utils.encode_cell({ r: hIdx, c })];
        let name = cell ? String(cell.v).trim() : '';
        if (name) {
          headerCount[name] = (headerCount[name] || 0) + 1;
          if (headerCount[name] > 1) name = `${name}_${headerCount[name] - 1}`;
        }
        colHeaders.push(name);
      }

      // 피평가자 이름: 헤더 행 바로 위에서 찾기
      let targetName = '';
      for (let i = hIdx - 1; i >= Math.max(0, hIdx - 5); i--) {
        const rowVals = [];
        for (let c = upRange.s.c; c <= upRange.e.c; c++) {
          const cell = upwardSheet[XLSX.utils.encode_cell({ r: i, c })];
          if (cell && cell.v != null && String(cell.v).trim()) rowVals.push(String(cell.v).trim());
        }
        if (rowVals.length >= 1 && rowVals.length <= 3) {
          const val = rowVals[0].replace(/\s*\(.*\)/, '').trim();
          if (/^[A-Za-z가-힣]{2,}/.test(val)) { targetName = val; break; }
        }
      }

      // 데이터 행: hIdx+1 ~ nextHIdx-1 직접 읽기
      for (let r = hIdx + 1; r < nextHIdx; r++) {
        const row = {};
        let hasData = false;
        for (let c = upRange.s.c; c <= upRange.e.c; c++) {
          const cell = upwardSheet[XLSX.utils.encode_cell({ r, c })];
          const key = colHeaders[c - upRange.s.c];
          if (key) {
            row[key] = cell ? cell.v : null;
            if (cell && cell.v != null && cell.v !== '') hasData = true;
          }
        }
        if (!hasData) continue;

        // 점수/보완역량 등 실제 평가 데이터가 없으면 빈 행으로 스킵
        const hasScore = row['점수'] != null || row['총점'] != null || row['종합 평균'] != null;
        const hasImprov = Object.keys(row).some(k => (k.includes('보완') || k.includes('성장이 필요')) && !k.endsWith('_1') && row[k] != null && row[k] !== '');
        const hasPositive = Object.keys(row).some(k => k.includes('긍정적') && row[k] != null && row[k] !== '');
        if (!hasScore && !hasImprov && !hasPositive) continue;

        const findCol = (keywords) => {
          const key = Object.keys(row).find(k => keywords.some(kw => k.includes(kw)));
          return key ? row[key] : null;
        };

        const reviewee = row['리뷰 대상자'] || row['리뷰대상자'] || targetName;
        if (!reviewee) continue;

        // 보완 역량: _1 suffix 없는 키
        const improvAreaKey = Object.keys(row).find(k =>
          (k.includes('보완') || k.includes('성장이 필요')) && !k.endsWith('_1')
        );
        const improvArea = improvAreaKey ? row[improvAreaKey] : null;
        const improvReasonKey = improvAreaKey ? (improvAreaKey + '_1') : null;
        const improvReason = (improvReasonKey && row[improvReasonKey]) || findCol(['선택한 이유', '이유']) || null;

        AppState.upwardFeedbacks.push({
          targetName: String(reviewee),
          department: row['부서'] || row['조직'] || '',
          nickname: row['닉네임'] || '',
          expertise: row['전문성'] || null,
          innovation: row['도전과 혁신'] || row['도전과혁신'] || null,
          attitude: row['태도'] || null,
          teamwork: row['팀워크'] || null,
          leadership: row['리더쉽'] || row['리더십'] || null,
          memberMgmt: row['구성원 관리'] || row['구성원관리'] || null,
          performance: row['퍼포먼스'] || null,
          cultureImprovement: row['팀문화 개선'] || row['팀문화개선'] || null,
          total: row['점수'] || row['총점'] || row['종합 평균'] || findCol(['총점', '평균', '점수']) || null,
          improvementArea: improvArea ? String(improvArea) : '',
          improvementReason: improvReason ? String(improvReason) : '',
          positiveImpact: findCol(['긍정적인 영향', '긍정적 영향', '구성원들에게', '긍정적']) || '',
          extraFeedback: findCol(['추가로']) || '',
        });
      }
    });

    // ── 챕터 리드 피드백 시트 추가 파싱 (프로덕트/개발 조직) ──
    // 시트명에 '챕터' 또는 '리더 피드백' 포함하는 시트 찾기
    const chapterSheetName = workbook.SheetNames.find(sn =>
      sn.includes('챕터') || sn === '리더 피드백'
    );
    if (chapterSheetName) {
      const chapterSheet = workbook.Sheets[chapterSheetName];
      // 헤더 행 자동 탐색: '리뷰 대상자' 또는 '조직' 컬럼이 있는 행
      const cRange = XLSX.utils.decode_range(chapterSheet['!ref'] || 'A1:A1');
      let chapterHeaderRow = 0;
      outer: for (let r = 0; r <= Math.min(5, cRange.e.r); r++) {
        for (let c = cRange.s.c; c <= cRange.e.c; c++) {
          const cell = chapterSheet[XLSX.utils.encode_cell({ r, c })];
          if (cell && ['리뷰 대상자', '조직', '평균', '점수'].includes(String(cell.v).trim())) {
            chapterHeaderRow = r; break outer;
          }
        }
      }
      const { rows: chapterRows } = sheetToRowsFromHeader(chapterSheet, chapterHeaderRow);
      chapterRows
        .filter(r => r['리뷰 대상자'] || r['리뷰대상자'])
        .forEach(r => {
          const findCol = (keywords) => {
            const key = Object.keys(r).find(k => keywords.some(kw => k.includes(kw)));
            return key ? r[key] : null;
          };
          const improvAreaKey = Object.keys(r).find(k =>
            (k.includes('보완') || k.includes('성장이 필요')) && !k.endsWith('_1')
          );
          const improvArea = improvAreaKey ? r[improvAreaKey] : null;
          const improvReasonKey = improvAreaKey ? (improvAreaKey + '_1') : null;
          const improvReason = (improvReasonKey && r[improvReasonKey]) || findCol(['선택한 이유', '이유']) || null;

          AppState.upwardFeedbacks.push({
            targetName: String(r['리뷰 대상자'] || r['리뷰대상자']),
            department: r['조직'] || r['부서'] || '',
            nickname: r['닉네임'] || '',
            total: r['점수'] || r['평균'] || r['종합 평균'] || findCol(['점수', '평균']) || null,
            improvementArea: improvArea ? String(improvArea) : '',
            improvementReason: improvReason ? String(improvReason) : '',
            positiveImpact: findCol(['긍정적인 영향', '긍정적 영향', '구성원들에게', '긍정적']) || '',
            extraFeedback: findCol(['추가로']) || '',
          });
        });
    }
  }

  // ── 추가 시트 파싱 (알려진 시트 외 나머지) ──
  AppState.extraSheets = [];
  const usedSheetNames = new Set();
  // 이미 사용한 시트들 수집
  [SHEET_PATTERNS.mainReview, SHEET_PATTERNS.leaderSelf, SHEET_PATTERNS.competency,
   SHEET_PATTERNS.peerFeedback, SHEET_PATTERNS.upwardFeedback].forEach(patterns => {
    workbook.SheetNames.forEach(sn => {
      if (patterns.some(p => sn === p || sn.includes(p) || p.includes(sn) ||
          sn.replace(/\s/g,'').includes(p.replace(/\s/g,'').substring(0,4)))) {
        usedSheetNames.add(sn);
      }
    });
  });
  workbook.SheetNames.forEach(sn => {
    if (usedSheetNames.has(sn)) return;
    // 알려진 패턴 키워드 포함 시트 제외
    const knownKeywords = ['핵심인재보상', '3개년 성과'];
    if (knownKeywords.some(k => sn.includes(k))) return;
    // 추가 시트: 헤더 자동 탐색 (첫 3행 중 데이터 있는 행)
    try {
      const sheet = workbook.Sheets[sn];
      // 헤더 행 탐색: 첫 10행 중 가장 많은 값이 있는 행
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
      let bestRow = 0, bestCount = 0;
      for (let r = 0; r <= Math.min(9, range.e.r); r++) {
        let count = 0;
        for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 20); c++) {
          const cell = sheet[XLSX.utils.encode_cell({ r, c })];
          if (cell && cell.v) count++;
        }
        if (count > bestCount) { bestCount = count; bestRow = r; }
      }
      const { headers, rows } = sheetToRowsFromHeader(sheet, bestRow);
      if (rows.length > 0) {
        AppState.extraSheets.push({ sheetName: sn, headers: headers.filter(Boolean), rows });
      }
    } catch (e) { /* 파싱 실패 시 무시 */ }
  });

  return { success: true };
}

function formatDate(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    return `${val.getFullYear()}.${String(val.getMonth() + 1).padStart(2, '0')}.${String(val.getDate()).padStart(2, '0')}`;
  }
  return String(val);
}

// =============================================
// 이력 파일 파싱 (평가 대시보드.xlsx)
// =============================================
function parseHistoryWorkbook(workbook) {
  // 시트명: '3개년 성과' 또는 첫 번째 시트
  const sheetName = workbook.SheetNames.includes('3개년 성과') ? '3개년 성과' : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return;

  // 헤더 행 자동 탐색: '성명' 컬럼이 있는 행 찾기
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  let headerRow = 2; // 기본값 (0-indexed)
  for (let r = 0; r <= Math.min(10, range.e.r); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && String(cell.v).trim() === '성명') { headerRow = r; break; }
    }
  }

  const { rows } = sheetToRowsFromHeader(sheet, headerRow);
  AppState.historyData = rows
    .filter(r => r['성명'])
    .map(r => ({
      name: r['성명'] || '',
      nickname: r['닉네임'] || '',
      division: r['부문'] || '',
      dept: r['부서'] || '',
      team: r['팀'] || '',
      career: r['25년 경력'] || r['경력'] || null,
      tenure: r['26년 근속'] || null,
      grade22: r['22년 평가'] || '',
      grade23: r['23년 평가'] || '',
      grade24: r['24년 평가'] || '',
    }));
}

// =============================================
// 차트 유틸
// =============================================
function destroyChart(id) {
  if (chartInstances[id]) {
    chartInstances[id].destroy();
    delete chartInstances[id];
  }
}

function renderGradeDonut(canvasId, gradeCounts) {
  destroyChart(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const labels = Object.keys(gradeCounts).filter(k => gradeCounts[k] > 0);
  if (labels.length === 0) {
    canvas.style.display = 'none';
    const nd = document.getElementById(canvasId.replace('-chart', '-no-data'));
    if (nd) nd.style.display = 'block';
    return;
  }
  const data = labels.map(k => gradeCounts[k]);
  const total = data.reduce((a, b) => a + b, 0);
  const gradeColors = { S: '#F59E0B', A: '#3B82F6', B: '#22C55E', C: '#F87171' };
  const bgColors = labels.map(l => gradeColors[l] || '#8FADD4');

  chartInstances[canvasId] = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: bgColors, borderWidth: 2, borderColor: '#fff' }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            generateLabels: (chart) => {
              return chart.data.labels.map((label, i) => ({
                text: `${label}  ${data[i]}명 (${Math.round(data[i] / total * 100)}%)`,
                fillStyle: bgColors[i],
                strokeStyle: '#fff',
                lineWidth: 2,
                index: i,
              }));
            }
          }
        },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const val = ctx.parsed;
              const pct = Math.round(val / total * 100);
              return ` ${ctx.label}: ${val}명 (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderHorizontalBar(canvasId, labels, data, maxValue) {
  destroyChart(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  chartInstances[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ data, backgroundColor: '#4A6FA5', borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { min: 0, max: maxValue || 5, ticks: { stepSize: 1 } } },
      plugins: { legend: { display: false } }
    }
  });
}

function renderAchievementBar(canvasId, achievementList) {
  destroyChart(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const counts = [0, 0, 0, 0, 0];
  achievementList.forEach(v => { if (v >= 1 && v <= 5) counts[Math.round(v) - 1]++; });
  chartInstances[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: ['1점', '2점', '3점', '4점', '5점'],
      datasets: [{ data: counts, backgroundColor: '#4A6FA5', borderRadius: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
    }
  });
}

// =============================================
// 요약 페이지 렌더링
// =============================================
function renderSummaryPage() {
  const emps = AppState.employees;
  if (!emps.length) return;

  document.getElementById('stat-division').textContent = AppState.divisionName || '-';
  document.getElementById('stat-total').textContent = emps.length;

  // 동적 섹션 레이블 (파일마다 다름)
  const secLabels = AppState.reviewSectionLabels || [];
  const useDynamic = secLabels.length > 0;

  // 달성도/등급 추출 헬퍼
  function getAchievement(emp) {
    if (useDynamic && emp.reviewSections.length > 0) {
      const f = emp.reviewSections.find(s => s.achievement != null);
      return f ? f.achievement : null;
    }
    return emp.achievement;
  }
  function getGrade(emp, secIdx) {
    if (useDynamic && emp.reviewSections.length > secIdx) return emp.reviewSections[secIdx].grade;
    return secIdx === 0 ? emp.grade1 : '';
  }

  const achievements = emps.map(e => parseFloat(getAchievement(e))).filter(v => !isNaN(v));
  const avg = achievements.length ? (achievements.reduce((a, b) => a + b, 0) / achievements.length).toFixed(1) : '-';
  document.getElementById('stat-avg').textContent = avg !== '-' ? `${avg} / 5` : '-';
  const completed = emps.filter(e => {
    const fg = String(e.finalGrade || '').trim();
    return fg !== '' && fg !== '-';
  }).length;
  document.getElementById('stat-completed').textContent = completed;

  // ── 최종 등급 집계 ──
  const finalGradeCounts = { S: 0, A: 0, B: 0, C: 0 };
  emps.forEach(e => {
    const gradeVal = e.finalGrade || '';
    const g = String(gradeVal).trim().charAt(0).toUpperCase();
    if (finalGradeCounts[g] !== undefined) finalGradeCounts[g]++;
  });

  const hasAnyFinalGrade = Object.values(finalGradeCounts).some(v => v > 0);

  // 등급 데이터 존재 여부 (테이블 컬럼용)
  const hasAnyGrade = emps.some(e => {
    if (useDynamic && e.reviewSections.length > 0) {
      return e.reviewSections.some(s => s.grade && /^[SABC]/i.test(String(s.grade).trim()));
    }
    return e.grade1 && /^[SABC]/i.test(String(e.grade1).trim());
  });

  // ── 차트 타이틀 동적 설정 ──
  const chartFinalTitle = document.getElementById('chart-final-title');
  if (chartFinalTitle) {
    const label = AppState.hasFinalGrade
      ? (emps[0]?.finalGradeLabel || '최종')
      : '최종';
    chartFinalTitle.textContent = `등급 분포 (${label})`;
  }

  // ── 최종 등급 도넛 차트 ──
  if (hasAnyFinalGrade) {
    renderGradeDonut('grade-chart-final', finalGradeCounts);
    document.getElementById('grade-final-no-data').style.display = 'none';
  } else {
    destroyChart('grade-chart-final');
    document.getElementById('grade-final-no-data').style.display = '';
  }

  // ── 부서별 최종 등급 스택 바 차트 ──
  if (hasAnyFinalGrade) {
    // 부서별 집계
    const deptMap = {};
    emps.forEach(e => {
      const dept = e.department || '미지정';
      if (!deptMap[dept]) deptMap[dept] = { S: 0, A: 0, B: 0, C: 0 };
      const g = String(e.finalGrade || '').trim().charAt(0).toUpperCase();
      if (deptMap[dept][g] !== undefined) deptMap[dept][g]++;
    });
    const deptLabels = Object.keys(deptMap);
    const gradeColors = { S: '#F59E0B', A: '#3B82F6', B: '#22C55E', C: '#EF4444' };
    const datasets = ['S', 'A', 'B', 'C'].map(grade => ({
      label: grade,
      data: deptLabels.map(d => deptMap[d][grade]),
      backgroundColor: gradeColors[grade],
      borderRadius: 2,
    }));

    destroyChart('dept-grade-chart');
    const ctx = document.getElementById('dept-grade-chart');
    chartInstances['dept-grade-chart'] = new Chart(ctx, {
      type: 'bar',
      data: { labels: deptLabels, datasets },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
        },
        scales: {
          x: { stacked: true, beginAtZero: true, ticks: { stepSize: 1 } },
          y: { stacked: true, ticks: { font: { size: 11 } } },
        },
      },
    });
    document.getElementById('dept-grade-no-data').style.display = 'none';
  } else {
    destroyChart('dept-grade-chart');
    document.getElementById('dept-grade-no-data').style.display = '';
  }

  // ── 테이블 헤더 동적 생성 ──
  const thead = document.querySelector('#summary-table thead tr');
  const colgroup = document.querySelector('#summary-table colgroup');

  // 고정 컬럼: 이름, 직책, 소속, 입사일, 총경력, 근속
  const fixedCols = [
    { label: '이름',   width: '80px' },
    { label: '직책',   width: '90px' },
    { label: '소속',   width: '130px' },
    { label: '입사일', width: '100px' },
    { label: '총경력', width: '80px' },
    { label: '근속',   width: '70px' },
  ];

  // 1차 종합 등급 컬럼
  const gradeColLabel = (() => {
    if (hasAnyGrade) {
      const gradeSecLabel = useDynamic && secLabels.length > 0
        ? (secLabels.find(l => /1차|종합/.test(l)) || secLabels[0])
        : '1차 종합';
      return `등급\n(${gradeSecLabel})`;
    }
    return '등급';
  })();

  const hasGradeCol = hasAnyGrade;

  // 최종 등급 컬럼 존재 여부
  const hasFinalGradeCol = AppState.hasFinalGrade && emps.some(e => e.finalGrade && String(e.finalGrade).trim() !== '');

  const dynCols = [
    ...(hasGradeCol ? [{ label: gradeColLabel, width: '110px' }] : []),
    ...(hasFinalGradeCol ? [{ label: '최종 등급', width: '90px' }] : []),
  ];

  const allCols = [...fixedCols, ...dynCols];

  colgroup.innerHTML = allCols.map(c => `<col style="width:${c.width}">`).join('');
  thead.innerHTML = allCols.map(c =>
    `<th style="white-space:pre-wrap;word-break:keep-all;">${c.label}</th>`
  ).join('');

  // ── 테이블 바디 ──
  document.getElementById('table-count').textContent = `${emps.length}명`;
  const tbody = document.getElementById('summary-tbody');
  tbody.innerHTML = '';

  emps.forEach((emp, idx) => {
    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';

    const fixedCells = `
      <td><strong>${emp.name}</strong></td>
      <td>${emp.position || '-'}</td>
      <td>${emp.department || '-'}</td>
      <td>${emp.joinDate || '-'}</td>
      <td>${emp.totalCareer || '-'}</td>
      <td>${emp.tenure || '-'}</td>
    `;

    // 1차 종합 등급 컬럼
    let gradeCell = '';
    if (hasGradeCol) {
      let gradeVal = '';
      if (useDynamic && emp.reviewSections.length > 0) {
        const gradeSection = emp.reviewSections.find(s => s.grade && String(s.grade).trim() !== '');
        gradeVal = gradeSection ? gradeSection.grade : '';
      } else {
        gradeVal = emp.grade1;
      }
      gradeCell = `<td>${gradeOrScoreBadge(gradeVal)}</td>`;
    }

    // 최종 등급 컬럼
    let finalGradeCell = '';
    if (hasFinalGradeCol) {
      finalGradeCell = `<td>${gradeOrScoreBadge(emp.finalGrade)}</td>`;
    }

    tr.innerHTML = fixedCells + gradeCell + finalGradeCell;
    tr.addEventListener('click', () => navigateDetail(idx));
    tbody.appendChild(tr);
  });
}

// =============================================
// 상세 페이지 렌더링
// =============================================
function renderDetailPage(index) {
  const emp = AppState.employees[index];
  if (!emp) return;

  updateNavButtons();

  // 이름 매칭 헬퍼: 괄호/공백 무시하고 포함 여부 확인
  function nameMatch(a, b) {
    if (!a || !b) return false;
    const clean = s => String(s).replace(/\s/g, '').split('(')[0];
    return clean(a) === clean(b) || clean(a).includes(clean(b)) || clean(b).includes(clean(a));
  }

  const leaderData = AppState.leaderReviews.filter(r => nameMatch(r.reviewee, emp.name));
  const peerData = AppState.peerFeedbacks.filter(r => nameMatch(r.reviewee, emp.name));
  // 상향 피드백: targetName(피평가자)으로 매칭
  const upwardData = AppState.upwardFeedbacks.filter(r => nameMatch(r.targetName, emp.name));
  // 역량 리뷰 제거
  const histData = (() => {
    const empName = emp.name.toLowerCase();
    const empDept = (emp.department || '').replace(/\s/g, '');

    // 소속 일치 여부: 이력파일의 팀 또는 부서와 평가파일 소속 비교
    function deptMatch(r) {
      const rTeam = (r.team || '').replace(/\s/g, '');
      const rDept = (r.dept || '').replace(/\s/g, '');
      return (rTeam && empDept && (rTeam === empDept || empDept.includes(rTeam) || rTeam.includes(empDept))) ||
             (rDept && empDept && (rDept === empDept || empDept.includes(rDept) || rDept.includes(empDept)));
    }

    // 1순위: 닉네임 + 팀/부서 동시 일치
    const byNickAndDept = AppState.historyData.find(r =>
      r.nickname && r.nickname.toLowerCase() === empName && deptMatch(r)
    );
    if (byNickAndDept) return byNickAndDept;

    // 2순위: 한글 성명 일치
    const byName = AppState.historyData.find(r => nameMatch(r.name, emp.name));
    if (byName) return byName;

    // 3순위: 닉네임만 일치 (중복 없는 경우에만)
    const byNick = AppState.historyData.filter(r =>
      r.nickname && r.nickname.toLowerCase() === empName
    );
    return byNick.length === 1 ? byNick[0] : null;
  })();

  const content = document.getElementById('detail-content');
  content.innerHTML = '';
  const sectionA = renderSectionA(emp, histData);
  (Array.isArray(sectionA) ? sectionA : [sectionA]).forEach(el => content.appendChild(el));
  content.appendChild(renderSectionB(emp, leaderData));
  const sectionC = renderSectionC(emp, peerData, upwardData);
  if (sectionC) content.appendChild(sectionC);
  // 역량 리뷰 제거
  // 추가 시트 렌더링
  AppState.extraSheets.forEach(extra => {
    const sectionE = renderExtraSheet(emp, extra);
    if (sectionE) content.appendChild(sectionE);
  });

  if (peerData.length > 0) {
    // 차트 제거 - 테이블로 대체됨
  }
  if (upwardData.length > 0) {
    // 차트 제거 - 테이블로 대체됨
  }
  // 역량 리뷰 차트 제거
}

// 등급 또는 숫자 점수를 배지로 렌더링
function gradeOrScoreBadge(val, size) {
  if (!val && val !== 0) return '<span class="grade-badge">-</span>';
  const s = String(val).trim();
  const lgClass = size === 'lg' ? ' lg' : '';
  // 숫자 점수 (1~5)
  if (/^\d+(\.\d+)?$/.test(s)) {
    return `<span class="grade-badge score-badge${lgClass}">${s}</span>`;
  }
  // 등급 문자: A(우수), B(만족) 등 → 첫 글자만 등급으로 사용
  const gradeChar = s.charAt(0).toUpperCase();
  const gradeClass = ['S','A','B','C'].includes(gradeChar) ? ` grade-${gradeChar}` : '';
  // 괄호 안 설명 제거하고 등급만 표시
  const displayVal = /^[SABC]\(/.test(s) ? gradeChar : s;
  return `<span class="grade-badge${gradeClass}${lgClass}">${displayVal}</span>`;
}

function renderSectionA(emp, histData) {
  const div = document.createElement('div');
  div.className = 'section-card';

  const career = emp.totalCareer || (histData ? histData.career : '') || '-';
  const tenure = emp.tenure || (histData ? histData.tenure : '') || '-';

  div.innerHTML = `
    <h3 class="section-title">기본 정보</h3>
    <table class="info-table">
      <thead>
        <tr>
          <th>소속</th><th>이름</th><th>직책</th>
          <th>입사일</th><th>총경력</th><th>근속</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>${emp.department || '-'}</td>
          <td><strong>${emp.name || '-'}</strong></td>
          <td>${emp.position || '-'}</td>
          <td>${emp.joinDate || '-'}</td>
          <td>${career}</td>
          <td>${tenure}</td>
        </tr>
      </tbody>
    </table>
  `;

  // 3개년 이력 별도 카드
  if (histData) {
    const histCard = document.createElement('div');
    histCard.className = 'section-card history-card';
    histCard.innerHTML = `
      <h3 class="section-title">성과 이력</h3>
      <div class="history-grades">
        <div class="history-grade-item">
          <div class="history-grade-year">2022</div>
          <div class="history-grade-badge">${gradeOrScoreBadge(histData.grade22, 'lg')}</div>
        </div>
        <div class="history-grade-arrow">→</div>
        <div class="history-grade-item">
          <div class="history-grade-year">2023</div>
          <div class="history-grade-badge">${gradeOrScoreBadge(histData.grade23, 'lg')}</div>
        </div>
        <div class="history-grade-arrow">→</div>
        <div class="history-grade-item">
          <div class="history-grade-year">2024</div>
          <div class="history-grade-badge">${gradeOrScoreBadge(histData.grade24, 'lg')}</div>
        </div>
      </div>
    `;

    const cards = [div, histCard];

    // 최종 등급 + 코멘트 카드 (성과 이력 바로 아래)
    if (emp.finalGrade && String(emp.finalGrade).trim() !== '') {
      const finalCard = document.createElement('div');
      finalCard.className = 'section-card';
      const label = emp.finalGradeLabel || '최종 등급';
      const commentHtml = emp.finalComment && String(emp.finalComment).trim()
        ? `<div class="eval-review-box" style="margin-top:14px;">${String(emp.finalComment).trim()}</div>`
        : '';
      finalCard.innerHTML = `
        <h3 class="section-title">${label}</h3>
        <div style="display:flex;align-items:center;gap:12px;">
          <span style="font-size:13px;color:var(--text-muted);">등급</span>
          ${gradeOrScoreBadge(emp.finalGrade, 'lg')}
        </div>
        ${commentHtml}
      `;
      cards.push(finalCard);
    }

    return cards;
  }

  // 최종 등급 카드 (이력 데이터 없어도 최종 등급이 있으면 표시)
  if (emp.finalGrade && String(emp.finalGrade).trim() !== '') {
    const finalCard = document.createElement('div');
    finalCard.className = 'section-card';
    const label = emp.finalGradeLabel || '최종 등급';
    const commentHtml = emp.finalComment && String(emp.finalComment).trim()
      ? `<div class="eval-review-box" style="margin-top:14px;">${String(emp.finalComment).trim()}</div>`
      : '';
    finalCard.innerHTML = `
      <h3 class="section-title">${label}</h3>
      <div style="display:flex;align-items:center;gap:12px;">
        <span style="font-size:13px;color:var(--text-muted);">등급</span>
        ${gradeOrScoreBadge(emp.finalGrade, 'lg')}
      </div>
      ${commentHtml}
    `;
    return [div, finalCard];
  }

  return [div];
}

function renderSectionB(emp, leaderData) {
  const div = document.createElement('div');
  div.className = 'section-card';

  const leaderSelfHtml = leaderData.length > 0 ? `
    <div class="subsection">
      <div class="collapsible-header" onclick="toggleCollapse(this)">
        <h4 style="margin:0;">셀프 리뷰 (리더) <span class="collapse-count">${leaderData.length}건</span></h4>
        <span class="collapse-icon">▲</span>
      </div>
      <div class="collapsible-body">
        <table class="leader-self-table">
          <thead><tr>
            <th class="col-task">주 업무 내용</th>
            <th class="col-weight">가중치</th>
            <th class="col-achieve">달성도</th>
            <th class="col-comment">코멘트</th>
          </tr></thead>
          <tbody>
            ${leaderData.map(r => {
              const w = r.weight != null ? parseFloat(r.weight) : null;
              const weightStr = w != null ? (w < 2 ? Math.round(w * 100) + '%' : w + '%') : '-';
              return `<tr>
                <td class="col-task">${r.task || '-'}</td>
                <td class="col-weight">${weightStr}</td>
                <td class="col-achieve">${r.achievement != null ? r.achievement : '-'}</td>
                <td class="col-comment">${r.comment || '-'}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>
  ` : '';

  // 동적 섹션이 있으면 그걸 사용, 없으면 기존 workReview/review1 방식
  let evalCardsHtml = '';
  const sections = emp.reviewSections && emp.reviewSections.filter(s =>
    s.achievement != null || s.grade || s.review
  );

  if (sections && sections.length > 0) {
    evalCardsHtml = `<div class="eval-cards">` + sections.map(s => {
      const achVal = s.achievement != null ? parseFloat(s.achievement) : null;
      const achHtml = achVal != null
        ? `<div class="eval-score"><span class="score-label">달성도</span><span class="achievement-box lg">${achVal} / 5</span></div>`
        : '';
      const gradeHtml = s.grade
        ? `<div class="eval-grade">${gradeOrScoreBadge(s.grade, 'lg')}</div>`
        : '';
      return `
        <div class="eval-card">
          <div class="eval-card-title">${s.label}</div>
          ${achHtml}${gradeHtml}
          <div class="eval-review-box">${s.review || '<span class="no-val">내용 없음</span>'}</div>
        </div>`;
    }).join('') + `</div>`;
  } else {
    // 기존 방식 (fallback)
    const achievementVal = emp.achievement != null ? parseFloat(emp.achievement) : null;
    const achievementHtml = achievementVal != null
      ? `<span class="achievement-box lg">${achievementVal} / 5</span>`
      : '<span class="no-val">-</span>';
    evalCardsHtml = `
      <div class="eval-cards">
        <div class="eval-card">
          <div class="eval-card-title">1차 의견 (팀장)</div>
          <div class="eval-score"><span class="score-label">달성도</span>${achievementHtml}</div>
          <div class="eval-review-box">${emp.workReview || '<span class="no-val">내용 없음</span>'}</div>
        </div>
        <div class="eval-card">
          <div class="eval-card-title">1차 종합 (실장/부문장)</div>
          <div class="eval-grade">${gradeOrScoreBadge(emp.grade1, 'lg')}</div>
          <div class="eval-review-box">${emp.review1 || '<span class="no-val">내용 없음</span>'}</div>
        </div>
      </div>`;
  }

  div.innerHTML = `
    <h3 class="section-title">성과 평가</h3>
    ${evalCardsHtml}
    ${leaderSelfHtml}
  `;
  return div;
}

function renderSectionC(emp, peerData, upwardData) {
  const hasPeer = peerData.length > 0;
  const hasUpward = upwardData.length > 0;
  if (!hasPeer && !hasUpward) return null;

  const div = document.createElement('div');
  div.className = 'section-card';
  div.innerHTML = `<h3 class="section-title">360도 평가</h3>`;

  if (hasPeer) {
    const avgAll = peerData.map(r => parseFloat(r.average || r.score)).filter(v => !isNaN(v));
    const overallAvg = avgAll.length ? (avgAll.reduce((a, b) => a + b, 0) / avgAll.length).toFixed(2) : '-';

    // 보완역량 데이터 존재 여부로 양식 구분
    const hasImprov = peerData.some(r => r.improvementArea && r.improvementArea.trim());
    const hasCollab = peerData.some(r => r.positiveImpact && r.positiveImpact.trim());

    const reviewerRows = peerData.map((r, i) => {
      const score = parseFloat(r.average || r.score);
      const scoreStr = !isNaN(score) ? `<span class="achievement-box">${score.toFixed(2)}</span>` : '-';
      if (hasImprov) {
        return `<tr>
          <td>평가자 ${i + 1}</td>
          <td>${scoreStr}</td>
          <td class="text-cell">${r.improvementArea || '-'}</td>
          <td class="text-cell">${r.improvementReason || '-'}</td>
          <td class="text-cell">${r.positiveImpact || '-'}</td>
        </tr>`;
      } else {
        return `<tr>
          <td>평가자 ${i + 1}</td>
          <td>${scoreStr}</td>
          <td class="text-cell">${r.positiveImpact || '-'}</td>
        </tr>`;
      }
    }).join('');

    const tableHeaders = hasImprov
      ? `<th style="width:70px;">평가자</th><th style="width:70px;">평균</th><th>보완 필요 역량</th><th>선택 이유</th><th>긍정적 영향</th>`
      : `<th style="width:70px;">평가자</th><th style="width:70px;">평균</th><th>협업 의견</th>`;

    const peerDiv = document.createElement('div');
    peerDiv.className = 'subsection';
    peerDiv.innerHTML = `
      <h4>동료 리뷰</h4>
      <div class="avg-highlight">종합 평균 <strong>${overallAvg}</strong> / 5 &nbsp;<span style="font-size:12px;opacity:0.7;">(${peerData.length}명)</span></div>
      <table class="data-table" style="margin-top:10px;">
        <thead><tr>${tableHeaders}</tr></thead>
        <tbody>${reviewerRows}</tbody>
      </table>
    `;
    div.appendChild(peerDiv);
  }

  if (hasUpward) {
    const totals = upwardData.map(r => parseFloat(r.total)).filter(v => !isNaN(v));
    const avgTotal = totals.length ? (totals.reduce((a, b) => a + b, 0) / totals.length).toFixed(1) : '-';

    // 평가자별 평균 목록 (상향 리뷰도 동료 리뷰와 동일한 방식)
    const hasExtraFeedback = upwardData.some(r => r.extraFeedback);
    const reviewerRows = upwardData.map((r, i) => {
      const score = parseFloat(r.total);
      const scoreStr = !isNaN(score) ? `<span class="achievement-box">${score.toFixed(1)}</span>` : '-';
      return `<tr>
        <td>평가자 ${i + 1}</td>
        <td>${scoreStr}</td>
        <td class="text-cell">${r.improvementArea || '-'}</td>
        <td class="text-cell">${r.improvementReason || '-'}</td>
        <td class="text-cell">${r.positiveImpact || '-'}</td>
        ${hasExtraFeedback ? `<td class="text-cell">${r.extraFeedback || '-'}</td>` : ''}
      </tr>`;
    }).join('');

    const upDiv = document.createElement('div');
    upDiv.className = 'subsection';
    upDiv.innerHTML = `
      <h4>상향 리뷰 (리더 평가)</h4>
      <div class="avg-highlight">총점 평균 <strong>${avgTotal}</strong> &nbsp;<span style="font-size:12px;opacity:0.7;">(${upwardData.length}명)</span></div>
      <table class="data-table" style="margin-top:10px;">
        <thead><tr>
          <th style="width:70px;">평가자</th>
          <th style="width:70px;">총점</th>
          <th>보완 필요 역량</th>
          <th>선택 이유</th>
          <th>긍정적 영향</th>
          ${hasExtraFeedback ? '<th>추가 피드백</th>' : ''}
        </tr></thead>
        <tbody>${reviewerRows}</tbody>
      </table>
    `;
    div.appendChild(upDiv);
  }

  return div;
}

function renderSectionD(emp, compData) {
  if (!compData.length) return null;
  const comp = compData[0];
  const div = document.createElement('div');
  div.className = 'section-card';
  div.innerHTML = `
    <h3 class="section-title">역량 리뷰</h3>
    <div class="chart-wrap" style="height:200px;"><canvas id="comp-chart"></canvas></div>
    ${comp.growthLevel ? `<div class="feedback-section"><div class="feedback-label">성장 수준 평가</div><p class="feedback-text">${comp.growthLevel}</p></div>` : ''}
    ${comp.advice ? `<div class="feedback-section"><div class="feedback-label">건설적 조언</div><p class="feedback-text">${comp.advice}</p></div>` : ''}
  `;
  return div;
}

// =============================================
// 추가 시트 렌더링 (챕터 리드 피드백 등)
// =============================================
function renderExtraSheet(emp, extra) {
  // 이 직원과 관련된 행만 필터링
  // 이름/닉네임 컬럼 후보 찾기
  const nameColCandidates = ['리뷰 대상자', '리뷰대상자', '성명', '이름', '닉네임', '팀원'];
  const nameCol = extra.headers.find(h => nameColCandidates.includes(h));

  let rows = extra.rows;
  if (nameCol) {
    rows = extra.rows.filter(r => {
      const val = r[nameCol];
      if (!val) return false;
      const clean = s => String(s).replace(/\s/g, '').split('(')[0].toLowerCase();
      return clean(val) === clean(emp.name) || clean(val).includes(clean(emp.name)) || clean(emp.name).includes(clean(val));
    });
  }
  if (rows.length === 0) return null;

  const div = document.createElement('div');
  div.className = 'section-card';
  div.innerHTML = `<h3 class="section-title">${extra.sheetName}</h3>`;

  // 테이블로 렌더링
  const visibleHeaders = extra.headers.filter(h => h && h !== nameCol);
  const table = document.createElement('table');
  table.className = 'data-table';
  table.innerHTML = `
    <thead><tr>${visibleHeaders.map(h => `<th>${h}</th>`).join('')}</tr></thead>
    <tbody>
      ${rows.map(r => `<tr>${visibleHeaders.map(h => {
        const v = r[h];
        const isLong = v && String(v).length > 30;
        return `<td class="${isLong ? 'text-cell' : ''}">${v != null ? String(v) : '-'}</td>`;
      }).join('')}</tr>`).join('')}
    </tbody>
  `;
  div.appendChild(table);
  return div;
}

// =============================================
// 접기/펼치기
// =============================================
function toggleCollapse(header) {
  const body = header.nextElementSibling;
  const icon = header.querySelector('.collapse-icon');
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'block';
  icon.textContent = isOpen ? '▼' : '▲';
}

// =============================================
// 초기화
// =============================================
document.addEventListener('DOMContentLoaded', () => {
  initUploader();
  showPage('upload');
});
