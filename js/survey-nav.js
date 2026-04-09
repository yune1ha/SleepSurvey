/**
 * survey-nav.js v3.0 — 설문 내비게이션 · 결과 저장 · 다중 환자 · 통합 Excel 출력
 * 수면 설문지 시스템
 *
 * 경로 기준: surveys/ (index.html 은 ../index.html)
 *
 * [v2.1] file:// localStorage 격리 → URL 파라미터(sn_s, sn_r) 전달로 우회
 * [v3.0] 다중 환자 누적(sn_a), 점수 계산 버튼 하단 이동, 새 환자 버튼, Excel 전체 환자 출력
 */

const SN_SESSION_KEY = 'sleepSurvey_session';
const SN_RESULTS_KEY = 'sleepSurvey_results';
const SN_ALL_KEY     = 'sleepSurvey_allResults';   // 완료된 환자 누적 배열

const SN_SURVEY_META = {
    'SDQ.html':       { name: 'SDQ',       fullName: '강점·난점 설문지 (SDQ)' },
    'K-CSHQ.html':    { name: 'K-CSHQ',    fullName: '소아 수면 습관 설문지 (K-CSHQ)' },
    'PED_MIDAS.html': { name: 'PED-MIDAS', fullName: '두통 영향 설문지 (PED-MIDAS)' },
    'BISQ-R.html':    { name: 'BISQ-R',    fullName: '영아 수면 설문지 (BISQ-R)' },
    'PDSS.html':      { name: 'PDSS',      fullName: '주간 졸림 척도 (PDSS)' }
};

const SN_RESULT_KEYS = {
    'SDQ.html':       'SDQ',
    'K-CSHQ.html':   'KCSHQ',
    'PED_MIDAS.html': 'PEDMIDAS',
    'BISQ-R.html':   'BISQR',
    'PDSS.html':      'PDSS'
};

// ── URL → localStorage 복원 (스크립트 로드 즉시 실행) ─────────────────────────
;(function _snRestoreFromUrl() {
    try {
        const params = new URLSearchParams(window.location.search);
        const rawS = params.get('sn_s');
        const rawR = params.get('sn_r');
        const rawA = params.get('sn_a');
        if (rawS) localStorage.setItem(SN_SESSION_KEY, decodeURIComponent(atob(rawS)));
        if (rawR) localStorage.setItem(SN_RESULTS_KEY, decodeURIComponent(atob(rawR)));
        if (rawA) localStorage.setItem(SN_ALL_KEY,     decodeURIComponent(atob(rawA)));
        if (rawS || rawR || rawA) {
            history.replaceState(null, '', window.location.pathname + window.location.hash);
        }
    } catch (e) { /* 파라미터 없거나 파싱 실패 시 무시 */ }
})();

// ── 세션 관리 ─────────────────────────────────────────────────────────────────
function snGetSession() {
    try { return JSON.parse(localStorage.getItem(SN_SESSION_KEY)); }
    catch (e) { return null; }
}
function snSaveSession(data) {
    localStorage.setItem(SN_SESSION_KEY, JSON.stringify(data));
}

// ── 결과 저장 ─────────────────────────────────────────────────────────────────
function snGetResults() {
    try { return JSON.parse(localStorage.getItem(SN_RESULTS_KEY)) || {}; }
    catch (e) { return {}; }
}

function snSaveResult(surveyFilename, data) {
    const key = SN_RESULT_KEYS[surveyFilename];
    if (!key) return;
    const results = snGetResults();
    results[key] = data;
    localStorage.setItem(SN_RESULTS_KEY, JSON.stringify(results));
    _snUpdateTab(surveyFilename);
}

// ── 다중 환자 누적 ────────────────────────────────────────────────────────────
function snGetAllResults() {
    try { return JSON.parse(localStorage.getItem(SN_ALL_KEY)) || []; }
    catch (e) { return []; }
}
function snSaveAllResults(arr) {
    localStorage.setItem(SN_ALL_KEY, JSON.stringify(arr));
}

// ── 자동 입력 ─────────────────────────────────────────────────────────────────
function snAutoFill(fieldId) {
    const s = snGetSession();
    if (!s || !s.patientId) return;
    const el = document.getElementById(fieldId);
    if (el && !el.value) el.value = s.patientId;
}

// ── 내비게이션 ────────────────────────────────────────────────────────────────
function _snBuildParams() {
    const s = localStorage.getItem(SN_SESSION_KEY);
    const r = localStorage.getItem(SN_RESULTS_KEY) || '{}';  // always pass, even if empty
    const a = localStorage.getItem(SN_ALL_KEY);
    const params = new URLSearchParams();
    if (s) params.set('sn_s', btoa(encodeURIComponent(s)));
    params.set('sn_r', btoa(encodeURIComponent(r)));          // always overwrite destination
    if (a) params.set('sn_a', btoa(encodeURIComponent(a)));
    return params.toString();
}

function snGoTo(filename) {
    const query = _snBuildParams();
    window.location.href = filename + (query ? '?' + query : '');
}

function snGoHome() {
    const query = _snBuildParams();
    window.location.href = '../index.html' + (query ? '?' + query : '');
}

function snGetAdjacentSurveys(currentFilename) {
    const s = snGetSession();
    if (!s) return { prev: null, next: null };
    const idx = s.surveys.indexOf(currentFilename);
    return {
        prev: idx > 0                    ? s.surveys[idx - 1] : null,
        next: idx < s.surveys.length - 1 ? s.surveys[idx + 1] : null
    };
}

function snGoToPrev(currentFilename) {
    const { prev } = snGetAdjacentSurveys(currentFilename);
    if (prev) snGoTo(prev);
}

function snGoToNext(currentFilename) {
    const { next } = snGetAdjacentSurveys(currentFilename);
    if (next) snGoTo(next);
}

/**
 * 이전/다음 설문 버튼을 .btn-calc 앞에 삽입.
 * 마지막 설문이면 .btn-calc 뒤에 "새 환자 추가" 버튼도 삽입.
 */
function snInjectSurveyNavButtons(currentFilename, containerId) {
    if (document.getElementById('sn-survey-nav')) return;
    const container = document.getElementById(containerId);
    if (!container) return;

    const s       = snGetSession();
    const surveys = s ? s.surveys : [];
    const idx     = surveys.indexOf(currentFilename);
    const isFirst = idx <= 0;
    const isLast  = idx === -1 || idx >= surveys.length - 1;

    const calcBtn = container.querySelector('.btn-calc');

    // ── 이전/다음 버튼 wrapper ──
    const wrapper = document.createElement('div');
    wrapper.id = 'sn-survey-nav';
    wrapper.style.cssText = 'display:flex;gap:10px;margin-top:10px;margin-bottom:4px;';

    if (!isFirst) {
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '◀ 이전 설문';
        prevBtn.className   = 'btn-nav';
        prevBtn.style.flex  = '1';
        prevBtn.onclick     = () => snGoToPrev(currentFilename);
        wrapper.appendChild(prevBtn);
    }
    if (!isLast) {
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '다음 설문 ▶';
        nextBtn.className   = 'btn-nav';
        nextBtn.style.flex  = '1';
        nextBtn.onclick     = () => snGoToNext(currentFilename);
        wrapper.appendChild(nextBtn);
    }

    if (wrapper.children.length > 0) {
        if (calcBtn) container.insertBefore(wrapper, calcBtn);
        else         container.appendChild(wrapper);
    }

    // ── 새 환자 추가 버튼 (마지막 설문만) ──
    if (isLast) {
        const newPatBtn = document.createElement('button');
        newPatBtn.id          = 'sn-new-patient-btn';
        newPatBtn.textContent = '+ 새 환자 추가';
        newPatBtn.className   = 'btn-nav';
        newPatBtn.style.cssText = 'width:100%;margin-top:8px;background:#17a2b8;';
        newPatBtn.onclick = snAddNewPatient;
        if (calcBtn) calcBtn.insertAdjacentElement('afterend', newPatBtn);
        else         container.appendChild(newPatBtn);
    }
}

/**
 * 현재 환자 데이터를 누적 배열에 저장하고 홈으로 이동.
 */
function snAddNewPatient() {
    if (!confirm('현재 환자 데이터를 저장하고 새 환자를 입력하시겠습니까?')) return;

    const session = snGetSession();
    const results = snGetResults();

    if (session && Object.keys(results).length > 0) {
        const all = snGetAllResults();
        all.push({ session, results });
        snSaveAllResults(all);
    }

    // 현재 환자 세션/결과 초기화, 누적 데이터는 유지
    localStorage.removeItem(SN_SESSION_KEY);
    localStorage.removeItem(SN_RESULTS_KEY);

    // sn_a 들고 홈으로 (sn_r='{}'로 빈 결과를 명시적으로 전달해 이전 환자 캐시 덮어쓰기)
    const a = localStorage.getItem(SN_ALL_KEY);
    const params = new URLSearchParams();
    if (a) params.set('sn_a', btoa(encodeURIComponent(a)));
    params.set('sn_r', btoa(encodeURIComponent('{}')));
    const query = params.toString();
    window.location.href = '../index.html' + (query ? '?' + query : '');
}

// ── 내비게이션 바 ─────────────────────────────────────────────────────────────
function snInjectNavBar(currentFilename) {
    if (document.getElementById('sn-navbar')) return;

    const s       = snGetSession();
    const results = snGetResults();
    const surveys = s ? s.surveys : [];

    document.body.style.paddingTop = '54px';

    const bar = document.createElement('div');
    bar.id = 'sn-navbar';
    bar.style.cssText = [
        'position:fixed', 'top:0', 'left:0', 'right:0', 'z-index:9999',
        'background:#2c3e50', 'color:#fff',
        'display:flex', 'align-items:center',
        'padding:0 14px', 'height:54px', 'gap:10px',
        "font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif",
        'box-shadow:0 2px 10px rgba(0,0,0,0.4)'
    ].join(';');

    bar.appendChild(_snBtn('← 메인', 'rgba(255,255,255,0.15)', snGoHome, '12px', 'nowrap'));
    bar.appendChild(_snSep());

    if (s) {
        const info = document.createElement('div');
        info.style.cssText = 'font-size:13px;white-space:nowrap;line-height:1.3;';
        info.innerHTML =
            `<strong>${s.patientId}</strong> &nbsp;·&nbsp; 만 <strong>${s.age}세</strong>` +
            `<br><span style="opacity:0.55;font-size:11px">${s.birthdate}</span>`;
        bar.appendChild(info);
        bar.appendChild(_snSep());
    }

    const tabWrap = document.createElement('div');
    tabWrap.id = 'sn-tabs';
    tabWrap.style.cssText = 'display:flex;gap:5px;flex:1;align-items:center;';

    surveys.forEach((file, i) => {
        const meta    = SN_SURVEY_META[file] || { name: file };
        const key     = SN_RESULT_KEYS[file];
        const saved   = !!(key && results[key]);
        const current = file === currentFilename;

        const tab = document.createElement('button');
        tab.id            = `sn-tab-${file.replace(/\./g, '_')}`;
        tab.dataset.idx   = i + 1;
        tab.dataset.file  = file;
        tab.style.cssText = _snTabStyle(current, saved);
        tab.innerHTML     = `${i + 1}.&nbsp;${meta.name}${saved ? '&nbsp;<span style="color:#2ecc71;font-size:13px">✓</span>' : ''}`;
        tab.onclick       = () => snGoTo(file);

        if (!current) {
            tab.onmouseover = () => { tab.style.background = 'rgba(255,255,255,0.25)'; };
            tab.onmouseout  = () => { tab.style.background = 'rgba(255,255,255,0.12)'; };
        }
        tabWrap.appendChild(tab);
    });
    bar.appendChild(tabWrap);

    bar.appendChild(_snSep());
    const viewBtn = _snBtn('결과 보기', '#8e44ad', snShowResultsModal, '12.5px', 'nowrap');
    viewBtn.style.fontWeight = 'bold';
    bar.appendChild(viewBtn);
    const expBtn = _snBtn('결과 출력 (Excel)', '#27ae60', snExportAllResults, '12.5px', 'nowrap');
    expBtn.style.fontWeight = 'bold';
    bar.appendChild(expBtn);

    document.body.insertBefore(bar, document.body.firstChild);
}

function _snUpdateTab(savedFilename) {
    const tab = document.getElementById(`sn-tab-${savedFilename.replace(/\./g, '_')}`);
    if (!tab) return;
    const meta = SN_SURVEY_META[savedFilename] || { name: savedFilename };
    tab.innerHTML = `${tab.dataset.idx}.&nbsp;${meta.name}&nbsp;<span style="color:#2ecc71;font-size:13px">✓</span>`;
}

function _snBtn(text, bg, onclick, fontSize, whiteSpace) {
    const b = document.createElement('button');
    b.innerHTML = text;
    b.style.cssText = [
        `background:${bg}`, 'border:none', 'color:#fff',
        'padding:6px 12px', 'border-radius:5px', 'cursor:pointer',
        `font-size:${fontSize || '12px'}`,
        `white-space:${whiteSpace || 'normal'}`,
        'font-family:inherit', 'transition:opacity 0.15s'
    ].join(';');
    b.onmouseover = () => { b.style.opacity = '0.82'; };
    b.onmouseout  = () => { b.style.opacity = '1'; };
    b.onclick = onclick;
    return b;
}

function _snSep() {
    const d = document.createElement('div');
    d.style.cssText = 'width:1px;height:30px;background:rgba(255,255,255,0.2);flex-shrink:0;';
    return d;
}

function _snTabStyle(isCurrent, isSaved) {
    return [
        'border:none', 'border-radius:5px', 'cursor:pointer',
        'padding:5px 11px', 'font-size:12px', 'font-weight:bold',
        'font-family:inherit', 'color:#fff',
        `background:${isCurrent ? '#3498db' : 'rgba(255,255,255,0.12)'}`,
        `box-shadow:${isCurrent ? 'inset 0 0 0 2px rgba(255,255,255,0.4)' : 'none'}`
    ].join(';');
}

// ── 결과 보기 모달 ────────────────────────────────────────────────────────────
function snShowResultsModal() {
    if (document.getElementById('sn-results-modal')) return;

    const all = snGetAllResults();
    const curSession = snGetSession();
    const curResults = snGetResults();
    const patients = [...all];
    if (curSession && Object.keys(curResults).length > 0) {
        patients.push({ session: curSession, results: curResults });
    }

    // ── 오버레이 ──
    const overlay = document.createElement('div');
    overlay.id = 'sn-results-modal';
    overlay.style.cssText = [
        'position:fixed', 'inset:0', 'z-index:99999',
        'background:rgba(0,0,0,0.55)',
        'display:flex', 'align-items:center', 'justify-content:center',
        'padding:20px', "font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif"
    ].join(';');
    overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

    // ── 다이얼로그 박스 ──
    const box = document.createElement('div');
    box.style.cssText = [
        'background:#fff', 'border-radius:12px',
        'box-shadow:0 8px 40px rgba(0,0,0,0.35)',
        'max-width:95vw', 'max-height:85vh',
        'overflow:auto', 'padding:28px 30px', 'min-width:600px'
    ].join(';');
    overlay.appendChild(box);

    // ── 헤더 ──
    const hdr = document.createElement('div');
    hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;';
    const title = document.createElement('h2');
    title.textContent = '현재까지 입력된 결과';
    title.style.cssText = 'margin:0;color:#2c3e50;font-size:1.2em;';
    const closeBtn = document.createElement('button');
    closeBtn.textContent = '✕';
    closeBtn.style.cssText = [
        'background:none', 'border:none', 'font-size:1.4em',
        'cursor:pointer', 'color:#888', 'padding:4px 8px', 'border-radius:4px'
    ].join(';');
    closeBtn.onmouseover = () => { closeBtn.style.background = '#f0f0f0'; };
    closeBtn.onmouseout  = () => { closeBtn.style.background = 'none'; };
    closeBtn.onclick = () => overlay.remove();
    hdr.appendChild(title);
    hdr.appendChild(closeBtn);
    box.appendChild(hdr);

    if (patients.length === 0) {
        const empty = document.createElement('p');
        empty.textContent = '저장된 결과가 없습니다.';
        empty.style.cssText = 'color:#888;text-align:center;padding:30px 0;';
        box.appendChild(empty);
        document.body.appendChild(overlay);
        return;
    }

    // ── 어떤 설문 컬럼을 보여줄지 판단 ──
    const hasSurvey = { SDQ: false, KCSHQ: false, PEDMIDAS: false, PDSS: false, BISQR: false };
    patients.forEach(({ results: r }) => {
        if (r.SDQ)      hasSurvey.SDQ      = true;
        if (r.KCSHQ)    hasSurvey.KCSHQ    = true;
        if (r.PEDMIDAS) hasSurvey.PEDMIDAS = true;
        if (r.PDSS)     hasSurvey.PDSS     = true;
        if (r.BISQR)    hasSurvey.BISQR    = true;
    });

    // ── 테이블 ──
    const table = document.createElement('table');
    table.style.cssText = 'width:100%;border-collapse:collapse;font-size:0.9em;';

    // 헤더 행
    const thead = document.createElement('thead');
    const hrow  = document.createElement('tr');
    const cols = ['#', '환자번호', '생년월일', '만나이', '두통'];
    if (hasSurvey.SDQ)      cols.push('SDQ');
    if (hasSurvey.KCSHQ)    cols.push('K-CSHQ A', 'K-CSHQ B', 'K-CSHQ C', 'K-CSHQ D', 'K-CSHQ E', 'K-CSHQ 합');
    if (hasSurvey.PEDMIDAS) cols.push('PED-MIDAS');
    if (hasSurvey.PDSS)     cols.push('PDSS');
    if (hasSurvey.BISQR)    cols.push('BISQ-R');
    cols.forEach(c => {
        const th = document.createElement('th');
        th.textContent = c;
        th.style.cssText = [
            'background:#2c3e50', 'color:#fff', 'padding:10px 12px',
            'text-align:center', 'white-space:nowrap',
            'border:1px solid #34495e', 'font-size:0.85em'
        ].join(';');
        hrow.appendChild(th);
    });
    thead.appendChild(hrow);
    table.appendChild(thead);

    // 데이터 행
    const tbody = document.createElement('tbody');
    patients.forEach(({ session: s, results: r }, i) => {
        const tr = document.createElement('tr');
        tr.style.background = i % 2 === 0 ? '#fff' : '#f8f9fa';

        const cells = [
            i + 1,
            s.patientId,
            s.birthdate,
            s.age + '세',
            s.hasHeadache ? '있음' : '없음'
        ];
        if (hasSurvey.SDQ)      cells.push(r.SDQ      ? r.SDQ.score               : '-');
        if (hasSurvey.KCSHQ) {
            const sc = r.KCSHQ ? r.KCSHQ.scores : null;
            cells.push(sc ? sc.A : '-', sc ? sc.B : '-', sc ? sc.C : '-',
                       sc ? sc.D : '-', sc ? sc.E : '-', sc ? sc.total : '-');
        }
        if (hasSurvey.PEDMIDAS) cells.push(r.PEDMIDAS ? r.PEDMIDAS.score           : '-');
        if (hasSurvey.PDSS)     cells.push(r.PDSS     ? r.PDSS.score               : '-');
        if (hasSurvey.BISQR)    cells.push(r.BISQR    ? '있음'                     : '-');

        cells.forEach((val, ci) => {
            const td = document.createElement('td');
            td.textContent = val;
            td.style.cssText = [
                'padding:9px 12px', 'border:1px solid #dee2e6',
                'text-align:center', 'white-space:nowrap',
                ci === 1 ? 'font-weight:bold;' : ''
            ].join(';');
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    box.appendChild(table);

    // 환자 수 요약
    const footer = document.createElement('p');
    footer.textContent = `총 ${patients.length}명`;
    footer.style.cssText = 'margin-top:12px;font-size:0.85em;color:#888;text-align:right;';
    box.appendChild(footer);

    document.body.appendChild(overlay);
}

// ── 통합 Excel 출력 (전체 환자) ───────────────────────────────────────────────
function snExportAllResults() {
    if (typeof XLSX === 'undefined') {
        alert('Excel 라이브러리 로딩 중입니다. 잠시 후 다시 시도해주세요.');
        return;
    }

    // 누적 환자 + 현재 환자 합산
    const all = snGetAllResults();
    const curSession = snGetSession();
    const curResults = snGetResults();
    const patients = [...all];
    if (curSession && Object.keys(curResults).length > 0) {
        patients.push({ session: curSession, results: curResults });
    }

    if (patients.length === 0) {
        alert('저장된 설문 결과가 없습니다.\n각 설문에서 점수 계산/저장 버튼을 먼저 눌러주세요.');
        return;
    }

    const wb    = XLSX.utils.book_new();
    const today = new Date().toISOString().split('T')[0];

    // ── 시트 1: 결과 요약 (환자별 1행) ──
    const summaryRows = patients.map(({ session: s, results: r }) => {
        const row = {
            '환자번호': s.patientId,
            '생년월일': s.birthdate,
            '만나이':   s.age,
            '두통여부': s.hasHeadache ? '예' : '아니오',
            '검사일':   today
        };
        if (r.SDQ)      { row['SDQ_총점']      = r.SDQ.score; }
        if (r.KCSHQ)    {
            row['KCSHQ_A']    = r.KCSHQ.scores?.A;
            row['KCSHQ_B']    = r.KCSHQ.scores?.B;
            row['KCSHQ_C']    = r.KCSHQ.scores?.C;
            row['KCSHQ_D']    = r.KCSHQ.scores?.D;
            row['KCSHQ_E']    = r.KCSHQ.scores?.E;
            row['KCSHQ_총점'] = r.KCSHQ.scores?.total;
        }
        if (r.PEDMIDAS) { row['PEDMIDAS_총점'] = r.PEDMIDAS.score; }
        if (r.PDSS)     { row['PDSS_총점']     = r.PDSS.score; }
        if (r.BISQR)    { row['BISQR_저장여부'] = '있음'; }
        return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), '결과요약');

    // ── 시트 2: SDQ 상세 ──
    const sdqRows = patients.filter(p => p.results.SDQ).map(({ results: r }) => {
        const row = { '환자번호': r.SDQ.patientId, '생년월일': r.SDQ.birthdate, '만나이': r.SDQ.age };
        Object.entries(r.SDQ.answers || {}).forEach(([k, v]) => { row[`Q${k}`] = v; });
        row['총점'] = r.SDQ.score;
        return row;
    });
    if (sdqRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sdqRows), 'SDQ');

    // ── 시트 3: K-CSHQ 상세 ──
    const kcshqRows = patients.filter(p => p.results.KCSHQ).map(({ results: r }) => {
        const row = { '환자번호': r.KCSHQ.patientId, '생년월일': r.KCSHQ.birthdate, '만나이': r.KCSHQ.age };
        const sc  = r.KCSHQ.scores || {};
        Object.assign(row, { 'A영역': sc.A, 'B영역': sc.B, 'C영역': sc.C, 'D영역': sc.D, 'E영역': sc.E, '총점': sc.total });
        const sub = r.KCSHQ.subjective || {};
        if (sub.bedtime)  row['취침시각']   = sub.bedtime;
        if (sub.duration) row['수면시간']   = sub.duration;
        if (sub.pain)     row['통증부위']   = sub.pain;
        if (sub.awake)    row['깬시간(분)'] = sub.awake;
        if (sub.wakeup)   row['기상시각']   = sub.wakeup;
        Object.entries(r.KCSHQ.answers || {}).forEach(([k, v]) => { row[`Q${k}`] = v; });
        return row;
    });
    if (kcshqRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kcshqRows), 'K-CSHQ');

    // ── 시트 4: PED-MIDAS 상세 ──
    const pedRows = patients.filter(p => p.results.PEDMIDAS).map(({ results: r }) => {
        const row = { '환자번호': r.PEDMIDAS.patientId, '생년월일': r.PEDMIDAS.birthdate, '만나이': r.PEDMIDAS.age };
        Object.entries(r.PEDMIDAS.answers || {}).forEach(([k, v]) => { row[`Q${k}`] = v; });
        row['총점'] = r.PEDMIDAS.score;
        return row;
    });
    if (pedRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pedRows), 'PED-MIDAS');

    // ── 시트 5: BISQ-R 상세 ──
    const bisqrRows = patients.filter(p => p.results.BISQR).map(({ results: r }) => {
        const row = { '환자번호': r.BISQR.patientId, '생년월일': r.BISQR.birthdate, '만나이': r.BISQR.age };
        Object.entries(r.BISQR.answers || {}).forEach(([k, v]) => { row[k] = v; });
        return row;
    });
    if (bisqrRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bisqrRows), 'BISQ-R');

    // ── 시트 6: PDSS 상세 ──
    const pdssRows = patients.filter(p => p.results.PDSS).map(({ results: r }) => {
        const row = { '환자번호': r.PDSS.patientId, '생년월일': r.PDSS.birthdate, '만나이': r.PDSS.age };
        Object.entries(r.PDSS.answers || {}).forEach(([k, v]) => { row[`Q${k}`] = v; });
        row['총점'] = r.PDSS.score;
        return row;
    });
    if (pdssRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pdssRows), 'PDSS');

    XLSX.writeFile(wb, `SleepSurvey_${today}.xlsx`);

    // 출력 후 누적 데이터 초기화
    localStorage.removeItem(SN_ALL_KEY);
}
