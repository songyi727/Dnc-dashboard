// block_common.js - index.html / manage.html 공용 로직
// window.BLOCK_DETAIL (block_data.js) 을 각 페이지가 쓰는 형태로 변환/집계한다.

const BlockCommon = (function () {
  const GROUPS = ["나보타", "브이올렛", "필러군", "리프팅실", "리알로"];

  const fmt = n => Math.round(n || 0).toLocaleString("ko-KR");
  const fmt1 = n => (n || 0).toLocaleString("ko-KR", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const pct = (a, b) => (!b ? 0 : Math.round((a / b) * 100));

  function quarterOfMonth(m) {
    return "Q" + Math.ceil(m / 3);
  }

  const QUARTER_INDEX = { "1Q": 0, "2Q": 1, "3Q": 2, "4Q": 3 };

  // quarters(예: ["1Q","3Q"])에 해당하는 월들만 골라서 actual_by_month 합산. quarters가 비어있으면(전체) 전부 합산.
  function sumActualForQuarters(actualByMonth, quarters) {
    let sum = 0;
    Object.keys(actualByMonth || {}).forEach(ym => {
      if (!quarters || !quarters.length) { sum += actualByMonth[ym]; return; }
      const m = Number(ym.split("-")[1]);
      const qLabel = Math.ceil(m / 3) + "Q";
      if (quarters.includes(qLabel)) sum += actualByMonth[ym];
    });
    return sum;
  }

  // mbo_q([1Q,2Q,3Q,4Q] 배열)에서 quarters에 해당하는 것만 합산. quarters가 비어있으면 전체 합산.
  function sumMboForQuarters(mboQ, quarters) {
    if (!mboQ) return 0;
    if (!quarters || !quarters.length) return mboQ.reduce((a, b) => a + b, 0);
    return quarters.reduce((a, ql) => a + (mboQ[QUARTER_INDEX[ql]] || 0), 0);
  }

  function niceCeil(v) {
    if (v <= 0) return 100;
    const exp = Math.floor(Math.log10(v));
    const base = Math.pow(10, exp);
    const mult = v / base;
    let nice;
    if (mult <= 1) nice = 1; else if (mult <= 2) nice = 2; else if (mult <= 5) nice = 5; else nice = 10;
    return nice * base;
  }

  // 특정 연도(year='all'이면 가장 최근 연도)의 품목군 스냅샷을 고른다
  function pickYearSnapshot(grp, year) {
    const byYear = (grp && grp.by_year) || {};
    const keys = Object.keys(byYear);
    if (!keys.length) return null;
    if (year === "all" || year == null) {
      const latest = keys.sort().pop();
      return byYear[latest];
    }
    return byYear[String(year)] || null;
  }

  // 이미 특정 연도로 스코프된 actual_by_month를 분기별로 집계 (연도 필터링 불필요)
  function actualByQuarterAny(monthMap) {
    const q = { 1: 0, 2: 0, 3: 0, 4: 0 };
    Object.keys(monthMap || {}).forEach(ym => {
      const m = Number(ym.split("-")[1]);
      q[Math.ceil(m / 3)] += monthMap[ym];
    });
    return q;
  }

  // recalled: 관리자가 실제로 회수 처리했는지 여부 (수동 입력값, MBO/기간 계산과 무관하게 최우선)
  function evalStatus(monthsSinceContract, totals, rate, recalled) {
    if (recalled) return { key: "recalled", label: "회수" };
    if (monthsSinceContract == null) return { key: "hold", label: "보류" };
    const activeGroups = Object.values(totals).filter(g => g.actual > 0).length;
    if (monthsSinceContract < 3) return { key: "hold", label: "보류" };
    if (monthsSinceContract >= 6 && activeGroups <= 1 && rate < 50) return { key: "recall", label: "회수대상" };
    if (monthsSinceContract >= 3 && activeGroups <= 1 && rate < 50) return { key: "warn", label: "1차경고" };
    return { key: "keep", label: "유지" };
  }

  // 블록처 1건(BLOCK_DETAIL의 원소)을 manage.html이 쓰는 형태로 변환.
  // year: 'all'(최신 연도 스냅샷) 또는 특정 연도 문자열/숫자 - 그 연도에 실제 계약된 품목군만 반영됨.
  function transformBlock(record, meta, year, refMonth) {
    const yearSel = year == null ? "all" : year;
    const monthKey = refMonth || meta.current_month;
    const q = {};
    const totals = {};
    const contractedFlags = {};

    const yearRemarks = [];
    GROUPS.forEach(g => {
      const grp = record.groups[g] || { by_year: {}, actual_by_month: {} };
      const snap = pickYearSnapshot(grp, yearSel);
      const mboQ = snap ? snap.mbo_q : [0, 0, 0, 0];
      const yearActuals = snap ? snap.actual_by_month : {};
      const qa = actualByQuarterAny(yearActuals);
      q[g] = {
        "1Q": { mbo: mboQ[0] || 0, actual: qa[1] },
        "2Q": { mbo: mboQ[1] || 0, actual: qa[2] },
        "3Q": { mbo: mboQ[2] || 0, actual: qa[3] },
        "4Q": { mbo: mboQ[3] || 0, actual: qa[4] },
      };
      const mboSum = (mboQ || []).reduce((a, b) => a + b, 0);
      const actualSum = qa[1] + qa[2] + qa[3] + qa[4];
      totals[g] = { mbo: mboSum, actual: actualSum };
      contractedFlags[g] = !!snap;
      if (snap && snap.remark) yearRemarks.push(snap.remark);
    });

    // 선택 연도(yearSel) 안에서의 비고 - "회수"가 하나라도 있으면 최우선으로 표시
    const yearRemark = yearRemarks.find(r => r.includes("회수")) || yearRemarks[0] || "";
    const yearRecalled = yearRemark.includes("회수");

    const mbo = GROUPS.reduce((a, g) => a + totals[g].mbo, 0);
    const ytd = GROUPS.reduce((a, g) => a + totals[g].actual, 0);
    const groupMonth = {};
    GROUPS.forEach(g => {
      const grp = record.groups[g] || { actual_by_month: {} };
      groupMonth[g] = contractedFlags[g] ? (grp.actual_by_month[monthKey] || 0) : 0;
    });
    const monthActual = GROUPS.reduce((a, g) => a + groupMonth[g], 0);
    const trend = meta.recent_months.map(ym => {
      const v = GROUPS.reduce((a, g) => {
        if (!contractedFlags[g]) return a; // 계약 안 된 품목군의 매출은 제외
        const grp = record.groups[g] || { actual_by_month: {} };
        return a + (grp.actual_by_month[ym] || 0);
      }, 0);
      const [y, m] = ym.split("-");
      return { m: Number(m) + "월", v };
    });

    return {
      manager: record.manager,
      team: record.team || "-",
      rep: record.rep || "-",
      name: record.name,
      biz: record.biz,
      contractYears: record.client_contract_years && record.client_contract_years.length
        ? record.client_contract_years
        : (record.contract_date ? [Number(String(record.contract_date).slice(0, 4))] : []),
      mbo, ytd, month: monthActual,
      months: record.months_since_contract,
      recalled: yearRecalled,
      remark: yearRemark,
      q, totals, contractedFlags, groupMonth,
      trend,
    };
  }

  function transformAll(detail, meta, year, refMonth) {
    return detail.map(r => transformBlock(r, meta, year, refMonth));
  }

  // 담당자(블록담당자) 단위 요약 - index.html 현황 요약용
  // companyInfo가 있으면 "업체정보"에 등록된 전 담당자를 기준으로 순회한다(가결 건이 0개여도 표시).
  // year: 'all'(최신 연도 스냅샷) 또는 특정 연도 - 그 연도에 실제 계약된 품목군 기준으로 MBO/평가 계산.
  function aggregateSummary(detail, meta, companyInfo, year, quarters) {
    const yearSel = year == null ? "all" : year;
    const qtrs = quarters && quarters.length ? quarters : null;
    const byManagerBiz = {};
    detail.forEach(r => {
      const key = r.manager_biz != null ? r.manager_biz : r.manager;
      if (!byManagerBiz[key]) byManagerBiz[key] = [];
      byManagerBiz[key].push(r);
    });

    const managerKeys = companyInfo ? Object.keys(companyInfo) : Object.keys(byManagerBiz);

    const rows = [];
    managerKeys.forEach(bizKey => {
      const allRecords = byManagerBiz[bizKey] || [];
      const records = yearSel === "all"
        ? allRecords
        : allRecords.filter(rec => (rec.client_contract_years || []).some(y => String(y) === String(yearSel)));
      const info = companyInfo ? companyInfo[bizKey] : null;
      const name = allRecords[0] ? allRecords[0].manager : (info ? info.name : bizKey);
      const contractYears = allRecords[0] ? (allRecords[0].manager_contract_years || []) : (info ? info.years : []);

      let mbo = 0;
      let earliestDate = null;
      let validBlocks = 0;
      const itemsByYear = {};
      const mboByGroup = { 나보타: 0, 브이올렛: 0, 필러군: 0, 리프팅실: 0, 리알로파인: 0 };
      const comp = { 3: 0, 2: 0, 1: 0, 0: 0 };
      const ev = { 보류: 0, 경고: 0, 회수대상: 0, 회수: 0, 유지: 0 };

      records.forEach(rec => {
        if (yearSel === "all") {
          if (rec.contract_date && (!earliestDate || rec.contract_date < earliestDate)) earliestDate = rec.contract_date;
        } else {
          GROUPS.forEach(g => {
            const grp = rec.groups[g] || { by_year: {} };
            const snap = (grp.by_year || {})[String(yearSel)];
            if (snap && snap.contract_date && (!earliestDate || snap.contract_date < earliestDate)) {
              earliestDate = snap.contract_date;
            }
          });
        }

        let recMboTotal = 0;
        const groupTotalsThisYear = {};
        const recYearRemarks = [];
        GROUPS.forEach(g => {
          const grp = rec.groups[g] || { by_year: {}, actual_by_month: {} };
          const snap = pickYearSnapshot(grp, yearSel);
          const mboSum = snap ? sumMboForQuarters(snap.mbo_q, qtrs) : 0;
          recMboTotal += mboSum;
          const actualForYear = snap
            ? sumActualForQuarters(snap.actual_by_month, qtrs)
            : 0;
          groupTotalsThisYear[g] = { mbo: mboSum, actual: actualForYear };
          const gKey = g === "리알로" ? "리알로파인" : g;
          mboByGroup[gKey] += mboSum;
          if (snap && snap.remark) recYearRemarks.push(snap.remark);

          Object.entries(grp.by_year || {}).forEach(([y, ySnap]) => {
            if (!itemsByYear[y]) itemsByYear[y] = { 나보타: 0, 브이올렛: 0, 필러군: 0, 리프팅실: 0, 리알로파인: 0 };
            const yearSum = sumActualForQuarters(ySnap.actual_by_month, qtrs);
            itemsByYear[y][gKey] += yearSum;
          });
        });
        mbo += recMboTotal;
        const recRecalled = recYearRemarks.some(r => r.includes("회수"));

        const activeGroups = Object.values(groupTotalsThisYear).filter(x => x.actual > 0).length;
        if (activeGroups >= 1) validBlocks += 1;
        const compLevel = Math.min(3, Math.max(0, activeGroups - 1));
        comp[compLevel] += 1;

        const rate = pct(
          Object.values(groupTotalsThisYear).reduce((a, x) => a + x.actual, 0),
          recMboTotal
        );
        const status = evalStatus(rec.months_since_contract, groupTotalsThisYear, rate, recRecalled);
        if (status.key === "hold") ev.보류 += 1;
        else if (status.key === "warn") ev.경고 += 1;
        else if (status.key === "recall") ev.회수대상 += 1;
        else if (status.key === "recalled") ev.회수 += 1;
        else ev.유지 += 1;
      });

      rows.push({
        name,
        contractYears,
        blocks: records.length,
        validBlocks,
        mbo: mbo / 1000000, // 원 단위 -> 백만원 단위 (요약 테이블 스케일에 맞춤)
        mboByGroup: Object.fromEntries(Object.entries(mboByGroup).map(([k, v]) => [k, v / 1000000])),
        date: earliestDate ? earliestDate.replace(/-/g, ".") : "-",
        itemsByYear: Object.fromEntries(
          Object.entries(itemsByYear).map(([y, v]) => [y, Object.fromEntries(Object.entries(v).map(([k, vv]) => [k, vv / 1000000]))])
        ),
        comp,
        ev,
      });
    });

    return rows;
  }

  // rec가 특정 연도(year) 안에서 처음 계약된 날짜("YYYY-MM-DD") - 품목군 중 가장 이른 by_year 계약일
  function yearEarliestContractDate(rec, year) {
    let earliest = null;
    GROUPS.forEach(g => {
      const snap = ((rec.groups[g] || {}).by_year || {})[String(year)];
      if (snap && snap.contract_date && (!earliest || snap.contract_date < earliest)) earliest = snap.contract_date;
    });
    return earliest;
  }

  // 계약월 누적 필터: yearSel이 특정 연도이고 cMonthSel이 특정 월(라벨 예: "07월" 또는 "7")이면
  // 그 달까지(1~cMonthSel월) 계약된 처만 남긴다. 'all'이면 전체 그대로 반환.
  function scopeDetailByContractMonth(detail, yearSel, cMonthSel) {
    if (yearSel === "all" || !cMonthSel || cMonthSel === "all") return detail;
    return detail.filter(rec => {
      const cd = yearEarliestContractDate(rec, yearSel);
      if (!cd) return false;
      return parseInt(cd.slice(5, 7), 10) <= parseInt(cMonthSel, 10);
    });
  }

  return {
    GROUPS, fmt, fmt1, pct, quarterOfMonth, niceCeil, evalStatus,
    transformBlock, transformAll, aggregateSummary,
    yearEarliestContractDate, scopeDetailByContractMonth,
    sumActualForQuarters, sumMboForQuarters,
  };
})();
