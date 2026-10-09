const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, PageOrientation, BorderStyle, LevelFormat, Footer, PageNumber,
  TableOfContents, PageBreak,
} = require('docx');

const FONT = '맑은 고딕';
const W = 14838; // landscape A4 content width (DXA)
const NAVY = '1F3864', LIGHT = 'DCE6F2', GOLD = 'FFF2CC', GRAY = 'F2F2F2';

const p = (text, opt = {}) => new Paragraph({
  spacing: { after: 80 }, ...opt.para,
  children: [].concat(text).map(t => typeof t === 'string'
    ? new TextRun({ text: t, font: FONT, size: opt.size || 19, bold: opt.bold, color: opt.color })
    : t),
});
const r = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 19, bold: o.bold, color: o.color, italics: o.italics });
const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: t, font: FONT, size: 30, bold: true, color: NAVY })] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 180, after: 100 }, children: [new TextRun({ text: t, font: FONT, size: 24, bold: true, color: NAVY })] });
const bullet = (text, lvl = 0) => new Paragraph({ numbering: { reference: 'b', level: lvl }, spacing: { after: 40 },
  children: [].concat(text).map(t => typeof t === 'string' ? r(t) : t) });
const note = t => p([r(t, { size: 16, color: '595959', italics: true })]);

function table(headers, rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const scale = W / total;
  const ws = widths.map(w => Math.floor(w * scale));
  ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'A6A6A6' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (txt, i, isHead, shade) => new TableCell({
    width: { size: ws[i], type: WidthType.DXA }, borders,
    margins: { top: 50, bottom: 50, left: 70, right: 70 },
    shading: isHead ? { type: ShadingType.CLEAR, color: 'auto', fill: NAVY } : (shade ? { type: ShadingType.CLEAR, color: 'auto', fill: shade } : undefined),
    children: String(txt).split('\n').map(line => new Paragraph({
      alignment: (opts.center || []).includes(i) || isHead ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: line, font: FONT, size: opts.size || 16, bold: isHead || (opts.boldFirst && i === 0), color: isHead ? 'FFFFFF' : undefined })],
    })),
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: ws,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((row, ri) => new TableRow({ children: row.map((c, i) => cell(c, i, false, (opts.highlight || []).includes(ri) ? GOLD : (ri % 2 ? GRAY : undefined))) })),
    ],
  });
}
const gap = () => new Paragraph({ spacing: { after: 60 }, children: [] });

// ---------------- 데이터 (단위: 억원, 34평 = 전용 약 84㎡) ----------------
// est: 추정치 여부
// ---------------- 데이터: scripts/data/complexes.json ----------------
const path = require('path');
const DATA = JSON.parse(fs.readFileSync(process.argv[3] || path.join(__dirname, 'data', 'complexes.json'), 'utf8'));
const ASOF = DATA.asOf;
const D = DATA.complexes.map(c => ({
  k: c.name, short: c.short, gu: c.gu, hh: c.households, yr: c.moveIn, sale: c.sale, saleEst: c.saleEst,
  jeonse: c.jeonse, jEst: c.jeonseEst, saleBasis: c.saleBasis, jBasis: c.jeonseBasis, hist: c.history,
  g: c.growth.base, gBull: c.growth.bull, gBear: c.growth.bear,
}));
const ACQ = DATA.assumptions.acqTaxRate, BROKER = DATA.assumptions.brokerRate; // 취득세+지방교육세(1주택/비조정 2주택, 6억 전후 약 1.1%), 중개보수 상한 0.4%
const f = x => x.toFixed(2);
const pct = x => (x * 100).toFixed(1) + '%';
D.forEach(d => {
  d.cost = d.sale * (ACQ + BROKER);
  d.total = d.sale + d.cost;
  d.inv = d.sale - d.jeonse + d.cost;
  d.ratio = d.jeonse / d.sale;
  d.jPerInv = d.jeonse / d.inv;
  ['g', 'gBull', 'gBear'].forEach(s => {
    d[s + 'P'] = d.sale * (1 + d[s]);
    d[s + 'Profit'] = d[s + 'P'] - d.sale - d.cost;
    d[s + 'Ret'] = d[s + 'Profit'] / d.inv;
  });
});

// 점수(가중치: 가격상승기대30 입지·교통20 갭·전세가율15 상품성15 학군·초품아10 공급리스크10)
const S = Object.fromEntries(DATA.complexes.map(c => [c.name, c.scores]));
const M = D.find(d => d.k === '더샵명지퍼스트월드3단지') || D[3];
const ranked = D.map(d => ({ ...d, sc: S[d.k], tot: S[d.k].reduce((a, b) => a + b, 0) })).sort((a, b) => b.tot - a.tot);

// ---------------- 문서 ----------------
const kids = [];
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 200 }, children: [r('부산 서부권 34평 아파트 7개 단지', { size: 36, bold: true, color: NAVY })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [r('매매·전세·갭·수익률 비교 및 매수전략 보고서', { size: 30, bold: true, color: NAVY })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [r(`사하구 괴정동 · 강서구 명지동 · 서구 서대신동 | 작성일 ${ASOF}`, { size: 20, color: '595959' })] }));

kids.push(table(['핵심 결론 (결론 → 근거 → 실행)'], [[
  `① 향후 가격 인상 기대 1위: ${ranked[0].k} (종합 ${ranked[0].tot}점). 전세가율 ${pct(ranked[0].ratio)}로 갭(투자금 약 ${f(ranked[0].inv)}억)이 작음.\n` +
  `② 2~4위: ${ranked[1].k}(${ranked[1].tot}점) · ${ranked[2].k}(${ranked[2].tot}점) · ${ranked[3].k}(${ranked[3].tot}점). 서대신동은 재개발로 신축 브랜드 타운이 형성되고 있고 전세 매물이 부족함.\n` +
  '③ 명지(2·3단지): 하단~녹산선 호재가 가장 크지만, 강서구 매매가 주간 -0.15%(부동산원 9/14)이고 에코델타시티 입주로 공급 부담이 있음. 전세가율 55~58%라 갭이 큼 → "착공 확인 후 급매" 전략.\n' +
  '④ 부산 전체 매매는 약보합(-0.02%, 9/14)이고 전세는 상승 중. 단기 급등보다는 "전세가 오르며 갭이 줄어드는 구간"을 노리는 저가 분할매수가 유효.'
]], [1], { size: 18 }));
kids.push(gap());
kids.push(table(['⚠ 데이터 신뢰도 고지 (반드시 확인)'], [[
  '• 국토부 실거래가 공개시스템·KB부동산·호갱노노 원문은 작성 환경의 네트워크 정책으로 직접 조회할 수 없었음. 수치는 웹 검색 결과에 노출된 실거래·시세 스니펫을 교차 확인한 값임.\n' +
  '• 표에서 "(추정)"으로 표시한 값(일부 전세가, 명지3단지 매매가)은 애널리스트 추정치임. 대지지분·거래량·회전율·분양가(프리미엄)·3년 실거래 추이 대부분은 "확인 필요"로 남김.\n' +
  '• 예상시세·수익률은 가정에 따른 시뮬레이션이며 투자 권유가 아님. 계약 전 rt.molit.go.kr(국토부 실거래가)과 등기부등본을 반드시 확인할 것.'
]], [1], { size: 17 }));

kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(h1('목차'));
kids.push(new TableOfContents('목차', { hyperlink: true, headingStyleRange: '1-2' }));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// 1. 시장 환경
kids.push(h1('1. 시장 환경 (부산·금리·세제)'));
kids.push(table(['항목', '수치/내용', '기준 시점', '출처', '시사점'], [
  ['부산 아파트 매매', '-0.02% (2주 연속 동일 낙폭)', '2026.9.14 주간', '한국부동산원(보도)', '약보합. 급매 위주 거래'],
  ['강서구 / 사하구 매매', '-0.15% / -0.05%', '2026.9.14 주간', '한국부동산원(보도)', '명지 약세 > 사하 약세'],
  ['서구 매매', '확인 필요', '-', '-', '원자료 확인 필요'],
  ['부산 전세', '상승 (해운대·강서 주도)', '2026.9월', '브릿지경제 9/18 보도', '전세 상승 → 갭 축소 구간'],
  ['부산 입주 물량', '2026년 1만3,207가구 / 2027년 1만6,346가구', '2025.12 전망', '보도(다음뉴스)', '2027년 증가. 강서(에코델타·명지)·남구에 집중'],
  ['기준금리', '3.00% (8/27, 2회 연속 인상)', '2026.8.27', '한국은행', '주담대 부담↑ → 갭투자 수요 위축'],
  ['세제 개편안(정부안)', '장기보유특별공제를 거주 기준으로 재편, 종부세 1주택 공제 실거주 14억/비거주 9억', '2026.8 발표', '기재부(보도)', '6억대 부산 주택은 종부세 영향 작음. 양도세는 실거주 여부가 핵심'],
], [14, 22, 11, 14, 22], { center: [2] }));
kids.push(note('※ 부산은 조정대상지역이 아님(비규제) → 2주택까지 취득세 일반세율(1~3%) 적용. 3주택부터 중과(8%). 규제 변경 여부는 계약 시점에 재확인 필요.'));

// 2. 단지 개요
kids.push(h1('2. 단지 개요 · 상품성 · 입지 · 학군 · 인프라'));
kids.push(note('※ 사용자가 말한 "힐스테이트 사하"는 사하구 괴정동 "힐스테이트사하역"으로, "대신 해모로"는 "대신해모로센트럴"로 해석함.'));
kids.push(table(['단지', '위치', '세대수', '입주', '상품성', '입지(교통)', '학군', '인프라', '초품아'], [
  ['힐스테이트사하역', '사하구 괴정동\n(괴정로 166)', '1,314', '2022.12', '★★★★★\n신축 대단지·현대 브랜드\n34/40/46평', '★★★★\n1호선 사하역 역세권\n하단역 환승권', '★★★\n괴정권 일반학군', '★★★★\n괴정 상권·전통시장', '확인 필요'],
  ['사하역비스타동원', '사하구 괴정동', '513', '2021.01', '★★★\n중형 단지', '★★★★\n1호선 사하역 역세권\n(입주민 후기 "역세권")', '★★★', '★★★★\n괴정 상권 공유', '확인 필요'],
  ['명지퍼스트월드2단지', '강서구 명지동\n(명지국제7로 37)', '1,406', '2019~2020\n(자료 상이)', '★★★★\n주상복합형 대단지\n(대지지분 작을 가능성)', '★★★\n현재 철도 없음\n하단~녹산선 예정', '★★★★\n명지 학원가', '★★★★★\n국제신도시 상업·녹지', '확인 필요'],
  ['명지퍼스트월드3단지', '강서구 명지동\n(명지국제2로 41)', '1,530', '2020.07', '★★★★\n대단지·준신축', '★★★\n하단~녹산선 예정\n(명지 2.3km 지하화)', '★★★★\n명지 학원가', '★★★★★', '확인 필요'],
  ['대신푸르지오1차', '서구 서대신동1가', '959', '2018.04', '★★★\n준신축', '★★★★\n1호선 서대신역권\n원도심 접근', '★★★★\n대신동 전통 학군지\n(경남고·동아대 소재)', '★★★', '확인 필요'],
  ['대신2차푸르지오', '서구 서대신동2가\n(서대신6구역 재개발)', '815', '2020.09', '★★★★\n신축급', '★★★★\n1호선 서대신역권', '★★★★\n분양 당시 "학군·교통·생활 인프라" 강조', '★★★', '확인 필요'],
  ['대신해모로센트럴', '서구 서대신동2가', '733', '2022.07', '★★★★\n신축(34평 255세대)', '★★★★\n입주민 후기 "교통 좋음"', '★★★★\n대신초 도보 9분(619m)', '★★★', '아님\n(대신초 619m)'],
], [13, 12, 6, 7, 13, 14, 12, 10, 8], { center: [2, 3, 8], boldFirst: true }));
kids.push(note('※ 별점은 공개 정보를 바탕으로 한 애널리스트의 정성 평가임. "초품아(초등학교를 품은 아파트)" 여부는 부산시교육청 통학구역으로 확인해야 함.'));

// 3. 가격
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(h1('3. 매매가 · 전세가 (최신 실거래·시세 기준, 34평)'));
kids.push(table(['단지', '매매가(억)', '매매 근거', '전세가(억)', '전세 근거', '전세가율', '최근 3년 추이'],
  D.map(d => [d.k, f(d.sale) + (d.saleEst ? '\n(추정)' : ''), d.saleBasis, f(d.jeonse) + (d.jEst ? '\n(추정)' : ''), d.jBasis, pct(d.ratio), d.hist]),
  [12, 7, 24, 7, 22, 7, 14], { center: [1, 3, 5], boldFirst: true }));
kids.push(note('단위: 억원 | 기준: 2026년 1~9월 실거래·시세(검색 스니펫) | 출처: 국토부 실거래가(호갱노노·KB부동산·리치고 등 재가공 정보)'));
kids.push(table(['프리미엄(분양가 대비)', '내용'], [
  ['산정 불가 → 확인 필요', '각 단지의 34평 분양가 원자료를 확인하지 못함. 분양가는 청약홈·분양 당시 입주자모집공고로 확인한 뒤 "현재 매매가 − 분양가(+옵션)"로 계산할 것. 힐스테이트사하역은 2023.3 거래가 4.84억(분양권/초기 거래 추정)과 비교하면 약 +1.06억(+22%) 상승.'],
], [20, 80]));

// 4. 투자분석
kids.push(h1('4. 투자금 · 회수율 · 예상시세 · 수익률 (갭투자 시뮬레이션)'));
kids.push(p([r('산식: ', { bold: true }), r('총매수가 = 매매가 + 취득세·지방교육세(1.1%) + 중개보수(0.4%) | 투자금(갭) = 매매가 − 전세가 + 부대비용 | 회수율 = 전세가 ÷ 매매가(매입가 중 전세로 회수되는 비율) | 투자금 대비 전세가 = 전세가 ÷ 투자금 (배)')]));
kids.push(table(['단지', '매매가', '전세가', '총매수가', '갭투자금', '회수율\n(전세가율)', '투자금 대비\n전세가(배)', '레버리지 평가'],
  D.map(d => [d.k, f(d.sale), f(d.jeonse), f(d.total), f(d.inv), pct(d.ratio), d.jPerInv.toFixed(2) + '배', d.inv < 1.8 ? '소액 갭 ◎' : d.inv < 2.3 ? '중간 ○' : '갭 큼 △']),
  [16, 8, 8, 9, 9, 10, 10, 10], { center: [1, 2, 3, 4, 5, 6, 7], boldFirst: true, highlight: D.map((d, i) => d.inv < 1.8 ? i : -1) }));
kids.push(note('단위: 억원 | 2주택 취득 시에도 부산(비조정) 6억 전후 취득세는 약 1.1%. 3주택 이상은 8% 중과여서 위 수익률이 성립하지 않음.'));
kids.push(h2('4-1. 2년 후 예상시세 · 수익금 · 수익률 (3개 시나리오)'));
kids.push(table(['단지', '가정 상승률\n(보수/기본/낙관)', '예상시세(기본)', '수익금(기본)', '수익률(기본)', '예상시세(낙관)', '수익률(낙관)', '예상시세(보수)', '수익률(보수)'],
  D.map(d => [d.k, `${pct(d.gBear)} / ${pct(d.g)} / ${pct(d.gBull)}`, f(d.gP), f(d.gProfit), pct(d.gRet), f(d.gBullP), pct(d.gBullRet), f(d.gBearP), pct(d.gBearRet)]),
  [16, 14, 9, 9, 9, 9, 9, 9, 9], { center: [1, 2, 3, 4, 5, 6, 7, 8], boldFirst: true }));
kids.push(note('수익금 = 예상시세 − 매매가 − 취득·중개비용 (양도세·보유세·중개비(매도)·기회비용 차감 전, 세전) | 상승률은 애널리스트 가정(추정)'));
kids.push(p([r('해석: ', { bold: true }), r('갭이 작은 서대신동·사하역권은 기본 시나리오 수익률이 12~13%로 명지(3%대)의 약 3.5배이고, 낙관 시나리오에서도 약 1.8배임(가정 상승률 차이와 갭 차이가 함께 반영된 결과). 반면 명지는 하단~녹산선 착공이 확정되면 낙관 시나리오 탄력이 가장 크지만, 갭이 크고 공급 부담이 있어 하방 손실폭도 가장 큼.')]));

// 5. 대지지분·거래량
kids.push(h1('5. 대지지분 · 거래량 · 회전율 · 가격인상 · 저평가'));
kids.push(table(['단지', '대지지분', '거래량·회전율(대용지표)', '가격인상 동력', '저평가 판단'], [
  ['힐스테이트사하역', '확인 필요(등기부)', '매물 매매 605·전세 53·월세 19건(조회 시점 미상) → 매물이 많아 매수자 우위 가능', '신축 대장·괴정 재개발(괴정11구역 2,432세대 계획) 주거환경 개선', '사하구 대장 대비 적정~약간 저평가(추정)'],
  ['사하역비스타동원', '확인 필요', '확인 필요', '힐스테이트사하역의 가격을 따라가는 효과', '대장 대비 0.1억 내외 차이 → 저평가 폭 작음'],
  ['명지퍼스트월드2단지', '확인 필요(주상복합 → 작을 가능성)', '2026.8~9월 실거래 2건 확인(거래는 살아 있음)', '하단~녹산선(2026 착공 목표·2029.12 준공 목표)', '호재 대비 저평가이나 공급 리스크 반영'],
  ['명지퍼스트월드3단지', '확인 필요', '확인 필요', '하단~녹산선, 명지 구간 지하화', '상동'],
  ['대신푸르지오1차', '확인 필요', '확인 필요', '서대신동 재개발 9개 구역 → 신축 타운화', '연식 대비 가격 높음(평당 2,400만) → 저평가 아님'],
  ['대신2차푸르지오', '확인 필요', '매매 매물 360건 vs 전세 8건(2026.5) → 전세 품귀', '전세 품귀 → 전세가 상승 → 매매가 지지', '1차보다 신축인데 가격이 비슷 → 상대 저평가(추정)'],
  ['대신해모로센트럴', '확인 필요', '확인 필요', '가장 최근 신축, 대신해모로센트럴2차(서대신동3가) 분양 연계', '서구 신축 중 가격 경쟁력 있음(추정)'],
], [15, 14, 26, 24, 21], { boldFirst: true }));
kids.push(note('※ 회전율 = 연간 거래건수 ÷ 세대수. 국토부 실거래 원자료를 확보하지 못해 매물 수를 대용지표로 사용함.'));

// 6. 개별 분석
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(h1('6. 단지별 개별 분석 (특징 · 호재 · 리스크)'));
const indiv = Object.fromEntries(DATA.complexes.map(c => [c.name, Object.entries(c.points)]));
ranked.forEach((d, i) => {
  kids.push(h2(`${i + 1}. ${d.k}  (종합 ${d.tot}점)`));
  kids.push(table(['구분', '내용'], [
    ['가격 요약', `매매 ${f(d.sale)}억${d.saleEst ? '(추정)' : ''} / 전세 ${f(d.jeonse)}억${d.jEst ? '(추정)' : ''} / 전세가율 ${pct(d.ratio)} / 갭투자금 ${f(d.inv)}억 / 2년 기본 수익률 ${pct(d.gRet)}`],
    ...indiv[d.k],
  ], [12, 88], { boldFirst: true }));
});

// 7. 순위
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(h1('7. 종합 순위 (향후 가격 인상 기대 기준) 1~4위'));
kids.push(p([r('평가 가중치(선공개): ', { bold: true }), r('가격상승 기대 30 · 입지·교통 20 · 갭·전세가율 15 · 상품성(연식·세대수) 15 · 학군·초품아 10 · 공급 리스크(역점수) 10 = 100점')]));
kids.push(table(['순위', '단지', '상승기대\n(30)', '입지·교통\n(20)', '갭·전세가율\n(15)', '상품성\n(15)', '학군\n(10)', '공급리스크\n(10)', '합계', '한줄 평'],
  ranked.map((d, i) => [String(i + 1), d.k, ...d.sc.map(String), String(d.tot), {
    '힐스테이트사하역': '사하 대장·역세권·소액 갭. 가격 인상 기대 1순위',
    '대신2차푸르지오': '전세 품귀 + 재개발 타운 중심. 갭 축소가 기대됨',
    '대신해모로센트럴': '서구 최신축·최소 갭. 초품아 아님',
    '대신푸르지오1차': '학군·입지 양호하나 연식 대비 가격 부담',
    '사하역비스타동원': '대장 대체재. 환금성 약점',
    '더샵명지퍼스트월드3단지': '호재 최대, 공급·갭 부담 → 타이밍 매수',
    '더샵명지퍼스트월드2단지': '3단지와 동일 논리, 주상복합 할인',
  }[d.k]]),
  [5, 17, 7, 7, 8, 7, 6, 8, 6, 29], { center: [0, 2, 3, 4, 5, 6, 7, 8], highlight: [0, 1, 2, 3] }));
kids.push(note('※ 점수는 애널리스트 정성 평가(추정). 1~4위는 노란색으로 표시.'));

// 8. 매수전략
kids.push(h1('8. 매수전략 1안 · 2안 · 3안'));
kids.push(table(['구분', '1안: 소액 갭투자 (수익률 극대화)', '2안: 실거주 후 비과세 (안정형)', '3안: 호재 선점 타이밍 매수 (공격형)'], [
  ['대상', `${ranked[0].k} (대안: 대신해모로센트럴)`, '대신2차푸르지오 또는 대신해모로센트럴', '더샵명지퍼스트월드3단지 (대안: 2단지)'],
  ['투자금(34평)', `약 ${f(ranked[0].inv)}억 (전세 4.3억 승계 기준)`, '자기자본 + 주담대 (DSR 한도 내)', `약 ${f(M.inv)}억 (갭 투자 시)`],
  ['매수 조건', '전세가율 72% 이상 매물 / 전세 만기 1년 이상 남은 매물 / 저층·비선호동 급매 우선', '매도 매물 360건 수준인 구간에서 "호가 − 3% 이상" 협상 / 로열동·중층', '하단~녹산선 착공(2026 목표) 공식 확인 후 진입 / 강서구 주간 변동률이 보합(0.00%)으로 전환되는 것을 확인'],
  ['실행 순서', '① 국토부 실거래로 최근 3개월 동일평형 최저가 확인 ② 전세계약서·확정일자 확인 ③ 잔금일에 전세보증금 승계 ④ 2년 뒤 전세 갱신 시 증액분 회수', '① 주담대 사전심사 ② 매수 후 전입·2년 거주 ③ 1주택 비과세(12억 이하) 요건 충족 ④ 개편안 통과 시 "거주 기간" 장기보유특별공제로 절세', '① 착공 뉴스 모니터링 ② 미분양·급매 탐색 ③ 분할매수(1채 → 추가 검토) ④ 개통(2029) 전후 매도 검토'],
  ['기대수익(2년, 기본)', `수익률 ${pct(ranked[0].gRet)} (세전)`, '시세차익 + 양도세 비과세 효과', `기본 ${pct(M.gRet)} / 낙관 ${pct(M.gBullRet)}`],
  ['핵심 리스크', '금리 추가 인상 시 전세 수요 감소 → 역전세', '대출이자 부담(기준금리 3%)', '착공 지연, 에코델타시티 입주 물량'],
  ['직장인 체크', '2주택: 부산 비조정 → 취득세 1~3%, 양도세 일반세율. 3주택부터 중과', '생애최초·신생아 특례 등 정책대출 활용 여부 확인', '투자금이 큼 → 비상자금 6개월치 확보 후 진입'],
], [10, 30, 30, 30], { boldFirst: true }));

// 9. 체크리스트
kids.push(h1('9. 계약 전 체크리스트'));
[
  '국토부 실거래가(rt.molit.go.kr)에서 동일 단지·34평 최근 3개월 실거래 3건 이상 확인 (이 보고서의 추정치를 대체)',
  '등기부등본: 대지권 비율(대지지분), 근저당, 신탁 여부 확인',
  '전세 승계 매물: 임차인 계약갱신청구권 사용 여부, 만기일, 보증금 반환 계획 확인',
  '부산시교육청 통학구역 조회로 초등학교 배정 확인 (초품아 여부)',
  '하단~녹산선 착공 공고와 괴정·서대신동 재개발 사업단계(조합설립·관리처분) 확인',
  '관리비(특히 주상복합), 경사지형, 일조권은 현장 임장으로 확인',
].forEach(t => kids.push(bullet(t)));

// 10. 출처
kids.push(h1('10. 출처'));
[
  '호갱노노 – 힐스테이트사하역 https://hogangnono.com/apt/dF761 / 사하역비스타동원 https://hogangnono.com/apt/dr2d1',
  '호갱노노 – 더샵명지퍼스트월드2단지 https://hogangnono.com/apt/ctp37 / 3단지 https://hogangnono.com/apt/cMG83',
  '호갱노노 – 대신푸르지오1차 https://hogangnono.com/apt/aO353 / 대신2차푸르지오 https://hogangnono.com/apt/ctv85 / 대신해모로센트럴 https://hogangnono.com/apt/dr12f',
  'KB부동산 – 대신2차푸르지오 https://kbland.kr/se/c/38311 / 명지퍼스트월드3단지 https://kbland.kr/se/c/39467',
  '리치고 – 사하역비스타동원 https://m.richgo.ai/realty/danji/a5ZUPGI / 명지3단지 https://m.richgo.ai/realty/danji/a5abXZV',
  '부산 매매·구별 변동률(한국부동산원 9월 2주) – 브릿지경제 https://www.viva100.com/article/20260918500475',
  '부산 입주물량 – https://v.daum.net/v/20251222185923780',
  '하단~녹산선 – 한국일보 https://www.hankookilbo.com/news/article/A2024082711120004670 / 부산일보 https://www.busan.com/view/busan/view.php?code=2024100618055948659',
  '괴정11구역 – 하우징타임즈 https://www.housingtimes.co.kr/news/articleView.html?idxno=1396',
  '서대신6구역(대신2차푸르지오) – 이데일리 https://edaily.co.kr/News/Read?mediaCodeNo=257&newsId=02400966615998784 / 부산일보 https://www.busan.com/view/busan/view.php?code=20170726000142',
].forEach(t => kids.push(bullet(t)));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 19 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 250 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 900, bottom: 900, left: 1000, right: 1000 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [r('부산 서부권 34평 비교 보고서 · ', { size: 16, color: '808080' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children: kids,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('written'); });
