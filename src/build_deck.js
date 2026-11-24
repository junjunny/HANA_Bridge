// HANA Bridge 발표 자료 생성기 (pptxgenjs)
// 실행: NODE_PATH=<node_modules> node build_deck.js <out.pptx> <apply_theme.js 경로>
const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fi = require('react-icons/fi');
const path = require('path');

const OUT = process.argv[2] || 'deck.pptx';
const { applyTheme } = require(process.argv[3]);

const HEX = {
  ink: '1D2629', deep: '00534F', green: '008485', mint: 'E6F4F3', mint2: 'CDEAE8',
  blue: '2F6FB5', amber: 'C77700', red: 'E0364B', sub: '55636A', soft: 'F3F7F7', line: 'DCE3E5', redbg: 'FDECEE',
};
const THEME = {
  name: 'HANA Bridge', headFontFace: '맑은 고딕', bodyFontFace: '맑은 고딕',
  colors: {
    dk1: HEX.ink, lt1: 'FFFFFF', dk2: HEX.deep, lt2: HEX.mint,
    accent1: HEX.green, accent2: HEX.blue, accent3: HEX.amber, accent4: HEX.red, accent5: HEX.sub, accent6: HEX.mint2,
    hlink: HEX.green, folHlink: HEX.deep,
  },
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.title = 'HANA Bridge';
pres.author = '지원자';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

async function icon(name, color = 'FFFFFF') {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fi[name], { color: '#' + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: 'DARK', background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: 0.8, y: 1.6, w: 11.7, h: 0.5, fontSize: 16, bold: true, color: C.accent6 }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.15, w: 11.7, h: 1.9, fontSize: 44, bold: true, color: C.background1, valign: 'top', align: 'left' }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.15, w: 11.7, h: 1.2, fontSize: 20, color: C.accent6, valign: 'top' }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'SECTION', background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: 0.8, y: 2.4, w: 11.7, h: 0.6, fontSize: 20, bold: true, color: C.accent6 }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 3.0, w: 11.7, h: 1.2, fontSize: 40, bold: true, color: C.background1, valign: 'top', align: 'left' }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: 0.6, y: 0.35, w: 12.1, h: 0.35, fontSize: 13, bold: true, color: C.accent1, margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.72, w: 12.1, h: 0.75, fontSize: 30, bold: true, color: C.text1, margin: 0, valign: 'top', align: 'left' }, text: '' } },
    { text: { text: 'HANA Bridge · 낯선 자산을 익숙한 금융과 연결하다', options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: 'right' },
});

const S = (sectionTitle, master = 'CONTENT') => pres.addSlide({ masterName: master, sectionTitle });
const head = (s, kicker, title) => { s.addText(kicker, { placeholder: 'kicker' }); s.addText(title, { placeholder: 'title' }); };
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 15, color: C.text1, valign: 'top' }, o));
const card = (s, x, y, w, h, color = C.background2, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.15, fill: { color }, line: { color, width: 0 }, objectName: name });
const circle = (s, x, y, d, color, name) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 }, objectName: name });

(async () => {
  const IC = {};
  for (const n of ['FiFileText', 'FiLink2', 'FiList', 'FiShield', 'FiSearch', 'FiCheckCircle', 'FiAlertTriangle', 'FiLayers', 'FiTarget'])
    IC[n] = await icon(n);
  const tbl = (s, rows, o) => s.addTable(rows, Object.assign({ x: 0.6, y: 1.8, w: 12.1, fontSize: 13, fontFace: '+mn-lt', color: HEX.ink, border: { type: 'solid', pt: 0.75, color: HEX.line }, valign: 'middle', margin: 0.08 }, o));
  const H = (t) => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: HEX.deep } } });
  const B = (t, color) => ({ text: t, options: { bold: true, color: color || HEX.ink } });
  const VD = { 채택: HEX.deep, '참고 한정': '8A5A00', 기각: 'C0263B', '분석 불가': HEX.sub };

  // 1 Title
  pres.addSection({ title: '도입' });
  let s = S('도입', 'DARK');
  s.addText('하나증권 청년 금융아이디어 공모전', { placeholder: 'kicker' });
  s.addText('HANA Bridge\n낯선 자산을 익숙한 금융과 연결하다', { placeholder: 'title' });
  s.addText('토큰증권 가치 이해 및 금융시장 연결 서비스  |  지원자명', { placeholder: 'body' });
  s.addNotes('안녕하십니까. 저는 실제 투자계약에서 손익 결정 변수를 찾아 검증된 금융정보와 연결하고, 손님이 무엇을 먼저 확인할지 알려주는 서비스, HANA Bridge를 제안합니다.');

  // 2 Summary
  s = S('도입');
  head(s, '한 장 요약', '정보는 있습니다. 무엇을 봐야 하는지 모를 뿐입니다');
  T(s, '하나증권이 계좌관리기관으로 참여한 한돈 투자계약증권 제1호를 공개 자료만으로 직접 분석했습니다. 손익 조건은 투자설명서 곳곳에 흩어져 있었고, 보도와 공시의 수치는 서로 달랐습니다. HANA Bridge는 계약 원문에서 출발해 무엇을 먼저 확인할지 알려줍니다.',
    { x: 0.6, y: 1.75, w: 12.1, h: 1.3, fontSize: 17, lineSpacingMultiple: 1.2 });
  [['43만 자', '1호 투자설명서 본문 분량', 'DART 2026.1.29, 공백 포함 427,026자'], ['6개', '손익 요인을 찾는 데 쓴 출처', '공시 2종 · KAPE · 핀돈 · 보도 2건'], ['4건', '서로 어긋난 수치', '청약률 · 수익률 · 원가 비중 · 청약 시기']].forEach(([v, l, n], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 3.4, 3.85, 2.9, C.background2, 'stat' + i);
    T(s, v, { x: x + 0.35, y: 3.7, w: 3.2, h: 1.0, fontSize: 44, bold: true, color: C.accent1, fit: 'shrink' });
    T(s, l, { x: x + 0.35, y: 4.8, w: 3.2, h: 0.6, fontSize: 16, bold: true });
    T(s, n, { x: x + 0.35, y: 5.5, w: 3.2, h: 0.6, fontSize: 12, color: C.accent5 });
  });
  T(s, '제안자 1인 탐색 관찰(2026.11.24), 손님 실험 아님', { x: 0.6, y: 6.45, w: 12.1, h: 0.3, fontSize: 11, color: C.accent5 });
  s.addNotes('결론부터 말씀드리겠습니다. 실제 상품 하나를 공개 자료로 분석해 보니, 정보는 충분히 있었지만 43만 자 투자설명서와 6개 출처에 흩어져 있었고, 보도와 공시 사이에 어긋난 수치가 4건 있었습니다. 손님에게 필요한 것은 더 많은 정보가 아니라 무엇을 어떤 기준으로 봐야 하는지입니다.');

  // 3 Section
  pres.addSection({ title: '문제와 기회' });
  s = S('문제와 기회', 'SECTION');
  s.addText('01  제안 배경', { placeholder: 'kicker' });
  s.addText('실제 상품 하나로 문제를 확인했습니다', { placeholder: 'title' });
  s.addNotes('먼저 실제 상품으로 확인한 문제입니다.');

  // 4 Real product terms
  s = S('문제와 기회');
  head(s, '한돈 가축투자계약증권 제1호 (발행인 데이터젠 · 계좌관리기관 하나증권)', '내 손익을 정하는 조건 네 가지, 모두 투자설명서 안에 있습니다');
  [['매각가 결정', '출하일 전후 3일 돼지 경락가격(KAPE, 탕박·제주 제외) 평균으로 식육포장처리업체에 직매', '본문 Ⅱ 적용 경락가격'], ['비용', '사료비는 실비 정산 (초과분 차감·미사용분 반환), 사육관리비 두당 58,000원', '본문 Ⅱ 사료비·사육관리비'], ['수수료', '사업손익이 모집금액의 104%를 넘으면 초과분의 30%가 청산수수료. 손해는 증권 가액 한도', '본문 Ⅱ 청산수수료'], ['위험·제약', '폐사·질병, 운영자 도산 위험. 제3자 양도 불가, 매각 전 현금화 어려움', '유의사항 3 · 본문 Ⅱ 양도제한']].forEach(([t, d, src], i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 1.75 + Math.floor(i / 2) * 2.3;
    card(s, x, y, 5.95, 2.1, C.background2, 'term' + i);
    T(s, t, { x: x + 0.35, y: y + 0.25, w: 5.3, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    T(s, d, { x: x + 0.35, y: y + 0.75, w: 5.3, h: 0.85, fontSize: 14 });
    T(s, '원문: ' + src, { x: x + 0.35, y: y + 1.65, w: 5.3, h: 0.3, fontSize: 11, color: C.accent5 });
  });
  T(s, '출처: DART 데이터젠 투자설명서(가축투자계약증권 제1호, 2026.1.29)', { x: 0.6, y: 6.45, w: 12.1, h: 0.3, fontSize: 11, color: C.accent5 });
  s.addNotes('한돈 1호의 손익 조건입니다. 매각가는 출하일 전후 3일 경락가격 평균, 사료비는 실비 정산, 104%를 넘는 이익의 30%는 운영자 수수료, 그리고 증권은 양도할 수 없습니다. 모두 투자설명서에 있지만 43만 자 문서 곳곳에 흩어져 있습니다.');

  // 5 Discrepancies
  s = S('문제와 기회');
  head(s, '탐색 관찰 결과', '같은 상품을 두고 공시와 보도의 숫자가 달랐습니다');
  tbl(s, [
    [H('항목'), H('A'), H('B'), H('원인')],
    [B('1호 청약률'), '보도 282%', '일반투자자 경쟁률 464%', '보도는 운영자 선배정 50% 포함 (25,098+5,406)÷10,812'],
    [B('1호 수익률'), '18.4% (스마트투데이)', '13.1% (파이낸셜뉴스)', '산정 기준 미표기. 매각금액 기준 vs 정산식 다른 단계'],
    [B('원가·제비용 비중'), '요약정보 95.4% / 4.6%', '본문 99.4% / 0.6%', '같은 투자설명서 안에서 요약과 본문이 다름'],
    [B('1호 청약 시기'), '"1월 청약" 보도', '"2월 공모" 보도', '실제 청약 2026.1.29~2.11 (증권발행실적보고서)'],
  ], { colW: [2.2, 2.6, 2.9, 4.4], rowH: 0.75, fontSize: 14 });
  card(s, 0.6, 5.75, 12.1, 0.85, C.background2, 'obsMsg');
  T(s, '손님이 기준이 다른 숫자를 그대로 비교하면 상품을 잘못 이해할 수 있습니다.', { x: 0.95, y: 5.75, w: 11.4, h: 0.85, fontSize: 17, bold: true, color: C.text2, valign: 'middle' });
  s.addNotes('탐색하면서 어긋난 수치 4건을 찾았습니다. 보도된 청약률 282%는 운영자가 먼저 가져간 50%를 포함한 값이었고, 일반투자자 기준 경쟁률은 464%였습니다. 수익률도 18.4%와 13.1%로 기준 없이 달리 보도됐습니다. 정보가 없는 게 아니라 기준을 알기 어렵다는 것이 문제입니다.');

  // 6 Why now / why Hana
  s = S('문제와 기회');
  head(s, '기회 · 왜 하나증권인가', '거래 인프라가 준비되는 지금, 이해 인프라가 필요합니다');
  s.addShape(pres.shapes.LINE, { x: 1.2, y: 2.85, w: 10.9, h: 0, line: { color: HEX.mint2, width: 4 }, objectName: 'axis' });
  [['2026.01', '한돈 1호 공모\n하나증권 계좌관리기관'], ['2026.05', '하나은행 두나무\n지분 6.55% 취득 결정'], ['2026.09', '하나증권V 출시\nAI 브리핑'], ['2027.02', '개정 전자증권법 시행\n투자계약증권 증권사 유통']].forEach(([d, l], i) => {
    const cx = 1.2 + i * (10.9 / 3), law = i === 3, lx = Math.max(0.6, Math.min(cx - 1.45, 12.73 - 2.9));
    circle(s, cx - 0.2, 2.65, 0.4, law ? C.accent4 : C.accent1, 'dot' + i);
    T(s, d, { x: lx, y: 2.0, w: 2.9, h: 0.45, fontSize: 18, bold: true, color: law ? C.accent4 : C.text2, align: 'center' });
    T(s, l, { x: lx, y: 3.25, w: 2.9, h: 0.9, fontSize: 14, align: 'center' });
  });
  card(s, 0.6, 4.6, 12.1, 1.9, C.background2, 'why');
  T(s, '토큰화는 발행·유통 방식의 변화일 뿐, 기초자산의 손익 변수는 이미 익숙한 시장에 있습니다. 하나증권은 한돈 상품의 청약금·매각대금 관리와 정산 이체를 맡아 상품 구조를 가까이서 다뤘고, 하나증권V와 리서치 조직을 갖추고 있습니다.',
    { x: 0.95, y: 4.75, w: 11.4, h: 1.6, fontSize: 16, lineSpacingMultiple: 1.2, valign: 'middle' });
  T(s, '출처: DART 투자설명서·하나금융지주 공시(2026.5.15), SBS Biz(2026.9.7), 법률신문', { x: 0.6, y: 6.6, w: 12.1, h: 0.3, fontSize: 11, color: C.accent5 });
  s.addNotes('하나증권은 이미 이 상품의 계좌관리기관으로 돈의 흐름을 다뤘고, 하나증권V와 리서치를 갖췄습니다. 2027년 2월 개정법 시행으로 투자계약증권도 증권사를 통해 유통됩니다. 거래 인프라와 함께 이해 인프라를 준비할 시점입니다.');

  // 7 Section
  pres.addSection({ title: '솔루션' });
  s = S('솔루션', 'SECTION');
  s.addText('02  HANA Bridge', { placeholder: 'kicker' });
  s.addText('추천이 아니라, 무엇을 봐야 하는지 알려줍니다', { placeholder: 'title' });
  s.addNotes('이제 해결책입니다.');

  // 8 Core idea vs AI
  s = S('솔루션');
  head(s, '핵심 아이디어', '계약 원문에서 출발하는 정보 매칭');
  card(s, 0.6, 1.75, 12.1, 1.3, C.text2, 'quote');
  T(s, '"실제 투자계약에서 손익 결정 변수를 찾아 검증된 금융정보와 연결하고, 손님이 무엇을 먼저 확인할지 알려주는 서비스"', { x: 1.0, y: 1.85, w: 11.3, h: 1.1, fontSize: 19, bold: true, color: C.background1, valign: 'middle', align: 'center' });
  tbl(s, [
    [H('구분'), H('일반 AI 요약 (ChatGPT 등에 질문)'), { text: 'HANA Bridge', options: { bold: true, color: 'FFFFFF', fill: { color: HEX.green } } }],
    [B('근거'), '답변마다 출처가 달라지고 원문 위치가 남지 않음', '투자설명서 원문 위치를 요인마다 고정 표시'],
    [B('연결 기준'), '질문할 때마다 다른 지표를 연결할 수 있음', '담당자가 승인한 연결 기준을 모든 상품에 같은 방식으로 적용, 기각 사유도 공개'],
    [B('갱신·책임'), '계약 변경·공시를 스스로 추적하지 않음', '공시·지표 갱신, 오류 정정, 재승인 기록, 하나증권 리서치 연결'],
  ], { y: 3.35, colW: [1.9, 4.6, 5.6], rowH: 0.75, fontSize: 14 });
  s.addNotes('ChatGPT에게 물어보면 되지 않느냐는 질문에 대한 답입니다. HANA Bridge는 원문 위치를 근거로 남기고, 담당자가 승인한 같은 기준을 반복 적용하며, 공시가 바뀌면 갱신하고 오류를 정정합니다.');

  // 9 Flow
  s = S('솔루션');
  head(s, '아이디어 구성', '네 단계로 이해를 쌓습니다');
  [['계약 원문', '투자설명서·투자계약서·공시', IC.FiFileText], ['손익 요인 추출', '권리·정산식·비용·가격 결정·위험·양도 가능성', IC.FiSearch], ['정보 매칭·승인', 'A 직접 · B 외부 지표 · C 비교 상품, D 통계 상태', IC.FiLink2], ['확인 우선순위', '개별 공시 → 직접 지표 → 비교 정보', IC.FiList]].forEach(([t, d, ic], i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.8, 2.88, 2.8, C.background2, 'step' + i);
    circle(s, x + 0.3, 2.05, 0.75, C.accent1, 'stepIc' + i);
    s.addImage({ data: ic, x: x + 0.47, y: 2.22, w: 0.41, h: 0.41, altText: '' });
    T(s, `${i + 1}. ${t}`, { x: x + 0.3, y: 3.0, w: 2.4, h: 0.45, fontSize: 17, bold: true });
    T(s, d, { x: x + 0.3, y: 3.5, w: 2.4, h: 0.95, fontSize: 13, color: C.accent5 });
  });
  card(s, 0.6, 4.9, 7.9, 1.6, C.background2, 'verdict');
  T(s, '판정 4가지', { x: 0.95, y: 5.05, w: 7.2, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  T(s, '채택 · 참고 한정 · 기각 · 보류. 점수나 가중치를 임의로 주지 않고, 원문으로 설명되는 경로만 채택합니다.', { x: 0.95, y: 5.5, w: 7.3, h: 0.9, fontSize: 14 });
  card(s, 8.75, 4.9, 3.95, 1.6, 'EEF1F2', 'ext');
  T(s, '확장 기능 (핵심 검증 후)', { x: 9.05, y: 5.05, w: 3.4, h: 0.4, fontSize: 15, bold: true, color: C.accent5 });
  T(s, '공통 위험 진단, 통계적 관련성, 외부 자산 연동', { x: 9.05, y: 5.5, w: 3.4, h: 0.9, fontSize: 13, color: C.accent5 });
  s.addNotes('구성은 네 단계입니다. 계약 원문에서 손익 요인을 추출하고, 정보를 유형별로 나눠 담당자가 승인한 것만 연결하며, 마지막으로 무엇을 먼저 확인할지 순서를 알려줍니다. 공통 위험 진단 같은 기능은 핵심을 검증한 뒤 확장합니다.');

  // 10 Phones
  s = S('솔루션');
  head(s, '화면 시안 · 한돈 1호 공시 내용 적용', '원문 근거 → 확인 순서 → 연결 상세');
  const ph = 4.75, pw = ph * 653 / 1220;
  ['① 손익 요인 카드', '② 확인 우선순위', '③ 연결 상세 (경락가격)'].forEach((c, i) => {
    const x = 0.75 + i * 3.0;
    s.addImage({ path: path.join(__dirname, `img/phone${i + 1}.png`), x, y: 1.65, w: pw, h: ph, altText: c + ' 화면 시안' });
    T(s, c, { x: x - 0.25, y: 6.5, w: pw + 0.5, h: 0.4, fontSize: 13, bold: true, align: 'center' });
  });
  card(s, 9.85, 1.75, 2.85, 4.95, C.background2, 'rules');
  T(s, '표시하지 않는 것', { x: 10.1, y: 1.95, w: 2.4, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  T(s, [
    { text: '근거 없는 영향력 막대', options: { bullet: true, breakLine: true } },
    { text: '위험 점수·상관계수', options: { bullet: true, breakLine: true } },
    { text: '가격 방향 예측', options: { bullet: true, breakLine: true } },
    { text: '매수 버튼', options: { bullet: true } },
  ], { x: 10.1, y: 2.5, w: 2.45, h: 2.4, fontSize: 14, paraSpaceAfter: 10 });
  T(s, '화면 속 수치는 모두 투자설명서 원문 값', { x: 10.1, y: 5.2, w: 2.45, h: 1.2, fontSize: 13, bold: true, color: C.accent1 });
  s.addNotes('실제 공시 내용을 넣은 화면입니다. 첫 화면은 원문 위치가 붙은 손익 요인, 두 번째는 오늘 확인할 순서, 세 번째는 경락가격 노드를 눌렀을 때 보이는 연결 이유, 전달 경로, 한계, 출처입니다.');

  // 11 Matching verdicts
  s = S('솔루션');
  head(s, '실제 상품 적용 결과', '후보 정보 10개 중 4개 채택, 2개 기각');
  const vrow = (name, type, v, why) => [B(name), type, { text: v, options: { bold: true, color: VD[v] } }, why];
  tbl(s, [
    [H('후보 정보'), H('유형'), H('판정'), H('근거 (투자설명서 원문)')],
    vrow('돼지 경락가격 (KAPE)', 'B 외부 지표', '채택', '계약상 매각가 산정 기준. 발행사 추정 민감도 100원당 약 1.8%'),
    vrow('출하·등급판정 공시', 'A 직접 정보', '채택', '출하 다음 날 핀돈 게시 → 확인 우선순위 1번'),
    vrow('사료 곡물가격·환율', 'B 외부 지표', '참고 한정', '사료비 실비 정산으로 전달 가능, 공급가 조정 방식 미공시'),
    vrow('기준금리·국고채', 'B 외부 지표', '기각', '정산식에 금리 항목 없음'),
    vrow('사료·축산 상장사 주가', 'C 비교 상품', '기각', '투자자 손익의 대용치가 될 수 없음'),
    vrow('가격 관련성 통계', 'D 통계', '분석 불가', '양도 금지·비상장으로 거래가격 시계열 없음'),
  ], { colW: [2.9, 1.8, 1.5, 5.9], rowH: 0.62, fontSize: 13.5 });
  T(s, '전체 판정: 채택 4 · 참고 한정 3 · 기각 2 · 분석 불가 1 (보충자료 2)', { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
  s.addNotes('실제 상품에 적용한 결과입니다. 경락가격은 계약상 매각가 기준이라 채택했고, 금리와 사료기업 주가는 정산식에 전달 경로가 없어 기각했습니다. 기각 사례가 있다는 것이 기준이 실제로 작동한다는 증거입니다.');

  // 12 CX
  s = S('솔루션');
  head(s, '손님 경험의 변화', '목표는 구매가 아니라 이해하고 판단할 수 있는 상태');
  T(s, '적용 전', { x: 2.75, y: 1.7, w: 4.8, h: 0.45, fontSize: 17, bold: true, color: C.accent4 });
  T(s, '적용 후', { x: 7.85, y: 1.7, w: 4.8, h: 0.45, fontSize: 17, bold: true, color: C.accent1 });
  [['공모 소식', '"청약률 282%, 수익률 18%" 보도만 보고 관심', '청약률·수익률의 산정 기준을 함께 확인'], ['구조 이해', '43만 자 투자설명서에서 직접 찾아야 함', '손익 요인 카드로 매각가 기준·수수료·양도 불가 이해'], ['투자 판단', '사료기업 주가·금리 뉴스 등 무관한 정보로 판단', '기각 사유를 보고 무관한 지표를 걸러냄. 투자 안 하는 결정도 성공'], ['보유·정산', '무엇이 손익을 바꾸는지 몰라 공시를 놓침', '출하·등급 공시 → 경락가격 순서로 확인']].forEach(([st, b, a], i) => {
    const y = 2.25 + i * 1.12;
    card(s, 0.6, y, 1.95, 0.95, C.text2, 'st' + i);
    T(s, st, { x: 0.6, y, w: 1.95, h: 0.95, fontSize: 15, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    card(s, 2.75, y, 4.9, 0.95, 'FBF3F4', 'b' + i);
    T(s, b, { x: 2.95, y, w: 4.5, h: 0.95, fontSize: 14, valign: 'middle' });
    card(s, 7.85, y, 4.85, 0.95, C.background2, 'a' + i);
    T(s, a, { x: 8.05, y, w: 4.45, h: 0.95, fontSize: 14, valign: 'middle' });
  });
  s.addNotes('손님 경험은 이렇게 바뀝니다. 기준이 다른 숫자를 구분하고, 손익 요인을 원문 근거로 이해하며, 무관한 지표를 걸러냅니다. 투자하지 않기로 결정하는 것도 성공적인 경험입니다.');

  // 13 Section
  pres.addSection({ title: '실행과 효과' });
  s = S('실행과 효과', 'SECTION');
  s.addText('03  차별성 · 실행 · 효과', { placeholder: 'kicker' });
  s.addText('조건이 충족될 때 한 단계씩 넓힙니다', { placeholder: 'title' });
  s.addNotes('마지막으로 차별성과 실행 방안입니다.');

  // 14 Competitors
  s = S('실행과 효과');
  head(s, '기존 서비스 공개 기능 조사 (2026.11.24)', '계약 기반 지표 매칭과 확인 순서는 발견되지 않았습니다');
  const ok = (t) => ({ text: t, options: { color: HEX.deep, bold: true } }), nf = { text: '조사 범위에서 미발견', options: { color: '8A979C' } };
  tbl(s, [
    [H('항목'), H('카사'), H('뮤직카우'), H('핀돈'), H('하나증권V')],
    [B('상품 구조 설명'), ok('확인됨'), ok('확인됨'), ok('확인됨'), nf],
    [B('상품별 외부 지표'), nf, nf, ok('부분 (경락가 정산)'), nf],
    [B('비교 상품과 이유'), nf, nf, nf, ok('다른 목적 (ETF 추천)')],
    [B('한계·오인 안내'), nf, ok('일반 문구'), nf, nf],
    [B('확인 우선순위'), nf, ok('일반 문구'), nf, nf],
  ], { colW: [2.5, 2.4, 2.4, 2.4, 2.4], rowH: 0.6, fontSize: 13.5 });
  T(s, '자료: 앱스토어 소개(카사·뮤직카우), findon.kr, SBS Biz(하나증권V). "미발견"은 공개 자료에서 찾지 못했다는 뜻이며 기능이 없다는 단정이 아닙니다. 펀블은 2026.7 파산 선고로 제외.', { x: 0.6, y: 5.6, w: 12.1, h: 0.7, fontSize: 12, color: C.accent5 });
  s.addNotes('공개 자료로 조사한 범위에서는 계약 기반 지표 매칭과 확인 순서를 제공하는 서비스를 찾지 못했습니다. 다만 미발견이 기능이 없다는 뜻은 아니므로 최초나 유일이라는 표현은 쓰지 않습니다.');

  // 15 Roadmap & ops
  s = S('실행과 효과');
  head(s, '실행 방안', '날짜가 아니라 조건이 충족되면 다음 단계로');
  [['1단계 · PoC', '한돈 1개 상품', '요인 카드·매칭·우선순위 적용, 손님 실험 10~20명', '통과: 오인율↓ + 정답률↑'], ['2단계 · 시범', '하나증권V 탭', '비개인화 MVP, 고객정보 연동 없음', '출시: 상품·데이터 권한·준법·사용자 검증'], ['3단계 · 확장', '유통형 토큰증권', '부동산, 공통 위험 진단, 동의 기반 개인화', '조건: 동의·철회·삭제 정책']].forEach(([n, t, d, g], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.85, 2.75, C.background2, 'rd' + i);
    T(s, n, { x: x + 0.3, y: 1.95, w: 3.3, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
    T(s, t, { x: x + 0.3, y: 2.3, w: 3.3, h: 0.45, fontSize: 18, bold: true });
    T(s, d, { x: x + 0.3, y: 2.8, w: 3.3, h: 0.85, fontSize: 13.5 });
    T(s, g, { x: x + 0.3, y: 3.75, w: 3.3, h: 0.6, fontSize: 13, bold: true, color: C.text2 });
  });
  tbl(s, [
    [H('역할 (가정)'), H('책임')],
    [B('상품 담당'), '계약 원문, 변경 공시, 상품별 예외'], [B('리서치'), '연결 기준·전달 경로·비교 대상 승인'],
    [B('준법·소비자보호'), '투자권유 여부(노출 순서·문구·이해상충), 오인 검증'], [B('디지털·데이터 / 정보보호'), '수집·기준일·갱신 / 접근권한·외부 AI 전송 통제'],
  ], { y: 4.8, colW: [3.0, 9.1], rowH: 0.4, fontSize: 13 });
  s.addNotes('실행은 날짜가 아니라 조건 기반입니다. 먼저 한돈 1개 상품으로 PoC를 하고, 오인율이 줄고 이해도가 오르는지 확인합니다. 상품, 데이터 권한, 준법 승인, 사용자 검증이 모두 갖춰지면 하나증권V에 시범 도입합니다.');

  // 16 Compliance & data
  s = S('실행과 효과');
  head(s, '실현 가능성과 고려사항', '오류는 숨기고, 고치고, 알리고, 기록합니다');
  ['연결 숨김', '원문 재대조 후 정정', '필요 시 손님 통지', '재승인 기록'].forEach((t, i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.8, 2.88, 0.95, 'FDECEE', 'err' + i);
    T(s, `${i + 1}. ${t}`, { x: x + 0.25, y: 1.8, w: 2.5, h: 0.95, fontSize: 15, bold: true, color: C.accent4, valign: 'middle' });
  });
  [['데이터', 'KAPE 경락가격 Open API 이용허락범위 "제한 없음"(공공데이터포털), DART 공시, 한국은행 ECOS. 발행사 공시 재게시는 협약으로 확보'], ['운영', 'AI 추출은 원문 위치를 보존하고 승인 전 게시 금지. 데이터 지연 시 기준일·지연 상태 표시, 오래된 정보로 알림 생성 금지'], ['규제', '매수 버튼 분리·면책 문구만으로 투자권유 여부를 판단하지 않고 준법 검토. 외부 AI로 개인정보 전송 금지'], ['비용', '템플릿 재사용으로 줄어드는 분석 시간과 늘어나는 검토 시간을 PoC에서 측정']].forEach(([k, v], i) => {
    const y = 3.0 + i * 0.92;
    circle(s, 0.6, y + 0.13, 0.5, C.accent1, 'chk' + i);
    s.addImage({ data: IC.FiCheckCircle, x: 0.68, y: y + 0.21, w: 0.34, h: 0.34, altText: '' });
    T(s, k, { x: 1.3, y: y + 0.1, w: 1.4, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: 'middle' });
    T(s, v, { x: 2.7, y: y + 0.05, w: 10.0, h: 0.75, fontSize: 14, valign: 'middle' });
  });
  s.addNotes('잘못된 연결로 손님이 오해하면 어떻게 하느냐는 질문에 대한 답입니다. 오류가 발견되면 연결을 숨기고, 정정하고, 필요하면 손님에게 알리고, 재승인 기록을 남깁니다. 데이터는 이용 조건이 확인된 공개 데이터로 시작합니다.');

  // 17 KPI
  s = S('실행과 효과');
  head(s, '기대효과와 성과 지표', '거래 증가가 아니라 이해와 오인 감소로 평가합니다');
  [['핵심', '위험 이해 정답률', '같은 과제, 조건 A(공개 자료) vs B(HANA Bridge)'], ['핵심', '오인율', '주가·청약률·비교상품 가격을 손익으로 오해한 비율'], ['보조', '과제 완료시간', '이해 없이 빨라지는 것은 성공으로 보지 않음']].forEach(([tag, t, d], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.8, 3.85, 2.2, i < 2 ? C.background1 : C.background2, 'kpi' + i);
    if (i < 2) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.8, w: 3.85, h: 2.2, rectRadius: 0.15, fill: { color: C.background1 }, line: { color: HEX.green, width: 1.5 }, objectName: 'kpiLine' + i });
    T(s, tag + ' 지표', { x: x + 0.35, y: 2.0, w: 3.2, h: 0.35, fontSize: 13, bold: true, color: i < 2 ? C.accent1 : C.accent5 });
    T(s, t, { x: x + 0.35, y: 2.4, w: 3.2, h: 0.55, fontSize: 22, bold: true });
    T(s, d, { x: x + 0.35, y: 3.0, w: 3.2, h: 0.85, fontSize: 13.5, color: C.accent5 });
  });
  card(s, 0.6, 4.3, 5.9, 2.2, C.background2, 'cust');
  T(s, '손님', { x: 0.95, y: 4.45, w: 5.2, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
  T(s, '흩어진 공시·통계·보도를 원문 근거와 함께 한 화면에서. 정산 구조·양도 제한을 투자 전에 이해. 기각 사유와 한계까지 확인', { x: 0.95, y: 4.9, w: 5.3, h: 1.4, fontSize: 14.5 });
  card(s, 6.85, 4.3, 5.85, 2.2, C.background2, 'comp');
  T(s, '하나증권 (가설과 검증 지표)', { x: 7.2, y: 4.45, w: 5.2, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
  T(s, '리서치 활용 확대 → 리서치 연결 이용률\n문의 대응 효율 → 반복 문의 감소\n손님 신뢰 → 재방문율', { x: 7.2, y: 4.9, w: 5.25, h: 1.4, fontSize: 14.5 });
  s.addNotes('성과는 거래 증가가 아니라 이해도와 오인 감소로 평가합니다. 회사 효과도 가설로 두고 리서치 이용률, 반복 문의, 재방문율로 검증하겠습니다.');

  // 18 Closing
  pres.addSection({ title: '마무리' });
  s = S('마무리', 'DARK');
  s.addText('HANA Bridge', { placeholder: 'kicker' });
  s.addText('새 상품을 "거래하는" 플랫폼에서,\n"이해하는" 플랫폼으로', { placeholder: 'title' });
  s.addText('감사합니다', { placeholder: 'body' });
  s.addNotes('HANA Bridge는 실제 계약에서 출발해 손님이 무엇을 봐야 하는지 알려주는 서비스입니다. 하나증권을 새로운 상품을 거래하는 플랫폼에서 이해할 수 있는 플랫폼으로 넓히겠습니다. 감사합니다.');

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('wrote', OUT);
})();
