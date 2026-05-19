import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ───────────────── WORD DATA (백업용 — 앱에서는 미사용) ───────────────── */
const DEFAULT_WORDS = [
  {s:18,w:"department",m:"부서, 부"},{s:18,w:"with regard to A",m:"A와 관련하여"},{s:18,w:"state",m:"주"},{s:18,w:"funding",m:"자금, 예산"},{s:18,w:"construction",m:"건설"},{s:18,w:"functional",m:"기능하는"},{s:18,w:"despite",m:"~에도 불구하고"},{s:18,w:"submit",m:"제출하다"},{s:18,w:"require",m:"요구하다"},{s:18,w:"documentation",m:"서류"},{s:18,w:"notification",m:"통지"},{s:18,w:"considerable",m:"상당한"},{s:18,w:"consequence",m:"결과"},{s:18,w:"budgetary",m:"예산의"},{s:18,w:"constraint",m:"제한, 제약"},{s:18,w:"therefore",m:"그러므로"},{s:18,w:"in order to +V",m:"V하기 위해서"},{s:18,w:"proceed",m:"진행하다"},{s:18,w:"request",m:"요청하다, 요청"},{s:18,w:"notify",m:"통지하다"},{s:18,w:"regarding A",m:"A와 관련하여"},{s:18,w:"look forward to A",m:"A를 고대하다"},
  {s:19,w:"figure",m:"생각하다"},{s:19,w:"select",m:"선택하다"},{s:19,w:"anticipation",m:"기대, 예상"},{s:19,w:"fall apart",m:"산산이 무너지다"},
  {s:20,w:"challenge",m:"도전, 어려움"},{s:20,w:"commitment",m:"전념"},{s:20,w:"obligation",m:"의무"},{s:20,w:"room",m:"여유, 여지"},{s:20,w:"packed",m:"빡빡한"},{s:20,w:"dedicated",m:"전념하는"},{s:20,w:"workout",m:"운동"},{s:20,w:"what if",m:"~라면 어떨까"},{s:20,w:"in the midst of",m:"~의 한가운데에서"},{s:20,w:"routine",m:"일상생활, 일상"},{s:20,w:"integrate",m:"통합하다"},{s:20,w:"chore",m:"집안일, 허드렛일"},{s:20,w:"inevitable",m:"불가피한"},{s:20,w:"trash",m:"쓰레기"},{s:20,w:"essential",m:"필수적인"},{s:20,w:"rather than A",m:"A라기 보다는"},{s:20,w:"purely",m:"순전히, 순수하게"},{s:20,w:"obligatory",m:"의무적인"},{s:20,w:"seize",m:"붙잡다, 움켜쥐다"},{s:20,w:"engage in A",m:"~을 하다, ~에 참여하다"},{s:20,w:"kettle",m:"주전자"},{s:20,w:"incorporate",m:"포함시키다"},
  {s:21,w:"concept",m:"개념"},{s:21,w:"recode",m:"재부호화하다"},{s:21,w:"reconstruct",m:"재구성하다"},{s:21,w:"rather than",m:"~라기 보다는"},{s:21,w:"retrieve",m:"상기하다"},{s:21,w:"trait",m:"특성, 특질"},{s:21,w:"efficient",m:"효율적인"},{s:21,w:"store",m:"저장하다"},{s:21,w:"a bit",m:"약간"},{s:21,w:"optimal",m:"최적의"},{s:21,w:"compression",m:"압축"},{s:21,w:"such as A",m:"A와 같은"},{s:21,w:"lack",m:"부족하다, 부족"},{s:21,w:"detail",m:"세부사항"},{s:21,w:"struggle",m:"고군분투하다"},{s:21,w:"generalize",m:"일반화하다"},{s:21,w:"represent",m:"나타내다"},{s:21,w:"abstract",m:"추상적인"},{s:21,w:"feature",m:"특징"},{s:21,w:"at a cost of A",m:"A을 희생해서"},{s:21,w:"instead",m:"그 대신에"},{s:21,w:"assumption",m:"가정"},{s:21,w:"motivation",m:"동기"},{s:21,w:"discovery",m:"발견"},{s:21,w:"representation",m:"재현, 표현"},{s:21,w:"completely",m:"완전히"},{s:21,w:"undermine",m:"손상시키다, 훼손하다"},{s:21,w:"legal",m:"법적인"},{s:21,w:"primacy",m:"우위성"},{s:21,w:"eyewitness",m:"목격자"},{s:21,w:"testimony",m:"증언, 증언하다"},
  {s:22,w:"laboratory",m:"실험실"},{s:22,w:"conduct",m:"수행하다"},{s:22,w:"evidence",m:"증거"},{s:22,w:"deliberation",m:"숙고"},{s:22,w:"facial",m:"얼굴의"},{s:22,w:"detailed",m:"자세한, 상세한"},{s:22,w:"discussion",m:"논의, 토의"},{s:22,w:"selection",m:"선택"},{s:22,w:"procedure",m:"절차"},{s:22,w:"inaccurate",m:"부정확한"},{s:22,w:"identification",m:"식별"},{s:22,w:"recognize",m:"인식하다, 알아보다"},{s:22,w:"pop out at A",m:"A에게 탁 떠오르다"},{s:22,w:"tend to-V",m:"V하는 경향이 있다"},{s:22,w:"accurate",m:"정확한"},{s:22,w:"immediate",m:"즉각적인"},{s:22,w:"reaction",m:"반응"},{s:22,w:"interpretation",m:"해석"},{s:22,w:"impression",m:"인상"},{s:22,w:"deliberative",m:"신중한"},{s:22,w:"reason",m:"추론하다"},{s:22,w:"narrative",m:"서사, 이야기"},{s:22,w:"describe",m:"묘사하다"},{s:22,w:"a variety of",m:"다양한"},{s:22,w:"behavioral",m:"행동의"},{s:22,w:"demonstrate",m:"보여주다, 입증하다"},{s:22,w:"relatively",m:"상대적으로, 비교적"},{s:22,w:"unconscious",m:"무의식적인"},{s:22,w:"guess",m:"추측하다"},{s:22,w:"compare",m:"비교하다"},{s:22,w:"logical",m:"논리적인"},{s:22,w:"justification",m:"정당화"},
  {s:23,w:"quantitative",m:"양적인"},{s:23,w:"measure",m:"측정하다"},{s:23,w:"the number of A",m:"~의 수"},{s:23,w:"interaction",m:"상호 작용"},{s:23,w:"individual",m:"개인"},{s:23,w:"recall",m:"회상하다"},{s:23,w:"yield",m:"산출하다"},{s:23,w:"reliable",m:"신뢰할 만한"},{s:23,w:"respondent",m:"응답자"},{s:23,w:"recollect",m:"회상하다"},{s:23,w:"definition",m:"정의"},{s:23,w:"vary",m:"다양하다, 다르다"},{s:23,w:"call in sick",m:"병가를 내다"},{s:23,w:"prompt",m:"유발하다"},{s:23,w:"concrete",m:"구체적인"},{s:23,w:"similarly",m:"비슷하게, 유사하게"},{s:23,w:"instead of",m:"~대신에"},{s:23,w:"rate",m:"평가하다"},{s:23,w:"procrastinator",m:"미루는 사람"},{s:23,w:"utility bill",m:"공과금 고지서"},{s:23,w:"bill",m:"청구서, 법안, 부리, 지폐"},{s:23,w:"currently",m:"현재에"},{s:23,w:"even though",m:"~에도 불구하고"},{s:23,w:"afford to+V",m:"V할 여유가 되다"},{s:23,w:"seek",m:"찾다, 추구하다"},{s:23,w:"response",m:"응답, 반응"},{s:23,w:"ensure",m:"보장하다"},{s:23,w:"consistency",m:"일관성"},
  {s:24,w:"evolution",m:"진화"},{s:24,w:"singularity",m:"특이점"},{s:24,w:"refer to",m:"~을 가리키다"},{s:24,w:"exceed",m:"~을 넘어서다, 초과하다"},{s:24,w:"intelligence",m:"지능"},{s:24,w:"predict",m:"예측하다"},{s:24,w:"improve",m:"개선하다"},{s:24,w:"evolve",m:"진화하다"},{s:24,w:"accelerate",m:"가속화하다"},{s:24,w:"pace",m:"속도"},{s:24,w:"aware",m:"인식하는, 아는"},{s:24,w:"pursue",m:"추구하다"},{s:24,w:"conscious",m:"의식적인"},{s:24,w:"a being",m:"존재"},{s:24,w:"consciousness",m:"의식"},{s:24,w:"dimension",m:"차원"},{s:24,w:"provide A with B",m:"A에게 B를 제공하다"},{s:24,w:"intellectual",m:"지적인, 영리한"},{s:24,w:"stimulation",m:"자극"},{s:24,w:"inspire",m:"불어넣다, 영감을 주다"},{s:24,w:"insight",m:"통찰력"},{s:24,w:"conversely",m:"반대로"},{s:24,w:"significant",m:"중대한"},{s:24,w:"impact",m:"영향"},{s:24,w:"direction",m:"방향"},{s:24,w:"depend on A",m:"A에 달려있다, 의존하다"},{s:24,w:"ethics",m:"윤리"},{s:24,w:"incorporate",m:"통합하다"},{s:24,w:"mutual",m:"서로의, 상호의"},{s:24,w:"coexistence",m:"공존"},{s:24,w:"right",m:"권리"},
  {s:25,w:"generation",m:"발전, 생성, 세대"},{s:25,w:"fossil fuel",m:"화석 연료"},{s:25,w:"nuclear energy",m:"핵에너지"},{s:25,w:"renewables",m:"재생 가능 에너지"},{s:25,w:"in terms of",m:"~라는 점에서"},{s:25,w:"a third",m:"3분의 1"},{s:25,w:"generate",m:"발전하다, 생성하다"},{s:25,w:"combine",m:"결합하다"},
  {s:26,w:"highly",m:"매우"},{s:26,w:"portrait",m:"인물 사진, 초상화"},{s:26,w:"celebrity",m:"유명 인사"},{s:26,w:"eagerly",m:"간절히, 열렬히"},{s:26,w:"await",m:"~을 기다리다"},{s:26,w:"weekly",m:"매주의, 주간의"},{s:26,w:"discuss",m:"토의하다, 논의하다"},{s:26,w:"contain",m:"담고있다, 포함하다"},{s:26,w:"prospect",m:"전망"},{s:26,w:"graduate",m:"졸업하다"},{s:26,w:"hire",m:"고용하다"},{s:26,w:"iconic",m:"상징적인"},{s:26,w:"instantly",m:"즉시"},{s:26,w:"spend time ~ing",m:"~하면서 시간을 보내다"},{s:26,w:"shoot",m:"촬영하다"},{s:26,w:"across",m:"~전체에"},{s:26,w:"exotic",m:"이국적인"},{s:26,w:"location",m:"위치, 장소"},{s:26,w:"up to",m:"(최대)~까지"},{s:26,w:"dozen",m:"12"},
  {s:29,w:"essentially",m:"본질적으로"},{s:29,w:"metaphor",m:"은유"},{s:29,w:"be different from",m:"~와 다르다"},{s:29,w:"linguistic",m:"언어적인"},{s:29,w:"passive",m:"수동적인"},{s:29,w:"in the sense that~",m:"~라는 점에서"},{s:29,w:"audience",m:"독자, 관객"},{s:29,w:"propose",m:"제시하다, 제안하다"},{s:29,w:"beggar",m:"구걸하는 사람, 거지"},{s:29,w:"be likely to-V",m:"~하기 쉽다"},{s:29,w:"cognitive",m:"인지적인"},{s:29,w:"effort",m:"노력"},{s:29,w:"further",m:"더"},{s:29,w:"engage",m:"끌어들이다"},{s:29,w:"prose",m:"산문"},{s:29,w:"on the other hand",m:"반면에"},{s:29,w:"impose",m:"강요하다, 부과하다"},{s:29,w:"realize",m:"실현하다, 깨닫다"},{s:29,w:"artifact",m:"인공물"},{s:29,w:"forcefully",m:"강력하게"},{s:29,w:"horizon",m:"지평"},{s:29,w:"afford to-V",m:"~할 여유가 되다"},{s:29,w:"potential",m:"잠재적인"},{s:29,w:"wonder",m:"궁금해하다"},{s:29,w:"normally",m:"일반적으로, 보통"},{s:29,w:"selling point",m:"매력, 장점"},{s:29,w:"usefulness",m:"유용성"},{s:29,w:"obvious",m:"분명한, 명백한"},{s:29,w:"at first glance",m:"첫 눈에"},{s:29,w:"beloved",m:"사랑받는"},{s:29,w:"in part",m:"부분적으로"},{s:29,w:"immediately",m:"즉각적으로"},
  {s:30,w:"limitation",m:"한계"},{s:30,w:"capacity",m:"능력"},{s:30,w:"resources",m:"자원"},{s:30,w:"compute",m:"계산하다"},{s:30,w:"independently",m:"독립적으로"},{s:30,w:"distribute",m:"분배하다"},{s:30,w:"computation",m:"계산"},{s:30,w:"flush",m:"(변기)물을 내리다"},{s:30,w:"interact",m:"상호 작용하다"},{s:30,w:"indeed",m:"실제로, 사실상"},{s:30,w:"dedicate A to B",m:"A를 B에 바치다"},{s:30,w:"calculation",m:"계산"},{s:30,w:"benefit from",m:"~에서 이득을 얻다"},{s:30,w:"contribute to A",m:"A에 기여하다"},{s:30,w:"whole",m:"전체"},{s:30,w:"do one's part",m:"~의 역할을 하다"},{s:30,w:"collective",m:"집합적인"},
  {s:31,w:"defence",m:"방어"},{s:31,w:"shallow",m:"얕은"},{s:31,w:"flounder",m:"넙치"},{s:31,w:"fold",m:"접다"},{s:31,w:"stare",m:"보다, 응시하다"},{s:31,w:"illusion",m:"착시"},{s:31,w:"excellent",m:"뛰어난, 우수한"},{s:31,w:"asset",m:"자산"},{s:31,w:"turn A into B",m:"A를 B로 바꾸다"},{s:31,w:"creature",m:"생물"},{s:31,w:"transform",m:"바꾸다, 변형시키다"},{s:31,w:"mimic",m:"모방하다"},{s:31,w:"spectacle",m:"광경, 장관의"},{s:31,w:"stick",m:"찌르다, 내밀다"},{s:31,w:"grab",m:"붙잡다"},
  {s:32,w:"suffer",m:"고통받다"},{s:32,w:"relate to",m:"~와 관련되다"},{s:32,w:"frame",m:"구성하다"},{s:32,w:"extreme",m:"극심한"},{s:32,w:"lung",m:"폐"},{s:32,w:"explode",m:"폭발하다"},{s:32,w:"deficit",m:"부족, 결핍"},{s:32,w:"normal",m:"정상적인"},{s:32,w:"boundary",m:"경계"},{s:32,w:"endurance",m:"인내력"},{s:32,w:"distance",m:"거리"},{s:32,w:"patient",m:"환자"},{s:32,w:"back pain",m:"등 통증, 요통"},{s:32,w:"therapy",m:"치료"},{s:32,w:"practitioner",m:"의사, 치료사"},{s:32,w:"mobilise",m:"풀어주다"},{s:32,w:"pressure",m:"압박"},{s:32,w:"sore",m:"아픈, 쓰린"},{s:32,w:"tissue",m:"조직"},
  {s:33,w:"electronics",m:"전자제품"},{s:33,w:"manufacture",m:"제조하다"},{s:33,w:"goods",m:"상품"},{s:33,w:"might have pp",m:"~했을지도 모른다"},{s:33,w:"be willing to-V",m:"기꺼이~하다"},{s:33,w:"income",m:"수입"},{s:33,w:"offer",m:"제공하다"},{s:33,w:"manipulate",m:"조종하다, 조작하다"},{s:33,w:"presence",m:"존재, 있음"},
  {s:34,w:"climate",m:"기후"},{s:34,w:"disaster",m:"재난"},{s:34,w:"scope",m:"범위"},{s:34,w:"transformation",m:"변화"},{s:34,w:"eliminate",m:"제거하다, 없애다"},{s:34,w:"marginal",m:"주변적인"},{s:34,w:"temperature",m:"온도"},{s:34,w:"degree",m:"도, 온도, 정도, 각도"},{s:34,w:"hardly",m:"거의~아닌"},{s:34,w:"isolate",m:"고립시키다"},{s:34,w:"expand",m:"확장되다, 확대하다"},{s:34,w:"cease",m:"멈추다"},{s:34,w:"plainly",m:"뚜렷하게, 명백히"},{s:34,w:"illustrate",m:"보여주다"},{s:34,w:"escapist",m:"도피적인"},{s:34,w:"pleasure",m:"즐거움"},{s:34,w:"even if",m:"~라고 하더라도"},{s:34,w:"horror",m:"공포"},{s:34,w:"no longer",m:"더 이상~아닌"},{s:34,w:"pretend",m:"~라고 가장하다, 인척하다"},{s:34,w:"suffering",m:"고통"},{s:34,w:"distant",m:"먼"},{s:34,w:"stop ~ing",m:"~하던 것을 멈추다"},{s:34,w:"within",m:"~이내에"},
  {s:35,w:"crisis",m:"위기"},{s:35,w:"political",m:"정치적인"},{s:35,w:"inevitable",m:"피할 수 없는"},{s:35,w:"fix",m:"고치다"},{s:35,w:"functionally",m:"기능적으로"},{s:35,w:"elective",m:"선택의"},{s:35,w:"nevertheless",m:"그럼에도 불구하고"},{s:35,w:"distressing",m:"괴로운, 고통스러운"},{s:35,w:"abundant",m:"풍족한, 풍부한"},{s:35,w:"resource",m:"자원"},{s:35,w:"scarce",m:"희소한"},{s:35,w:"governmental",m:"정부의"},{s:35,w:"neglect",m:"소홀함, 방치"},{s:35,w:"indifference",m:"무관심"},{s:35,w:"infrastructure",m:"사회 기반 시설"},{s:35,w:"contamination",m:"오염"},{s:35,w:"careless",m:"부주의한"},{s:35,w:"urbanization",m:"도시화"},{s:35,w:"in other words",m:"다시 말해서"},{s:35,w:"address",m:"다루다, 처리하다"},{s:35,w:"leak",m:"누수"},{s:35,w:"deliver",m:"전달하다"},{s:35,w:"theft",m:"도난"},{s:35,w:"account for",m:"~을 차지하다"},{s:35,w:"estimate",m:"추정하다"},{s:35,w:"loss",m:"손실"},{s:35,w:"freshwater",m:"담수"},{s:35,w:"selective",m:"선택적인"},{s:35,w:"scarcity",m:"부족, 희소성"},{s:35,w:"highlight",m:"강조하다"},{s:35,w:"inequity",m:"불평등"},{s:35,w:"leave",m:"남겨두다, 떠나다"},{s:35,w:"billion",m:"십억"},{s:35,w:"proper",m:"적절한"},{s:35,w:"sanitation",m:"위생"},{s:35,w:"worldwide",m:"전 세계적으로"},
  {s:36,w:"thrive",m:"성공하다, 번창하다"},{s:36,w:"depend on",m:"~에 달려 있다"},{s:36,w:"navigate",m:"다루다"},{s:36,w:"value",m:"~를 가치 있게 여기다"},{s:36,w:"count on",m:"~을 기대하다"},{s:36,w:"probably",m:"아마도"},{s:36,w:"merit",m:"이점"},{s:36,w:"physically",m:"신체적으로"},{s:36,w:"genetically",m:"유전적으로"},{s:36,w:"million",m:"백만"},{s:36,w:"skilled",m:"능숙한"},{s:36,w:"maximize",m:"최대화하다"},{s:36,w:"standing",m:"지위"},{s:36,w:"tendency",m:"경향"},{s:36,w:"unconsciously",m:"무의식적으로"},{s:36,w:"monitor",m:"관찰하다, 감시하다"},{s:36,w:"perceive",m:"인식하다"},{s:36,w:"process",m:"처리하다, 과정"},{s:36,w:"self-esteem",m:"자존감"},{s:36,w:"pride",m:"자존심, 자부심"},{s:36,w:"shame",m:"수치심"},{s:36,w:"insecurity",m:"불안"},{s:36,w:"compel",m:"강요하다"},{s:36,w:"crucially",m:"결정적으로"},{s:36,w:"realize",m:"깨닫다"},{s:36,w:"respond to",m:"~에 반응하다"},{s:36,w:"performance",m:"수행"},{s:36,w:"come off",m:"나타나다"},{s:36,w:"grudging",m:"투덜대는"},{s:36,w:"cynical",m:"냉소적인"},{s:36,w:"persuasive",m:"설득력이 있는"},
];

/* ───────────────── EMPTY INITIAL STATE ───────────────── */
const DEFAULT_TABS = [];

const ACCESS_CODE = "141";

// ① 스프레드시트 > 확장 프로그램 > Apps Script에 gas-code.js 내용 붙여넣기
// ② [배포] → [새 배포] → 유형: 웹 앱, 액세스: 모든 사용자 → 배포
// ③ 배포 URL을 프로젝트 루트의 .env 파일에 저장:
//    VITE_SHEET_API_URL=https://script.google.com/macros/s/배포ID/exec
const SHEET_API_URL = import.meta.env?.VITE_SHEET_API_URL || "";

function normalizeSheetWords(data) {
  if (!Array.isArray(data)) return [];

  const str = (v) => (v == null ? "" : String(v).trim());

  return data
    .map((item) => {
      const section = item.s ?? item.section ?? item.no ?? item["지문번호"] ?? item["번호"] ?? item["범위"] ?? 1;
      const word    = item.w ?? item.word ?? item.WORD ?? item["영단어"] ?? item["단어"] ?? item["영어"] ?? "";
      const meaning = item.m ?? item.meaning ?? item.MEANING ?? item["한글뜻"] ?? item["뜻"] ?? item["의미"] ?? item["한국어"] ?? "";
      // 오답 보기 (스프레드시트 컬럼에서 직접 가져옴)
      const ew = [item["영어오답1"], item["영어오답2"], item["영어오답3"], item["영어오답4"]].map(str).filter(Boolean);
      const kw = [item["한글오답1"], item["한글오답2"], item["한글오답3"], item["한글오답4"]].map(str).filter(Boolean);
      return {
        s: isNaN(Number(section)) ? str(section) : Number(section),
        w: str(word),
        m: str(meaning),
        ew, // 영어 오답 보기 (뜻→영단어 문제용)
        kw, // 한글 오답 보기 (영단어→뜻 문제용)
      };
    })
    .filter((item) => item.w && item.m);
}

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function getSections(words) {
  return [...new Set(words.map(w => w.s))].sort((a, b) => {
    const an = typeof a === "number" ? a : parseInt(a) || 0;
    const bn = typeof b === "number" ? b : parseInt(b) || 0;
    if (an !== bn) return an - bn;
    return String(a).localeCompare(String(b));
  });
}

let _audioCtx = null;
function getAudioCtx() {
  if (!_audioCtx || _audioCtx.state === "closed") {
    _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  }
  if (_audioCtx.state === "suspended") _audioCtx.resume();
  return _audioCtx;
}

function playShutter() {
  try {
    const ctx = getAudioCtx();
    const now = ctx.currentTime;
    const sr = ctx.sampleRate;

    const click = (t, freq, q, vol, decay) => {
      const len = Math.floor(sr * Math.min(decay * 10, 0.12));
      const buf = ctx.createBuffer(1, len, sr);
      const d = buf.getChannelData(0);
      for (let i = 0; i < len; i++) d[i] = (Math.random() * 2 - 1) * Math.exp(-(i / sr) / decay);
      const f = ctx.createBiquadFilter();
      f.type = "bandpass"; f.frequency.value = freq; f.Q.value = q;
      const g = ctx.createGain(); g.gain.value = vol;
      const s = ctx.createBufferSource(); s.buffer = buf;
      s.connect(f); f.connect(g); g.connect(ctx.destination);
      s.start(t);
    };

    // 첫 번째 커튼: 날카로운 고음 + 바디 울림
    click(now,        4200, 0.6, 2.0, 0.004);
    click(now + 0.003, 1400, 1.8, 0.7, 0.018);
    // 두 번째 커튼 (80ms 후): 조금 낮고 부드럽게
    click(now + 0.08, 3000, 0.7, 1.4, 0.005);
    click(now + 0.084, 1100, 1.5, 0.5, 0.022);
  } catch (_) {}
}

function playCorrect() {
  try {
    const ctx = getAudioCtx();
    const now = ctx.currentTime;
    // 밝은 두 음: C5 → E5 (정답 딩동)
    [[523.25, 0], [659.25, 0.12]].forEach(([freq, t]) => {
      const osc = ctx.createOscillator();
      const g   = ctx.createGain();
      osc.type = "sine"; osc.frequency.value = freq;
      g.gain.setValueAtTime(0.4, now + t);
      g.gain.exponentialRampToValueAtTime(0.001, now + t + 0.25);
      osc.connect(g); g.connect(ctx.destination);
      osc.start(now + t); osc.stop(now + t + 0.3);
    });
  } catch (_) {}
}

function playWrong() {
  try {
    const ctx = getAudioCtx();
    const now = ctx.currentTime;
    // 낮고 짧은 버저음: 두 번 (오답 삐-삐)
    [0, 0.14].forEach(t => {
      const osc = ctx.createOscillator();
      const g   = ctx.createGain();
      osc.type = "sawtooth"; osc.frequency.value = 180;
      g.gain.setValueAtTime(0.3, now + t);
      g.gain.exponentialRampToValueAtTime(0.001, now + t + 0.1);
      osc.connect(g); g.connect(ctx.destination);
      osc.start(now + t); osc.stop(now + t + 0.12);
    });
  } catch (_) {}
}

/* ───────────────── MCQ TEST (객관식 4지선다) ───────────────── */
function MCQTest({ allWords, allSections }) {
  const [testSections, setTestSections] = useState(new Set(allSections));
  const [direction, setDirection]       = useState("kor"); // kor=영단어→뜻, eng=뜻→영단어
  const [screen, setScreen]             = useState("setup");
  const [deck, setDeck]                 = useState([]);   // [{word, choices, answerIdx}]
  const [qIdx, setQIdx]                 = useState(0);
  const [chosen, setChosen]             = useState(null); // 선택한 보기 인덱스
  const [score, setScore]               = useState(0);
  const [wrongList, setWrongList]       = useState([]);

  const filteredWords = useMemo(
    () => allWords.filter(w => testSections.has(w.s)),
    [allWords, testSections]
  );

  const toggleSec = (s) => setTestSections(prev => {
    const n = new Set(prev); n.has(s) ? n.delete(s) : n.add(s); return n;
  });

  const buildDeck = (words, dir) => {
    const shuffled = shuffle(words);
    return shuffled.map(word => {
      // 오답 보기: 시트에서 가져온 값 우선, 부족하면 다른 단어에서 랜덤 보충
      const rawWrong = dir === "kor"
        ? (word.kw && word.kw.length > 0 ? word.kw : [])
        : (word.ew && word.ew.length > 0 ? word.ew : []);
      const fallbackPool = words
        .filter(w => w.w !== word.w)
        .map(w => dir === "kor" ? w.m : w.w);
      const wrongs = [...rawWrong];
      for (const fb of shuffle(fallbackPool)) {
        if (wrongs.length >= 3) break;
        if (!wrongs.includes(fb)) wrongs.push(fb);
      }
      const correct = dir === "kor" ? word.m : word.w;
      const choices = shuffle([correct, ...wrongs.slice(0, 3)]);
      return { word, choices, answerIdx: choices.indexOf(correct) };
    });
  };

  const startTest = (dir = direction) => {
    if (filteredWords.length < 2) return;
    setDeck(buildDeck(filteredWords, dir));
    setQIdx(0); setChosen(null); setScore(0); setWrongList([]);
    setScreen("quiz");
  };

  const retryWrong = () => {
    if (wrongList.length === 0) return;
    setDeck(buildDeck(wrongList, direction));
    setQIdx(0); setChosen(null); setScore(0); setWrongList([]);
    setScreen("quiz");
  };

  const handleChoice = (idx) => {
    if (chosen !== null) return; // 이미 선택함
    setChosen(idx);
    const correct = idx === deck[qIdx].answerIdx;
    if (correct) { playCorrect(); setScore(s => s + 1); }
    else         { playWrong();  setWrongList(wl => [...wl, deck[qIdx].word]); }
  };

  const next = () => {
    if (qIdx + 1 < deck.length) { setQIdx(i => i + 1); setChosen(null); }
    else setScreen("done");
  };

  /* ── setup ── */
  if (screen === "setup") {
    return (
      <div style={{ animation: "fadeUp 0.4s ease-out" }}>
        <SectionChips label="✏️ 테스트 범위" sections={allSections} allWords={allWords}
          selected={testSections} onToggle={toggleSec}
          onAll={() => setTestSections(new Set(allSections))} onNone={() => setTestSections(new Set())} />
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 11, color: "#a0978a", marginBottom: 8, fontWeight: 500 }}>문제 방향</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className={`pill ${direction==="kor"?"active":""}`}
              onClick={() => setDirection("kor")}>영단어 → 뜻</button>
            <button className={`pill ${direction==="eng"?"active":""}`}
              onClick={() => setDirection("eng")}>뜻 → 영단어</button>
          </div>
        </div>
        <button className="action-btn primary" disabled={filteredWords.length < 2}
          onClick={() => startTest(direction)}>
          테스트 시작 ({filteredWords.length}문제)
        </button>
      </div>
    );
  }

  /* ── done ── */
  if (screen === "done") {
    const pct = Math.round(score / deck.length * 100);
    return (
      <div style={{ padding: "40px 20px", maxWidth: 500, margin: "0 auto", animation: "fadeUp 0.4s ease-out" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>{pct >= 80 ? "🎉" : pct >= 50 ? "🙂" : "😢"}</div>
          <h2 style={{ fontSize: 22, fontWeight: 700, color: "#2c2824", marginBottom: 6 }}>테스트 완료!</h2>
          <p style={{ fontSize: 28, fontWeight: 800, color: pct>=80?"#5a9a6a":"#d4644a", marginBottom: 4 }}>
            {score} <span style={{ fontSize: 16, color: "#a0978a" }}>/ {deck.length}</span>
          </p>
          <p style={{ fontSize: 14, color: "#a0978a" }}>정답률 {pct}%</p>
          <div style={{ display: "flex", gap: 10, justifyContent: "center", marginTop: 24, flexWrap: "wrap" }}>
            <button className="action-btn primary" onClick={() => startTest(direction)}>🔄 다시 풀기</button>
            {wrongList.length > 0 &&
              <button className="action-btn secondary" onClick={retryWrong}>❌ 틀린 것만 ({wrongList.length})</button>}
            <button className="action-btn secondary" onClick={() => setScreen("setup")}>설정으로</button>
          </div>
        </div>
        {wrongList.length > 0 && (
          <div style={{ borderTop: "1px solid #e8e3db", paddingTop: 20 }}>
            <div style={{ fontSize: 11, color: "#a0978a", fontWeight: 600, letterSpacing: 3, marginBottom: 12, textTransform: "uppercase" }}>틀린 단어</div>
            {wrongList.map((w, i) => (
              <div key={i} style={{ display: "flex", gap: 12, padding: "7px 0", borderBottom: "1px solid #f0ece5", alignItems: "baseline" }}>
                <span style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: 13, color: "#2c2824", fontWeight: 600, flex: 1 }}>{w.w}</span>
                <span style={{ fontSize: 13, color: "#6b6259" }}>{w.m}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  /* ── quiz ── */
  const q = deck[qIdx];
  const isKor = direction === "kor";
  return (
    <div style={{ animation: "fadeUp 0.3s ease-out", maxWidth: 560, margin: "0 auto" }}>
      {/* 진행 바 */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <div style={{ flex: 1, height: 4, background: "#e8e3db", borderRadius: 2, overflow: "hidden" }}>
          <div style={{ width: `${(qIdx / deck.length) * 100}%`, height: "100%", background: "#c4a46c", transition: "width 0.3s" }} />
        </div>
        <span style={{ fontSize: 12, color: "#a0978a", whiteSpace: "nowrap" }}>{qIdx + 1} / {deck.length}</span>
        <span style={{ fontSize: 12, color: "#5a9a6a", fontWeight: 600 }}>{score}점</span>
      </div>
      {/* 문제 */}
      <div style={{ background: "#fff", borderRadius: 16, padding: "32px 28px", marginBottom: 16, textAlign: "center", boxShadow: "0 2px 12px rgba(44,40,36,0.07)" }}>
        <div style={{ fontSize: 11, color: "#a0978a", letterSpacing: 3, marginBottom: 12, textTransform: "uppercase" }}>
          {isKor ? "Word" : "뜻"}
        </div>
        <div style={{ fontSize: isKor ? 28 : 20, fontWeight: 700, color: "#2c2824",
          fontFamily: isKor ? "'JetBrains Mono',monospace" : "inherit", wordBreak: "break-word", lineHeight: 1.4 }}>
          {isKor ? q.word.w : q.word.m}
        </div>
        <div style={{ fontSize: 11, color: "#c4a46c", marginTop: 10 }}>#{q.word.s}번</div>
      </div>
      {/* 보기 */}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {q.choices.map((choice, i) => {
          const isCorrect = i === q.answerIdx;
          const isChosen  = i === chosen;
          let bg = "#fff", border = "#e5e2dc", color = "#2c2824";
          if (chosen !== null) {
            if (isCorrect)      { bg = "#f0f7f2"; border = "#5a9a6a"; color = "#3a7a4a"; }
            else if (isChosen)  { bg = "#fdf5f3"; border = "#d4644a"; color = "#d4644a"; }
          }
          return (
            <button key={i} onClick={() => handleChoice(i)}
              style={{ width: "100%", padding: "14px 20px", background: bg, border: `2px solid ${border}`,
                borderRadius: 12, cursor: chosen !== null ? "default" : "pointer",
                fontSize: 15, color, fontWeight: isChosen || (chosen !== null && isCorrect) ? 700 : 400,
                textAlign: "left", transition: "all 0.15s",
                fontFamily: !isKor ? "'JetBrains Mono',monospace" : "inherit",
                display: "flex", alignItems: "center", gap: 12 }}>
              <span style={{ minWidth: 24, height: 24, borderRadius: "50%",
                background: chosen !== null && isCorrect ? "#5a9a6a" : chosen !== null && isChosen ? "#d4644a" : "#f0ece5",
                color: chosen !== null && (isCorrect || isChosen) ? "#fff" : "#a0978a",
                display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700 }}>
                {String.fromCharCode(9312 + i) /* ①②③④ */}
              </span>
              {choice}
            </button>
          );
        })}
      </div>
      {/* 다음 버튼 */}
      {chosen !== null && (
        <div style={{ marginTop: 20, textAlign: "center" }}>
          <button className="action-btn primary" onClick={next}>
            {qIdx + 1 < deck.length ? "다음 →" : "결과 보기"}
          </button>
        </div>
      )}
    </div>
  );
}

/* ───────────────── CARD TEST MODE ───────────────── */
function CardTestMode({ allWords, allSections }) {
  const [testSections, setTestSections] = useState(new Set(allSections));
  const [wordCount, setWordCount] = useState(20);
  const [screen, setScreen] = useState("setup");
  const [deck, setDeck] = useState([]);
  const [idx, setIdx] = useState(0);
  const [countdown, setCountdown] = useState(3);
  const timerRef = useRef(null);

  const clearTimers = () => {
    if (timerRef.current) {
      clearTimeout(timerRef.current.to);
      clearInterval(timerRef.current.iv);
      timerRef.current = null;
    }
  };

  useEffect(() => () => clearTimers(), []);

  const filteredWords = useMemo(
    () => allWords.filter(w => testSections.has(w.s)),
    [allWords, testSections]
  );

  const goBlack = (curIdx, curDeck) => {
    clearTimers();
    setScreen("black");
    let c = 5;
    setCountdown(c);
    const iv = setInterval(() => { c--; setCountdown(c); }, 1000);
    const to = setTimeout(() => {
      clearInterval(iv);
      if (curIdx + 1 < curDeck.length) {
        goWord(curIdx + 1, curDeck);
      } else {
        clearTimers();
        setScreen("done");
      }
    }, 5000);
    timerRef.current = { to, iv };
  };

  const goWord = (curIdx, curDeck) => {
    clearTimers();
    setIdx(curIdx);
    setScreen("word");
    playShutter();
    let c = 3;
    setCountdown(c);
    const iv = setInterval(() => { c--; setCountdown(c); }, 1000);
    const to = setTimeout(() => {
      clearInterval(iv);
      goBlack(curIdx, curDeck);
    }, 3000);
    timerRef.current = { to, iv };
  };

  const startTest = () => {
    const actualCount = wordCount === "all" ? filteredWords.length : Math.min(wordCount, filteredWords.length);
    if (actualCount === 0) return;
    const newDeck = shuffle(filteredWords).slice(0, actualCount);
    setDeck(newDeck);
    goWord(0, newDeck);
  };

  const stopTest = () => { clearTimers(); setScreen("setup"); };

  const toggleSec = (s) => setTestSections(prev => {
    const n = new Set(prev);
    n.has(s) ? n.delete(s) : n.add(s);
    return n;
  });

  if (screen === "word") {
    const word = deck[idx];
    return (
      <div style={{ position: "fixed", inset: 0, background: "#faf9f6", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", zIndex: 50, fontFamily: "'Noto Sans KR', sans-serif" }}>
        <div style={{ position: "absolute", top: 24, left: 24, right: 24, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 13, color: "#a0978a", fontWeight: 500 }}>{idx + 1} / {deck.length}</span>
          <button onClick={stopTest} style={{ padding: "6px 16px", border: "1px solid #d4cfc6", borderRadius: 6, background: "#fff", cursor: "pointer", fontSize: 13, fontFamily: "'Noto Sans KR'" }}>중단</button>
        </div>
        <div style={{ fontSize: 11, color: "#a0978a", letterSpacing: 4, marginBottom: 16, textTransform: "uppercase" }}>Word</div>
        <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 40, fontWeight: 700, color: "#2c2824", textAlign: "center", wordBreak: "break-word", maxWidth: "80vw", lineHeight: 1.2 }}>{word.w}</div>
        <div style={{ fontSize: 12, color: "#c4a46c", marginTop: 14 }}>#{word.s}번</div>
        <div style={{ marginTop: 48, width: 68, height: 68, borderRadius: "50%", background: "#2c2824", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 16px rgba(44,40,36,0.15)" }}>
          <span style={{ fontSize: 30, fontWeight: 700, color: "#fff" }}>{Math.max(countdown, 0)}</span>
        </div>
      </div>
    );
  }

  if (screen === "black") {
    return (
      <div style={{ position: "fixed", inset: 0, background: "#080808", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", zIndex: 50 }}>
        <div style={{ position: "absolute", top: 24, left: 24, right: 24, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 13, color: "#3a3a3a" }}>{idx + 1} / {deck.length}</span>
          <button onClick={stopTest} style={{ padding: "6px 16px", border: "1px solid #2a2a2a", borderRadius: 6, background: "transparent", cursor: "pointer", fontSize: 13, color: "#3a3a3a", fontFamily: "'Noto Sans KR'" }}>중단</button>
        </div>
        <div style={{ fontSize: 12, color: "#3b3b3b", marginBottom: 20, fontFamily: "'Noto Sans KR'" }}>한글 뜻을 쓰세요</div>
        <div style={{ width: 60, height: 60, borderRadius: "50%", border: "2px solid #2a2a2a", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <span style={{ fontSize: 26, fontWeight: 600, color: "#2a2a2a" }}>{Math.max(countdown, 0)}</span>
        </div>
      </div>
    );
  }

  if (screen === "done") {
    return (
      <div style={{ padding: "40px 20px", maxWidth: 600, margin: "0 auto", animation: "fadeUp 0.4s ease-out" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontSize: 52, marginBottom: 16 }}>✅</div>
          <h2 style={{ fontSize: 22, fontWeight: 700, color: "#2c2824", marginBottom: 8 }}>테스트 완료!</h2>
          <p style={{ fontSize: 14, color: "#8a8278", marginBottom: 24 }}>총 {deck.length}개 단어를 모두 봤어요.</p>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", justifyContent: "center" }}>
            <button className="action-btn primary" onClick={() => goWord(0, deck)}>🔄 다시 시작</button>
            <button className="action-btn secondary" onClick={stopTest}>설정으로</button>
          </div>
        </div>
        <div style={{ borderTop: "1px solid #e8e3db", paddingTop: 24 }}>
          <div style={{ fontSize: 11, color: "#a0978a", fontWeight: 600, letterSpacing: 3, marginBottom: 16, textTransform: "uppercase" }}>Answer Sheet</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0" }}>
            {deck.map((word, i) => (
              <div key={i} style={{ display: "flex", alignItems: "baseline", gap: 8, padding: "7px 10px", borderBottom: "1px solid #f0ece5", background: i % 2 === 0 ? "#faf9f6" : "#fff" }}>
                <span style={{ fontSize: 11, color: "#c4a46c", minWidth: 20, textAlign: "right" }}>{i + 1}</span>
                <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: "#2c2824", fontWeight: 600, minWidth: 0, flex: 1 }}>{word.w}</span>
                <span style={{ fontSize: 13, color: "#6b6259", flexShrink: 0 }}>{word.m}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // setup
  const actualCount = wordCount === "all" ? filteredWords.length : Math.min(wordCount, filteredWords.length);
  return (
    <div style={{ animation: "fadeUp 0.4s ease-out" }}>
      <SectionChips label="📂 범위 선택" sections={allSections} allWords={allWords} selected={testSections}
        onToggle={toggleSec} onAll={() => setTestSections(new Set(allSections))} onNone={() => setTestSections(new Set())} />
      <div style={{ marginBottom: 20 }}>
        <div style={{ fontSize: 11, color: "#a0978a", marginBottom: 8, fontWeight: 500 }}>테스트 단어 수</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[10, 20, 30, "all"].map(n => (
            <button key={n} className={`pill ${wordCount === n ? "active" : ""}`} onClick={() => setWordCount(n)}>
              {n === "all" ? `전체 (${filteredWords.length}개)` : `${n}개`}
            </button>
          ))}
        </div>
        <p style={{ fontSize: 12, color: "#a0978a", marginTop: 10 }}>
          선택 범위 {filteredWords.length}개 중 {actualCount}개 출제
        </p>
      </div>
      <div style={{ marginBottom: 24, padding: "14px 18px", background: "rgba(196,164,108,0.08)", borderRadius: 10, border: "1px solid rgba(196,164,108,0.2)", fontSize: 13, color: "#6b655c", lineHeight: 1.9 }}>
        📸 영단어 <strong>3초</strong> 표시 후 → 검은 화면 <strong>5초</strong> (종이에 한글 뜻 쓰기)<br/>
        찰칵 소리와 함께 다음 단어로 넘어갑니다.
      </div>
      <button className="action-btn primary" onClick={startTest} disabled={filteredWords.length === 0}
        style={{ opacity: filteredWords.length === 0 ? 0.4 : 1 }}>
        테스트 시작
      </button>
    </div>
  );
}

/* ───────────────── LOGIN ───────────────── */
function LoginScreen({ onLogin }) {
  const [code, setCode] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);

  const handleSubmit = () => {
    if (code.trim().toLowerCase() === ACCESS_CODE) {
      onLogin();
    } else {
      setError(true);
      setShake(true);
      setTimeout(() => setShake(false), 500);
    }
  };

  return (
    <div className="login-shell">
      <div className="login-card" style={{ animation: shake ? "shakeX 0.5s" : "fadeUp 0.6s ease-out" }}>
        <img className="brand-logo" src="/oneforone-logo.jpeg" alt="원포원영어학원" style={{ width: 210, margin: "0 auto 22px" }} />
        <div style={{ width: 44, height: 3, background: "#214f2d", borderRadius: 999, margin: "0 auto 24px" }} />
        <p style={{ fontSize: 14, color: "#657083", marginBottom: 28 }}>학원에서 배부한 코드를 입력하세요</p>
        <div style={{ display: "flex", gap: 8, justifyContent: "center", maxWidth: 340, margin: "0 auto" }}>
          <input type="password" value={code}
            onChange={e => { setCode(e.target.value); setError(false); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            placeholder="접속 코드"
            style={{
              flex: 1, minWidth: 0, padding: "12px 16px", border: `1.5px solid ${error ? "#d4644a" : "#d8dee8"}`,
              borderRadius: 10, fontSize: 15, outline: "none", background: "#fff", color: "#172033",
              fontFamily: "var(--font-mono)",
            }}
          />
          <button onClick={handleSubmit} style={{
            padding: "12px 24px", border: "none", borderRadius: 10, background: "#172033",
            color: "#ffffff", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "var(--font-sans)", minWidth: 72, whiteSpace: "nowrap",
          }}>입장</button>
        </div>
        {error && <p style={{ color: "#d4644a", fontSize: 13, marginTop: 12 }}>코드가 올바르지 않습니다.</p>}
      </div>
    </div>
  );
}

/* ───────────────── PRINT VIEW ───────────────── */

function PrintView({ tabs, activeTabIdx, onClose }) {
  const [selTabIdx, setSelTabIdx] = useState(activeTabIdx);
  const [printMode, setPrintMode] = useState("full");
  const [selSections, setSelSections] = useState(new Set());

  const curTab = tabs[selTabIdx] || tabs[0];
  const tabSections = useMemo(() => getSections(curTab.words), [curTab]);

  useEffect(() => {
    setSelSections(new Set(getSections(curTab.words)));
  }, [selTabIdx]);

  useEffect(() => {
    document.body.style.overflow = "auto";
    return () => { document.body.style.overflow = ""; };
  }, []);

  const printWords = useMemo(
    () => curTab.words.filter(w => selSections.has(w.s)),
    [curTab.words, selSections]
  );
  const printSections = useMemo(
    () => tabSections.filter(s => selSections.has(s)),
    [tabSections, selSections]
  );
  const toggleSec = (s) => setSelSections(prev => {
    const n = new Set(prev); n.has(s) ? n.delete(s) : n.add(s); return n;
  });

  const renderSection = (sec) => {
    const secWords = printWords.filter(w => w.s === sec);
    if (!secWords.length) return null;
    return (
      <div key={sec} style={{ marginBottom: 14, pageBreakInside: "avoid", breakInside: "avoid" }}>
        <div style={{ fontSize: 10, fontWeight: 700, color: "#c4a46c", padding: "3px 0", borderBottom: "1.5px solid #c4a46c", marginBottom: 2 }}>#{sec}번</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5, tableLayout: "fixed" }}>
          <tbody>
            {secWords.map((w, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #f0ede6", pageBreakInside: "avoid", breakInside: "avoid" }}>
                <td style={{ padding: "3px 4px", width: 18, color: "#ccc", textAlign: "right", fontSize: 10 }}>{i+1}</td>
                <td style={{ padding: "3px 4px", fontFamily: "'JetBrains Mono', monospace", fontWeight: 500, width: "48%", color: "#1a1a1a", overflow: "hidden" }}>
                  {printMode==="meaningOnly"
                    ? <span style={{ borderBottom: "1px dotted #ccc", display: "inline-block", minWidth: 65 }}>&nbsp;</span>
                    : w.w}
                </td>
                <td style={{ padding: "3px 4px", color: "#444", fontSize: 11, overflow: "hidden" }}>
                  {printMode==="wordOnly"
                    ? <span style={{ borderBottom: "1px dotted #ccc", display: "inline-block", minWidth: 65 }}>&nbsp;</span>
                    : w.m}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div id="print-area" style={{ position: "fixed", inset: 0, zIndex: 9999, background: "#fff", color: "#222", overflow: "auto", fontFamily: "'Noto Sans KR', sans-serif" }}>
      <style>{`
        @page { size: A4 portrait; margin: 12mm 14mm; }
        @media print {
          .no-print   { display: none !important; }
          #print-area { position: static !important; overflow: visible !important; background: #fff !important; }
        }
      `}</style>

      {/* ── 상단 툴바 ── */}
      <div className="no-print" style={{ position: "sticky", top: 0, zIndex: 10, background: "#2c2824", padding: "10px 18px", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <button onClick={onClose} style={{ padding: "7px 14px", border: "1.5px solid rgba(255,255,255,0.35)", borderRadius: 6, background: "transparent", cursor: "pointer", fontSize: 13, fontFamily: "'Noto Sans KR'", color: "#faf9f6", fontWeight: 500 }}>← 돌아가기</button>
        <span style={{ width: 1, height: 20, background: "rgba(255,255,255,0.2)" }} />
        {tabs.map((tab, idx) => (
          <button key={idx} onClick={() => setSelTabIdx(idx)} style={{ padding: "7px 14px", border: `1.5px solid ${idx===selTabIdx?"#c4a46c":"rgba(255,255,255,0.2)"}`, borderRadius: 6, background: idx===selTabIdx?"#c4a46c":"transparent", cursor: "pointer", fontSize: 13, fontFamily: "'Noto Sans KR'", color: idx===selTabIdx?"#fff":"rgba(255,255,255,0.55)", fontWeight: idx===selTabIdx?600:400 }}>{tab.name}</button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 6, alignItems: "center" }}>
          {[["full","전체"],["wordOnly","뜻 빈칸"],["meaningOnly","단어 빈칸"]].map(([v,l]) => (
            <button key={v} onClick={() => setPrintMode(v)} style={{ padding: "7px 12px", border: `1.5px solid ${printMode===v?"#c4a46c":"rgba(255,255,255,0.2)"}`, borderRadius: 6, background: printMode===v?"#c4a46c":"transparent", color: printMode===v?"#fff":"rgba(255,255,255,0.55)", cursor: "pointer", fontSize: 12, fontFamily: "'Noto Sans KR'" }}>{l}</button>
          ))}
          <button onClick={() => window.print()} style={{ padding: "7px 18px", border: "none", borderRadius: 6, background: "#c4a46c", color: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "'Noto Sans KR'", marginLeft: 4 }}>🖨 프린트</button>
        </div>
      </div>

      {/* ── 섹션 선택 바 ── */}
      <div className="no-print" style={{ background: "#f4f1eb", padding: "10px 18px", borderBottom: "1px solid #e5e2dc", display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
        <span style={{ fontSize: 11, color: "#a0978a", fontWeight: 500, marginRight: 2 }}>섹션 선택</span>
        <button onClick={() => setSelSections(new Set(tabSections))} style={{ padding: "3px 10px", border: "1px solid #d4cfc6", borderRadius: 12, background: "#fff", cursor: "pointer", fontSize: 11, fontFamily: "'Noto Sans KR'", color: "#6b655c" }}>전체</button>
        <button onClick={() => setSelSections(new Set())} style={{ padding: "3px 10px", border: "1px solid #d4cfc6", borderRadius: 12, background: "#fff", cursor: "pointer", fontSize: 11, fontFamily: "'Noto Sans KR'", color: "#6b655c" }}>해제</button>
        <span style={{ width: 1, height: 16, background: "#d4cfc6" }} />
        {tabSections.map(s => (
          <button key={s} onClick={() => toggleSec(s)} style={{ padding: "3px 10px", border: `1px solid ${selSections.has(s)?"#2c2824":"#d4cfc6"}`, borderRadius: 12, background: selSections.has(s)?"#2c2824":"#fff", color: selSections.has(s)?"#faf9f6":"#8a8278", cursor: "pointer", fontSize: 11, fontFamily: "'Noto Sans KR'" }}>
            #{s} <span style={{ opacity: 0.55 }}>({curTab.words.filter(w => w.s === s).length})</span>
          </button>
        ))}
        <span style={{ marginLeft: "auto", fontSize: 11, color: "#a0978a" }}>총 {printWords.length}개</span>
      </div>

      {/* ── 인쇄 본문 (A4 기준) ── */}
      <div className="print-body" style={{ padding: "20px 24px", maxWidth: 794, margin: "0 auto" }}>
        {/* 타이틀 */}
        <div style={{ textAlign: "center", marginBottom: 16 }}>
          <div style={{ fontSize: 10, letterSpacing: 4, color: "#a0978a", textTransform: "uppercase", marginBottom: 3 }}>Vocabulary</div>
          <h2 style={{ fontSize: 16, fontWeight: 700, color: "#2c2824", marginBottom: 3 }}>{curTab.name}</h2>
          <p style={{ fontSize: 10, color: "#aaa" }}>총 {printWords.length}개 · {printMode==="full"?"전체":printMode==="wordOnly"?"뜻 빈칸":"단어 빈칸"}</p>
          <div style={{ width: 32, height: 1.5, background: "#c4a46c", margin: "7px auto 0" }} />
        </div>

        {printWords.length === 0
          ? <div style={{ textAlign: "center", padding: 60, color: "#a0978a" }}>섹션을 선택해주세요.</div>
          : printSections.map(renderSection)
        }
      </div>
    </div>
  );
}

/* ───────────────── SECTION CHIPS ───────────────── */
function SectionChips({ sections, allWords, selected, onToggle, onAll, onNone, label }) {
  return (
    <div style={{ marginBottom: 14 }}>
      {label && <div style={{ fontSize: 11, color: "#a0978a", marginBottom: 6, fontWeight: 500 }}>{label}</div>}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
        <button className="pill" onClick={onAll} style={{ fontSize: 11 }}>전체</button>
        <button className="pill" onClick={onNone} style={{ fontSize: 11 }}>해제</button>
        <span style={{ width: 1, height: 20, background: "#d4cfc6", margin: "0 2px" }} />
        {sections.map(s => (
          <button key={s} className={`pill ${selected.has(s) ? "active" : ""}`} onClick={() => onToggle(s)} style={{ fontSize: 11 }}>
            #{s} <span style={{ opacity: 0.6, marginLeft: 2 }}>({allWords.filter(w => w.s === s).length})</span>
          </button>
        ))}
      </div>
    </div>
  );
}

/* ───────────────── MAIN APP ───────────────── */
export default function VocabApp() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [tabs, setTabs] = useState([]);
  const [activeTabIdx, setActiveTabIdx] = useState(0);
  const [allWords, setAllWords] = useState([]);
  const [sheetLoading, setSheetLoading] = useState(false);
  const [sheetError, setSheetError] = useState("");
  const [dataSource, setDataSource] = useState("로딩 중...");
  const [mode, setMode] = useState("list");
  const [selectedSections, setSelectedSections] = useState(new Set());
  const [hideMode, setHideMode] = useState("none");
  const [starred, setStarred] = useState(() => {
    try {
      const saved = localStorage.getItem("vocab-starred");
      return saved ? new Set(JSON.parse(saved)) : new Set();
    } catch { return new Set(); }
  });
  const [filterStarred, setFilterStarred] = useState(false);
  const [showPrint, setShowPrint] = useState(false);
  const [cardIdx, setCardIdx] = useState(0);
  const [flipped, setFlipped] = useState(false);
  const [cardDeck, setCardDeck] = useState([]);
  const [cardSections, setCardSections] = useState(new Set());
  const [showUpload, setShowUpload] = useState(false);
  const [uploadMsg, setUploadMsg] = useState("");
  const fileRef = useRef(null);
  const swipeRef = useRef(null);

  const allSections = useMemo(() => getSections(allWords), [allWords]);

  // ⭐ 즐겨찾기 localStorage 자동 저장
  useEffect(() => {
    try { localStorage.setItem("vocab-starred", JSON.stringify([...starred])); } catch {}
  }, [starred]);

  const resetWordState = (words) => {
    const secs = new Set(getSections(words));
    setAllWords(words);
    setSelectedSections(secs);
    setCardSections(secs);
    // starred는 초기화하지 않음 — localStorage에서 유지됨
    setFilterStarred(false);
    setCardDeck([]);
    setCardIdx(0);
    setFlipped(false);
  };

  const applyTabs = (newTabs, sourceLabel) => {
    if (newTabs.length === 0) {
      setSheetError("단어 데이터가 없습니다. 구글시트를 확인해주세요.");
      setDataSource("연결 실패");
      return;
    }
    setTabs(newTabs);
    setActiveTabIdx(0);
    resetWordState(newTabs[0].words);
    setDataSource(sourceLabel);
  };

  const switchTab = (idx) => {
    if (idx === activeTabIdx) return;
    setActiveTabIdx(idx);
    resetWordState(tabs[idx].words);
  };

  const loadSheetWords = async () => {
    if (!SHEET_API_URL) {
      setDataSource("기본 내장 단어");
      return;
    }

    setSheetLoading(true);
    setSheetError("");
    try {
      const response = await fetch(`${SHEET_API_URL}${SHEET_API_URL.includes("?") ? "&" : "?"}t=${Date.now()}`);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const json = await response.json();

      // 멀티탭 형식: { sheets: [{ name, words }] }
      if (json.sheets && Array.isArray(json.sheets)) {
        const parsedTabs = json.sheets
          .map(sheet => ({ name: sheet.name, words: normalizeSheetWords(sheet.words) }))
          .filter(tab => tab.words.length > 0);
        if (parsedTabs.length === 0) throw new Error("구글시트에서 단어 데이터를 찾지 못했습니다.");
        applyTabs(parsedTabs, "구글시트 연동");
      } else {
        // 단일 배열 형식 (하위 호환)
        const parsed = normalizeSheetWords(Array.isArray(json) ? json : json.words);
        if (parsed.length === 0) throw new Error("구글시트에서 단어 데이터를 찾지 못했습니다.");
        applyTabs([{ name: "전체", words: parsed }], `구글시트 연동 · ${parsed.length}개`);
      }
    } catch (error) {
      console.error(error);
      setSheetError("구글시트 데이터를 불러오지 못했습니다. '시트 새로고침'을 눌러 다시 시도해주세요.");
      setDataSource("연결 실패");
    } finally {
      setSheetLoading(false);
    }
  };

  useEffect(() => {
    loadSheetWords();
  }, []);

  const filteredWords = useMemo(() => {
    let w = allWords.filter(x => selectedSections.has(x.s));
    if (filterStarred) w = w.filter(x => starred.has(x.w));
    return w;
  }, [allWords, selectedSections, filterStarred, starred]);

  const filteredSections = useMemo(() => allSections.filter(s => selectedSections.has(s)), [allSections, selectedSections]);

  const cardFilteredWords = useMemo(() => allWords.filter(x => cardSections.has(x.s)), [allWords, cardSections]);

  const toggleSection = (s) => setSelectedSections(prev => { const n = new Set(prev); n.has(s)?n.delete(s):n.add(s); return n; });
  const toggleCardSection = (s) => setCardSections(prev => { const n = new Set(prev); n.has(s)?n.delete(s):n.add(s); return n; });

  const [cardShuffle, setCardShuffle] = useState(true);

  const buildCardDeck = (words, doShuffle = cardShuffle) => {
    const deck = doShuffle ? shuffle(words) : [...words];
    setCardDeck(deck);
    setCardIdx(0);
    setFlipped(false);
  };

  const startCards = (doShuffle = cardShuffle) => buildCardDeck(cardFilteredWords, doShuffle);

  useEffect(() => {
    if (mode === "card") buildCardDeck(allWords.filter(x => cardSections.has(x.s)), cardShuffle);
  }, [cardSections]);

  const toggleStar = (wordKey) => setStarred(prev => { const n = new Set(prev); n.has(wordKey)?n.delete(wordKey):n.add(wordKey); return n; });

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let startRow = 0;
        let colWord = -1, colMeaning = -1, colSection = -1;

        for (let r = 0; r < Math.min(5, rows.length); r++) {
          const row = (rows[r] || []).map(c => String(c || "").toLowerCase().trim());
          for (let c = 0; c < row.length; c++) {
            if (/word|단어|영어|english|vocabulary/i.test(row[c])) colWord = c;
            if (/meaning|뜻|의미|한국어|korean/i.test(row[c])) colMeaning = c;
            if (/section|번호|번|no|구분|범위/i.test(row[c])) colSection = c;
          }
          if (colWord >= 0 && colMeaning >= 0) { startRow = r + 1; break; }
        }

        if (colWord < 0 || colMeaning < 0) {
          const firstRow = rows[0] || [];
          if (firstRow.length >= 3) {
            colSection = 0; colWord = 1; colMeaning = 2; startRow = 0;
            if (isNaN(Number(firstRow[0])) && /[a-zA-Z가-힣]/.test(String(firstRow[0]))) startRow = 1;
          } else if (firstRow.length >= 2) {
            colWord = 0; colMeaning = 1; startRow = 0;
            if (/word|단어/i.test(String(firstRow[0]))) startRow = 1;
          }
        }

        const parsed = [];
        let currentSection = 1;
        for (let r = startRow; r < rows.length; r++) {
          const row = rows[r];
          if (!row || row.length === 0) continue;
          const word = String(row[colWord] ?? "").trim();
          const meaning = String(row[colMeaning] ?? "").trim();
          if (!word && !meaning) continue;
          if (colSection >= 0) {
            const sv = row[colSection];
            if (sv !== undefined && sv !== null && String(sv).trim() !== "") {
              currentSection = isNaN(Number(sv)) ? String(sv).trim() : Number(sv);
            }
          }
          if (word) parsed.push({ s: currentSection, w: word, m: meaning });
        }

        if (parsed.length === 0) {
          setUploadMsg("단어를 찾을 수 없습니다. 열 순서를 확인해주세요.");
          return;
        }

        applyTabs([{ name: "업로드", words: parsed }], `업로드 파일 · ${parsed.length}개`);
        setUploadMsg(`${parsed.length}개 단어가 등록되었습니다!`);
        setShowUpload(false);
        setMode("list");
      } catch (err) {
        setUploadMsg("파일을 읽을 수 없습니다. xlsx/csv 파일을 확인해주세요.");
      }
    };
    reader.readAsArrayBuffer(file);
    if (fileRef.current) fileRef.current.value = "";
  };

  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} />;
  if (showPrint) return <PrintView tabs={tabs} activeTabIdx={activeTabIdx} onClose={() => setShowPrint(false)} />;

  return (
    <div className="app-shell">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Noto+Sans+KR:wght@300;400;500;700;800&display=swap');
        *{box-sizing:border-box}
        @keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
        @keyframes shakeX{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-4px)}80%{transform:translateX(4px)}}
        .app-shell{min-height:100vh;background:linear-gradient(145deg,#fbfcff 0%,#eef3f8 48%,#f7f4ec 100%);font-family:var(--font-sans)}
        .login-shell{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;background:linear-gradient(145deg,#fbfcff 0%,#edf3f8 52%,#f7f4ec 100%);font-family:var(--font-sans)}
        .login-card{width:min(420px,100%);text-align:center;background:rgba(255,255,255,0.82);border:1px solid rgba(216,222,232,0.9);border-radius:18px;padding:42px 28px;box-shadow:0 24px 60px rgba(23,32,51,0.12),0 2px 12px rgba(23,32,51,0.05);backdrop-filter:blur(16px)}
        .app-frame{max-width:900px;margin:0 auto;padding:32px 20px 0}
        .surface{background:rgba(255,255,255,0.78);border:1px solid rgba(216,222,232,0.95);box-shadow:0 18px 46px rgba(23,32,51,0.08);backdrop-filter:blur(14px)}
        .header-panel{border-radius:22px;padding:24px 24px 0;overflow:hidden}
        .pill{padding:8px 14px;border:1px solid #d8dee8;background:#fff;color:#657083;border-radius:999px;cursor:pointer;font-size:12px;font-family:var(--font-sans);font-weight:600;transition:all 0.16s ease;white-space:nowrap;box-shadow:0 1px 2px rgba(23,32,51,0.04)}
        .pill:hover{border-color:#2f7f7a;color:#2f7f7a;transform:translateY(-1px)}
        .pill.active{border-color:#172033;background:#172033;color:#fff;box-shadow:0 8px 22px rgba(23,32,51,0.16)}
        .tab-btn{padding:13px 18px;border:none;background:transparent;color:#657083;font-size:14px;font-weight:700;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;font-family:var(--font-sans);position:relative}
        .tab-btn:hover{color:#172033;background:rgba(47,127,122,0.05)}
        .tab-btn.active{color:#172033;border-bottom-color:#2f7f7a;background:#fff}
        .vocab-row{display:grid;grid-template-columns:40px minmax(0,1.1fr) minmax(0,1fr);align-items:center;padding:13px 18px;border-bottom:1px solid #edf0f5;transition:background 0.12s,transform 0.12s;gap:12px}
        .vocab-row:hover{background:#f7fbfb}
        .vocab-row span{min-width:0;overflow-wrap:anywhere}
        .star-btn{background:#f6f8fb;border:1px solid transparent;border-radius:999px;width:28px;height:28px;cursor:pointer;font-size:15px;transition:transform 0.15s,border-color 0.15s,background 0.15s;display:flex;align-items:center;justify-content:center}
        .star-btn:hover{transform:scale(1.08);border-color:#d8dee8;background:#fff}
        .blur-text{filter:blur(5px);transition:filter 0.15s;user-select:none;cursor:pointer}
        .blur-text:hover{filter:blur(0)}
        .card-container{perspective:900px;width:100%;max-width:460px;height:280px;margin:0 auto;cursor:pointer}
        .card-inner{position:relative;width:100%;height:100%;transition:transform 0.5s cubic-bezier(0.4,0,0.2,1);transform-style:preserve-3d}
        .card-inner.flipped{transform:rotateY(180deg)}
        .card-face{position:absolute;width:100%;height:100%;backface-visibility:hidden;border-radius:20px;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:32px}
        .card-front{background:#fff;border:1px solid #d8dee8;box-shadow:0 20px 48px rgba(23,32,51,0.1)}
        .card-back{background:linear-gradient(145deg,#172033,#2f4e68);transform:rotateY(180deg);color:#fff}
        .test-input{background:#fff;border:1.5px solid #d8dee8;color:#172033;padding:11px 14px;border-radius:10px;font-family:var(--font-mono);font-size:14px;width:100%;max-width:240px;outline:none;transition:border-color 0.2s,box-shadow 0.2s}
        .test-input:focus{border-color:#2f7f7a;box-shadow:0 0 0 4px rgba(47,127,122,0.12)}
        .test-input.correct{border-color:#5a9a6a;background:#f0f7f2}
        .test-input.wrong{border-color:#d4644a;background:#fdf5f3}
        .action-btn{padding:11px 22px;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;font-family:var(--font-sans);transition:all 0.2s}
        .action-btn.primary{background:#172033;color:#fff;box-shadow:0 12px 26px rgba(23,32,51,0.18)}
        .action-btn.primary:hover{background:#26334b;transform:translateY(-1px)}
        .action-btn.secondary{background:#e9eef5;color:#344153}
        .action-btn.secondary:hover{background:#dfe7f0}
        .nav-circle{width:42px;height:42px;border-radius:50%;border:1px solid #d8dee8;background:#fff;color:#657083;font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.15s}
        .nav-circle:hover{border-color:#2f7f7a;color:#2f7f7a;transform:translateY(-1px)}
        .nav-circle:disabled{opacity:0.3;cursor:not-allowed}
        .modal-overlay{position:fixed;inset:0;background:rgba(23,32,51,0.36);z-index:100;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px)}
        .modal-box{background:#fff;border-radius:20px;padding:28px;max-width:520px;width:100%;max-height:80vh;overflow:auto;box-shadow:0 24px 72px rgba(23,32,51,0.22)}
        .drop-zone{border:2px dashed #cbd5e1;border-radius:16px;padding:42px 20px;text-align:center;cursor:pointer;transition:all 0.2s;background:#f8fafc}
        .drop-zone:hover{border-color:#2f7f7a;background:#f2faf9}
        .section-card{background:#fff;border:1px solid #dfe5ee;border-radius:16px;overflow:hidden;box-shadow:0 10px 26px rgba(23,32,51,0.05)}
        .section-title{font-size:12px;font-weight:800;color:#2f7f7a;padding:10px 18px;background:#f2faf9;border-bottom:1px solid #d5e8e5;letter-spacing:1px}
        .exam-tabs{display:flex;gap:0;overflow-x:auto;overflow-y:hidden;padding-top:16px;flex-wrap:nowrap}
        .exam-tab{padding:9px 18px;border:1px solid #d8dee8;border-bottom:none;background:#f6f8fb;color:#657083;border-radius:10px 10px 0 0;cursor:pointer;font-size:13px;font-weight:500;font-family:var(--font-sans);white-space:nowrap;transition:all 0.15s;margin-right:2px;position:relative}
        .exam-tab.active{border-color:#2f7f7a;background:#fff;color:#172033;font-weight:800;z-index:2}
        .mode-strip{display:flex;gap:0;border:1px solid #d8dee8;border-top-color:#2f7f7a;margin-bottom:16px;background:#fff;overflow-x:auto;border-radius:0 0 14px 14px;box-shadow:0 8px 18px rgba(23,32,51,0.03)}
        @media(max-width:700px){
          .app-frame{padding:18px 12px 0}
          .header-panel{border-radius:18px;padding:20px 16px 0}
          .header-panel .brand-logo{width:92px!important;height:56px!important}
          .exam-tab{padding:9px 14px}
          .vocab-row{grid-template-columns:34px minmax(0,1fr);padding:12px 14px;font-size:13px}
          .vocab-row span:last-child{grid-column:2;color:#657083}
          .card-container{height:240px}
          .mode-strip{margin-inline:0}
        }
      `}</style>

      {/* Header */}
      <div className="app-frame">
        <div className="surface header-panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16, minWidth: 0 }}>
            <img className="brand-logo" src="/oneforone-logo.jpeg" alt="원포원영어학원" style={{ width: 118, height: 72, flex: "0 0 auto" }} />
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: 11, letterSpacing: 4, color: "#657083", marginBottom: 6, textTransform: "uppercase" }}>vocabulary</div>
              <h1 style={{ fontFamily: "var(--font-sans)", fontSize: 28, fontWeight: 800, color: "#214f2d", letterSpacing: 0, margin: 0 }}>원포원영어학원</h1>
              <p style={{ fontSize: 13, color: "#657083", marginTop: 6 }}>함께 걷는 영어, 원포원이 응원합니다</p>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            {SHEET_API_URL && <button onClick={loadSheetWords} disabled={sheetLoading} className="pill">{sheetLoading ? "불러오는 중..." : "🔄 시트 새로고침"}</button>}
            <button onClick={() => setShowUpload(true)} className="pill">📥 단어 업로드</button>
            <button onClick={() => setShowPrint(true)} className="pill">🖨 프린트</button>
          </div>
        </div>
        {uploadMsg && (
          <div style={{ marginTop: 12, padding: "10px 16px", borderRadius: 8, background: uploadMsg.includes("등록") ? "#f0f7f2" : "#fdf5f3", color: uploadMsg.includes("등록") ? "#5a9a6a" : "#d4644a", fontSize: 13, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            {uploadMsg}
            <button onClick={() => setUploadMsg("")} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: "inherit" }}>✕</button>
          </div>
        )}
        <div style={{ marginTop: 14, padding: "10px 14px", borderRadius: 12, background: sheetError ? "#fdf5f3" : "#f2faf9", color: sheetError ? "#d4644a" : "#456260", fontSize: 12, display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap", border: `1px solid ${sheetError ? "#f4c7bd" : "#d5e8e5"}` }}>
          <span>데이터: {sheetLoading ? "구글시트 불러오는 중..." : dataSource}</span>
          {sheetError && <span>{sheetError}</span>}
        </div>
        <div style={{ width: "100%", height: 1, background: "linear-gradient(90deg, #2f7f7a, #d8dee8, transparent)", margin: "20px 0 0" }} />

        {/* 단어 탭 */}
        <div className="exam-tabs">
          {tabs.map((tab, idx) => (
            <button key={idx} onClick={() => switchTab(idx)} className={`exam-tab ${idx === activeTabIdx ? "active" : ""}`}>
              {tab.name}
              <span style={{ marginLeft: 5, fontSize: 11, opacity: 0.6 }}>({tab.words.length})</span>
            </button>
          ))}
        </div>

        {/* 모드 탭 */}
        <div className="mode-strip">
          {[["list","📋 목록"],["card","🃏 카드"],["cardtest","📸 카드테스트"],["test","✏️ 테스트"]].map(([k,l]) => (
            <button key={k} className={`tab-btn ${mode===k?"active":""}`}
              onClick={() => { setMode(k); if(k==="card") startCards(); }}
              style={{ whiteSpace: "nowrap" }}
            >{l}</button>
          ))}
        </div>

        {mode === "list" && (
          <>
            <SectionChips label="📂 범위 선택" sections={allSections} allWords={allWords} selected={selectedSections}
              onToggle={toggleSection} onAll={() => setSelectedSections(new Set(allSections))} onNone={() => setSelectedSections(new Set())} />
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
              <button className={`pill ${hideMode==="meaning"?"active":""}`} onClick={() => setHideMode(hideMode==="meaning"?"none":"meaning")}>뜻 가리기</button>
              <button className={`pill ${hideMode==="word"?"active":""}`} onClick={() => setHideMode(hideMode==="word"?"none":"word")}>단어 가리기</button>
              <button className={`pill ${filterStarred?"active":""}`} onClick={() => setFilterStarred(!filterStarred)}>⭐ {starred.size}</button>
              <span style={{ marginLeft: "auto", fontSize: 12, color: "#657083", fontWeight: 700 }}>{filteredWords.length}개</span>
            </div>
          </>
        )}
        </div>
      </div>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "18px 20px 48px" }}>

        {/* LIST */}
        {mode === "list" && (
          <div>
            {filteredSections.map(sec => {
              const secWords = filteredWords.filter(w => w.s === sec);
              if (!secWords.length) return null;
              return (
                <div key={sec} className="section-card" style={{ marginBottom: 18 }}>
                  <div className="section-title">#{sec}번</div>
                  <div style={{ background: "#fff", overflow: "hidden" }}>
                    {secWords.map((w, i) => (
                      <div className="vocab-row" key={`${sec}-${i}`}>
                        <button className="star-btn" onClick={() => toggleStar(w.w)}>{starred.has(w.w) ? "⭐" : <span style={{ color: "#c4a46c" }}>☆</span>}</button>
                        <span className={hideMode==="word"?"blur-text":""} style={{ fontFamily: "var(--font-mono)", fontSize: 14, fontWeight: 600, color: "#172033" }}>{w.w}</span>
                        <span className={hideMode==="meaning"?"blur-text":""} style={{ fontSize: 13, color: "#4f5b6b" }}>{w.m}</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
            {filteredWords.length === 0 && (
              <div style={{ textAlign: "center", padding: 60, color: "#a0978a" }}>
                {filterStarred ? "즐겨찾기한 단어가 없습니다."
                  : allWords.length === 0
                    ? (sheetLoading ? "구글시트에서 단어를 불러오는 중..." : "상단의 '시트 새로고침' 버튼을 눌러 단어를 불러오세요.")
                    : "범위를 선택해주세요."}
              </div>
            )}
          </div>
        )}

        {/* CARD */}
        {mode === "card" && (
          <div style={{ animation: "fadeUp 0.4s ease-out" }}>
            <SectionChips label="🃏 카드 범위 선택" sections={allSections} allWords={allWords} selected={cardSections}
              onToggle={toggleCardSection} onAll={() => setCardSections(new Set(allSections))} onNone={() => setCardSections(new Set())} />
            {/* 순서/셔플 토글 */}
            <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
              <button className={`pill ${!cardShuffle?"active":""}`} onClick={() => { setCardShuffle(false); startCards(false); }}>순서대로</button>
              <button className={`pill ${cardShuffle?"active":""}`} onClick={() => { setCardShuffle(true); startCards(true); }}>🔀 셔플</button>
            </div>
            {cardDeck.length > 0 ? (
              <>
                <div style={{ textAlign: "center", marginBottom: 20, color: "#a0978a", fontSize: 13 }}>
                  {cardIdx + 1} / {cardDeck.length}
                </div>
                {/* 스와이프 + 클릭 둘 다 지원 */}
                <div
                  className="card-container"
                  onClick={() => setFlipped(f => !f)}
                  onTouchStart={(e) => { swipeRef.current = { x: e.touches[0].clientX, y: e.touches[0].clientY }; }}
                  onTouchEnd={(e) => {
                    const start = swipeRef.current;
                    if (!start) return;
                    const dx = e.changedTouches[0].clientX - start.x;
                    const dy = Math.abs(e.changedTouches[0].clientY - start.y);
                    swipeRef.current = null;
                    if (Math.abs(dx) < 10 && dy < 10) return; // 탭은 onClick이 처리
                    if (Math.abs(dx) > 40 && dy < 60) { // 좌우 스와이프 → 다음/이전
                      e.preventDefault();
                      if (dx < 0 && cardIdx < cardDeck.length - 1) { setCardIdx(i => i+1); setFlipped(false); }
                      if (dx > 0 && cardIdx > 0) { setCardIdx(i => i-1); setFlipped(false); }
                    }
                  }}
                >
                  <div className={`card-inner ${flipped?"flipped":""}`}>
                    <div className="card-face card-front">
                      <div style={{ fontSize: 11, color: "#a0978a", marginBottom: 10, letterSpacing: 3, textTransform: "uppercase" }}>Word</div>
                      <div style={{ fontFamily: "'JetBrains Mono'", fontSize: 28, fontWeight: 600, color: "#2c2824", textAlign: "center", wordBreak: "break-word" }}>{cardDeck[cardIdx].w}</div>
                      <div style={{ fontSize: 11, color: "#c4a46c", marginTop: 16 }}>#{cardDeck[cardIdx].s}번</div>
                    </div>
                    <div className="card-face card-back">
                      <div style={{ fontSize: 11, color: "#a0978a", marginBottom: 10, letterSpacing: 3, textTransform: "uppercase" }}>뜻</div>
                      <div style={{ fontSize: 22, fontWeight: 700, textAlign: "center" }}>{cardDeck[cardIdx].m}</div>
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", justifyContent: "center", gap: 14, marginTop: 28, alignItems: "center" }}>
                  <button className="nav-circle" disabled={cardIdx===0} onClick={() => {setCardIdx(cardIdx-1);setFlipped(false);}}>←</button>
                  <button className="nav-circle" disabled={cardIdx>=cardDeck.length-1} onClick={() => {setCardIdx(cardIdx+1);setFlipped(false);}}>→</button>
                </div>
              </>
            ) : (
              <div style={{ textAlign: "center", padding: 60, color: "#a0978a" }}>범위를 선택해주세요.</div>
            )}
          </div>
        )}

        {/* CARD TEST */}
        {mode === "cardtest" && (
          <CardTestMode allWords={allWords} allSections={allSections} />
        )}

        {/* TEST — 객관식 4지선다 */}
        {mode === "test" && (
          <MCQTest allWords={allWords} allSections={allSections} key={activeTabIdx} />
        )}
      </div>

      {/* Upload Modal */}
      {showUpload && (
        <div className="modal-overlay" onClick={() => setShowUpload(false)}>
          <div className="modal-box" onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <h3 style={{ fontSize: 16, fontWeight: 700, color: "#2c2824" }}>📥 스프레드시트 업로드</h3>
              <button onClick={() => setShowUpload(false)} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#a0978a" }}>✕</button>
            </div>
            <div className="drop-zone" onClick={() => fileRef.current?.click()}>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFileUpload} />
              <div style={{ fontSize: 32, marginBottom: 12 }}>📄</div>
              <p style={{ fontSize: 14, color: "#5a554d", fontWeight: 500 }}>클릭하여 파일 선택</p>
              <p style={{ fontSize: 12, color: "#a0978a", marginTop: 6 }}>.xlsx, .xls, .csv 파일 지원</p>
            </div>
            <div style={{ marginTop: 20, padding: 16, background: "#faf9f6", borderRadius: 10, border: "1px solid #e5e2dc" }}>
              <p style={{ fontSize: 13, fontWeight: 600, color: "#2c2824", marginBottom: 8 }}>스프레드시트 양식 안내</p>
              <p style={{ fontSize: 12, color: "#6b655c", lineHeight: 1.7 }}>아래 3개 열을 순서대로 작성해주세요.</p>
              <div style={{ marginTop: 10, fontFamily: "'JetBrains Mono'", fontSize: 11, color: "#5a554d", background: "#fff", padding: 12, borderRadius: 6, border: "1px solid #e5e2dc", lineHeight: 1.8 }}>
                <div style={{ display: "grid", gridTemplateColumns: "60px 1fr 1fr", gap: 8, fontWeight: 600, borderBottom: "1px solid #e5e2dc", paddingBottom: 6, marginBottom: 6 }}>
                  <span>번호</span><span>WORD</span><span>MEANING</span>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "60px 1fr 1fr", gap: 8 }}><span>18</span><span>department</span><span>부서, 부</span></div>
                <div style={{ display: "grid", gridTemplateColumns: "60px 1fr 1fr", gap: 8 }}><span>18</span><span>submit</span><span>제출하다</span></div>
                <div style={{ display: "grid", gridTemplateColumns: "60px 1fr 1fr", gap: 8 }}><span>19</span><span>figure</span><span>생각하다</span></div>
              </div>
              <p style={{ fontSize: 11, color: "#a0978a", marginTop: 8 }}>
                * 헤더 행이 있으면 자동 인식합니다.<br/>
                * 번호 열이 없으면 단어/뜻 2열만으로도 가능합니다.
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
