import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ───────────────── WORD DATA ───────────────── */
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

/* ───────────────── TAB DATA ───────────────── */
const DEFAULT_TABS = [
  { name: "18~21번", words: DEFAULT_WORDS.filter(w => w.s >= 18 && w.s <= 21) },
  { name: "22~25번", words: DEFAULT_WORDS.filter(w => w.s >= 22 && w.s <= 25) },
  { name: "26~30번", words: DEFAULT_WORDS.filter(w => w.s >= 26 && w.s <= 30) },
  { name: "31~36번", words: DEFAULT_WORDS.filter(w => w.s >= 31) },
];

const ACCESS_CODE = "141";

// ① 스프레드시트 > 확장 프로그램 > Apps Script에 gas-code.js 내용 붙여넣기
// ② [배포] → [새 배포] → 유형: 웹 앱, 액세스: 모든 사용자 → 배포
// ③ 배포 URL을 프로젝트 루트의 .env 파일에 저장:
//    VITE_SHEET_API_URL=https://script.google.com/macros/s/배포ID/exec
const SHEET_API_URL = import.meta.env?.VITE_SHEET_API_URL || "";

function normalizeSheetWords(data) {
  if (!Array.isArray(data)) return [];

  return data
    .map((item) => {
      const section = item.s ?? item.section ?? item.no ?? item["번호"] ?? item["범위"] ?? 1;
      const word = item.w ?? item.word ?? item.WORD ?? item["단어"] ?? item["영어"] ?? "";
      const meaning = item.m ?? item.meaning ?? item.MEANING ?? item["뜻"] ?? item["의미"] ?? item["한국어"] ?? "";
      return {
        s: isNaN(Number(section)) ? String(section).trim() : Number(section),
        w: String(word).trim(),
        m: String(meaning).trim(),
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
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "60px 20px", textAlign: "center", animation: "fadeUp 0.4s ease-out" }}>
        <div style={{ fontSize: 52, marginBottom: 16 }}>✅</div>
        <h2 style={{ fontSize: 22, fontWeight: 700, color: "#2c2824", marginBottom: 8 }}>테스트 완료!</h2>
        <p style={{ fontSize: 14, color: "#8a8278", marginBottom: 32 }}>총 {deck.length}개 단어를 모두 봤어요.</p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", justifyContent: "center" }}>
          <button className="action-btn primary" onClick={() => goWord(0, deck)}>🔄 다시 시작</button>
          <button className="action-btn secondary" onClick={stopTest}>설정으로</button>
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
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center",
      background: "linear-gradient(145deg, #faf9f6 0%, #f0ede6 100%)",
      fontFamily: "'Noto Sans KR', sans-serif",
    }}>
      <div style={{ textAlign: "center", animation: shake ? "shakeX 0.5s" : "fadeUp 0.6s ease-out" }}>
        <div style={{ fontSize: 11, letterSpacing: 6, color: "#a0978a", marginBottom: 12, textTransform: "uppercase" }}>
          원포원 영어학원 내신 단어장
        </div>
        <h1 style={{ fontFamily: "'Playfair Display', Georgia, serif", fontSize: 42, fontWeight: 700, color: "#2c2824", marginBottom: 8, letterSpacing: -1 }}>
          VOCABULARY
        </h1>
        <div style={{ width: 40, height: 2, background: "#c4a46c", margin: "0 auto 32px" }} />
        <p style={{ fontSize: 14, color: "#8a8278", marginBottom: 28 }}>학원에서 배부한 코드를 입력하세요</p>
        <div style={{ display: "flex", gap: 8, justifyContent: "center", maxWidth: 340, margin: "0 auto" }}>
          <input type="password" value={code}
            onChange={e => { setCode(e.target.value); setError(false); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            placeholder="접속 코드"
            style={{
              flex: 1, padding: "12px 16px", border: `1.5px solid ${error ? "#d4644a" : "#d4cfc6"}`,
              borderRadius: 8, fontSize: 15, outline: "none", background: "#fff", color: "#2c2824",
              fontFamily: "'JetBrains Mono', monospace",
            }}
          />
          <button onClick={handleSubmit} style={{
            padding: "12px 24px", border: "none", borderRadius: 8, background: "#2c2824",
            color: "#faf9f6", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "'Noto Sans KR'",
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
  const [tabs, setTabs] = useState(DEFAULT_TABS);
  const [activeTabIdx, setActiveTabIdx] = useState(0);
  const [allWords, setAllWords] = useState(DEFAULT_TABS[0].words);
  const [sheetLoading, setSheetLoading] = useState(false);
  const [sheetError, setSheetError] = useState("");
  const [dataSource, setDataSource] = useState(SHEET_API_URL ? "구글시트 연결 대기" : "기본 내장 단어");
  const [mode, setMode] = useState("list");
  const [selectedSections, setSelectedSections] = useState(new Set(getSections(DEFAULT_TABS[0].words)));
  const [hideMode, setHideMode] = useState("none");
  const [starred, setStarred] = useState(new Set());
  const [filterStarred, setFilterStarred] = useState(false);
  const [showPrint, setShowPrint] = useState(false);
  const [cardIdx, setCardIdx] = useState(0);
  const [flipped, setFlipped] = useState(false);
  const [cardDeck, setCardDeck] = useState([]);
  const [cardSections, setCardSections] = useState(new Set(getSections(DEFAULT_TABS[0].words)));
  const [testWords, setTestWords] = useState([]);
  const [testAnswers, setTestAnswers] = useState({});
  const [testSubmitted, setTestSubmitted] = useState(false);
  const [testType, setTestType] = useState("eng");
  const [showUpload, setShowUpload] = useState(false);
  const [uploadMsg, setUploadMsg] = useState("");
  const fileRef = useRef(null);

  const allSections = useMemo(() => getSections(allWords), [allWords]);

  const resetWordState = (words) => {
    const secs = new Set(getSections(words));
    setAllWords(words);
    setSelectedSections(secs);
    setCardSections(secs);
    setStarred(new Set());
    setFilterStarred(false);
    setCardDeck([]);
    setCardIdx(0);
    setFlipped(false);
    setTestWords([]);
    setTestAnswers({});
    setTestSubmitted(false);
  };

  const applyTabs = (newTabs, sourceLabel) => {
    const validTabs = newTabs.length > 0 ? newTabs : DEFAULT_TABS;
    setTabs(validTabs);
    setActiveTabIdx(0);
    resetWordState(validTabs[0].words);
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
      setSheetError("구글시트 데이터를 불러오지 못해 기본 내장 단어를 표시합니다.");
      applyTabs(DEFAULT_TABS, "기본 내장 단어");
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

  const buildCardDeck = (words) => {
    const deck = shuffle(words);
    setCardDeck(deck);
    setCardIdx(0);
    setFlipped(false);
  };

  const startCards = () => buildCardDeck(cardFilteredWords);

  useEffect(() => {
    if (mode === "card") buildCardDeck(allWords.filter(x => cardSections.has(x.s)));
  }, [cardSections]);

  const startTest = (type = "eng") => {
    setTestType(type);
    setTestWords(shuffle(filteredWords).slice(0, Math.min(20, filteredWords.length)));
    setTestAnswers({});
    setTestSubmitted(false);
    setMode("test");
  };

  const toggleStar = (wordKey) => setStarred(prev => { const n = new Set(prev); n.has(wordKey)?n.delete(wordKey):n.add(wordKey); return n; });

  const testScore = testSubmitted ? testWords.filter((w, i) => {
    const ans = (testAnswers[i] || "").trim().toLowerCase();
    if (testType === "eng") return ans === w.w.toLowerCase();
    return w.m.split(",").some(m => ans === m.trim());
  }).length : 0;

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
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg, #faf9f6 0%, #f4f1eb 50%, #ebe7df 100%)", fontFamily: "'Noto Sans KR', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Noto+Sans+KR:wght@300;400;500;700&family=Playfair+Display:wght@600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        @keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
        @keyframes shakeX{0%,100%{transform:translateX(0)}20%{transform:translateX(-8px)}40%{transform:translateX(8px)}60%{transform:translateX(-4px)}80%{transform:translateX(4px)}}
        .pill{padding:7px 14px;border:1px solid #d4cfc6;background:#fff;color:#8a8278;border-radius:20px;cursor:pointer;font-size:12px;font-family:'Noto Sans KR';transition:all 0.15s;white-space:nowrap}
        .pill:hover{border-color:#c4a46c;color:#c4a46c}
        .pill.active{border-color:#2c2824;background:#2c2824;color:#faf9f6}
        .tab-btn{padding:10px 20px;border:none;background:transparent;color:#a0978a;font-size:14px;font-weight:500;cursor:pointer;border-bottom:2px solid transparent;transition:all 0.2s;font-family:'Noto Sans KR'}
        .tab-btn:hover{color:#2c2824}
        .tab-btn.active{color:#2c2824;border-bottom-color:#c4a46c}
        .vocab-row{display:grid;grid-template-columns:36px 1.2fr 1fr;align-items:center;padding:11px 16px;border-bottom:1px solid #ebe7df;transition:background 0.12s;gap:8px}
        .vocab-row:hover{background:rgba(196,164,108,0.06)}
        .star-btn{background:none;border:none;cursor:pointer;font-size:16px;transition:transform 0.15s}
        .star-btn:hover{transform:scale(1.25)}
        .blur-text{filter:blur(5px);transition:filter 0.15s;user-select:none;cursor:pointer}
        .blur-text:hover{filter:blur(0)}
        .card-container{perspective:900px;width:100%;max-width:420px;height:260px;margin:0 auto;cursor:pointer}
        .card-inner{position:relative;width:100%;height:100%;transition:transform 0.5s cubic-bezier(0.4,0,0.2,1);transform-style:preserve-3d}
        .card-inner.flipped{transform:rotateY(180deg)}
        .card-face{position:absolute;width:100%;height:100%;backface-visibility:hidden;border-radius:14px;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:28px}
        .card-front{background:#fff;border:1px solid #e5e2dc;box-shadow:0 4px 24px rgba(44,40,36,0.06)}
        .card-back{background:#2c2824;transform:rotateY(180deg);color:#faf9f6}
        .test-input{background:#fff;border:1.5px solid #d4cfc6;color:#2c2824;padding:10px 14px;border-radius:8px;font-family:'JetBrains Mono',monospace;font-size:14px;width:100%;max-width:220px;outline:none;transition:border-color 0.2s}
        .test-input:focus{border-color:#c4a46c}
        .test-input.correct{border-color:#5a9a6a;background:#f0f7f2}
        .test-input.wrong{border-color:#d4644a;background:#fdf5f3}
        .action-btn{padding:10px 22px;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;font-family:'Noto Sans KR';transition:all 0.2s}
        .action-btn.primary{background:#2c2824;color:#faf9f6}
        .action-btn.primary:hover{background:#3d3630}
        .action-btn.secondary{background:#e5e2dc;color:#5a554d}
        .action-btn.secondary:hover{background:#d4cfc6}
        .nav-circle{width:40px;height:40px;border-radius:50%;border:1px solid #d4cfc6;background:#fff;color:#8a8278;font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.15s}
        .nav-circle:hover{border-color:#c4a46c;color:#c4a46c}
        .nav-circle:disabled{opacity:0.3;cursor:not-allowed}
        .modal-overlay{position:fixed;inset:0;background:rgba(44,40,36,0.3);z-index:100;display:flex;align-items:center;justify-content:center}
        .modal-box{background:#fff;border-radius:16px;padding:28px;max-width:500px;width:90%;max-height:80vh;overflow:auto;box-shadow:0 16px 48px rgba(44,40,36,0.15)}
        .drop-zone{border:2px dashed #d4cfc6;border-radius:12px;padding:40px 20px;text-align:center;cursor:pointer;transition:all 0.2s;background:#faf9f6}
        .drop-zone:hover{border-color:#c4a46c;background:#f8f5ef}
        @media(max-width:600px){.vocab-row{grid-template-columns:32px 1fr 1fr;padding:10px 12px;font-size:13px}}
      `}</style>

      {/* Header */}
      <div style={{ maxWidth: 780, margin: "0 auto", padding: "36px 20px 0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: 5, color: "#a0978a", marginBottom: 6, textTransform: "uppercase" }}>vocabulary</div>
            <h1 style={{ fontFamily: "'Playfair Display', Georgia, serif", fontSize: 24, fontWeight: 700, color: "#2c2824", letterSpacing: -0.5 }}>원포원영어학원</h1>
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
        <div style={{ marginTop: 12, padding: "9px 14px", borderRadius: 8, background: sheetError ? "#fdf5f3" : "rgba(196,164,108,0.08)", color: sheetError ? "#d4644a" : "#8a8278", fontSize: 12, display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <span>데이터: {sheetLoading ? "구글시트 불러오는 중..." : dataSource}</span>
          {sheetError && <span>{sheetError}</span>}
        </div>
        <div style={{ width: "100%", height: 1, background: "linear-gradient(90deg, #c4a46c, #d4cfc6, transparent)", margin: "20px 0 0" }} />

        {/* 단어 탭 */}
        <div style={{ display: "flex", gap: 0, overflowX: "auto", paddingTop: 16, flexWrap: "nowrap" }}>
          {tabs.map((tab, idx) => (
            <button key={idx} onClick={() => switchTab(idx)} style={{
              padding: "8px 18px",
              border: `1px solid ${idx === activeTabIdx ? "#c4a46c" : "#d4cfc6"}`,
              borderBottom: idx === activeTabIdx ? "2px solid #fff" : "1px solid #d4cfc6",
              background: idx === activeTabIdx ? "#fff" : "#f4f1eb",
              color: idx === activeTabIdx ? "#2c2824" : "#8a8278",
              borderRadius: "8px 8px 0 0",
              cursor: "pointer",
              fontSize: 13,
              fontWeight: idx === activeTabIdx ? 600 : 400,
              fontFamily: "'Noto Sans KR'",
              whiteSpace: "nowrap",
              transition: "all 0.15s",
              marginRight: 2,
              position: "relative",
              bottom: -1,
            }}>
              {tab.name}
              <span style={{ marginLeft: 5, fontSize: 11, opacity: 0.6 }}>({tab.words.length})</span>
            </button>
          ))}
        </div>
        <div style={{ borderBottom: "1px solid #d4cfc6", marginBottom: 0 }} />

        {/* 모드 탭 */}
        <div style={{ display: "flex", gap: 4, borderBottom: "1px solid #e5e2dc", marginBottom: 16, background: "#fff", padding: "0 4px", overflowX: "auto" }}>
          {[["list","📋 목록"],["card","🃏 카드"],["cardtest","📸 카드테스트"],["test","✏️ 테스트"]].map(([k,l]) => (
            <button key={k} className={`tab-btn ${mode===k?"active":""}`}
              onClick={() => { setMode(k); if(k==="card") startCards(); if(k==="test") startTest(); }}
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
              <span style={{ marginLeft: "auto", fontSize: 12, color: "#a0978a" }}>{filteredWords.length}개</span>
            </div>
          </>
        )}
      </div>

      <div style={{ maxWidth: 780, margin: "0 auto", padding: "0 20px 48px" }}>

        {/* LIST */}
        {mode === "list" && (
          <div>
            {filteredSections.map(sec => {
              const secWords = filteredWords.filter(w => w.s === sec);
              if (!secWords.length) return null;
              return (
                <div key={sec} style={{ marginBottom: 20 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#c4a46c", padding: "8px 16px", background: "rgba(196,164,108,0.08)", borderRadius: "10px 10px 0 0", borderBottom: "1.5px solid #c4a46c", letterSpacing: 1 }}>#{sec}번</div>
                  <div style={{ background: "#fff", border: "1px solid #e5e2dc", borderTop: "none", borderRadius: "0 0 10px 10px", overflow: "hidden" }}>
                    {secWords.map((w, i) => (
                      <div className="vocab-row" key={`${sec}-${i}`}>
                        <button className="star-btn" onClick={() => toggleStar(w.w)}>{starred.has(w.w) ? "⭐" : <span style={{ color: "#c4a46c" }}>☆</span>}</button>
                        <span className={hideMode==="word"?"blur-text":""} style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, fontWeight: 500, color: "#2c2824" }}>{w.w}</span>
                        <span className={hideMode==="meaning"?"blur-text":""} style={{ fontSize: 13, color: "#6b655c" }}>{w.m}</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
            {filteredWords.length === 0 && <div style={{ textAlign: "center", padding: 60, color: "#a0978a" }}>{filterStarred ? "즐겨찾기한 단어가 없습니다." : "범위를 선택해주세요."}</div>}
          </div>
        )}

        {/* CARD */}
        {mode === "card" && (
          <div style={{ animation: "fadeUp 0.4s ease-out" }}>
            <SectionChips label="🃏 카드 범위 선택" sections={allSections} allWords={allWords} selected={cardSections}
              onToggle={toggleCardSection} onAll={() => setCardSections(new Set(allSections))} onNone={() => setCardSections(new Set())} />
            {cardDeck.length > 0 ? (
              <>
                <div style={{ textAlign: "center", marginBottom: 20, color: "#a0978a", fontSize: 13 }}>
                  {cardIdx + 1} / {cardDeck.length} <span style={{ marginLeft: 8, fontSize: 12 }}>(클릭하여 뒤집기)</span>
                </div>
                <div className="card-container" onClick={() => setFlipped(!flipped)}>
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
                  <button className="action-btn secondary" onClick={startCards}>🔀 셔플</button>
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

        {/* TEST */}
        {mode === "test" && (
          <div style={{ animation: "fadeUp 0.4s ease-out" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
              <div style={{ display: "flex", gap: 8 }}>
                <button className={`pill ${testType==="eng"?"active":""}`} onClick={() => startTest("eng")}>뜻 → 영단어</button>
                <button className={`pill ${testType==="kor"?"active":""}`} onClick={() => startTest("kor")}>영단어 → 뜻</button>
              </div>
              {testSubmitted && (
                <span style={{ padding: "5px 14px", borderRadius: 20, fontSize: 13, fontWeight: 700, background: testScore >= testWords.length * 0.8 ? "#f0f7f2" : "#fdf5f3", color: testScore >= testWords.length * 0.8 ? "#5a9a6a" : "#d4644a" }}>
                  {testScore} / {testWords.length}
                </span>
              )}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {testWords.map((w, i) => {
                const ans = (testAnswers[i] || "").trim().toLowerCase();
                const isCorrect = testSubmitted && (testType === "eng" ? ans === w.w.toLowerCase() : w.m.split(",").some(m => ans === m.trim()));
                const isWrong = testSubmitted && !isCorrect;
                return (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", background: "#fff", borderRadius: 10, border: `1px solid ${isCorrect?"#5a9a6a":isWrong?"#d4644a":"#e5e2dc"}` }}>
                    <span style={{ fontSize: 12, color: "#bbb", width: 24, textAlign: "right" }}>{i+1}</span>
                    <span style={{ flex: 1, fontSize: 14, fontWeight: 500, color: "#2c2824", minWidth: 80, fontFamily: testType === "kor" ? "'JetBrains Mono', monospace" : "inherit" }}>
                      {testType === "eng" ? w.m : w.w}
                    </span>
                    <input className={`test-input ${testSubmitted?(isCorrect?"correct":"wrong"):""}`}
                      placeholder={testType === "eng" ? "영단어 입력" : "뜻 입력"}
                      value={testAnswers[i] || ""} disabled={testSubmitted}
                      onChange={e => setTestAnswers({...testAnswers, [i]: e.target.value})}
                      onKeyDown={e => { if (e.key === "Enter" && !testSubmitted) { const next = e.target.closest("div").nextElementSibling?.querySelector("input"); if (next) next.focus(); else setTestSubmitted(true); } }}
                    />
                    {isWrong && <span style={{ fontSize: 12, color: "#d4644a", fontWeight: 600, minWidth: 60, fontFamily: testType === "eng" ? "'JetBrains Mono'" : "inherit" }}>{testType === "eng" ? w.w : w.m}</span>}
                  </div>
                );
              })}
            </div>
            <div style={{ display: "flex", gap: 12, marginTop: 24 }}>
              {!testSubmitted ? (
                <button className="action-btn primary" onClick={() => setTestSubmitted(true)}>채점하기</button>
              ) : (
                <>
                  <button className="action-btn primary" onClick={() => startTest(testType)}>다시 풀기</button>
                  <button className="action-btn secondary" onClick={() => {
                    const wrong = testWords.filter((w,i) => { const a = (testAnswers[i]||"").trim().toLowerCase(); return testType==="eng"?a!==w.w.toLowerCase():!w.m.split(",").some(m=>a===m.trim()); });
                    if (wrong.length > 0) { setTestWords(shuffle(wrong)); setTestAnswers({}); setTestSubmitted(false); }
                  }}>❌ 틀린 것만 다시</button>
                </>
              )}
            </div>
          </div>
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
