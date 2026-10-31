// rules.mjs — 사전에 넣지 않는 글 (merge·extract --check가 같이 쓴다)
const JOSA = /^(?:\(?[으이]?\)?[은는이가을를와과로도의에만]?(?:\/[은는이가을를와과로])?|\(으\)로|으로(?:\/로)?|[가-힣])$/;
const KEEP = new Set(['나']); // 한 글자지만 화면 글
const CODE = (k) => /^[\^/(<]|[\\$]|<\/?(?!b>)[a-z:]+[\s>]|\{(?!\d+\})[^}]+\}/i.test(k) || /^\s*(?:const|function|return|if)\b/.test(k);
const ONE = new Set(['열', '행', '쪽', '칸', '장', '개', '후', '전', '한', '나']); // 한 글자 낱말 (표 머리·배지)

export function skipKey(k) {
  const core = k.trim();
  if (ONE.has(core) || KEEP.has(core)) return false;
  if (JOSA.test(core)) return true;
  return CODE(k);
}

