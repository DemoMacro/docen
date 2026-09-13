// w:numFmt display formatting (OOXML ST_NumberFormat): a counter value → its
// rendered text, shared by the list-number substitution and the page-number
// field evaluation. Formats beyond a numeral system's glyph table (circled
// numbers past ⑳, the kana/jamo runs past their counts) fall back to decimal,
// matching Word's behavior for out-of-range values.

const CJK_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
const CJK_LEGAL_DIGITS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
const CJK_UNITS = ["", "十", "百", "千"];
const CJK_LEGAL_UNITS = ["", "拾", "佰", "仟"];

/** chineseCounting composition (零 fill between non-zero groups; the 10-19
 *  range drops the leading 一). `legal` swaps in the financial numerals
 *  (壹拾贰) — chineseLegalSimplified. */
function chineseNumeral(n: number, legal = false): string {
  if (n < 1 || n > 9999) return String(n);
  const digits: number[] = [];
  for (let rest = n; rest > 0; rest = Math.floor(rest / 10)) digits.unshift(rest % 10);
  const numerals = legal ? CJK_LEGAL_DIGITS : CJK_DIGITS;
  const units = legal ? CJK_LEGAL_UNITS : CJK_UNITS;
  let out = "";
  let zeroPending = false;
  digits.forEach((d, i) => {
    const unit = units[digits.length - 1 - i];
    if (d === 0) {
      if (out) zeroPending = true;
      return;
    }
    if (zeroPending) {
      out += numerals[0];
      zeroPending = false;
    }
    // 10-19 is 十X, not 一十X (the legal forms keep the leading 壹).
    if (!(d === 1 && unit === units[1] && digits.length === 2 && !legal)) out += numerals[d];
    out += unit;
  });
  return out;
}

const ROMAN_PAIRS: [number, string][] = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

export function romanNumeral(n: number, upper: boolean): string {
  let rest = n;
  let out = "";
  for (const [value, glyph] of ROMAN_PAIRS) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return upper ? out : out.toLowerCase();
}

/** 1→a…26→z, 27→aa (spreadsheet-style, Word's letter numbering). */
function letterNumeral(n: number, upper: boolean): string {
  let out = "";
  let rest = n;
  while (rest > 0) {
    rest--;
    out = String.fromCharCode(97 + (rest % 26)) + out;
    rest = Math.floor(rest / 26);
  }
  return upper ? out.toUpperCase() : out;
}

/** The kana sequences (aiueo's gojūon straight run, iroha's poem order) and
 *  the Korean runs (koreanCounting's syllables, ganada's jamo) — fixed glyph
 *  tables a counter indexes into. */
const AIUEO =
  "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわゐゑをん";
const IROHA =
  "いろはにほへとちりぬるをわかよたれそつねならむうゐのおくやまけふこえてあさきゆめみしゑひもせす";
const KATAKANA =
  "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヰヱヲン";
const KATAKANA_HALF = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ";
const KOREAN_COUNTING = "가나다라마바사아자차카타파하";
const GANADA = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ";

/** Fixed-glyph run tokens (kana, jamo): the counter indexes the table, past
 *  its end the decimal digits return. */
function glyphRun(table: string, n: number): string {
  return n >= 1 && n <= table.length ? table[n - 1]! : String(n);
}

const ENGLISH_ONES = [
  "one",
  "two",
  "three",
  "four",
  "five",
  "six",
  "seven",
  "eight",
  "nine",
  "ten",
  "eleven",
  "twelve",
  "thirteen",
  "fourteen",
  "fifteen",
  "sixteen",
  "seventeen",
  "eighteen",
  "nineteen",
];
const ENGLISH_TENS = [
  "",
  "",
  "twenty",
  "thirty",
  "forty",
  "fifty",
  "sixty",
  "seventy",
  "eighty",
  "ninety",
];

/** one → nine hundred ninety-nine (spelled-out English; beyond that the
 *  decimal digits return — page counts never reach it). */
function cardinalText(n: number): string {
  if (n < 1 || n > 999) return String(n);
  if (n < 20) return ENGLISH_ONES[n - 1]!;
  if (n < 100) {
    const t = ENGLISH_TENS[Math.floor(n / 10)]!;
    const r = n % 10;
    return r === 0 ? t : `${t}-${ENGLISH_ONES[r - 1]!}`;
  }
  const h = `${ENGLISH_ONES[Math.floor(n / 100) - 1]!} hundred`;
  const r = n % 100;
  return r === 0 ? h : `${h} ${cardinalText(r)}`;
}

/** 1st/2nd/3rd… — eleven–thirteen all take th (11th, 112th). */
function ordinal(n: number): string {
  if (n % 100 >= 11 && n % 100 <= 13) return `${n}th`;
  return `${n}${["th", "st", "nd", "rd"][n % 10 > 3 ? 0 : n % 10]}`;
}

/** The spelled-out ordinal's suffix rules (one → first, twenty → twentieth). */
function ordinalWord(word: string): string {
  const irregular: Record<string, string> = {
    one: "first",
    two: "second",
    three: "third",
    five: "fifth",
    eight: "eighth",
    nine: "ninth",
    twelve: "twelfth",
  };
  return irregular[word] ?? (word.endsWith("y") ? `${word.slice(0, -1)}ieth` : `${word}th`);
}

/** first/second/third… — only the final word ordinalizes (twenty-one →
 *  twenty-first, one hundred → one hundredth), keeping the compound's hyphen. */
function ordinalText(n: number): string {
  if (n < 1 || n > 999) return String(n);
  const card = cardinalText(n);
  const cut = Math.max(card.lastIndexOf(" "), card.lastIndexOf("-"));
  return cut < 0
    ? ordinalWord(card)
    : card.slice(0, cut) + card[cut] + ordinalWord(card.slice(cut + 1));
}

/** One counter value under its w:numFmt token — the shared list/page number
 *  renderer. Unsupported and out-of-table values render decimal. */
export function formatNumber(format: string | undefined, n: number): string {
  if (n < 0) return String(n);
  switch (format) {
    case undefined:
    case "decimal":
      return String(n);
    case "lowerLetter":
      return letterNumeral(n, false);
    case "upperLetter":
      return letterNumeral(n, true);
    case "lowerRoman":
      return romanNumeral(n, false);
    case "upperRoman":
      return romanNumeral(n, true);
    // ①–⑳ (U+2460–U+2473); Word stops the circled run at twenty, too.
    case "decimalEnclosedCircle":
      return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case "chineseCounting":
    case "chineseCountingThousand":
    case "japaneseCounting":
      return chineseNumeral(n);
    case "chineseLegalSimplified":
      return chineseNumeral(n, true);
    case "aiueo":
      return glyphRun(AIUEO, n);
    case "iroha":
      return glyphRun(IROHA, n);
    case "katakana":
      return glyphRun(KATAKANA, n);
    case "katakanaHalf":
      return glyphRun(KATAKANA_HALF, n);
    case "koreanCounting":
      return glyphRun(KOREAN_COUNTING, n);
    case "ganada":
      return glyphRun(GANADA, n);
    case "numberInDash":
      return `- ${n} -`;
    case "ordinal":
      return ordinal(n);
    case "cardinalText":
      return cardinalText(n);
    case "ordinalText":
      return ordinalText(n);
    case "hex":
      return n.toString(16).toUpperCase();
    default:
      return String(n);
  }
}
