/**
 * Slack の営業報告 → 顧客管理シート（J〜U 列）自動転記
 *
 * 仕組み
 *  1. 設定タブの Slack チャンネルから、親メッセージ（「19時広告面談入ります！」など）とスレッドの返信を読む。
 *  2. 「投稿者＝担当者」「投稿日」「親メッセージの時刻」で、対象タブの行と何回目の商談かをコードで特定する。
 *     （お客様の名前は Slack では本名、シートでは LINE の表示名なので、名前では照合しない）
 *  3. 担当者の報告文だけを Gemini に読ませ、着座・結果・次回日時などを JSON で受け取り、コードで検査する。
 *  4. 空欄のセルと、このツールが前に書いたセル（自動転記ログで判断）だけを書き換える。
 *     人が入力した値と食い違う場合や、迷う場合は書き込まずに「要確認」タブへ回す。
 *
 * 秘密情報（スクリプト プロパティ）: SLACK_BOT_TOKEN, GEMINI_API_KEY（任意: GEMINI_MODEL）
 * そのほかの設定は、スプレッドシートの「設定」タブで変更する。
 */

// ===================== 固定の定義 =====================

var APP = {
  VERSION: '1.0.0',
  TZ: 'Asia/Tokyo',
  HEADER_ROW: 3,
  FIRST_DATA_ROW: 4,
  NUM_COLS: 21, // A〜U
  SETTINGS_SHEET: '設定',
  REVIEW_SHEET: '要確認',
  LOG_SHEET: '自動転記ログ',
  PREVIEW_SHEET: '試運転結果',
  TIME_BUDGET_MS: 4.5 * 60 * 1000, // Apps Script の上限（6分）より前に止める
  RECHECK_RECENT_MS: 48 * 3600 * 1000, // 直近48時間に動きがあったスレッドは、編集に気づけるよう毎回読み直す
  MAX_FAILURES: 3,
  REASON_MAX_LEN: 400,
  LOG_KIND_AUTO: '自動転記',
  OWNER_KINDS: ['初回の一括転記', '自動転記'],
};

var COLS = 'ABCDEFGHIJKLMNOPQRSTU';
function colIdx(letter) { return COLS.indexOf(letter); }

// 行の見出し（3行目）。列がずれていたら何も書かずに止める。
var EXPECTED_HEADERS = {
  A: 'LINE ID', B: '表示名', G: 'アポ予定日', H: 'アポ時間', I: '担当者',
  J: '初回商談実施予定日', K: '初回商談実施時間', L: '初回商談着座', M: '初回商談結果',
  N: '2回目商談実施予定日', O: '2回目商談実施時間', P: '2回目商談結果',
  Q: '3回目商談実施予定日', R: '3回目商談実施時間', S: '3回目商談結果',
  T: 'プラン・金額', U: '失注・保留理由',
};

var MEETINGS = [
  { no: 1, date: 'J', time: 'K', attend: 'L', result: 'M', next: { date: 'N', time: 'O' } },
  { no: 2, date: 'N', time: 'O', attend: null, result: 'P', next: { date: 'Q', time: 'R' } },
  { no: 3, date: 'Q', time: 'R', attend: null, result: 'S', next: null },
];

var ATTEND_VALUES = ['着座', '飛び', 'リスケ', 'キャンセル'];
var RESULT_VALUES = ['成約', '成約言質', '保留', '失注', '飛び', 'キャンセル', '未実施'];
var CLOSED_RESULTS = ['成約', '成約言質'];
var NOT_HELD_RESULTS = ['飛び', 'キャンセル', '未実施'];
var DATE_COLS = ['J', 'N', 'Q'];
var TIME_COLS = ['K', 'O', 'R'];
var MEETING_POST_RE = /入ります|入りまーす|面談|アポ|クロ|商談|予定変更|日程変更|リスケ/;

// ===================== 純粋な関数（テスト対象） =====================

function normalizeText(s) {
  return String(s == null ? '' : s).normalize('NFKC').trim();
}

function normalizeName(s) {
  return normalizeText(s).replace(/[\s　]/g, '');
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

/** 貼り付けたトークンやキーの前後の空白・改行・引用符を取り除く。 */
function cleanSecret(s) {
  return String(s == null ? '' : s).replace(/^[\s"'“”「」]+|[\s"'“”「」]+$/g, '');
}

/** トークンの中身は出さずに、種類（先頭5文字）と長さだけを説明する。 */
function describeToken(token) {
  var t = String(token || '');
  var kind = /^xoxb-/.test(t) ? 'ボットのトークン（xoxb-）' : /^xoxp-/.test(t) ? 'ユーザーのトークン（xoxp-）'
    : 'xoxb- で始まらない値（先頭: ' + t.slice(0, 5) + '）';
  return kind + '、' + t.length + '文字';
}

/** "2026/11/24" や "2026-11-24" → "2026/11/24"。読めなければ ''。 */
function normDate(s) {
  var m = normalizeText(s).match(/(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})/);
  if (!m) return '';
  var y = +m[1], mo = +m[2], d = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '/' + pad2(mo) + '/' + pad2(d);
}

/** "9:00" "09:00" "15:30（キャンセル済み）" → "HH:MM"。読めなければ ''。 */
function normTime(s) {
  var m = normalizeText(s).match(/(\d{1,2}):(\d{2})/);
  if (!m) return '';
  var h = +m[1], mi = +m[2];
  if (h > 23 || mi > 59) return '';
  return pad2(h) + ':' + pad2(mi);
}

/** Slack の書式（メンション・リンク・エスケープ）を読みやすい文字に戻す。 */
function cleanSlackText(text) {
  return String(text || '')
    .replace(/<@[A-Z0-9]+\|([^>]+)>/g, '@$1')
    .replace(/<@([A-Z0-9]+)>/g, '@$1')
    .replace(/<!(channel|here|everyone)>/g, '')
    .replace(/<(https?:[^|>]+)\|([^>]+)>/g, '$2')
    .replace(/<(https?:[^>]+)>/g, '$1')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

/**
 * 親メッセージから商談の時刻を読む。「19時」「17時半」「10時30」「15:30」「19:15～」など。
 * メンションや URL は先に取り除く。見つからなければ ''。
 */
function parseParentTime(text) {
  var t = normalizeText(String(text || '').replace(/<[^>]*>/g, ' '));
  // 「時」の直後の分は「30」のような2桁か「5分」の形だけ読む（「21時 2回目」の 2 を分と読まないため）
  var re = /(\d{1,2})\s*:\s*(\d{2})|(\d{1,2})時(半|(\d{2})(?!\d)|(\d{1,2})分)?/g;
  var m;
  while ((m = re.exec(t)) !== null) {
    var h, mi;
    if (m[1] != null) { h = +m[1]; mi = +m[2]; }
    else { h = +m[3]; mi = m[4] === '半' ? 30 : (m[5] != null ? +m[5] : (m[6] != null ? +m[6] : 0)); }
    if (h >= 0 && h <= 23 && mi >= 0 && mi <= 59) return pad2(h) + ':' + pad2(mi);
  }
  return '';
}

/** Slack の ts（秒）→ 日本時間の "yyyy/MM/dd"。 */
function jstDateFromTs(ts) {
  var d = new Date(Math.floor(parseFloat(ts) * 1000) + 9 * 3600 * 1000);
  return d.getUTCFullYear() + '/' + pad2(d.getUTCMonth() + 1) + '/' + pad2(d.getUTCDate());
}

/** "yyyy/MM/dd" に日数を足す。 */
function addDays(dateStr, days) {
  var p = dateStr.split('/');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]) + days * 86400000);
  return d.getUTCFullYear() + '/' + pad2(d.getUTCMonth() + 1) + '/' + pad2(d.getUTCDate());
}

/** "yyyy/MM/dd" 同士の差（日）。b - a。 */
function daysBetween(a, b) {
  var pa = a.split('/'), pb = b.split('/');
  return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}

/** 親メッセージの商談日。投稿日（日本時間）。「明日」とあれば翌日。 */
function meetingDateForParent(ts, text) {
  var d = jstDateFromTs(ts);
  return /明日/.test(normalizeText(text)) ? addDays(d, 1) : d;
}

function rowKeyOf(tab, row) {
  return tab + '|' + normalizeText(row[colIdx('A')]) + '|' + normDate(row[colIdx('G')]);
}

/**
 * 担当者・日付・時刻に合う行と、何回目の商談かを探す。
 * tabs: { タブ名: [[A..U の表示値], ...] }
 * 戻り値: [{ tab, rowIndex, meetingNo }]
 */
function findCandidates(tabs, repName, date, time) {
  var rep = normalizeName(repName);
  var out = [];
  Object.keys(tabs).forEach(function (tab) {
    tabs[tab].forEach(function (row, i) {
      if (!rep || normalizeName(row[colIdx('I')]) !== rep) return;
      var slots = [
        { no: 1, d: row[colIdx('G')], t: row[colIdx('H')] },
        { no: 1, d: row[colIdx('J')], t: row[colIdx('K')] },
        { no: 2, d: row[colIdx('N')], t: row[colIdx('O')] },
        { no: 3, d: row[colIdx('Q')], t: row[colIdx('R')] },
      ];
      var found = {};
      slots.forEach(function (s) {
        if (normDate(s.d) === date && normTime(s.t) === time) found[s.no] = true;
      });
      Object.keys(found).forEach(function (no) {
        out.push({ tab: tab, rowIndex: i, meetingNo: +no });
      });
    });
  });
  return out;
}

/** 担当者として登場する名前の集合（正規化済み）。 */
function repSet(tabs) {
  var set = {};
  Object.keys(tabs).forEach(function (tab) {
    tabs[tab].forEach(function (row) {
      var n = normalizeName(row[colIdx('I')]);
      if (n) set[n] = true;
    });
  });
  return set;
}

function formatYen(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 報告のプラン名を「設定」タブの正式なプラン名にそろえる（略称・表記ゆれ対策）。
 * 完全一致がなければ、前方一致（略称）が1つだけのときにその名前を使う。決められなければ null。
 */
function resolvePlanName(name, plans) {
  var n = normalizeName(name);
  if (!n) return null;
  var names = Object.keys(plans);
  for (var i = 0; i < names.length; i++) if (normalizeName(names[i]) === n) return names[i];
  var hits = names.filter(function (p) {
    var q = normalizeName(p);
    return q.indexOf(n) === 0 || n.indexOf(q) === 0;
  });
  return hits.length === 1 ? hits[0] : null;
}

function isoToSheetDate(s) {
  var m = normalizeText(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return m ? normDate(m[1] + '/' + m[2] + '/' + m[3]) : '';
}

/**
 * Gemini の読み取り結果を検査し、書き込むセルの一覧にする。
 * ex: 読み取り結果, meetingNo: 1〜3, meetingDate: "yyyy/MM/dd", parentTime: "HH:MM", plans: {名前: 金額}
 * 戻り値: { updates: [{col, value}], notes: [文字列] }
 */
function buildUpdates(ex, meetingNo, meetingDate, parentTime, plans) {
  var res = { updates: [], notes: [] };
  if (!ex || ex.is_report !== true) return res;
  var m = MEETINGS[meetingNo - 1];
  if (!m) { res.notes.push('商談の回数が読み取れません'); return res; }

  var attendance = ATTEND_VALUES.indexOf(ex.attendance) >= 0 ? ex.attendance : null;
  var result = RESULT_VALUES.indexOf(ex.result) >= 0 ? ex.result : null;
  if (ex.attendance && !attendance) res.notes.push('着座の値「' + ex.attendance + '」はプルダウンにありません');
  if (ex.result && !result) res.notes.push('結果の値「' + ex.result + '」はプルダウンにありません');

  if (meetingNo === 1) {
    if (!attendance && result === '飛び') attendance = '飛び';
    if (!attendance && result === 'キャンセル') attendance = 'キャンセル';
    if (attendance === '飛び' && !result) result = '飛び';
    if (attendance === 'リスケ' && !result) result = '未実施';
    if (attendance === 'キャンセル' && !result) result = 'キャンセル';
    if (!attendance && result && NOT_HELD_RESULTS.indexOf(result) < 0) attendance = '着座';
  }
  if (!attendance && !result) return res;

  var actual = ex.actual_time ? normTime(ex.actual_time) : '';
  if (ex.actual_time && !actual) res.notes.push('実施時刻「' + ex.actual_time + '」が読み取れません');
  var time = actual || parentTime;

  // 親メッセージから取った日付・時刻は「空欄のときだけ書く」（実施時刻など、前に書いた値を上書きしないため）
  res.updates.push({ col: m.date, value: meetingDate, fillOnly: true });
  if (time) res.updates.push({ col: m.time, value: time, fillOnly: !actual });
  if (m.attend && attendance) res.updates.push({ col: m.attend, value: attendance });
  if (result) res.updates.push({ col: m.result, value: result });

  var held = meetingNo === 1 ? attendance === '着座' : (result && NOT_HELD_RESULTS.indexOf(result) < 0);
  if (ex.next_meeting_date) {
    var nd = isoToSheetDate(ex.next_meeting_date);
    var nt = ex.next_meeting_time ? normTime(ex.next_meeting_time) : '';
    if (!held) {
      res.notes.push('実施されていない商談に次回日時（' + ex.next_meeting_date + '）があります。再調整後の日時は手入力してください');
    } else if (!nd) {
      res.notes.push('次回の日付「' + ex.next_meeting_date + '」が読み取れません');
    } else if (daysBetween(meetingDate, nd) < 0 || daysBetween(meetingDate, nd) > 90) {
      res.notes.push('次回の日付 ' + nd + ' が商談日 ' + meetingDate + ' から見て不自然です');
    } else if (!m.next) {
      res.notes.push('4回目以降の予定（' + nd + (nt ? ' ' + nt : '') + '）は入れる列がありません');
    } else {
      res.updates.push({ col: m.next.date, value: nd });
      if (nt) res.updates.push({ col: m.next.time, value: nt });
    }
  }

  if (result && CLOSED_RESULTS.indexOf(result) >= 0 && ex.plan) {
    var planName = resolvePlanName(ex.plan, plans);
    if (!planName) {
      res.notes.push('プラン「' + ex.plan + '」が設定タブのどのプラン名とも一致しません');
    } else {
      var price = (ex.amount && ex.amount > 0) ? ex.amount : plans[planName];
      res.updates.push({ col: 'T', value: planName + ' ' + formatYen(price) });
    }
  }

  if (ex.reason) {
    var reason = normalizeText(ex.reason).replace(/\s+/g, ' ');
    if (reason.length > APP.REASON_MAX_LEN) reason = reason.slice(0, APP.REASON_MAX_LEN - 1) + '…';
    if (reason) res.updates.push({ col: 'U', value: reason });
  }
  return res;
}

/** セルの値を比較用にそろえる。 */
function comparable(col, v) {
  if (DATE_COLS.indexOf(col) >= 0) return normDate(v) || normalizeText(v);
  if (TIME_COLS.indexOf(col) >= 0) return normTime(v) || normalizeText(v);
  return normalizeText(v);
}

/**
 * 1セルの判断。current: 今の表示値, value: 書きたい値, owned: このツールが最後に書いた値（なければ null）
 * 'write' | 'same' | 'conflict'
 */
function decideCell(col, current, value, owned) {
  var cur = comparable(col, current), val = comparable(col, value);
  if (cur === '') return 'write';
  if (cur === val) return 'same';
  if (owned != null && comparable(col, owned) === cur) return 'write';
  return 'conflict';
}

/** Slack の投稿リンクから「チャンネル:ts」を取り出す。読めなければ ''。 */
function threadKeyFromLink(link) {
  var m = String(link || '').match(/\/archives\/([A-Z0-9]+)\/p(\d{10})(\d{6})/);
  return m ? m[1] + ':' + m[2] + '.' + m[3] : '';
}

/**
 * 自動転記ログの行から、「どのスレッドが、どの行の何回目の商談を書いたか」を作る。
 * 回数は結果・着座の列（L/M→1回目、P→2回目、S→3回目）から決める（次回日時の列は前の回のスレッドが書くため使わない）。
 * logRows: [[記録日時, タブ, LINE ID, アポ予定日, 列, 値, リンク, 処理], ...]
 */
function buildThreadMap(logRows) {
  var noByCol = { L: 1, M: 1, P: 2, S: 3 };
  var map = {};
  logRows.forEach(function (r) {
    if (APP.OWNER_KINDS.indexOf(normalizeText(r[7])) < 0) return;
    var no = noByCol[normalizeText(r[4])];
    var tk = threadKeyFromLink(r[6]);
    if (!no || !tk) return;
    map[tk] = { tab: normalizeText(r[1]), rowKey: normalizeText(r[1]) + '|' + normalizeText(r[2]) + '|' + normDate(r[3]), no: no };
  });
  return map;
}

/** スレッドの中身のハッシュ用の文字列（編集の検出に使う）。 */
function threadFingerprint(messages) {
  return messages.map(function (m) {
    return [m.ts, m.user || '', m.text || '', (m.edited && m.edited.ts) || ''].join('\u0001');
  }).join('\u0002');
}

/** 使えるモデル名から、混雑しにくい安定版 Flash を優先して並べる。 */
function rankModels(names) {
  var skip = /(tts|image|audio|live|embedding|aqa|imagen|veo|learnlm|gemma|robotics|computer-use|native|thinking)/i;
  var seen = {};
  var parsed = [];
  names.slice().sort(function (a, b) { return a.length - b.length; }).forEach(function (n) {
    if (skip.test(n)) return;
    var m = n.match(/^gemini-(\d+(?:\.\d+)?)-(flash-lite|flash|pro)(?:-(.+))?$/);
    if (!m) return;
    var suffix = m[3] || '';
    var preview = /preview|exp/i.test(suffix);
    if (suffix && !preview && !/^\d{3}$/.test(suffix)) return;
    var k = m[1] + '-' + m[2] + '-' + preview;
    if (seen[k]) return;
    seen[k] = true;
    parsed.push({ n: n, v: parseFloat(m[1]), tier: m[2], preview: preview });
  });
  var tierRank = { flash: 0, 'flash-lite': 1, pro: 2 };
  parsed.sort(function (a, b) {
    return (a.preview - b.preview) || (tierRank[a.tier] - tierRank[b.tier]) || (b.v - a.v);
  });
  var stable = parsed.filter(function (p) { return !p.preview; }).map(function (p) { return p.n; });
  var aliases = ['gemini-flash-latest', 'gemini-flash-lite-latest'].filter(function (a) { return names.indexOf(a) >= 0; });
  return stable.concat(aliases);
}

/** Gemini に渡す報告文（ユーザー側）を組み立てる。 */
function buildPrompt(p) {
  var lines = [
    '面談日: ' + p.meetingDate.replace(/\//g, '-'),
    '親メッセージの時刻: ' + p.parentTime,
    '何回目の商談か: ' + p.meetingNo + '回目',
    '担当者: ' + p.rep,
    'シート上のお客様の表示名（LINE の表示名。Slack では本名のことが多い）: ' + (p.displayName || '不明'),
  ];
  if (p.currentReason) lines.push('シートの今の理由欄: ' + p.currentReason);
  lines.push('');
  p.messages.forEach(function (m, i) {
    lines.push((i === 0 ? '[親メッセージ] ' : '[返信] ') + m.name + (m.isRep ? '（担当者）' : '') + ' ' + m.when);
    lines.push(m.text);
    lines.push('');
  });
  return lines.join('\n');
}

var SYSTEM_PROMPT = [
  'あなたは営業チームの Slack 報告を読み、顧客管理シートに転記する値を JSON で返すアシスタントです。',
  '1つのスレッドは、1人のお客様の1回の商談についての報告です。親メッセージは担当者が商談の前に投稿する「○時面談入ります」などの連絡で、結果はスレッドの返信に書かれます。',
  '',
  'ルール:',
  '- 結果は担当者本人の返信をもとに判断する。ほかの人の応援コメントは無視する。ただし、ほかの人の返信に事実（例: 再調整を依頼した）があれば参考にしてよい。',
  '- スレッドの中で状況が変わった場合（例: 「○日まで検討」のあとに「失注」）は、最新の状態を返す。',
  '- 言い方と値の対応:',
  '  ・未着座、来ない、入室されない → attendance「飛び」、result「飛び」',
  '  ・予定変更の依頼、日程変更 → attendance「リスケ」、result「未実施」',
  '  ・キャンセル → attendance「キャンセル」、result「キャンセル」',
  '  ・ブリッジ、次回（2回目、クロージング、クロ、戦略設計など）の日時が決まった → attendance「着座」、result「保留」、次回の日時を next_meeting_date と next_meeting_time に入れる',
  '  ・二回目調整など、次回の日時が未定 → attendance「着座」、result「保留」、next_meeting_date は null',
  '  ・「○日まで検討」 → result「保留」',
  '  ・失注、見送り → result「失注」',
  '  ・スクリーニング → result「失注」',
  '  ・言質（例: 「（プラン名）言質」） → result「成約言質」',
  '  ・成約、契約完了、決済完了 → result「成約」',
  '- 同じ日に時間を変えて実施した場合（例: 12時は未着座で13時に実施）は、実施した時刻を actual_time に入れ、実施した商談の結果を返す。',
  '- 次にお客様と会う（話す）予定の日時が書かれていれば、next_meeting_date と next_meeting_time に入れる。2回目・クロージングなどの商談に限らず、成約言質のあとのローン手続きや契約・決済の手続き（例: 「27日15:00〜」）も含める。',
  '- 日付は面談日を基準に年と月を補い、YYYY-MM-DD で返す（例: 面談日 2026-11-21 で「24日」→ 2026-11-24）。時刻は 24 時間制の HH:MM で返す。',
  '- plan は成約・成約言質のときだけ、報告にプラン名があれば入れる。選べるプラン名のどれかで返し、報告が略称や表記ゆれ（語尾の省略など）でも正式なプラン名に直す。amount は報告に総額（円）がはっきり書かれているときだけ整数で入れる。',
  '- reason は保留・失注・成約言質のときに、Slack の報告に書かれている内容をもとに詳しく書く。結果の理由に加えて、その背景（お客様の状況・仕事や収入・懸念点・検討状況・担当者が考える次の打ち手など）を、報告にある範囲で 150〜300 文字の文章にまとめる。報告に書かれていないことは推測で書かない。',
  '- 入力に「シートの今の理由欄」がある場合は、それを土台に今回の報告の内容を足して、更新した全文を reason に返す（今回の結果と合わなくなった部分は書き換える）。',
  '- まだ結果が報告されていない（「入ります」や待機中の連絡、相談だけ）なら is_report を false にする。',
  '- 1つのスレッドに複数のお客様の結果がある、どの結果か判断できない、など迷う場合は confidence を low にし、note に理由を書く。',
].join('\n');

function responseSchema(planNames) {
  var plan = { type: 'STRING', nullable: true };
  if (planNames.length) plan['enum'] = planNames;
  return {
    type: 'OBJECT',
    properties: {
      is_report: { type: 'BOOLEAN' },
      customer_name: { type: 'STRING', nullable: true },
      attendance: { type: 'STRING', nullable: true, 'enum': ATTEND_VALUES },
      result: { type: 'STRING', nullable: true, 'enum': RESULT_VALUES },
      actual_time: { type: 'STRING', nullable: true },
      next_meeting_date: { type: 'STRING', nullable: true },
      next_meeting_time: { type: 'STRING', nullable: true },
      plan: plan,
      amount: { type: 'INTEGER', nullable: true },
      reason: { type: 'STRING', nullable: true },
      confidence: { type: 'STRING', 'enum': ['high', 'low'] },
      note: { type: 'STRING', nullable: true },
    },
    required: ['is_report', 'confidence'],
    propertyOrdering: ['is_report', 'customer_name', 'attendance', 'result', 'actual_time',
      'next_meeting_date', 'next_meeting_time', 'plan', 'amount', 'reason', 'confidence', 'note'],
  };
}

// ===================== メニュー =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('自動転記')
    .addItem('接続テスト', 'menuCheckSetup')
    .addItem('試運転（シートに書き込まない）', 'menuDryRun')
    .addSeparator()
    .addItem('今すぐ実行', 'menuRunNow')
    .addItem('定期実行を開始（30分ごと）', 'menuStartSchedule')
    .addItem('定期実行を止める', 'menuStopSchedule')
    .addToUi();
}

function menuCheckSetup() {
  var lines = checkSetup();
  SpreadsheetApp.getUi().alert('接続テスト', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuDryRun() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = runSync({ dryRun: true });
    ui.alert('試運転が終わりました',
      '「' + APP.PREVIEW_SHEET + '」タブに、書き込む予定の内容を出しました（シートは変更していません）。\n' + r.summary, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('試運転できませんでした', String(e.message || e), ui.ButtonSet.OK);
  }
}

function menuRunNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = runSync({ dryRun: false });
    ui.alert('実行しました', r.summary, ui.ButtonSet.OK);
  } catch (e) {
    if (isConfigError_(e)) reportFatal_(e);
    ui.alert('実行できませんでした', String(e.message || e), ui.ButtonSet.OK);
  }
}

function menuStartSchedule() {
  stopSchedule_();
  ScriptApp.newTrigger('scheduledRun').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert('定期実行を開始しました（30分ごと）。止めるときはメニューの「定期実行を止める」を使ってください。');
}

function menuStopSchedule() {
  var n = stopSchedule_();
  SpreadsheetApp.getUi().alert(n ? '定期実行を止めました。' : '定期実行は設定されていませんでした。');
}

function stopSchedule_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledRun') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/** 定期実行の入口。設定や接続の問題は例外にして、Google からの失敗通知メールで気づけるようにする。 */
function scheduledRun() {
  try {
    runSync({ dryRun: false });
  } catch (e) {
    if (isConfigError_(e)) reportFatal_(e);
    throw e;
  }
}

// ===================== 実行本体 =====================

function ConfigError(message) {
  this.name = 'ConfigError';
  this.message = message;
  this.stack = (new Error(message)).stack;
}
ConfigError.prototype = Object.create(Error.prototype);

function TimeUpError() {
  this.name = 'TimeUpError';
  this.message = '時間切れ';
}
TimeUpError.prototype = Object.create(Error.prototype);

function isConfigError_(e) { return !!e && (e instanceof ConfigError || e.name === 'ConfigError'); }

/** 設定や接続の問題を「要確認」タブにも残す（同じ内容は1日1回まで）。 */
function reportFatal_(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(APP.REVIEW_SHEET);
    if (!sh) return;
    var key = 'config:' + jstDateFromTs(Date.now() / 1000) + ':' + String(e.message).slice(0, 80);
    if (readReviewKeys_(sh)[key]) return;
    var row = lastFilledRow_(sh) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    sh.getRange(row, 1, 1, 9).setValues([[nowText_(), '', '', '', '自動転記が止まっています', String(e.message), '', false, key]]);
    sh.getRange(row, 8).insertCheckboxes();
  } catch (e2) {
    console.error('要確認タブへの記録にも失敗: ' + e2);
  }
}

function runSync(opts) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { summary: 'ほかの実行が動いているため、今回は何もしませんでした。' };
  try {
    var ctx = loadContext_(opts.dryRun);
    processThreads_(ctx);
    checkMissingReports_(ctx);
    if (ctx.dryRun) writePreview_(ctx);
    else pruneState_(ctx);
    var s = ctx.stats;
    var summary = 'スレッド ' + s.threads + ' 件を確認、セル ' + s.writes + ' 件を' + (ctx.dryRun ? '書き込む予定' : '書き込み') +
      '、要確認 ' + s.reviews + ' 件' + (s.timeUp ? '（時間切れのため残りは次回）' : '') + (s.errors ? '、読み取り失敗 ' + s.errors + ' 件' : '') + '。';
    console.log(summary);
    return { summary: summary };
  } finally {
    lock.releaseLock();
  }
}

function loadContext_(dryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var token = cleanSecret(props.getProperty('SLACK_BOT_TOKEN'));
  var geminiKey = cleanSecret(props.getProperty('GEMINI_API_KEY'));
  if (!token) throw new ConfigError('スクリプト プロパティに SLACK_BOT_TOKEN がありません');
  if (!geminiKey) throw new ConfigError('スクリプト プロパティに GEMINI_API_KEY がありません');

  var settings = readSettings_(ss);
  var tabs = {}, sheets = {};
  settings.tabs.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) throw new ConfigError('設定タブの「対象タブ」にある「' + name + '」が見つかりません');
    verifyHeaders_(sh, name);
    sheets[name] = sh;
    var last = sh.getLastRow();
    tabs[name] = last >= APP.FIRST_DATA_ROW
      ? sh.getRange(APP.FIRST_DATA_ROW, 1, last - APP.FIRST_DATA_ROW + 1, APP.NUM_COLS).getDisplayValues()
      : [];
  });

  var logSheet = ensureSheet_(ss, APP.LOG_SHEET, ['記録日時', 'タブ', 'LINE ID', 'アポ予定日', '列', '書き込んだ値', '元のSlack投稿', '処理']);
  var reviewSheet = ensureSheet_(ss, APP.REVIEW_SHEET, ['登録日時', 'タブ', '担当者', '面談日時', '内容', '理由', '元のSlack投稿', '対応済み', 'キー']);
  var logRows = readLogRows_(logSheet);

  return {
    ss: ss, props: props, token: token, geminiKey: geminiKey, geminiModel: props.getProperty('GEMINI_MODEL') || '',
    dryRun: !!dryRun, settings: settings, tabs: tabs, sheets: sheets, reps: repSet(tabs),
    logSheet: logSheet, reviewSheet: reviewSheet,
    owned: buildOwnership(logRows), threadMap: buildThreadMap(logRows), reviewKeys: readReviewKeys_(reviewSheet),
    users: {}, models: null, started: Date.now(),
    preview: [], stats: { threads: 0, writes: 0, reviews: 0, errors: 0, timeUp: false },
    today: jstDateFromTs(Date.now() / 1000),
  };
}

function readSettings_(ss) {
  var sh = ss.getSheetByName(APP.SETTINGS_SHEET);
  if (!sh) throw new ConfigError('「' + APP.SETTINGS_SHEET + '」タブがありません');
  var values = sh.getDataRange().getValues();
  function get(label) {
    for (var i = 0; i < values.length; i++) if (normalizeText(values[i][0]) === label) return values[i][1];
    return '';
  }
  var channels = normalizeText(get('対象チャンネル')).split(/[,、\s]+/).filter(Boolean);
  if (!channels.length) throw new ConfigError('設定タブの「対象チャンネル」が空です');
  channels.forEach(function (c) {
    if (!/^[CG][A-Z0-9]{6,}$/.test(c)) throw new ConfigError('「対象チャンネル」の「' + c + '」はチャンネル ID の形ではありません');
  });
  var tabs = normalizeText(get('対象タブ')).split(/[,、]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  if (!tabs.length) throw new ConfigError('設定タブの「対象タブ」が空です');

  var after = get('処理開始日時');
  var afterMs;
  if (after instanceof Date) afterMs = after.getTime();
  else {
    var d = normDate(after), t = normTime(after) || '00:00';
    if (!d) throw new ConfigError('設定タブの「処理開始日時」が日時として読めません');
    afterMs = new Date(d.replace(/\//g, '-') + 'T' + t + ':00+09:00').getTime();
  }
  var lookbackDays = Number(get('読み込む期間（日）')) || 60;
  var missingDays = Number(get('報告なしとみなす日数')) || 2;

  var plans = {};
  var start = -1;
  for (var i = 0; i < values.length; i++) {
    if (normalizeText(values[i][0]) === 'プラン' && normalizeText(values[i][1]) === '金額') { start = i + 1; break; }
  }
  if (start > 0) {
    for (var j = start; j < values.length; j++) {
      var name = normalizeText(values[j][0]);
      if (!name) break;
      var price = Number(String(values[j][1]).replace(/[,円\s]/g, ''));
      if (!(price > 0)) throw new ConfigError('設定タブのプラン「' + name + '」の金額が数字ではありません');
      plans[name] = price;
    }
  }
  return { channels: channels, tabs: tabs, processAfterMs: afterMs, lookbackDays: lookbackDays, missingDays: missingDays, plans: plans };
}

function verifyHeaders_(sh, name) {
  var header = sh.getRange(APP.HEADER_ROW, 1, 1, APP.NUM_COLS).getDisplayValues()[0];
  Object.keys(EXPECTED_HEADERS).forEach(function (col) {
    var got = normalizeName(header[colIdx(col)]);
    if (got !== normalizeName(EXPECTED_HEADERS[col])) {
      throw new ConfigError('「' + name + '」の ' + col + ' 列の見出しが「' + EXPECTED_HEADERS[col] + '」ではありません（今は「' + header[colIdx(col)] + '」）。列がずれている可能性があるため止めました');
    }
  });
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function readLogRows_(logSheet) {
  var last = logSheet.getLastRow();
  return last < 2 ? [] : logSheet.getRange(2, 1, last - 1, 8).getDisplayValues();
}

/** 自動転記ログから「このツールが最後に書いた値」を読む。キー: 行キー|列 */
function buildOwnership(logRows) {
  var owned = {};
  logRows.forEach(function (r) {
    if (APP.OWNER_KINDS.indexOf(normalizeText(r[7])) < 0) return;
    var key = normalizeText(r[1]) + '|' + normalizeText(r[2]) + '|' + normDate(r[3]) + '|' + normalizeText(r[4]);
    owned[key] = r[5];
  });
  return owned;
}

function readReviewKeys_(reviewSheet) {
  var keys = {};
  var last = lastFilledRow_(reviewSheet);
  if (last < 2) return keys;
  reviewSheet.getRange(2, 9, last - 1, 1).getDisplayValues().forEach(function (r) {
    if (r[0]) keys[r[0]] = true;
  });
  return keys;
}

/** A 列で最後に値がある行（チェックボックス列があると getLastRow が使えないため）。 */
function lastFilledRow_(sh) {
  var max = sh.getMaxRows();
  if (max < 2) return 1;
  var vals = sh.getRange(1, 1, max, 1).getDisplayValues();
  for (var i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '') return i + 1;
  return 1;
}

function timeLeft_(ctx) {
  return Date.now() - ctx.started < APP.TIME_BUDGET_MS;
}

function stateKey_(channel, ts) { return 't:' + channel + ':' + ts; }

function readState_(ctx, key) {
  var raw = ctx.props.getProperty(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveState_(ctx, key, state) {
  if (ctx.dryRun) return;
  ctx.props.setProperty(key, JSON.stringify(state));
}

function processThreads_(ctx) {
  var nowSec = Date.now() / 1000;
  var oldest = nowSec - ctx.settings.lookbackDays * 86400;
  var parents = [];
  ctx.settings.channels.forEach(function (ch) {
    slackHistory_(ctx, ch, oldest).forEach(function (m) { m._channel = ch; parents.push(m); });
  });
  parents.sort(function (a, b) { return parseFloat(a.ts) - parseFloat(b.ts); });

  for (var i = 0; i < parents.length; i++) {
    if (!timeLeft_(ctx)) { ctx.stats.timeUp = true; break; }
    var p = parents[i];
    if (p.subtype || p.bot_id || !p.user) continue;
    if (p.thread_ts && p.thread_ts !== p.ts) continue; // チャンネルにも送られた返信
    var activity = Math.max(parseFloat(p.ts), parseFloat(p.latest_reply || p.ts));
    if (!ctx.dryRun && activity * 1000 <= ctx.settings.processAfterMs) continue;

    var key = stateKey_(p._channel, p.ts);
    var state = readState_(ctx, key) || {};
    var lr = p.latest_reply || p.ts;
    var recent = (nowSec - activity) * 1000 < APP.RECHECK_RECENT_MS;
    if (!ctx.dryRun && state.lr === lr && !recent && !state.f) continue;

    ctx.beforeStart = activity * 1000 <= ctx.settings.processAfterMs;
    try {
      ctx.stats.threads++;
      handleThread_(ctx, p, key, state, lr);
    } catch (e) {
      if (isConfigError_(e)) throw e;
      if (e && e.name === 'TimeUpError') { ctx.stats.timeUp = true; break; }
      ctx.stats.errors++;
      console.error('スレッド ' + key + ' の処理に失敗: ' + (e && e.stack || e));
      state.f = (state.f || 0) + 1;
      if (state.f >= APP.MAX_FAILURES) {
        addReview_(ctx, {
          tab: '', rep: '', when: '', content: cleanSlackText(p.text).slice(0, 80),
          reason: '読み取りに ' + state.f + ' 回失敗しました（' + String(e && e.message || e).slice(0, 120) + '）',
          link: permalink_(ctx, p._channel, p.ts), key: 'thread:' + p._channel + ':' + p.ts + ':failed',
        });
        state.f = 0;
        state.lr = lr;
      }
      saveState_(ctx, key, state);
    }
  }
}

function handleThread_(ctx, p, key, state, lr) {
  var channel = p._channel;
  var rep = userName_(ctx, p.user);
  var parentText = cleanSlackText(p.text);
  if (!ctx.reps[normalizeName(rep)]) {
    // 商談の連絡に見えるのに、投稿者の名前がどのタブの担当者列にもない（名前の表記ちがいの見逃しを防ぐ）
    if (MEETING_POST_RE.test(parentText) && parseParentTime(p.text) && p.reply_count) {
      addReview_(ctx, {
        tab: '', rep: rep, when: jstWhen_(p.ts), content: parentText.slice(0, 80),
        reason: 'Slack の名前「' + rep + '」が、対象タブの担当者列（I 列）にありません。表記が違う場合はそろえてください',
        link: permalinkLazy_(ctx, channel, p.ts), key: 'user:' + p.user + ':not_rep',
      });
    }
    state.lr = lr;
    saveState_(ctx, key, state);
    return;
  }

  var messages = p.reply_count ? slackReplies_(ctx, channel, p.ts) : [p];
  var fp = threadFingerprint(messages);
  var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, fp, Utilities.Charset.UTF_8));
  if (!ctx.dryRun && state.h === hash && !state.f) { state.lr = lr; saveState_(ctx, key, state); return; }

  var replies = messages.filter(function (m) { return m.ts !== p.ts; });
  var repReplies = replies.filter(function (m) { return m.user === p.user; });
  var isMeetingPost = MEETING_POST_RE.test(parentText);
  if (!isMeetingPost && !repReplies.length) { finishState_(ctx, key, state, lr, hash); return; }

  var parentTime = parseParentTime(p.text);
  var meetingDate = meetingDateForParent(p.ts, parentText);
  var link = permalinkLazy_(ctx, channel, p.ts);
  var base = { rep: rep, when: meetingDate + (parentTime ? ' ' + parentTime : ''), content: parentText.slice(0, 80), link: link };

  if (!parentTime) {
    if (repReplies.length) addReview_(ctx, merge_(base, { tab: '', reason: '親メッセージから商談の時刻が読み取れません', key: 'thread:' + channel + ':' + p.ts + ':no_time' }));
    finishState_(ctx, key, state, lr, hash);
    return;
  }

  // 行の特定：前にこのスレッドから書いた行があればそれを使う（実施時刻を書き換えた後でも同じ行を指すため）。
  // 処理の記録（state.map）か、自動転記ログに残っているリンク（一括転記の分を含む）から探す。
  var target = null;
  var known = state.map || ctx.threadMap[channel + ':' + p.ts];
  if (known) {
    var rows = ctx.tabs[known.tab] || [];
    for (var i = 0; i < rows.length; i++) {
      if (rowKeyOf(known.tab, rows[i]) === known.rowKey) { target = { tab: known.tab, rowIndex: i, meetingNo: known.no }; break; }
    }
  }
  if (!target) {
    var cands = findCandidates(ctx.tabs, rep, meetingDate, parentTime);
    if (cands.length !== 1) {
      if (repReplies.length) {
        addReview_(ctx, merge_(base, {
          tab: cands.length ? cands.map(function (c) { return c.tab; }).join(', ') : '',
          reason: cands.length ? '担当者と日時に合う予定が ' + cands.length + ' 件あり、どれか決められません' : '担当者と日時に合う予定が、対象タブのどこにもありません',
          key: 'thread:' + channel + ':' + p.ts + ':' + (cands.length ? 'multi_match' : 'no_match'),
        }));
      }
      finishState_(ctx, key, state, lr, hash);
      return;
    }
    target = cands[0];
  }
  var row = ctx.tabs[target.tab][target.rowIndex];
  var rowKey = rowKeyOf(target.tab, row);
  state.map = { tab: target.tab, rowKey: rowKey, no: target.meetingNo };
  base.tab = target.tab;
  base.content = '【' + row[colIdx('B')] + ' ' + target.meetingNo + '回目】' + parentText.slice(0, 60);

  if (!repReplies.length) { finishState_(ctx, key, state, lr, hash); return; } // まだ報告がない

  var prompt = buildPrompt({
    meetingDate: meetingDate, parentTime: parentTime, meetingNo: target.meetingNo, rep: rep,
    displayName: row[colIdx('B')],
    currentReason: normalizeText(row[colIdx('U')]),
    messages: messages.map(function (m) {
      return { name: userName_(ctx, m.user), isRep: m.user === p.user, when: jstWhen_(m.ts), text: cleanSlackText(m.text) };
    }),
  });
  var ex = geminiExtract_(ctx, prompt);

  if (ex.is_report && ex.confidence === 'low') {
    addReview_(ctx, merge_(base, { reason: 'AI が判断に迷いました: ' + (ex.note || '理由なし'), key: 'thread:' + channel + ':' + p.ts + ':low:' + hash.slice(0, 8) }));
    finishState_(ctx, key, state, lr, hash);
    return;
  }

  var built = buildUpdates(ex, target.meetingNo, meetingDate, parentTime, ctx.settings.plans);
  built.notes.forEach(function (n, k) {
    addReview_(ctx, merge_(base, { reason: n, key: 'thread:' + channel + ':' + p.ts + ':note:' + hash.slice(0, 8) + ':' + k }));
  });
  applyUpdates_(ctx, target, row, rowKey, built.updates, base, channel, p.ts);
  finishState_(ctx, key, state, lr, hash);
}

function finishState_(ctx, key, state, lr, hash) {
  state.lr = lr;
  state.h = hash;
  state.f = 0;
  saveState_(ctx, key, state);
}

function applyUpdates_(ctx, target, row, rowKey, updates, base, channel, ts) {
  var sh = ctx.sheets[target.tab];
  var sheetRow = APP.FIRST_DATA_ROW + target.rowIndex;
  var expectedId = normalizeText(row[colIdx('A')]);
  var logRows = [];
  var lineId = expectedId, apo = normDate(row[colIdx('G')]);
  updates.forEach(function (u) {
    var ci = colIdx(u.col);
    var current = row[ci];
    var ownKey = rowKey + '|' + u.col;
    var owned = ctx.owned.hasOwnProperty(ownKey) ? ctx.owned[ownKey] : null;
    var decision = u.fillOnly && normalizeText(current) !== '' ? 'same' : decideCell(u.col, current, u.value, owned);
    if (ctx.dryRun) {
      var label = decision === 'write' ? '書き込み予定' : decision === 'same' ? '同じ値（変更なし）'
        : u.col === 'U' ? '人が書いた理由があるため変更しない' : '要確認（人が入力した値と違う）';
      if (ctx.beforeStart) label += '（参考: 処理開始日時より前のスレッドのため、本番では処理しない）';
      ctx.preview.push([target.tab, sheetRow, row[colIdx('B')], target.meetingNo + '回目', u.col, current, u.value, label, base.link()]);
      if (decision === 'write') { row[ci] = u.value; ctx.stats.writes++; }
      return;
    }
    if (decision === 'same') return;
    if (decision === 'conflict' && u.col === 'U') return; // 人が書いた理由は尊重し、要確認にも出さない
    if (decision === 'conflict') {
      addReview_(ctx, merge_(base, {
        reason: u.col + ' 列: シートは「' + current + '」、報告からは「' + u.value + '」。人が入力した値なので書き換えていません',
        key: 'thread:' + channel + ':' + ts + ':conflict:' + u.col + ':' + u.value,
      }));
      return;
    }
    // 書き込む直前に、行がずれていないか LINE ID で確かめる
    var idNow = normalizeText(sh.getRange(sheetRow, 1).getDisplayValue());
    if (idNow !== expectedId) {
      throw new ConfigError('「' + target.tab + '」' + sheetRow + '行目の LINE ID が読み込み時（' + expectedId + '）と今（' + idNow + '）で違います。行がずれた可能性があるため止めました');
    }
    var v = /^[=+\-@]/.test(u.value) ? "'" + u.value : u.value;
    sh.getRange(sheetRow, ci + 1).setValue(v);
    row[ci] = u.value;
    ctx.owned[ownKey] = u.value;
    ctx.stats.writes++;
    logRows.push([nowText_(), target.tab, "'" + lineId, "'" + apo, u.col, "'" + u.value, base.link(), APP.LOG_KIND_AUTO]);
  });
  if (logRows.length) {
    var ls = ctx.logSheet;
    var start = ls.getLastRow() + 1;
    if (start + logRows.length - 1 > ls.getMaxRows()) ls.insertRowsAfter(ls.getMaxRows(), logRows.length + 500);
    ls.getRange(start, 1, logRows.length, 8).setValues(logRows);
    SpreadsheetApp.flush();
  }
}

function checkMissingReports_(ctx) {
  var today = ctx.today;
  var floor = addDays(today, -ctx.settings.lookbackDays);
  Object.keys(ctx.tabs).forEach(function (tab) {
    ctx.tabs[tab].forEach(function (row, i) {
      var lineId = normalizeText(row[colIdx('A')]);
      if (!lineId) return;
      var apo = normDate(row[colIdx('G')]);
      MEETINGS.forEach(function (m) {
        var d = normDate(row[colIdx(m.date)]) || (m.no === 1 ? apo : '');
        var t = normTime(row[colIdx(m.time)]) || (m.no === 1 ? normTime(row[colIdx('H')]) : '');
        if (!d || d < floor) return;
        if (m.no === 1 && /キャンセル/.test(row[colIdx('H')])) return;
        if (normalizeText(row[colIdx(m.result)])) return;
        if (daysBetween(d, today) < ctx.settings.missingDays) return;
        addReview_(ctx, {
          tab: tab, rep: row[colIdx('I')], when: d + (t ? ' ' + t : ''),
          content: row[colIdx('B')] + 'さんの' + m.no + '回目商談', reason: '予定日を過ぎたが、Slack に結果の報告が見当たらない',
          link: '', key: 'missing:' + tab + ':' + lineId + ':' + apo + ':' + m.no,
        });
      });
    });
  });
}

function addReview_(ctx, r) {
  if (ctx.reviewKeys[r.key]) return;
  ctx.reviewKeys[r.key] = true;
  ctx.stats.reviews++;
  var link = typeof r.link === 'function' ? r.link() : (r.link || '');
  if (ctx.dryRun) {
    ctx.preview.push([r.tab || '', '', '', '', '', '', '', '要確認: ' + r.reason + (r.content ? '（' + r.content + '）' : ''), link]);
    return;
  }
  var sh = ctx.reviewSheet;
  var row = lastFilledRow_(sh) + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
  sh.getRange(row, 1, 1, 9).setValues([[nowText_(), r.tab || '', r.rep || '', "'" + (r.when || ''), r.content || '', r.reason, link, false, r.key]]);
  sh.getRange(row, 8).insertCheckboxes();
}

function writePreview_(ctx) {
  var sh = ctx.ss.getSheetByName(APP.PREVIEW_SHEET) || ctx.ss.insertSheet(APP.PREVIEW_SHEET);
  sh.clear();
  var header = ['タブ', '行', '表示名', '商談', '列', '今の値', '報告からの値', '判定', '元のSlack投稿'];
  var rows = [header].concat(ctx.preview.length ? ctx.preview : [['（書き込む予定の内容はありません）', '', '', '', '', '', '', '', '']]);
  sh.getRange(1, 1, rows.length, header.length).setValues(rows.map(function (r) {
    return r.map(function (v) { return typeof v === 'string' && /^\d/.test(v) ? "'" + v : v; });
  }));
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** 読み込み期間より古いスレッドの記録を消して、スクリプト プロパティが膨らまないようにする。 */
function pruneState_(ctx) {
  var limit = Date.now() / 1000 - (ctx.settings.lookbackDays + 30) * 86400;
  var all = ctx.props.getProperties();
  Object.keys(all).forEach(function (k) {
    var m = k.match(/^t:[^:]+:(\d+\.\d+)$/);
    if (m && parseFloat(m[1]) < limit) ctx.props.deleteProperty(k);
  });
}

function merge_(a, b) {
  var o = {};
  Object.keys(a).forEach(function (k) { o[k] = a[k]; });
  Object.keys(b).forEach(function (k) { o[k] = b[k]; });
  return o;
}

function nowText_() { return Utilities.formatDate(new Date(), APP.TZ, 'yyyy/MM/dd HH:mm:ss'); }
function jstWhen_(ts) { return Utilities.formatDate(new Date(parseFloat(ts) * 1000), APP.TZ, 'M/d HH:mm'); }

// ===================== Slack =====================

function slackApi_(ctx, method, params) {
  var q = Object.keys(params).filter(function (k) { return params[k] != null && params[k] !== ''; })
    .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
  var url = 'https://slack.com/api/' + method + (q ? '?' + q : '');
  for (var attempt = 0; attempt < 5; attempt++) {
    var res;
    try {
      res = UrlFetchApp.fetch(url, { method: 'get', headers: { Authorization: 'Bearer ' + ctx.token }, muteHttpExceptions: true });
    } catch (e) {
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }
    var code = res.getResponseCode();
    if (code === 429) {
      var h = res.getHeaders();
      var wait = Number(h['Retry-After'] || h['retry-after'] || 5);
      Utilities.sleep(Math.min(Math.max(wait, 1), 30) * 1000);
      continue;
    }
    if (code >= 500) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    var body = JSON.parse(res.getContentText());
    if (body.ok) return body;
    var err = body.error || 'unknown_error';
    if (/invalid_auth|not_authed|account_inactive|token_revoked|token_expired/.test(err)) {
      throw new ConfigError('Slack のトークンが使えません（' + err + '）。登録されているのは ' + describeToken(ctx.token) +
        '。Slack API の「Install App」にある Bot User OAuth Token を「Copy」して、SLACK_BOT_TOKEN に貼り直してください');
    }
    if (/missing_scope/.test(err)) {
      throw new ConfigError('Slack アプリの権限が足りません（' + method + ': ' + (body.needed || '') + '）。README の権限を追加して再インストールしてください');
    }
    if (/channel_not_found|not_in_channel/.test(err)) {
      throw new ConfigError('Slack アプリがチャンネル ' + params.channel + ' を読めません（' + err + '）。チャンネルにアプリを招待してください');
    }
    throw new Error('Slack ' + method + ' のエラー: ' + err);
  }
  throw new Error('Slack ' + method + ' に何度か失敗しました');
}

function slackHistory_(ctx, channel, oldest) {
  var out = [], cursor = '';
  do {
    var body = slackApi_(ctx, 'conversations.history', { channel: channel, oldest: String(oldest), limit: 200, cursor: cursor });
    out = out.concat(body.messages || []);
    cursor = body.response_metadata && body.response_metadata.next_cursor || '';
  } while (cursor);
  return out;
}

function slackReplies_(ctx, channel, ts) {
  var out = [], cursor = '';
  do {
    var body = slackApi_(ctx, 'conversations.replies', { channel: channel, ts: ts, limit: 200, cursor: cursor });
    out = out.concat(body.messages || []);
    cursor = body.response_metadata && body.response_metadata.next_cursor || '';
  } while (cursor);
  out.sort(function (a, b) { return parseFloat(a.ts) - parseFloat(b.ts); });
  return out;
}

function userName_(ctx, userId) {
  if (!userId) return '';
  if (ctx.users[userId]) return ctx.users[userId];
  var cache = CacheService.getScriptCache();
  var cached = cache.get('u:' + userId);
  if (cached) return (ctx.users[userId] = cached);
  var name = userId;
  try {
    var u = slackApi_(ctx, 'users.info', { user: userId }).user || {};
    var pr = u.profile || {};
    name = pr.real_name || u.real_name || pr.display_name || u.name || userId;
  } catch (e) {
    if (e instanceof ConfigError || e.name === 'ConfigError') throw e;
  }
  cache.put('u:' + userId, name, 21600);
  return (ctx.users[userId] = name);
}

function permalink_(ctx, channel, ts) {
  try {
    return slackApi_(ctx, 'chat.getPermalink', { channel: channel, message_ts: ts }).permalink || '';
  } catch (e) {
    if (e instanceof ConfigError || e.name === 'ConfigError') throw e;
    return '';
  }
}

/** 必要になったときだけリンクを取りに行く（API 呼び出しを減らすため）。 */
function permalinkLazy_(ctx, channel, ts) {
  var v = null;
  return function () { if (v === null) v = permalink_(ctx, channel, ts); return v; };
}

// ===================== Gemini =====================

var GEMINI_ROOT = 'https://generativelanguage.googleapis.com/v1beta/';
var STATIC_MODELS = ['gemini-2.5-flash', 'gemini-flash-latest', 'gemini-flash-lite-latest'];

function geminiModels_(ctx) {
  if (ctx.models) return ctx.models;
  var list = null;
  var cache = CacheService.getScriptCache();
  var cached = cache.get('gemini_models');
  if (cached) list = JSON.parse(cached);
  if (!list) {
    try {
      var res = UrlFetchApp.fetch(GEMINI_ROOT + 'models?pageSize=1000', { headers: { 'x-goog-api-key': ctx.geminiKey }, muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        var names = (JSON.parse(res.getContentText()).models || [])
          .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
          .map(function (m) { return String(m.name || '').replace(/^models\//, ''); });
        list = rankModels(names);
        if (list.length) cache.put('gemini_models', JSON.stringify(list), 21600);
      } else if (res.getResponseCode() === 400 || res.getResponseCode() === 403) {
        throw new ConfigError('Gemini の API キーが使えません（' + res.getResponseCode() + '）。GEMINI_API_KEY を確認してください');
      }
    } catch (e) {
      if (e instanceof ConfigError || e.name === 'ConfigError') throw e;
    }
  }
  if (!list || !list.length) list = STATIC_MODELS.slice();
  if (ctx.geminiModel) list = [ctx.geminiModel].concat(list.filter(function (m) { return m !== ctx.geminiModel; }));
  ctx.models = list.slice(0, 4);
  return ctx.models;
}

function geminiErrorInfo_(text) {
  var info = { message: '', retryMs: null, perDay: false };
  try {
    var e = (JSON.parse(text).error) || {};
    info.message = e.message || '';
    (e.details || []).forEach(function (d) {
      var type = String(d['@type'] || '');
      if (/RetryInfo/.test(type) && d.retryDelay) {
        var s = parseFloat(String(d.retryDelay));
        if (!isNaN(s)) info.retryMs = s * 1000;
      }
      if (/QuotaFailure/.test(type)) {
        (d.violations || []).forEach(function (v) { if (/PerDay/i.test(String(v.quotaId || ''))) info.perDay = true; });
      }
    });
  } catch (e2) { /* 本文が JSON でない */ }
  if (/per day|daily/i.test(info.message)) info.perDay = true;
  return info;
}

/** 報告文を Gemini に読ませて JSON を返す。混雑・上限のときは待つかモデルを替え、それでもだめなら例外（次回に再挑戦）。 */
function geminiExtract_(ctx, prompt) {
  var payload = JSON.stringify({
    systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: responseSchema(Object.keys(ctx.settings.plans)) },
  });
  var models = geminiModels_(ctx);
  var lastErr = '';
  for (var mi = 0; mi < models.length; mi++) {
    var model = models[mi];
    for (var attempt = 0; attempt < 3; attempt++) {
      if (!timeLeft_(ctx)) throw new TimeUpError();
      var res;
      try {
        res = UrlFetchApp.fetch(GEMINI_ROOT + 'models/' + model + ':generateContent', {
          method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': ctx.geminiKey }, payload: payload, muteHttpExceptions: true,
        });
      } catch (e) {
        lastErr = '通信エラー';
        Utilities.sleep(2000 * (attempt + 1));
        continue;
      }
      var code = res.getResponseCode();
      var text = res.getContentText();
      if (code === 200) {
        var parsed = parseGeminiJson_(text);
        if (parsed) return parsed;
        lastErr = model + ' の返答が JSON として読めません';
        continue;
      }
      var info = geminiErrorInfo_(text);
      if (code === 400 && /api key/i.test(info.message)) throw new ConfigError('Gemini の API キーが使えません。GEMINI_API_KEY を確認してください');
      if (code === 403) throw new ConfigError('Gemini の API キーに権限がありません（403）。GEMINI_API_KEY を確認してください');
      lastErr = model + ': ' + code + ' ' + info.message;
      if (code === 404) break;
      if (code === 429) {
        if (info.perDay) break;
        Utilities.sleep(Math.min(info.retryMs || 10000, 30000));
        continue;
      }
      if (code >= 500) { Utilities.sleep(3000 * (attempt + 1)); continue; }
      break;
    }
  }
  throw new Error('Gemini で読み取れませんでした（' + lastErr + '）');
}

function parseGeminiJson_(text) {
  try {
    var body = JSON.parse(text);
    var cand = (body.candidates || [])[0];
    if (!cand || !cand.content) return null;
    var out = (cand.content.parts || []).filter(function (p) { return !p.thought && p.text; }).map(function (p) { return p.text; }).join('');
    var obj = JSON.parse(out);
    if (typeof obj.is_report !== 'boolean') return null;
    return obj;
  } catch (e) {
    return null;
  }
}

// ===================== 接続テスト =====================

function checkSetup() {
  var lines = [];
  var ok = true;
  function line(good, text) { lines.push((good ? '✅ ' : '❌ ') + text); if (!good) ok = false; }
  var ctx;
  try {
    ctx = loadContext_(true);
    line(true, '設定タブと対象タブの見出し: OK（' + ctx.settings.tabs.join('、') + '）');
  } catch (e) {
    line(false, e.message);
    return lines;
  }
  try {
    var auth = slackApi_(ctx, 'auth.test', {});
    line(true, 'Slack: ' + (auth.team || '') + ' に「' + (auth.user || '') + '」として接続');
    ctx.settings.channels.forEach(function (ch) {
      try {
        slackApi_(ctx, 'conversations.history', { channel: ch, limit: 1 });
        line(true, 'チャンネル ' + ch + ' を読めます');
      } catch (e) { line(false, e.message); }
    });
  } catch (e) { line(false, e.message); }
  try {
    var models = geminiModels_(ctx);
    var probe = UrlFetchApp.fetch(GEMINI_ROOT + 'models/' + models[0] + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': ctx.geminiKey }, muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'OK とだけ返してください' }] }] }),
    });
    line(probe.getResponseCode() === 200, 'Gemini: ' + models[0] + (probe.getResponseCode() === 200 ? ' で応答あり' : ' が ' + probe.getResponseCode() + ' を返しました'));
  } catch (e) { line(false, e.message); }
  lines.push('');
  lines.push(ok ? 'すべて OK です。次に「試運転」を実行してください。' : '❌ の項目を直してから、もう一度実行してください。');
  return lines;
}

// Node でのテスト用（Apps Script では module は存在しない）
if (typeof module !== 'undefined') {
  module.exports = {
    normalizeText: normalizeText, normalizeName: normalizeName, normDate: normDate, normTime: normTime,
    cleanSlackText: cleanSlackText, parseParentTime: parseParentTime, jstDateFromTs: jstDateFromTs,
    addDays: addDays, daysBetween: daysBetween, meetingDateForParent: meetingDateForParent,
    findCandidates: findCandidates, repSet: repSet, buildUpdates: buildUpdates, decideCell: decideCell,
    comparable: comparable, rankModels: rankModels, buildPrompt: buildPrompt, responseSchema: responseSchema,
    rowKeyOf: rowKeyOf, formatYen: formatYen, threadFingerprint: threadFingerprint,
    threadKeyFromLink: threadKeyFromLink, buildThreadMap: buildThreadMap, buildOwnership: buildOwnership,
    resolvePlanName: resolvePlanName, cleanSecret: cleanSecret, describeToken: describeToken,
  };
}
