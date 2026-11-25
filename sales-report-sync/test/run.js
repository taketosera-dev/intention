/* eslint-disable no-console */
// Code.gs のテスト。Apps Script の各サービスを擬似的に用意して、照合→書き込み→ログ→要確認の流れを確かめる。
// 実行: node sales-report-sync/test/run.js
// データはすべて架空のもの（実在の顧客・社員の情報は使わない）。
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const assert = require('assert');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

let passed = 0;
// vm で読み込んだコードが返すオブジェクトは別の realm のものなので、JSON にしてから比べる
function eq(a, b, msg) { assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), b, msg); }
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('  ok  ' + name);
  } catch (e) {
    console.error('  NG  ' + name);
    throw e;
  }
}

// ---------- 擬似 Apps Script ----------

function colLetter(n) { return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[n - 1]; }

class FakeSheet {
  constructor(name, rows, maxRows) {
    this.name = name;
    this.cells = {}; // "r,c" -> display string
    this.maxRows = maxRows || Math.max(rows.length + 50, 100);
    rows.forEach((row, i) => row.forEach((v, j) => this.set(i + 1, j + 1, v)));
  }
  key(r, c) { return r + ',' + c; }
  get(r, c) { const v = this.cells[this.key(r, c)]; return v == null ? '' : v; }
  set(r, c, v) {
    if (r > this.maxRows) throw new Error(this.name + ': 行 ' + r + ' はグリッドの外です');
    if (typeof v === 'string' && v.startsWith("'")) v = v.slice(1);
    if (v === false) v = 'FALSE';
    if (v === true) v = 'TRUE';
    this.cells[this.key(r, c)] = v instanceof Date ? v : String(v);
  }
  getLastRow() {
    let last = 0;
    Object.keys(this.cells).forEach((k) => { if (this.cells[k] !== '') last = Math.max(last, +k.split(',')[0]); });
    return last;
  }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(_after, n) { this.maxRows += n; }
  getRange(r, c, nr, nc) { return new FakeRange(this, r, c, nr || 1, nc || 1); }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), 10); }
  setFrozenRows() {}
  clear() { this.cells = {}; }
  getName() { return this.name; }
}

class FakeRange {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getDisplayValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) {
        const v = this.sh.get(this.r + i, this.c + j);
        row.push(v instanceof Date ? v.toISOString() : v);
      }
      out.push(row);
    }
    return out;
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValue(v) { this.sh.set(this.r, this.c, v); return this; }
  setValues(vals) {
    if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error('setValues の大きさが範囲と合いません');
    vals.forEach((row, i) => row.forEach((v, j) => this.sh.set(this.r + i, this.c + j, v)));
    return this;
  }
  setFontWeight() { return this; }
  insertCheckboxes() { if (this.sh.get(this.r, this.c) === '') this.sh.set(this.r, this.c, 'FALSE'); return this; }
}

class FakeSpreadsheet {
  constructor(sheets) { this.sheets = {}; sheets.forEach((s) => { this.sheets[s.name] = s; }); }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { const s = new FakeSheet(n, []); this.sheets[n] = s; return s; }
}

function makeEnv(opts) {
  const props = Object.assign({}, opts.props || {});
  const cache = {};
  const calls = { slack: [], gemini: 0 };
  const env = {
    console: { log() {}, error() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => opts.ss, flush() {}, getUi: () => { throw new Error('UI なし'); } },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = String(v); },
        deleteProperty: (k) => { delete props[k]; },
        getProperties: () => Object.assign({}, props),
      }),
    },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: { getProjectTriggers: () => [], newTrigger() { throw new Error('テストでは使わない'); } },
    Utilities: {
      sleep() {},
      formatDate: (d, _tz, fmt) => {
        const j = new Date(d.getTime() + 9 * 3600 * 1000);
        const p = (n) => String(n).padStart(2, '0');
        if (fmt === 'M/d HH:mm') return (j.getUTCMonth() + 1) + '/' + j.getUTCDate() + ' ' + p(j.getUTCHours()) + ':' + p(j.getUTCMinutes());
        return j.getUTCFullYear() + '/' + p(j.getUTCMonth() + 1) + '/' + p(j.getUTCDate()) + ' ' + p(j.getUTCHours()) + ':' + p(j.getUTCMinutes()) + ':' + p(j.getUTCSeconds());
      },
      computeDigest: (_a, s) => Array.from(crypto.createHash('md5').update(s, 'utf8').digest()),
      base64Encode: (bytes) => Buffer.from(bytes.map((b) => (b + 256) % 256)).toString('base64'),
      DigestAlgorithm: { MD5: 'MD5' },
      Charset: { UTF_8: 'UTF-8' },
    },
    UrlFetchApp: {
      fetch: (url, o) => {
        if (url.startsWith('https://slack.com/api/')) {
          const u = new URL(url);
          const method = u.pathname.replace('/api/', '');
          const params = Object.fromEntries(u.searchParams.entries());
          calls.slack.push(method);
          return resp(200, opts.slack(method, params));
        }
        if (url.includes('generativelanguage.googleapis.com')) {
          if (url.includes('/models?')) return resp(200, { models: [{ name: 'models/gemini-2.5-flash', supportedGenerationMethods: ['generateContent'] }] });
          calls.gemini++;
          const body = JSON.parse(o.payload);
          return opts.gemini(body.contents[0].parts[0].text, calls.gemini);
        }
        throw new Error('想定外の URL: ' + url);
      },
    },
  };
  env.module = { exports: {} };
  vm.createContext(env);
  vm.runInContext(SRC, env, { filename: 'Code.gs' });
  env.__props = props;
  env.__calls = calls;
  return env;
}

function resp(code, body, headers) {
  const text = typeof body === 'string' ? body : JSON.stringify(body);
  return { getResponseCode: () => code, getContentText: () => text, getHeaders: () => headers || {} };
}

function geminiOk(obj) {
  return resp(200, { candidates: [{ content: { parts: [{ text: JSON.stringify(obj) }] } }] });
}

// ---------- 純粋な関数 ----------

const F = makeEnv({ ss: new FakeSpreadsheet([]), slack: () => ({ ok: true }), gemini: () => geminiOk({}) }).module.exports;

console.log('純粋な関数');

test('親メッセージの時刻を読む', () => {
  assert.strictEqual(F.parseParentTime('19時広告面談入ります！\n<@U000AAA|担当　太郎>'), '19:00');
  assert.strictEqual(F.parseParentTime('17時半広告面談入ります！'), '17:30');
  assert.strictEqual(F.parseParentTime('10時30スレッズアポ入ります！'), '10:30');
  assert.strictEqual(F.parseParentTime('15:30　スレッズアポ入ります！'), '15:30');
  assert.strictEqual(F.parseParentTime('<@U000AAA|担当>\n19:15～予定していた方から予定変更の依頼がありました。'), '19:15');
  assert.strictEqual(F.parseParentTime('9時広告面談入ります！'), '09:00');
  assert.strictEqual(F.parseParentTime('１９時　面談入ります！'), '19:00');
  assert.strictEqual(F.parseParentTime('21時　2回目入ります！\n11日の私のアポの後、12日に'), '21:00');
  assert.strictEqual(F.parseParentTime('21時2回目入ります！'), '21:00');
  assert.strictEqual(F.parseParentTime('19時 2回目アポ入ります！'), '19:00');
  assert.strictEqual(F.parseParentTime('9時5分から面談入ります'), '09:05');
  assert.strictEqual(F.parseParentTime('14時クロージング入ります！'), '14:00');
  assert.strictEqual(F.parseParentTime('12時広告クロージング入ります！\n未読なので未着座の可能性大です'), '12:00');
  assert.strictEqual(F.parseParentTime('<@U000AAA|担当>さんがチャンネルに参加しました'), '');
});

test('日付・時刻をそろえる', () => {
  assert.strictEqual(F.normDate('2026/9/1'), '2026/09/01');
  assert.strictEqual(F.normDate('2026-11-24'), '2026/11/24');
  assert.strictEqual(F.normDate(''), '');
  assert.strictEqual(F.normTime('9:00'), '09:00');
  assert.strictEqual(F.normTime('15:30（キャンセル済み）'), '15:30');
  assert.strictEqual(F.normTime('25:00'), '');
});

test('Slack の ts を日本時間の日付にする', () => {
  // 2026-11-24 23:30 JST = 2026-11-24 14:30 UTC
  const ts = String(Date.UTC(2026, 10, 24, 14, 30) / 1000) + '.000100';
  assert.strictEqual(F.jstDateFromTs(ts), '2026/11/24');
  // 2026-11-25 00:30 JST = 2026-11-24 15:30 UTC
  const ts2 = String(Date.UTC(2026, 10, 24, 15, 30) / 1000) + '.000100';
  assert.strictEqual(F.jstDateFromTs(ts2), '2026/11/25');
  assert.strictEqual(F.meetingDateForParent(ts, '明日10時面談入ります'), '2026/11/25');
});

function row(o) {
  const r = new Array(21).fill('');
  Object.keys(o).forEach((k) => { r['ABCDEFGHIJKLMNOPQRSTU'.indexOf(k)] = o[k]; });
  return r;
}

test('担当者×日時で行と回数を特定する', () => {
  const tabs = {
    TabA: [
      row({ A: '100', B: 'テストA', G: '2026/11/21', H: '17:30', I: '山田花子' }),
      row({ A: '101', B: 'テストB', G: '2026/11/21', H: '19:00', I: '山田花子', J: '2026/11/21', K: '19:00', N: '2026/11/24', O: '12:00' }),
      row({ A: '102', B: 'テストC', G: '2026/11/21', H: '17:30', I: '鈴木一郎' }),
    ],
    TabB: [row({ A: '200', B: 'テストD', G: '2026/10/29', H: '15:30（キャンセル済み）', I: '山田 花子' })],
  };
  eq(F.findCandidates(tabs, '山田　花子', '2026/11/21', '17:30'), [{ tab: 'TabA', rowIndex: 0, meetingNo: 1 }]);
  eq(F.findCandidates(tabs, '山田花子', '2026/11/24', '12:00'), [{ tab: 'TabA', rowIndex: 1, meetingNo: 2 }]);
  eq(F.findCandidates(tabs, '山田花子', '2026/10/29', '15:30'), [{ tab: 'TabB', rowIndex: 0, meetingNo: 1 }]);
  eq(F.findCandidates(tabs, '山田花子', '2026/11/22', '17:30'), []);
  eq(F.findCandidates(tabs, '佐藤', '2026/11/21', '17:30'), []);
});

const PLANS = { プランA: 300000, プランB: 200000, プランC: 100000 };
function upd(ex, no, date, time) { return F.buildUpdates(ex, no, date || '2026/11/21', time || '17:30', PLANS); }
function asMap(r) { const m = {}; r.updates.forEach((u) => { m[u.col] = u.value; }); return m; }

test('未着座は 着座・結果とも「飛び」', () => {
  const r = upd({ is_report: true, attendance: '飛び', result: '飛び', confidence: 'high' }, 1);
  eq(asMap(r), { J: '2026/11/21', K: '17:30', L: '飛び', M: '飛び' });
});

test('ブリッジは 着座・保留＋次回日時を2回目へ', () => {
  const r = upd({ is_report: true, attendance: '着座', result: '保留', next_meeting_date: '2026-11-24', next_meeting_time: '12:00', reason: '2回目へブリッジ', confidence: 'high' }, 1);
  eq(asMap(r), { J: '2026/11/21', K: '17:30', L: '着座', M: '保留', N: '2026/11/24', O: '12:00', U: '2回目へブリッジ' });
});

test('2回目の成約言質：実施時刻・プラン金額・理由', () => {
  const r = upd({ is_report: true, result: '成約言質', actual_time: '13:00', plan: 'プランB', reason: 'ローン審査待ち', confidence: 'high' }, 2, '2026/11/24', '12:00');
  eq(asMap(r), { N: '2026/11/24', O: '13:00', P: '成約言質', T: 'プランB 200,000', U: 'ローン審査待ち' });
});

test('成約言質のあとの手続きの日時は3回目の予定に入れる', () => {
  const r = upd({ is_report: true, result: '成約言質', actual_time: '13:00', next_meeting_date: '2026-11-27', next_meeting_time: '15:00', plan: 'プランB', confidence: 'high' }, 2, '2026/11/24', '12:00');
  const m = asMap(r);
  assert.strictEqual(m.Q, '2026/11/27');
  assert.strictEqual(m.R, '15:00');
  assert.strictEqual(r.notes.length, 0);
});

test('プラン名の略称は正式名にそろえる', () => {
  const plans = { ロングプラン: 300000, ショート: 100000 };
  assert.strictEqual(F.resolvePlanName('ロング', plans), 'ロングプラン');
  assert.strictEqual(F.resolvePlanName('ロングプラン', plans), 'ロングプラン');
  assert.strictEqual(F.resolvePlanName('ショート', plans), 'ショート');
  assert.strictEqual(F.resolvePlanName('ミドル', plans), null);
  assert.strictEqual(F.resolvePlanName('', plans), null);
  const r = F.buildUpdates({ is_report: true, result: '成約言質', plan: 'ロング', confidence: 'high' }, 2, '2026/11/24', '12:00', plans);
  assert.strictEqual(asMap(r).T, 'ロングプラン 300,000');
  const r2 = F.buildUpdates({ is_report: true, result: '成約言質', plan: 'ミドル', confidence: 'high' }, 2, '2026/11/24', '12:00', plans);
  assert.ok(!('T' in asMap(r2)));
  assert.strictEqual(r2.notes.length, 1);
});

test('報告に金額があればそちらを使う', () => {
  const r = upd({ is_report: true, result: '成約', plan: 'プランA', amount: 250000, confidence: 'high' }, 2, '2026/11/24', '12:00');
  assert.strictEqual(asMap(r).T, 'プランA 250,000');
});

test('3回目の次回日時は列がないので要確認', () => {
  const r = upd({ is_report: true, result: '保留', next_meeting_date: '2026-11-30', next_meeting_time: '19:00', confidence: 'high' }, 3, '2026/11/20', '19:00');
  eq(asMap(r), { Q: '2026/11/20', R: '19:00', S: '保留' });
  assert.strictEqual(r.notes.length, 1);
  assert.ok(/4回目/.test(r.notes[0]));
});

test('リスケ・飛びのときの次回日時は書かずに要確認', () => {
  const r = upd({ is_report: true, attendance: 'リスケ', next_meeting_date: '2026-11-28', confidence: 'high' }, 1);
  eq(asMap(r), { J: '2026/11/21', K: '17:30', L: 'リスケ', M: '未実施' });
  assert.strictEqual(r.notes.length, 1);
});

test('不自然な次回日付は書かない', () => {
  const r = upd({ is_report: true, attendance: '着座', result: '保留', next_meeting_date: '2026-10-01', confidence: 'high' }, 1);
  assert.ok(!('N' in asMap(r)));
  assert.strictEqual(r.notes.length, 1);
});

test('プルダウンにない値は書かずに要確認', () => {
  const r = upd({ is_report: true, attendance: '着座', result: '検討中', confidence: 'high' }, 1);
  assert.ok(!('M' in asMap(r)));
  assert.ok(r.notes.some((n) => /プルダウン/.test(n)));
});

test('報告でないものは何もしない', () => {
  eq(upd({ is_report: false, confidence: 'high' }, 1).updates, []);
});

test('長い理由は 400 文字で切る', () => {
  const r = upd({ is_report: true, attendance: '着座', result: '失注', reason: 'あ'.repeat(500), confidence: 'high' }, 1);
  assert.strictEqual(asMap(r).U.length, 400);
});

test('セルごとの判断', () => {
  assert.strictEqual(F.decideCell('M', '', '保留', null), 'write');
  assert.strictEqual(F.decideCell('M', '保留', '保留', null), 'same');
  assert.strictEqual(F.decideCell('K', '9:00', '09:00', null), 'same');
  assert.strictEqual(F.decideCell('J', '2026/9/1', '2026/09/01', null), 'same');
  assert.strictEqual(F.decideCell('M', '保留', '失注', '保留'), 'write'); // ツールが書いた値は更新してよい
  assert.strictEqual(F.decideCell('M', '保留', '失注', null), 'conflict'); // 人が入力した値
  assert.strictEqual(F.decideCell('M', '成約', '失注', '保留'), 'conflict'); // ツールの後に人が直した
});

test('ログのリンクから、スレッドと行・回数の対応を作る', () => {
  const link = 'https://example.slack.com/archives/C0TESTAAA1/p1795488841700539';
  eq(F.threadKeyFromLink(link + '?thread_ts=1795488841.700539&cid=C0TESTAAA1'), 'C0TESTAAA1:1795488841.700539');
  const map = F.buildThreadMap([
    ['', '顧客A', '100', '2026/11/21', 'N', '2026/11/24', 'https://example.slack.com/archives/C0TESTAAA1/p1795249533953229', '初回の一括転記'],
    ['', '顧客A', '100', '2026/11/21', 'M', '保留', 'https://example.slack.com/archives/C0TESTAAA1/p1795249533953229', '初回の一括転記'],
    ['', '顧客A', '100', '2026/11/21', 'P', '成約言質', link, '初回の一括転記'],
    ['', '顧客A', '100', '2026/11/21', 'S', '失注', 'https://example.slack.com/archives/C0TESTAAA1/p1795000000000001', '試運転'],
  ]);
  eq(map, {
    'C0TESTAAA1:1795249533.953229': { tab: '顧客A', rowKey: '顧客A|100|2026/11/21', no: 1 },
    'C0TESTAAA1:1795488841.700539': { tab: '顧客A', rowKey: '顧客A|100|2026/11/21', no: 2 },
  });
});

test('貼り付けたトークンの前後の空白・改行・引用符を取り除く', () => {
  assert.strictEqual(F.cleanSecret('  xoxb-123-abc\n'), 'xoxb-123-abc');
  assert.strictEqual(F.cleanSecret('"xoxb-123-abc"'), 'xoxb-123-abc');
  assert.strictEqual(F.cleanSecret(null), '');
  assert.ok(/ボットのトークン（xoxb-）、12文字/.test(F.describeToken('xoxb-123-abc')));
  assert.ok(!/123-abc/.test(F.describeToken('xoxb-123-abc')), 'トークンの中身は出さない');
  assert.ok(/xoxb- で始まらない値（先頭: abcde）/.test(F.describeToken('abcdefgh')));
});

test('モデルは安定版 Flash を優先', () => {
  const list = F.rankModels(['gemini-2.5-pro', 'gemini-2.0-flash', 'gemini-2.5-flash-preview-05-20', 'gemini-flash-latest', 'gemini-2.5-flash-lite', 'gemini-2.5-flash', 'gemini-2.5-flash-image']);
  assert.strictEqual(list[0], 'gemini-2.5-flash');
  assert.ok(list.indexOf('gemini-2.5-flash-image') < 0);
  assert.ok(list.indexOf('gemini-2.5-flash-preview-05-20') < 0);
});

// ---------- 全体の流れ ----------

console.log('全体の流れ（擬似 Slack・擬似 Gemini）');

const HEADER = ['LINE ID', '表示名', '友だち追加日', '流入経路', '広告名', '申込日', 'アポ予定日', 'アポ時間', '担当者',
  '初回商談\n実施予定日', '初回商談\n実施時間', '初回商談\n着座', '初回商談\n結果', '2回目商談\n実施予定日', '2回目商談\n実施時間', '2回目商談\n結果',
  '3回目商談\n実施予定日', '3回目商談\n実施時間', '3回目商談\n結果', 'プラン・金額', '失注・保留理由', 'メモ'];

function tsAt(y, mo, d, h, mi) { return (Date.UTC(y, mo - 1, d, h - 9, mi) / 1000).toFixed(6); }

function buildWorld() {
  const tabA = new FakeSheet('顧客A', [
    ['顧客管理（テスト）'], [], HEADER,
    ['100', 'テストA', '', '', '', '', '2026/11/21', '17:30', '山田花子'],
    ['101', 'テストB', '', '', '', '', '2026/11/21', '19:00', '山田花子'],
    ['102', 'テストC', '', '', '', '', '2026/11/22', '20:00', '鈴木一郎', '', '', '', '保留'], // M は人が入力済み
    ['103', 'テストD', '', '', '', '', '2026/11/10', '10:00', '鈴木一郎'], // 報告なし
  ]);
  const tabB = new FakeSheet('顧客B', [
    ['顧客管理（テスト2）'], [], HEADER,
    ['200', 'テストE', '', '', '', '', '2026/11/23', '15:30（キャンセル済み）', '山田花子'],
  ]);
  const settings = new FakeSheet('設定', [
    ['項目', '値', '説明'],
    ['対象チャンネル', 'C0TESTAAA1, C0TESTBBB2', ''],
    ['対象タブ', '顧客A, 顧客B', ''],
    ['処理開始日時', '2026/11/20 00:00', ''],
    ['読み込む期間（日）', '60', ''],
    ['報告なしとみなす日数', '2', ''],
    [],
    ['プラン', '金額', ''],
    ['プランA', '300000'], ['プランB', '200000'], ['プランC', '100000'],
  ]);
  const log = new FakeSheet('自動転記ログ', [['記録日時', 'タブ', 'LINE ID', 'アポ予定日', '列', '書き込んだ値', '元のSlack投稿', '処理']], 2000);
  const review = new FakeSheet('要確認', [['登録日時', 'タブ', '担当者', '面談日時', '内容', '理由', '元のSlack投稿', '対応済み', 'キー']], 500);
  for (let r = 2; r <= 500; r++) review.set(r, 8, 'FALSE'); // 本物と同じく、空行にもチェックボックスがある
  const ss = new FakeSpreadsheet([tabA, tabB, settings, log, review]);

  const users = { U1: '山田　花子', U2: '鈴木一郎', U9: '応援 次郎' };
  const T = {
    a1: tsAt(2026, 11, 21, 17, 25), // テストA 初回 → ブリッジ 11/24 12:00
    a2: tsAt(2026, 11, 24, 11, 54), // テストA 2回目 → 12時は未着座、13時に実施して成約言質
    b1: tsAt(2026, 11, 21, 18, 52), // テストB 初回 → 未着座
    c1: tsAt(2026, 11, 22, 19, 50), // テストC 初回 → 失注（M は人が「保留」と入力済み）
    x1: tsAt(2026, 11, 23, 9, 50), // 合う行がない
    e1: tsAt(2026, 11, 23, 15, 20), // テストE（別タブ・別チャンネル） → 未着座
    n1: tsAt(2026, 11, 23, 12, 0), // 担当者でない人の投稿
  };
  const msg = (ts, user, text, extra) => Object.assign({ ts, user, text, type: 'message' }, extra || {});
  const threads = {
    C0TESTAAA1: [
      [msg(T.a1, 'U1', '17時半広告面談入ります！\n<@U9|応援 次郎>'), msg(T.a1.replace(/.$/, '9'), 'U9', 'ファイト！'), msg(tsAt(2026, 11, 21, 18, 40), 'U1', 'テスト太郎さん→11/24（火）12:00～ブリッジ')],
      [msg(T.a2, 'U1', '12時広告クロージング入ります！'), msg(tsAt(2026, 11, 24, 12, 20), 'U1', '未着座でしたが13時開始で再度時間を取りました'), msg(tsAt(2026, 11, 24, 14, 10), 'U1', 'テスト太郎さん→27日15:00〜 プランB言質。ローン審査待ち')],
      [msg(T.b1, 'U1', '19時広告面談入ります！'), msg(tsAt(2026, 11, 21, 19, 15), 'U1', 'テスト花子→未着座でした')],
      [msg(T.c1, 'U2', '20時面談入ります！'), msg(tsAt(2026, 11, 22, 21, 10), 'U2', 'テスト三郎さん→失注。費用面で見送り')],
      [msg(T.x1, 'U2', '10時面談入ります！'), msg(tsAt(2026, 11, 23, 11, 0), 'U2', '→失注')],
      [msg(T.n1, 'U9', '13時に会議します、面談の振り返り'), msg(tsAt(2026, 11, 23, 12, 5), 'U9', 'よろしく')],
      [msg(tsAt(2026, 11, 20, 9, 0), 'U2', '', { subtype: 'channel_join' })],
    ],
    C0TESTBBB2: [
      [msg(T.e1, 'U1', '15:30　スレッズアポ入ります！'), msg(tsAt(2026, 11, 23, 15, 45), 'U1', '→未着座')],
    ],
  };
  const slack = (method, p) => {
    if (method === 'conversations.history') {
      return { ok: true, messages: threads[p.channel].map((th) => Object.assign({}, th[0], th.length > 1 ? { reply_count: th.length - 1, latest_reply: th[th.length - 1].ts, thread_ts: th[0].ts } : {})).reverse() };
    }
    if (method === 'conversations.replies') {
      const th = threads[p.channel].find((t) => t[0].ts === p.ts);
      return { ok: true, messages: th.map((m, i) => Object.assign({ thread_ts: th[0].ts }, m, i ? { parent_user_id: th[0].user } : {})) };
    }
    if (method === 'users.info') return { ok: true, user: { id: p.user, real_name: users[p.user], profile: { real_name: users[p.user] } } };
    if (method === 'chat.getPermalink') return { ok: true, permalink: 'https://example.slack.com/archives/' + p.channel + '/p' + p.message_ts.replace('.', '') };
    if (method === 'auth.test') return { ok: true, team: 'テスト', user: 'bot' };
    throw new Error('想定外の Slack API: ' + method);
  };
  const gemini = (prompt) => {
    // 親メッセージの内容で、それらしい読み取り結果を返す
    if (prompt.includes('17時半広告面談')) return geminiOk({ is_report: true, attendance: '着座', result: '保留', next_meeting_date: '2026-11-24', next_meeting_time: '12:00', reason: '2回目へブリッジ。金銭面の懸念あり', confidence: 'high' });
    if (prompt.includes('12時広告クロージング')) {
      assert.ok(prompt.includes('何回目の商談か: 2回目'), '2回目として渡す');
      assert.ok(prompt.includes('シートの今の理由欄: 2回目へブリッジ'), '今の理由欄を渡す');
      return geminiOk({ is_report: true, result: '成約言質', actual_time: '13:00', next_meeting_date: '2026-11-27', next_meeting_time: '15:00', plan: 'プランB', reason: 'プランBで言質。ローン審査待ち', confidence: 'high' });
    }
    if (prompt.includes('19時広告面談')) return geminiOk({ is_report: true, attendance: '飛び', result: '飛び', confidence: 'high' });
    if (prompt.includes('20時面談')) return geminiOk({ is_report: true, attendance: '着座', result: '失注', reason: '費用面で見送り', confidence: 'high' });
    if (prompt.includes('15:30')) return geminiOk({ is_report: true, attendance: '飛び', result: '飛び', confidence: 'high' });
    throw new Error('想定外の prompt: ' + prompt.slice(0, 80));
  };
  return { ss, tabA, tabB, log, review, slack, gemini, T };
}

function cell(sh, rowNo, col) { return sh.get(rowNo, 'ABCDEFGHIJKLMNOPQRSTU'.indexOf(col) + 1); }
function reviewRows(review) {
  const out = [];
  for (let r = 2; r <= review.getMaxRows(); r++) if (review.get(r, 1) !== '') out.push({ reason: review.get(r, 6), key: review.get(r, 9), tab: review.get(r, 2) });
  return out;
}

// 今日 = 2026/11/26 として動かす
const REAL_NOW = Date.now;
Date.now = () => Date.UTC(2026, 10, 26, 1, 0); // 2026/11/26 10:00 JST

try {
  const w = buildWorld();
  const env = makeEnv({ ss: w.ss, slack: w.slack, gemini: w.gemini, props: { SLACK_BOT_TOKEN: 'xoxb-test', GEMINI_API_KEY: 'test-key' } });

  test('試運転はシートを変えず、試運転結果タブに出す', () => {
    const before = JSON.stringify(w.tabA.cells);
    const r = env.runSync({ dryRun: true });
    assert.strictEqual(JSON.stringify(w.tabA.cells), before);
    const pv = w.ss.getSheetByName('試運転結果');
    assert.ok(pv && pv.getLastRow() > 1, '試運転結果がある');
    assert.strictEqual(w.log.getLastRow(), 1, 'ログは増えない');
    assert.strictEqual(reviewRows(w.review).length, 0, '要確認も増えない');
    assert.ok(Object.keys(env.__props).every((k) => !k.startsWith('t:')), '処理済みの記録も残さない');
    assert.ok(/書き込む予定/.test(r.summary));
  });

  test('本番：初回→2回目の流れ、未着座、別タブ・別チャンネル', () => {
    env.runSync({ dryRun: false });
    // テストA（4行目）
    assert.strictEqual(cell(w.tabA, 4, 'J'), '2026/11/21');
    assert.strictEqual(cell(w.tabA, 4, 'K'), '17:30');
    assert.strictEqual(cell(w.tabA, 4, 'L'), '着座');
    assert.strictEqual(cell(w.tabA, 4, 'M'), '保留');
    assert.strictEqual(cell(w.tabA, 4, 'N'), '2026/11/24');
    assert.strictEqual(cell(w.tabA, 4, 'O'), '13:00', '実施した時刻');
    assert.strictEqual(cell(w.tabA, 4, 'P'), '成約言質');
    assert.strictEqual(cell(w.tabA, 4, 'Q'), '2026/11/27', '成約言質のあとの手続きの日時も3回目に入れる');
    assert.strictEqual(cell(w.tabA, 4, 'R'), '15:00');
    assert.strictEqual(cell(w.tabA, 4, 'T'), 'プランB 200,000');
    assert.strictEqual(cell(w.tabA, 4, 'U'), 'プランBで言質。ローン審査待ち', 'ツールが書いた理由は後の報告で更新');
    // テストB（5行目）未着座
    assert.strictEqual(cell(w.tabA, 5, 'L'), '飛び');
    assert.strictEqual(cell(w.tabA, 5, 'M'), '飛び');
    // テストE（別タブ）
    assert.strictEqual(cell(w.tabB, 4, 'J'), '2026/11/23');
    assert.strictEqual(cell(w.tabB, 4, 'K'), '15:30');
    assert.strictEqual(cell(w.tabB, 4, 'M'), '飛び');
  });

  test('人が入力した値は書き換えずに要確認へ', () => {
    assert.strictEqual(cell(w.tabA, 6, 'M'), '保留');
    const rv = reviewRows(w.review);
    assert.ok(rv.some((r) => /M 列: シートは「保留」/.test(r.reason)), '食い違いが要確認にある');
    assert.strictEqual(cell(w.tabA, 6, 'L'), '着座', '空欄だった着座は書く');
  });

  test('合う行がない報告・報告なしの予定は要確認へ', () => {
    const rv = reviewRows(w.review);
    assert.ok(rv.some((r) => /合う予定が、対象タブのどこにもありません/.test(r.reason)));
    assert.ok(rv.some((r) => r.key === 'missing:顧客A:103:2026/11/10:1'), '報告なし');
    // 担当者列にない人の「商談らしい投稿」は、名前の表記ちがいの見逃しを防ぐため 1 人につき 1 件だけ要確認に出す
    eq(rv.filter((r) => /応援 次郎/.test(r.reason)).map((r) => r.key), ['user:U9:not_rep']);
  });

  test('ログに書いた値と元の投稿が残る', () => {
    const n = w.log.getLastRow() - 1;
    assert.ok(n >= 15, 'ログの件数 ' + n);
    assert.strictEqual(w.log.get(2, 8), '自動転記');
    assert.ok(/^https:\/\/example\.slack\.com\//.test(w.log.get(2, 7)));
  });

  test('2回目の実行では何も増えない（同じ処理を繰り返しても安全）', () => {
    const logBefore = w.log.getLastRow();
    const rvBefore = reviewRows(w.review).length;
    const cellsBefore = JSON.stringify(w.tabA.cells) + JSON.stringify(w.tabB.cells);
    const g = env.__calls.gemini;
    env.runSync({ dryRun: false });
    assert.strictEqual(w.log.getLastRow(), logBefore);
    assert.strictEqual(reviewRows(w.review).length, rvBefore);
    assert.strictEqual(JSON.stringify(w.tabA.cells) + JSON.stringify(w.tabB.cells), cellsBefore);
    assert.strictEqual(env.__calls.gemini, g, '変化のないスレッドで Gemini を呼ばない');
  });

  test('人がツールの値を直したら、その後は上書きしない', () => {
    w.tabA.set(5, 13, '保留'); // テストB の M を人が「飛び」→「保留」に直した
    const env2 = makeEnv({ ss: w.ss, slack: w.slack, gemini: w.gemini, props: env.__props });
    // スレッドを読み直させるため、処理済みの記録を消す
    Object.keys(env2.__props).filter((k) => k.startsWith('t:')).forEach((k) => delete env2.__props[k]);
    const g = env2.__calls.gemini;
    env2.runSync({ dryRun: false });
    assert.ok(env2.__calls.gemini > g, 'スレッドを読み直している');
    assert.strictEqual(cell(w.tabA, 5, 'M'), '保留');
    assert.ok(reviewRows(w.review).some((r) => /M 列: シートは「保留」、報告からは「飛び」/.test(r.reason)));
  });

  test('一括転記済みのスレッドに追記があったら、ログのリンクから同じ行を見つける', () => {
    const w5 = buildWorld();
    // テストA は一括転記済み：2回目は 12時の予定が 13時に実施され、O 列は 13:00（親メッセージの「12時」とは合わない）
    const pre = { J: '2026/11/21', K: '17:30', L: '着座', M: '保留', N: '2026/11/24', O: '13:00', P: '成約言質' };
    const link = (ts) => 'https://example.slack.com/archives/C0TESTAAA1/p' + ts.replace('.', '');
    let r = 2;
    Object.keys(pre).forEach((col) => {
      w5.tabA.set(4, 'ABCDEFGHIJKLMNOPQRSTU'.indexOf(col) + 1, pre[col]);
      // O（実施時刻）と P（2回目の結果）は 2回目のスレッド、それ以外は初回のスレッドから書いた
      const ts = (col === 'O' || col === 'P') ? w5.T.a2 : w5.T.a1;
      ['2026/11/26 01:11:45', '顧客A', '100', '2026/11/21', col, pre[col], link(ts), '初回の一括転記']
        .forEach((v, j) => w5.log.set(r, j + 1, v));
      r++;
    });
    // 処理開始日時より後に、2回目のスレッドへ追記があった
    const settings = w5.ss.getSheetByName('設定');
    settings.set(4, 2, '2026/11/25 00:00');
    const g = (prompt) => {
      if (prompt.includes('12時広告クロージング')) {
        assert.ok(prompt.includes('何回目の商談か: 2回目'));
        // 読み直しのとき、AI が実施時刻を返さないこともある。親メッセージの「12時」で 13:00 を上書きしないこと
        return geminiOk({ is_report: true, result: '成約', plan: 'プランB', confidence: 'high' });
      }
      return w5.gemini(prompt);
    };
    const slack2 = (method, p) => {
      const res = w5.slack(method, p);
      if (method === 'conversations.history' && p.channel === 'C0TESTAAA1') {
        res.messages.forEach((m) => { if (m.ts === w5.T.a2) m.latest_reply = tsAt(2026, 11, 25, 10, 0); });
      }
      return res;
    };
    const env7 = makeEnv({ ss: w5.ss, slack: slack2, gemini: g, props: { SLACK_BOT_TOKEN: 'x', GEMINI_API_KEY: 'y' } });
    env7.runSync({ dryRun: false });
    assert.strictEqual(cell(w5.tabA, 4, 'P'), '成約', 'ツールが書いた「成約言質」を「成約」に更新');
    assert.strictEqual(cell(w5.tabA, 4, 'O'), '13:00', '実施時刻はそのまま');
    assert.ok(!reviewRows(w5.review).some((x) => /合う予定/.test(x.reason) && x.key.includes(w5.T.a2)), '行が見つからない扱いにしない');
  });

  test('列の見出しがずれていたら何も書かずに止める', () => {
    const w2 = buildWorld();
    w2.tabA.set(3, 13, '結果'); // M 列の見出しを変える
    const env3 = makeEnv({ ss: w2.ss, slack: w2.slack, gemini: w2.gemini, props: { SLACK_BOT_TOKEN: 'x', GEMINI_API_KEY: 'y' } });
    assert.throws(() => env3.runSync({ dryRun: false }), /見出し/);
    assert.strictEqual(cell(w2.tabA, 4, 'J'), '');
  });

  test('Gemini が混雑しても例外で止まらず、次回に再挑戦する', () => {
    const w3 = buildWorld();
    let n = 0;
    const busy = () => { n++; return resp(503, { error: { message: 'overloaded' } }); };
    const env4 = makeEnv({ ss: w3.ss, slack: w3.slack, gemini: busy, props: { SLACK_BOT_TOKEN: 'x', GEMINI_API_KEY: 'y' } });
    const r = env4.runSync({ dryRun: false });
    assert.ok(/読み取り失敗/.test(r.summary));
    assert.strictEqual(cell(w3.tabA, 4, 'M'), '');
    // 次の実行でまた試す（失敗の記録があるため）
    const env5 = makeEnv({ ss: w3.ss, slack: w3.slack, gemini: w3.gemini, props: env4.__props });
    env5.runSync({ dryRun: false });
    assert.strictEqual(cell(w3.tabA, 4, 'M'), '保留');
  });

  test('Slack のトークンが無効なら設定エラーで止める', () => {
    const w4 = buildWorld();
    const bad = (method) => (method === 'conversations.history' ? { ok: false, error: 'invalid_auth' } : w4.slack(method, {}));
    const env6 = makeEnv({ ss: w4.ss, slack: bad, gemini: w4.gemini, props: { SLACK_BOT_TOKEN: 'x', GEMINI_API_KEY: 'y' } });
    assert.throws(() => env6.runSync({ dryRun: false }), /トークン/);
  });
} finally {
  Date.now = REAL_NOW;
}

console.log('\n' + passed + ' 件すべて成功');
