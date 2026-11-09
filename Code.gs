// 우리 일기장 - Apps Script 서버
var S = {
  users: ['회원', ['id','nickname','salt','hash','role','status','ver','must','pen','createdAt']],
  entries: ['일기', ['id','authorId','date','title','body','mood','weather','visibility','photos','createdAt','updatedAt']],
  comments: ['댓글', ['id','entryId','authorId','body','createdAt']],
  likes: ['공감', ['entryId','userId']]
};
var DAY = 86400000;

function doGet() {
  return ContentService.createTextOutput('우리 일기장 서버가 작동 중입니다.');
}

// 클라우드플레어 화면에서 오는 요청을 받습니다
function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  return ContentService.createTextOutput(JSON.stringify(api(b.action, b.token, b.p)))
    .setMimeType(ContentService.MimeType.JSON);
}

// 처음 한 번 실행: 시트와 초대코드를 만듭니다. 초대코드는 '설정' 시트 B1에 있습니다.
function setup() {
  Object.keys(S).forEach(sh_);
  folder_(); secret_();
  var s = settings_();
  if (!s.getRange('B1').getDisplayValue()) setInvite_();
  Logger.log('초대코드: ' + s.getRange('B1').getDisplayValue());
}

function settings_() {
  var ss = SpreadsheetApp.getActive(), s = ss.getSheetByName('설정');
  if (!s) { s = ss.insertSheet('설정'); s.getRange('A1').setValue('초대코드'); s.getRange('B1').setNumberFormat('@'); }
  return s;
}
function setInvite_() { var c = uid_().slice(0, 8).toUpperCase(); settings_().getRange('B1').setValue(c); return c; }
function invite_() { return settings_().getRange('B1').getDisplayValue(); }
function secret_() {
  var p = PropertiesService.getScriptProperties(), v = p.getProperty('SECRET');
  if (!v) { v = uid_() + uid_(); p.setProperty('SECRET', v); }
  return v;
}
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER');
  if (id) return DriveApp.getFolderById(id);
  var f = DriveApp.createFolder('우리 일기장 사진');
  p.setProperty('FOLDER', f.getId());
  return f;
}
function uid_() { return Utilities.getUuid().replace(/-/g, ''); }
function now_() { return new Date().toISOString(); }
function fail_(m, code) { throw { appError: m, code: code || 'ERR' }; }
function b64_(b) { return Utilities.base64EncodeWebSafe(b); }
function hash_(pw, salt) {
  var h = salt;
  for (var i = 0; i < 500; i++) h = b64_(Utilities.computeHmacSha256Signature(h, salt + pw));
  return h;
}

function sh_(k) {
  var d = S[k], ss = SpreadsheetApp.getActive(), s = ss.getSheetByName(d[0]);
  if (!s) {
    s = ss.insertSheet(d[0]);
    s.getRange(1, 1, 1, d[1].length).setValues([d[1]]);
    s.setFrozenRows(1);
  }
  return s;
}
function rows_(k) {
  var s = sh_(k), f = S[k][1], n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, f.length).getDisplayValues().map(function (r, i) {
    var o = { _r: i + 2 };
    f.forEach(function (x, j) { o[x] = r[j]; });
    return o;
  });
}
function write_(k, r, o) {
  var f = S[k][1], g = sh_(k).getRange(r, 1, 1, f.length);
  g.setNumberFormat('@');
  g.setValues([f.map(function (x) { return o[x] == null ? '' : String(o[x]); })]);
}
function add_(k, o) {
  var s = sh_(k), r = s.getLastRow() + 1;
  if (r > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 50);
  write_(k, r, o);
}
function del_(k, r) {
  var s = sh_(k);
  if (s.getMaxRows() <= 2) s.insertRowsAfter(s.getMaxRows(), 10);
  s.deleteRow(r);
}
function delWhere_(k, fn) {
  rows_(k).filter(fn).reverse().forEach(function (x) { del_(k, x._r); });
}
function one_(k, id) {
  return rows_(k).filter(function (x) { return x.id === id; })[0];
}

function makeToken_(u) {
  var p = b64_(u.id + '|' + (Date.now() + 30 * DAY) + '|' + u.ver);
  return p + '.' + b64_(Utilities.computeHmacSha256Signature(p, secret_()));
}
function auth_(t) {
  var a = String(t || '').split('.');
  if (a.length !== 2 || b64_(Utilities.computeHmacSha256Signature(a[0], secret_())) !== a[1]) fail_('다시 로그인해 주세요.', 'AUTH');
  var p = Utilities.newBlob(Utilities.base64DecodeWebSafe(a[0])).getDataAsString().split('|');
  var u = one_('users', p[0]);
  if (!u || u.status !== 'active' || u.ver !== p[2] || Number(p[1]) < Date.now()) fail_('다시 로그인해 주세요.', 'AUTH');
  return u;
}
function pub_(u) { return { id: u.id, nickname: u.nickname, role: u.role, status: u.status, pen: Number(u.pen) }; }
function canSee_(e, u) { return e.visibility === 'public' || e.authorId === u.id; }
function checkPw_(p) { if (String(p || '').length < 8) fail_('비밀번호는 8자 이상으로 입력하세요.'); }

// 화면에서 호출하는 유일한 입구
function api(action, token, p) {
  var lock = LockService.getScriptLock();
  try {
    if (['me', 'feed', 'calendar', 'photo'].indexOf(action) < 0) lock.waitLock(20000);
    p = p || {};
    if (action === 'login') return { ok: true, data: login_(p) };
    if (action === 'signup') return { ok: true, data: signup_(p) };
    var u = auth_(token);
    if (u.must === '1' && action !== 'me' && action !== 'changePw') fail_('비밀번호를 먼저 변경하세요.');
    if (!A[action]) fail_('잘못된 요청입니다.');
    if (action.indexOf('admin') === 0 && u.role !== 'admin') fail_('관리자만 할 수 있습니다.');
    return { ok: true, data: A[action](u, p) };
  } catch (e) {
    if (e && e.appError) return { ok: false, error: e.appError, code: e.code };
    console.error(e);
    return { ok: false, error: '처리하지 못했습니다. 다시 시도해 주세요.' };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function login_(p) {
  var id = String(p.id || '').trim().toLowerCase(), c = CacheService.getScriptCache(), key = 'f' + id;
  var n = Number(c.get(key) || 0);
  if (n >= 5) fail_('로그인 실패가 많아 10분간 잠겼습니다.');
  var u = one_('users', id);
  if (!u || u.status !== 'active' || hash_(String(p.password), u.salt) !== u.hash) {
    c.put(key, String(n + 1), 600);
    fail_('아이디 또는 비밀번호가 맞지 않습니다.');
  }
  c.remove(key);
  return { token: makeToken_(u) };
}
function signup_(p) {
  var c = CacheService.getScriptCache(), n = Number(c.get('inv') || 0);
  if (n >= 20) fail_('시도가 많아 10분간 가입이 잠겼습니다.');
  if (String(p.invite || '').trim().toUpperCase() !== invite_()) {
    c.put('inv', String(n + 1), 600);
    fail_('초대코드가 맞지 않습니다.');
  }
  var id = String(p.id || '').trim().toLowerCase(), nick = String(p.nickname || '').trim();
  if (!/^[a-z0-9_]{3,20}$/.test(id)) fail_('아이디는 영문 소문자, 숫자 3~20자로 입력하세요.');
  if (!nick || nick.length > 12) fail_('별명은 1~12자로 입력하세요.');
  checkPw_(p.password);
  var us = rows_('users');
  if (us.some(function (x) { return x.id === id; })) fail_('이미 쓰는 아이디입니다.');
  if (us.some(function (x) { return x.nickname === nick; })) fail_('이미 쓰는 별명입니다.');
  var salt = uid_();
  var u = { id: id, nickname: nick, salt: salt, hash: hash_(String(p.password), salt), role: us.length ? 'member' : 'admin',
    status: 'active', ver: '1', must: '0', pen: us.length % 8, createdAt: now_() };
  add_('users', u);
  return { token: makeToken_(u) };
}

function view_(e, u, cs, ls) {
  return { id: e.id, authorId: e.authorId, date: e.date, title: e.title, body: e.body, mood: e.mood, weather: e.weather,
    visibility: e.visibility, photos: JSON.parse(e.photos || '[]'), edited: e.updatedAt !== e.createdAt, createdAt: e.createdAt,
    comments: cs.filter(function (c) { return c.entryId === e.id; }).map(function (c) {
      return { id: c.id, authorId: c.authorId, body: c.body, createdAt: c.createdAt };
    }),
    likes: ls.filter(function (l) { return l.entryId === e.id; }).map(function (l) { return l.userId; }) };
}
function visible_(u) {
  return rows_('entries').filter(function (e) { return canSee_(e, u); }).sort(function (a, b) {
    return (b.date + b.createdAt).localeCompare(a.date + a.createdAt);
  });
}
function ownEntry_(u, id) {
  var e = one_('entries', id);
  if (!e || !canSee_(e, u)) fail_('글을 찾을 수 없습니다.');
  return e;
}
function trash_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (x) {} }

var A = {
  me: function (u) {
    return { me: pub_(u), must: u.must === '1', members: rows_('users').map(pub_) };
  },
  feed: function (u, p) {
    var list = visible_(u).filter(function (e) {
      return (!p.authorId || e.authorId === p.authorId) && (!p.date || e.date === p.date);
    });
    var off = Number(p.offset) || 0, cs = rows_('comments'), ls = rows_('likes');
    return { items: list.slice(off, off + 15).map(function (e) { return view_(e, u, cs, ls); }), more: list.length > off + 15 };
  },
  calendar: function (u, p) {
    var m = {};
    visible_(u).forEach(function (e) {
      if (e.date.indexOf(p.month) !== 0) return;
      m[e.date] = m[e.date] || [];
      if (m[e.date].indexOf(e.authorId) < 0) m[e.date].push(e.authorId);
    });
    return m;
  },
  photo: function (u, p) {
    var e = ownEntry_(u, p.entryId);
    if (JSON.parse(e.photos || '[]').indexOf(p.fileId) < 0) fail_('사진을 찾을 수 없습니다.');
    return 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(p.fileId).getBlob().getBytes());
  },
  save: function (u, p) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date || '')) fail_('날짜를 확인하세요.');
    var title = String(p.title || '').slice(0, 100), body = String(p.body || '');
    if (body.length > 20000) fail_('본문은 20,000자까지 쓸 수 있습니다.');
    var old = p.id ? one_('entries', p.id) : null;
    if (p.id && (!old || old.authorId !== u.id)) fail_('본인 글만 수정할 수 있습니다.');
    var before = old ? JSON.parse(old.photos || '[]') : [];
    var keep = (p.keep || []).filter(function (x) { return before.indexOf(x) >= 0; });
    var fresh = p.fresh || [];
    if (keep.length + fresh.length > 4) fail_('사진은 4장까지 올릴 수 있습니다.');
    if (!body.trim() && !title.trim() && !keep.length && !fresh.length) fail_('내용을 입력하세요.');
    var folder = folder_();
    var ids = keep.concat(fresh.map(function (d) {
      var b = String(d).split(',')[1] || '';
      if (b.length > 3000000) fail_('사진 용량이 너무 큽니다.');
      return folder.createFile(Utilities.newBlob(Utilities.base64Decode(b), 'image/jpeg', uid_() + '.jpg')).getId();
    }));
    before.filter(function (x) { return keep.indexOf(x) < 0; }).forEach(trash_);
    var t = now_(), e = { id: old ? old.id : uid_(), authorId: u.id, date: p.date, title: title, body: body, mood: p.mood || '', weather: p.weather || '',
      visibility: p.visibility === 'private' ? 'private' : 'public', photos: JSON.stringify(ids), createdAt: old ? old.createdAt : t, updatedAt: t };
    if (old) write_('entries', old._r, e); else add_('entries', e);
    return e.id;
  },
  remove: function (u, p) {
    var e = ownEntry_(u, p.id);
    if (e.authorId !== u.id && u.role !== 'admin') fail_('삭제 권한이 없습니다.');
    JSON.parse(e.photos || '[]').forEach(trash_);
    delWhere_('comments', function (c) { return c.entryId === e.id; });
    delWhere_('likes', function (l) { return l.entryId === e.id; });
    del_('entries', one_('entries', e.id)._r);
  },
  comment: function (u, p) {
    ownEntry_(u, p.entryId);
    var b = String(p.body || '').trim();
    if (!b || b.length > 1000) fail_('댓글은 1~1,000자로 입력하세요.');
    add_('comments', { id: uid_(), entryId: p.entryId, authorId: u.id, body: b, createdAt: now_() });
  },
  uncomment: function (u, p) {
    var c = one_('comments', p.id);
    if (!c) return;
    var e = ownEntry_(u, c.entryId);
    if (c.authorId !== u.id && e.authorId !== u.id && u.role !== 'admin') fail_('삭제 권한이 없습니다.');
    del_('comments', c._r);
  },
  like: function (u, p) {
    ownEntry_(u, p.entryId);
    var l = rows_('likes').filter(function (x) { return x.entryId === p.entryId && x.userId === u.id; })[0];
    if (l) del_('likes', l._r); else add_('likes', { entryId: p.entryId, userId: u.id });
  },
  changePw: function (u, p) {
    if (hash_(String(p.current), u.salt) !== u.hash) fail_('현재 비밀번호가 맞지 않습니다.');
    checkPw_(p.next);
    u.salt = uid_(); u.hash = hash_(String(p.next), u.salt); u.ver = String(Number(u.ver) + 1); u.must = '0';
    write_('users', u._r, u);
    return { token: makeToken_(u) };
  },
  adminInfo: function () { return { invite: invite_() }; },
  adminInvite: function () { return { invite: setInvite_() }; },
  adminStatus: function (u, p) {
    var t = one_('users', p.id);
    if (!t || t.id === u.id) fail_('변경할 수 없는 계정입니다.');
    t.status = t.status === 'active' ? 'suspended' : 'active';
    t.ver = String(Number(t.ver) + 1);
    write_('users', t._r, t);
  },
  adminReset: function (u, p) {
    var t = one_('users', p.id);
    if (!t) fail_('계정을 찾을 수 없습니다.');
    var temp = uid_().slice(0, 10);
    t.salt = uid_(); t.hash = hash_(temp, t.salt); t.ver = String(Number(t.ver) + 1); t.must = '1';
    write_('users', t._r, t);
    return { temp: temp };
  }
};
