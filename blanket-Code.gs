/** 담요 공용제 DB + (선택) 돌림판 로그 통합 웹앱 */
var BLANKET_SHEET_ID = '1xnmsPFoRoRpNiNYMSjDWAxBuEY17yXBZuSp0VP5Bqz8';
var LUCKY_SHEET_ID = '1xnmsPFoRoRpNiNYMSjDWAxBuEY17yXBZuSp0VP5Bqz8';
var ADMIN_PIN = '2825';

/** 시트 탭 이름 (한글) */
var SHEET = {
  settings: '설정',
  bookings: '예약',
  suggestions: '건의함',
  blacklist: '페널티',
  notices: '공지사항'
};

/** 설정 키 (시트 표시용 한글 ↔ API 영문) */
var SETTING_KEY = {
  notice: '공지사항',
  washDate: '세탁일',
  quota: '하루정원',
  year: '연도',
  month: '월',
  datesRepaired: '날짜보정완료'
};

/** 예약 칼럼 */
var BOOKING_HEADERS = [
  '예약ID', '날짜', '수령시간', '시간코드', '학번', '이름', '전화번호',
  '상태', '취소사유', '기기키', '등록시각', '수정시각'
];
var BOOKING_KEYS = [
  'id', 'date', 'time', 'timeKey', 'studentId', 'name', 'phone',
  'status', 'cancelReason', 'deviceKey', 'createdAt', 'updatedAt'
];

var SUGGEST_HEADERS = ['건의ID', '내용', '등록시각'];
var SUGGEST_KEYS = ['id', 'content', 'createdAt'];

var BLACK_HEADERS = ['학번', '사유', '등록일', '등록시각'];
var BLACK_KEYS = ['studentId', 'reason', 'date', 'createdAt'];

var SETTINGS_HEADERS = ['항목', '값'];
var NOTICE_HEADERS = ['내용', '등록시각'];

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function blanketSs_() {
  return SpreadsheetApp.openById(BLANKET_SHEET_ID);
}

function renameIfExists_(ss, fromName, toName) {
  var sh = ss.getSheetByName(fromName);
  if (sh && !ss.getSheetByName(toName)) {
    sh.setName(toName);
  }
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    var first = String(sh.getRange(1, 1).getValue() || '');
    // 영문 헤더였거나 비어 있으면 한글로 교체
    if (!first || first === 'key' || first === 'id' || first === 'studentId') {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function ensureBlanketSchema_() {
  var ss = blanketSs_();

  // 기존 영문 탭 → 한글 탭으로 이름 변경
  renameIfExists_(ss, 'Settings', SHEET.settings);
  renameIfExists_(ss, 'Bookings', SHEET.bookings);
  renameIfExists_(ss, 'Suggestions', SHEET.suggestions);
  renameIfExists_(ss, 'Blacklist', SHEET.blacklist);
  renameIfExists_(ss, '공지', SHEET.notices);

  ensureSheet_(ss, SHEET.settings, SETTINGS_HEADERS);
  ensureSheet_(ss, SHEET.bookings, BOOKING_HEADERS);
  ensureSheet_(ss, SHEET.suggestions, SUGGEST_HEADERS);
  ensureSheet_(ss, SHEET.blacklist, BLACK_HEADERS);
  ensureSheet_(ss, SHEET.notices, NOTICE_HEADERS);

  // 설정 키도 한글로 마이그레이션
  migrateSettingKeys_(ss.getSheetByName(SHEET.settings));

  var settings = ss.getSheetByName(SHEET.settings);
  var defaults = {
    notice: '담요는 당일 종례 전까지 반드시 반납해 주세요.',
    washDate: '10월 25일',
    quota: '9'
  };
  Object.keys(defaults).forEach(function (k) {
    if (getSetting_(settings, k) === '') setSetting_(settings, k, defaults[k]);
  });

  // 공지사항 탭이 비어 있으면 설정값으로 1회 시드 (이후엔 공지사항 탭 최상단 사용)
  seedNoticesFromSettings_(ss);

  // 날짜 보정은 1회만 (매 요청마다 하면 취소/예약이 멈춤)
  if (getSetting_(settings, 'datesRepaired') !== '1') {
    repairBookingDates_(ss);
    setSetting_(settings, 'datesRepaired', '1');
  }
  return ss;
}

function migrateSettingKeys_(sh) {
  var data = sh.getDataRange().getValues();
  var enToKo = {
    key: '항목',
    value: '값',
    notice: '공지사항',
    washDate: '세탁일',
    quota: '하루정원',
    year: '연도',
    month: '월'
  };
  for (var i = 0; i < data.length; i++) {
    var k = String(data[i][0] || '');
    if (enToKo[k]) {
      sh.getRange(i + 1, 1).setValue(enToKo[k]);
    }
  }
}

function settingLabel_(apiKey) {
  return SETTING_KEY[apiKey] || apiKey;
}

function getSetting_(sh, apiKey) {
  var label = settingLabel_(apiKey);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '');
    if (k === label || k === apiKey) {
      var v = data[i][1];
      if (v == null || v === '') return '';
      if (Object.prototype.toString.call(v) === '[object Date]') {
        return Utilities.formatDate(v, 'Asia/Seoul', 'M월 d일');
      }
      return String(v);
    }
  }
  return '';
}

function setSetting_(sh, apiKey, value) {
  var label = settingLabel_(apiKey);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '');
    if (k === label || k === apiKey) {
      sh.getRange(i + 1, 1).setValue(label);
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String(value));
      return;
    }
  }
  sh.appendRow([label, String(value)]);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@');
}

/** 설정 시트의 공지사항 키 중 가장 아래(최신) 값 */
function getLatestSettingNotice_(sh) {
  var label = settingLabel_('notice');
  var data = sh.getDataRange().getValues();
  var found = '';
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][0] || '');
    if (k !== label && k !== 'notice') continue;
    var v = data[i][1];
    if (v == null || v === '') continue;
    if (Object.prototype.toString.call(v) === '[object Date]') {
      found = Utilities.formatDate(v, 'Asia/Seoul', 'M월 d일');
    } else {
      found = String(v);
    }
  }
  return found;
}

function seedNoticesFromSettings_(ss) {
  var sh = ss.getSheetByName(SHEET.notices);
  if (!sh || sh.getLastRow() >= 2) return;
  var text = getLatestSettingNotice_(ss.getSheetByName(SHEET.settings));
  if (!text) return;
  sh.appendRow([text, new Date()]);
  sh.getRange(2, 1).setNumberFormat('@');
}

/**
 * 학생/관리자에 보여줄 공지:
 * 1) 공지사항 탭 최상단(2행)
 * 2) 없으면 설정 탭 공지사항 키 중 최신(가장 아래) 값
 */
function getLatestNotice_(ss) {
  var sh = ss.getSheetByName(SHEET.notices);
  if (sh && sh.getLastRow() >= 2) {
    var top = sh.getRange(2, 1).getDisplayValue();
    if (top != null && String(top).trim() !== '') return String(top).trim();
  }
  return getLatestSettingNotice_(ss.getSheetByName(SHEET.settings));
}

/** 새 공지를 공지사항 탭 최상단에 넣고, 설정값도 동기화 */
function pushNotice_(ss, text) {
  var value = String(text == null ? '' : text);
  var sh = ensureSheet_(ss, SHEET.notices, NOTICE_HEADERS);
  // 최상단과 내용이 같으면 중복 행만 만들지 않음
  if (sh.getLastRow() >= 2) {
    var cur = String(sh.getRange(2, 1).getDisplayValue() || '').trim();
    if (cur === value.trim()) {
      setSetting_(ss.getSheetByName(SHEET.settings), 'notice', value);
      return;
    }
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1).setNumberFormat('@').setValue(value);
  sh.getRange(2, 2).setValue(new Date());
  setSetting_(ss.getSheetByName(SHEET.settings), 'notice', value);
}

function uid_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

/** 시트 Date/문자열 → YYYY-MM-DD (서울 기준) */
function toYmd_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return s;
}

function asText_(v) {
  if (v == null) return '';
  return String(v);
}

function rowsToObjects_(sh, koHeaders, enKeys) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < enKeys.length; c++) {
      var ko = koHeaders[c];
      var en = enKeys[c];
      var col = headers.indexOf(ko);
      if (col < 0) col = headers.indexOf(en);
      if (col < 0) col = c;
      obj[en] = values[r][col];
    }
    out.push(obj);
  }
  return out;
}

function normalizeBooking_(b) {
  return {
    id: asText_(b.id),
    date: toYmd_(b.date),
    time: asText_(b.time),
    timeKey: asText_(b.timeKey),
    studentId: asText_(b.studentId),
    name: asText_(b.name),
    phone: asText_(b.phone),
    status: asText_(b.status) || '예약완료',
    cancelReason: asText_(b.cancelReason),
    deviceKey: asText_(b.deviceKey),
    createdAt: b.createdAt || '',
    updatedAt: b.updatedAt || ''
  };
}

function bookings_(ss) {
  return rowsToObjects_(ss.getSheetByName(SHEET.bookings), BOOKING_HEADERS, BOOKING_KEYS)
    .map(normalizeBooking_);
}

function suggestions_(ss) {
  return rowsToObjects_(ss.getSheetByName(SHEET.suggestions), SUGGEST_HEADERS, SUGGEST_KEYS).map(function (s) {
    return {
      id: asText_(s.id),
      content: asText_(s.content),
      createdAt: s.createdAt || ''
    };
  });
}

function blacklist_(ss) {
  return rowsToObjects_(ss.getSheetByName(SHEET.blacklist), BLACK_HEADERS, BLACK_KEYS).map(function (b) {
    return {
      studentId: asText_(b.studentId),
      reason: asText_(b.reason),
      date: asText_(b.date),
      createdAt: b.createdAt || ''
    };
  });
}

/** 기존 예약 날짜 칸을 텍스트(YYYY-MM-DD)로 보정 */
function repairBookingDates_(ss) {
  var sh = ss.getSheetByName(SHEET.bookings);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  for (var r = 1; r < values.length; r++) {
    var ymd = toYmd_(values[r][1]);
    if (ymd) {
      sh.getRange(r + 1, 2).setNumberFormat('@').setValue(ymd);
    }
    // 학번/전화도 텍스트
    if (values[r][4] !== '' && values[r][4] != null) {
      sh.getRange(r + 1, 5).setNumberFormat('@').setValue(asText_(values[r][4]));
    }
    if (values[r][6] !== '' && values[r][6] != null) {
      sh.getRange(r + 1, 7).setNumberFormat('@').setValue(asText_(values[r][6]));
    }
  }
}

function activeBookings_(bookings) {
  return bookings.filter(function (b) {
    return String(b.status) === '예약완료';
  });
}

function countByDate_(bookings) {
  var map = {};
  activeBookings_(bookings).forEach(function (b) {
    var d = toYmd_(b.date);
    if (!d) return;
    map[d] = (map[d] || 0) + 1;
  });
  return map;
}

function findBlacklist_(list, studentId) {
  studentId = String(studentId);
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].studentId) === studentId) return list[i];
  }
  return null;
}

function isoWeekKey_(dateStr) {
  var p = dateStr.split('-').map(Number);
  var d = new Date(p[0], p[1] - 1, p[2]);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  var week1 = new Date(d.getFullYear(), 0, 4);
  var week = 1 + Math.round(((d - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
  return d.getFullYear() + '-W' + week;
}

function requirePin_(pin) {
  return String(pin || '') === ADMIN_PIN;
}

function handleLucky_(data) {
  var ss = SpreadsheetApp.openById(LUCKY_SHEET_ID);
  var sheet = ss.getSheets()[0];
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.appendRow(['순번', '당첨일', '일련번호', '확인코드', '상품', '이모지', '기록시각', 'UserAgent']);
  }
  var winNo = Math.max(1, sheet.getLastRow());
  var total = data.totalWinners || 24;
  var serial = data.serial || '';
  sheet.appendRow([
    winNo + ' / ' + total,
    data.winDate || '',
    serial,
    data.code || '',
    data.prize || '',
    data.emoji || '',
    new Date(),
    data.userAgent || ''
  ]);
  return { ok: true, winNo: winNo, app: 'lucky-wheel' };
}

function findMyBooking_(bookings, deviceKey, studentId) {
  var active = activeBookings_(bookings);
  if (deviceKey) {
    for (var i = active.length - 1; i >= 0; i--) {
      if (String(active[i].deviceKey) === String(deviceKey)) return active[i];
    }
  }
  if (studentId) {
    for (var j = active.length - 1; j >= 0; j--) {
      if (String(active[j].studentId) === String(studentId)) return active[j];
    }
  }
  return null;
}

function currentYearMonth_() {
  var now = new Date();
  return {
    year: parseInt(Utilities.formatDate(now, 'Asia/Seoul', 'yyyy'), 10),
    month: parseInt(Utilities.formatDate(now, 'Asia/Seoul', 'M'), 10)
  };
}

function bootstrap_(deviceKey, studentId) {
  var ss = ensureBlanketSchema_();
  var settings = ss.getSheetByName(SHEET.settings);
  var bookings = bookings_(ss);
  var quota = parseInt(getSetting_(settings, 'quota') || '9', 10);
  var ym = currentYearMonth_();
  return {
    ok: true,
    notice: getLatestNotice_(ss),
    washDate: getSetting_(settings, 'washDate'),
    quota: quota,
    year: ym.year,
    month: ym.month,
    counts: countByDate_(bookings),
    myBooking: findMyBooking_(bookings, deviceKey, studentId)
  };
}

function lookupBooking_(data) {
  var ss = ensureBlanketSchema_();
  var studentId = String(data.studentId || '').trim();
  var name = String(data.name || '').trim();
  if (!studentId) return { ok: false, error: '학번을 입력해주세요.' };
  var bookings = bookings_(ss);
  var active = activeBookings_(bookings);
  var found = null;
  for (var i = active.length - 1; i >= 0; i--) {
    if (String(active[i].studentId) !== studentId) continue;
    if (name && String(active[i].name) !== name) continue;
    found = active[i];
    break;
  }
  if (!found) return { ok: false, error: '예약 내역이 없습니다.' };
  return { ok: true, booking: found };
}

function createBooking_(data) {
  var ss = ensureBlanketSchema_();
  var settings = ss.getSheetByName(SHEET.settings);
  var bookSh = ss.getSheetByName(SHEET.bookings);
  var list = bookings_(ss);
  var bl = blacklist_(ss);
  var quota = parseInt(getSetting_(settings, 'quota') || '9', 10);

  var date = String(data.date || '');
  var studentId = String(data.studentId || '').trim();
  var name = String(data.name || '').trim();
  var phone = String(data.phone || '').trim();
  var time = String(data.time || '');
  var timeKey = String(data.timeKey || '');
  var deviceKey = String(data.deviceKey || '');

  if (!date || !studentId || !name || !phone || !time || !timeKey || !deviceKey) {
    return { ok: false, error: '필수 값이 없습니다.' };
  }

  var blocked = findBlacklist_(bl, studentId);
  if (blocked) {
    return { ok: false, error: '페널티로 예약이 정지된 학번입니다. (' + blocked.reason + ')' };
  }

  var active = activeBookings_(list);
  if (active.some(function (b) { return String(b.deviceKey) === deviceKey; })) {
    return { ok: false, error: '이 기기에 이미 예약이 있습니다.' };
  }
  if (active.some(function (b) { return String(b.studentId) === studentId && toYmd_(b.date) === (toYmd_(date) || date); })) {
    return { ok: false, error: '같은 날짜에 이미 예약한 학번입니다. (1일 1회)' };
  }

  var weekKey = isoWeekKey_(toYmd_(date) || date);
  var weekCount = active.filter(function (b) {
    return String(b.studentId) === studentId && isoWeekKey_(toYmd_(b.date)) === weekKey;
  }).length;
  if (weekCount >= 2) {
    return { ok: false, error: '주 2회 대여 제한을 초과했습니다.' };
  }

  var dayCount = active.filter(function (b) { return toYmd_(b.date) === (toYmd_(date) || date); }).length;
  if (dayCount >= quota) {
    return { ok: false, error: '해당 날짜 잔여 담요가 없습니다.' };
  }

  var now = new Date();
  var id = uid_('bk');
  var ymd = toYmd_(date) || date;
  bookSh.appendRow([
    id, ymd, time, timeKey, studentId, name, phone,
    '예약완료', '', deviceKey, now, now
  ]);
  var row = bookSh.getLastRow();
  // 날짜/학번/전화는 반드시 텍스트로 저장 (시트가 Date/숫자로 바꾸지 못하게)
  bookSh.getRange(row, 2).setNumberFormat('@').setValue(ymd);
  bookSh.getRange(row, 5).setNumberFormat('@').setValue(studentId);
  bookSh.getRange(row, 7).setNumberFormat('@').setValue(phone);

  return {
    ok: true,
    booking: {
      id: id, date: ymd, time: time, timeKey: timeKey,
      studentId: studentId, name: name, phone: phone,
      status: '예약완료', deviceKey: deviceKey
    },
    counts: countByDate_(bookings_(ss))
  };
}

function cancelBooking_(data) {
  var ss = ensureBlanketSchema_();
  var bookSh = ss.getSheetByName(SHEET.bookings);
  var values = bookSh.getDataRange().getValues();
  var deviceKey = String(data.deviceKey || '');
  var bookingId = String(data.bookingId || '');
  var studentId = String(data.studentId || '');
  var reason = String(data.cancelReason || '');

  for (var r = 1; r < values.length; r++) {
    var id = String(values[r][0]);
    var status = String(values[r][7]);
    var dk = String(values[r][9]);
    var sid = String(values[r][4]);
    var match = false;
    if (bookingId && id === bookingId) match = true;
    else if (deviceKey && dk === deviceKey) match = true;
    else if (studentId && sid === studentId) match = true;
    if (match && status === '예약완료') {
      bookSh.getRange(r + 1, 8).setValue('예약취소');
      bookSh.getRange(r + 1, 9).setValue(reason);
      bookSh.getRange(r + 1, 12).setValue(new Date());
      return { ok: true, counts: countByDate_(bookings_(ss)) };
    }
  }
  return { ok: false, error: '취소할 예약을 찾지 못했습니다.' };
}

function addSuggestion_(data) {
  var ss = ensureBlanketSchema_();
  var content = String(data.content || '').trim();
  if (!content) return { ok: false, error: '내용을 입력해주세요.' };
  ss.getSheetByName(SHEET.suggestions).appendRow([uid_('sg'), content, new Date()]);
  return { ok: true };
}

function saveSettings_(data) {
  if (!requirePin_(data.pin)) return { ok: false, error: 'PIN 오류' };
  var ss = ensureBlanketSchema_();
  var sh = ss.getSheetByName(SHEET.settings);
  if (data.notice != null) pushNotice_(ss, String(data.notice));
  if (data.washDate != null) setSetting_(sh, 'washDate', String(data.washDate));
  if (data.quota != null) setSetting_(sh, 'quota', String(data.quota));
  // 연도/월은 시트 설정 없이 항상 현재일 기준
  return { ok: true };
}

function addPenalty_(data) {
  if (!requirePin_(data.pin)) return { ok: false, error: 'PIN 오류' };
  var ss = ensureBlanketSchema_();
  var studentId = String(data.studentId || '').trim();
  var reason = String(data.reason || '').trim();
  if (!studentId || !reason) return { ok: false, error: '학번/사유 필요' };
  var list = blacklist_(ss);
  if (findBlacklist_(list, studentId)) return { ok: false, error: '이미 등록된 학번입니다.' };
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'M월 d일');
  ss.getSheetByName(SHEET.blacklist).appendRow([studentId, reason, today, new Date()]);
  return { ok: true, blacklist: blacklist_(ss) };
}

function removePenalty_(data) {
  if (!requirePin_(data.pin)) return { ok: false, error: 'PIN 오류' };
  var ss = ensureBlanketSchema_();
  var sh = ss.getSheetByName(SHEET.blacklist);
  var values = sh.getDataRange().getValues();
  var studentId = String(data.studentId || '');
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][0]) === studentId) {
      sh.deleteRow(r + 1);
      return { ok: true, blacklist: blacklist_(ss) };
    }
  }
  return { ok: false, error: '해당 학번이 없습니다.' };
}

function updateBookingStatus_(data) {
  if (!requirePin_(data.pin)) return { ok: false, error: 'PIN 오류' };
  var ss = ensureBlanketSchema_();
  var sh = ss.getSheetByName(SHEET.bookings);
  var values = sh.getDataRange().getValues();
  var bookingId = String(data.bookingId || '');
  var status = String(data.status || '');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === bookingId) {
      sh.getRange(r + 1, 8).setValue(status);
      sh.getRange(r + 1, 12).setValue(new Date());
      return { ok: true, bookings: bookings_(ss) };
    }
  }
  return { ok: false, error: '예약을 찾지 못했습니다.' };
}

function adminBootstrap_(pin) {
  if (!requirePin_(pin)) return { ok: false, error: 'PIN 오류' };
  var ss = ensureBlanketSchema_();
  var settings = ss.getSheetByName(SHEET.settings);
  var list = bookings_(ss);
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var todayList = list.filter(function (b) { return toYmd_(b.date) === today; });
  var upcomingList = activeBookings_(list).filter(function (b) {
    return toYmd_(b.date) >= today;
  });
  var unreturnedList = activeBookings_(list).filter(function (b) {
    return toYmd_(b.date) < today;
  });
  var ym = currentYearMonth_();
  return {
    ok: true,
    notice: getLatestNotice_(ss),
    washDate: getSetting_(settings, 'washDate'),
    quota: parseInt(getSetting_(settings, 'quota') || '9', 10),
    year: ym.year,
    month: ym.month,
    today: today,
    bookings: list,
    todayBookings: todayList,
    upcomingBookings: upcomingList,
    unreturnedList: unreturnedList,
    unreturned: unreturnedList.length,
    blacklist: blacklist_(ss),
    suggestions: suggestions_(ss).reverse()
  };
}

function route_(payload, params) {
  var action = (payload && payload.action) || (params && params.action) || '';
  var data = payload || params || {};

  if (!action && (data.prize || data.winNo || data.serial)) {
    return handleLucky_(data);
  }

  switch (action) {
    case 'ping': return { ok: true, app: 'blanket' };
    case 'bootstrap': return bootstrap_(data.deviceKey, data.studentId);
    case 'lookupBooking': return lookupBooking_(data);
    case 'createBooking': return createBooking_(data);
    case 'cancelBooking': return cancelBooking_(data);
    case 'addSuggestion': return addSuggestion_(data);
    case 'saveSettings': return saveSettings_(data);
    case 'addPenalty': return addPenalty_(data);
    case 'removePenalty': return removePenalty_(data);
    case 'updateBookingStatus': return updateBookingStatus_(data);
    case 'adminBootstrap': return adminBootstrap_(data.pin);
    default: return { ok: false, error: 'unknown action: ' + action };
  }
}

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    return jsonOut_(route_(null, params));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var data = JSON.parse(raw);
    return jsonOut_(route_(data, {}));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
