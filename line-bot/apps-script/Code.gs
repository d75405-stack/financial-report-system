/**
 * 里辦公處 LINE 官方帳號後端（Google Apps Script）
 *
 * 一個 Web App 同時處理兩種請求：
 *   1. LINE Messaging API 的 webhook（body 含 events 陣列）
 *   2. LIFF 頁面與管理後台的 API 呼叫（body 含 action）
 *
 * 必要的指令碼屬性（專案設定 → 指令碼屬性）：
 *   LINE_CHANNEL_ACCESS_TOKEN  Messaging API 的 Channel access token
 *   LIFF_ID                    LIFF App ID，例如 1651234567-AbCdEfGh
 * 由 setup() 自動產生（也可自行修改）：
 *   ADMIN_TOKEN                管理後台登入密碼
 *   INVITE_CODE                工作人員綁定用邀請碼
 *   PHOTO_FOLDER_ID            回報照片存放的 Google Drive 資料夾
 * 選用：
 *   NOTIFY_ADMINS              'false' 可關閉新回報／新預約時推播給里長
 */

const TZ = 'Asia/Taipei';
const ADMIN_ROLES = ['里長', '管理員'];
const REPORT_CATEGORIES = ['環境清潔', '道路路燈', '治安安全', '長者關懷', '活動支援', '其他'];
const REPORT_STATUSES = ['待處理', '處理中', '已完成', '不處理'];
const BOOKING_STATUSES = ['已預約', '已完成', '已取消', '未到'];

const SHEETS = {
  成員: ['userId', 'name', 'phone', 'role', 'group', 'status', 'joinedAt'],
  回報: ['id', 'createdAt', 'userId', 'name', 'category', 'content', 'location', 'photoUrl', 'status', 'handler', 'note', 'updatedAt'],
  公告: ['id', 'createdAt', 'target', 'title', 'content', 'recipients'],
  律師時段: ['slotId', 'date', 'start', 'end', 'lawyer', 'capacity', 'note'],
  諮詢預約: ['id', 'createdAt', 'slotId', 'userId', 'name', 'phone', 'topic', 'detail', 'status', 'reminded'],
};

// ───────────────────────── 初始化 ─────────────────────────

/** 第一次部署前在編輯器手動執行一次。 */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_TOKEN')) props.setProperty('ADMIN_TOKEN', randomToken_(24));
  if (!props.getProperty('INVITE_CODE')) props.setProperty('INVITE_CODE', randomToken_(6).toUpperCase());
  if (!props.getProperty('PHOTO_FOLDER_ID')) {
    props.setProperty('PHOTO_FOLDER_ID', DriveApp.createFolder('LINE 回報照片').getId());
  }
  const hasTrigger = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sendBookingReminders');
  if (!hasTrigger) ScriptApp.newTrigger('sendBookingReminders').timeBased().everyDays(1).atHour(18).inTimezone(TZ).create();

  console.log('管理後台密碼 ADMIN_TOKEN = ' + props.getProperty('ADMIN_TOKEN'));
  console.log('工作人員邀請碼 INVITE_CODE = ' + props.getProperty('INVITE_CODE'));
  ['LINE_CHANNEL_ACCESS_TOKEN', 'LIFF_ID'].forEach(k => {
    if (!props.getProperty(k)) console.warn('尚未設定指令碼屬性 ' + k);
  });
}

// ───────────────────────── 進入點 ─────────────────────────

function doGet() {
  return json_({ ok: true, service: 'line-bot' });
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return json_({ ok: false, error: '無效的請求' });
  }
  if (Array.isArray(body.events)) {
    body.events.forEach(ev => {
      try { handleEvent_(ev); } catch (err) { console.error(err.stack || err); }
    });
    return json_({ ok: true });
  }
  try {
    return json_({ ok: true, data: handleApi_(body) });
  } catch (err) {
    console.error(err.stack || err);
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// ───────────────────────── LINE webhook ─────────────────────────

const HELP_TEXT = [
  '可以輸入以下關鍵字：',
  '・回報：開啟回報表單（工作人員）',
  '・回報 內容：直接用文字回報',
  '・律師諮詢：查看時段並預約',
  '・我的預約：查看預約紀錄',
  '・公告：最新宣達事項',
  '・綁定 邀請碼 姓名：工作人員綁定身分',
].join('\n');

function handleEvent_(ev) {
  const uid = ev.source && ev.source.userId;
  if (!uid || ev.source.type !== 'user') return;

  if (ev.type === 'follow') {
    reply_(ev.replyToken, [text_('感謝加入里辦公處官方帳號！\n\n' + HELP_TEXT)]);
    return;
  }
  if (ev.type !== 'message' || ev.message.type !== 'text') return;

  const t = ev.message.text.trim();
  const member = findMember_(uid);

  let m;
  if ((m = t.match(/^綁定\s+(\S+)\s+(.+)$/))) {
    if (m[1].toUpperCase() !== String(prop_('INVITE_CODE')).toUpperCase()) {
      reply_(ev.replyToken, [text_('邀請碼不正確，請向里長確認。')]);
      return;
    }
    const name = m[2].trim().slice(0, 30);
    // 先回覆成功才寫入：reply token 只有 LINE 平台會發，偽造的 webhook 會回覆失敗而不會寫入資料。
    if (reply_(ev.replyToken, [text_(`綁定完成，${name} 您好！\n之後可輸入「回報」開啟回報表單。`)])) {
      upsertMember_(uid, { name });
    }
    return;
  }

  if ((m = t.match(/^回報\s+([\s\S]+)$/))) {
    if (!isActiveMember_(member)) {
      reply_(ev.replyToken, [text_('您尚未綁定工作人員身分。請輸入「綁定 邀請碼 姓名」。')]);
      return;
    }
    const id = newId_('R');
    if (reply_(ev.replyToken, [text_(`已收到回報（編號 ${id}），謝謝！`)])) {
      createReport_(member, { id, category: '其他', content: m[1].trim() });
    }
    return;
  }

  switch (t) {
    case '回報':
      reply_(ev.replyToken, [isActiveMember_(member)
        ? linkButton_('填寫回報表單，可附照片與位置。', '開啟回報表單', liffUrl_('report'))
        : linkButton_('工作人員請先綁定身分，再使用回報功能。', '綁定身分', liffUrl_('register'))]);
      return;
    case '律師諮詢':
    case '法律諮詢':
    case '預約':
      reply_(ev.replyToken, [linkButton_('免費律師諮詢，請選擇時段預約。', '查看時段並預約', liffUrl_('booking'))]);
      return;
    case '我的預約':
      reply_(ev.replyToken, [text_(myBookingsText_(uid))]);
      return;
    case '公告':
    case '最新公告':
      reply_(ev.replyToken, [text_(latestAnnouncementsText_(member))]);
      return;
    case '說明':
    case '選單':
    case 'help':
      reply_(ev.replyToken, [text_(HELP_TEXT)]);
      return;
  }
  // 其他訊息不自動回覆，留給里辦人員在官方帳號後台以聊天回覆。
}

function myBookingsText_(uid) {
  const slots = indexBy_(readAll_('律師時段'), 'slotId');
  const list = readAll_('諮詢預約').filter(b => b.userId === uid && b.status === '已預約');
  if (!list.length) return '目前沒有預約。輸入「律師諮詢」即可預約。';
  return '您的律師諮詢預約：\n' + list.map(b => {
    const s = slots[b.slotId] || {};
    return `・${s.date} ${s.start}-${s.end}｜${b.topic}`;
  }).join('\n');
}

function latestAnnouncementsText_(member) {
  const list = readAll_('公告').filter(a => canSee_(a.target, member)).slice(-5).reverse();
  if (!list.length) return '目前沒有公告。';
  return list.map(a => `【${a.title}】${a.createdAt}\n${a.content}`).join('\n\n');
}

function canSee_(target, member) {
  if (target === 'all') return true;
  if (!isActiveMember_(member)) return false;
  if (target === 'members') return true;
  return target === 'group:' + member.group;
}

// ───────────────────────── API（LIFF 與管理後台） ─────────────────────────

function handleApi_(req) {
  const action = String(req.action || '');
  if (action.indexOf('admin.') === 0) {
    if (!req.adminToken || req.adminToken !== prop_('ADMIN_TOKEN')) throw new Error('管理密碼錯誤');
    return withLock_(() => adminApi_(action.slice(6), req));
  }
  const user = verifyIdToken_(req.idToken);
  return withLock_(() => liffApi_(action, req, user));
}

function liffApi_(action, req, user) {
  const member = findMember_(user.userId);
  switch (action) {
    case 'me':
      return {
        userId: user.userId,
        displayName: user.name,
        member: isActiveMember_(member) ? pick_(member, ['name', 'role', 'group']) : null,
        categories: REPORT_CATEGORIES,
      };

    case 'register': {
      if (String(req.inviteCode || '').trim().toUpperCase() !== String(prop_('INVITE_CODE')).toUpperCase()) {
        throw new Error('邀請碼不正確');
      }
      const name = requireText_(req.name, '姓名', 30);
      upsertMember_(user.userId, { name, phone: clean_(req.phone, 20) });
      return { ok: true };
    }

    case 'submitReport': {
      if (!isActiveMember_(member)) throw new Error('請先綁定工作人員身分');
      const category = REPORT_CATEGORIES.indexOf(req.category) >= 0 ? req.category : '其他';
      const content = requireText_(req.content, '回報內容', 2000);
      const id = newId_('R');
      const photoUrl = req.photo ? savePhoto_(id, req.photo) : '';
      createReport_(member, { id, category, content, location: clean_(req.location, 300), photoUrl });
      return { id };
    }

    case 'myReports':
      return readAll_('回報').filter(r => r.userId === user.userId).slice(-20).reverse()
        .map(r => pick_(r, ['id', 'createdAt', 'category', 'content', 'status', 'note']));

    case 'listSlots':
      return openSlots_();

    case 'book': {
      const slot = openSlots_().find(s => s.slotId === req.slotId);
      if (!slot) throw new Error('此時段已額滿或不存在');
      const bookings = readAll_('諮詢預約');
      if (bookings.some(b => b.slotId === slot.slotId && b.userId === user.userId && b.status === '已預約')) {
        throw new Error('您已預約此時段');
      }
      const b = {
        id: newId_('B'),
        createdAt: now_(),
        slotId: slot.slotId,
        userId: user.userId,
        name: requireText_(req.name, '姓名', 30),
        phone: requireText_(req.phone, '聯絡電話', 20),
        topic: requireText_(req.topic, '諮詢類別', 20),
        detail: clean_(req.detail, 1000),
        status: '已預約',
        reminded: '',
      };
      append_('諮詢預約', b);
      notifyAdmins_(`新律師諮詢預約\n${slot.date} ${slot.start}-${slot.end}\n${b.name}｜${b.topic}`);
      return { id: b.id, slot };
    }

    case 'myBookings': {
      const slots = indexBy_(readAll_('律師時段'), 'slotId');
      return readAll_('諮詢預約').filter(b => b.userId === user.userId).reverse().map(b => {
        const s = slots[b.slotId] || {};
        return { id: b.id, status: b.status, topic: b.topic, date: s.date, start: s.start, end: s.end, lawyer: s.lawyer };
      });
    }

    case 'cancelBooking': {
      const b = readAll_('諮詢預約').find(x => x.id === req.id && x.userId === user.userId);
      if (!b || b.status !== '已預約') throw new Error('找不到可取消的預約');
      update_('諮詢預約', b._row, { status: '已取消' });
      return { ok: true };
    }

    case 'announcements':
      return readAll_('公告').filter(a => canSee_(a.target, member)).slice(-20).reverse()
        .map(a => pick_(a, ['createdAt', 'title', 'content']));
  }
  throw new Error('未知的動作：' + action);
}

function adminApi_(action, req) {
  switch (action) {
    case 'summary': {
      const reports = readAll_('回報');
      const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
      return {
        members: readAll_('成員').filter(isActiveMember_).length,
        pendingReports: reports.filter(r => r.status === '待處理').length,
        inProgressReports: reports.filter(r => r.status === '處理中').length,
        upcomingBookings: readAll_('諮詢預約').filter(b => b.status === '已預約').length,
        openSlots: openSlots_().length,
        today,
        quota: messageQuota_(),
        inviteCode: prop_('INVITE_CODE'),
        liffUrl: liffUrl_(''),
      };
    }

    case 'options':
      return { reportCategories: REPORT_CATEGORIES, reportStatuses: REPORT_STATUSES, bookingStatuses: BOOKING_STATUSES, adminRoles: ADMIN_ROLES };

    case 'reports':
      return readAll_('回報').reverse().map(stripRow_);

    case 'updateReport': {
      const r = readAll_('回報').find(x => x.id === req.id);
      if (!r) throw new Error('找不到回報 ' + req.id);
      const patch = { updatedAt: now_() };
      if (REPORT_STATUSES.indexOf(req.status) >= 0) patch.status = req.status;
      if (req.handler !== undefined) patch.handler = clean_(req.handler, 30);
      if (req.note !== undefined) patch.note = clean_(req.note, 1000);
      update_('回報', r._row, patch);
      if (req.notify) {
        push_(r.userId, [text_(`您的回報（${r.id}）狀態更新為「${patch.status || r.status}」${patch.note ? '\n' + patch.note : ''}`)]);
      }
      return { ok: true };
    }

    case 'members':
      return readAll_('成員').map(stripRow_);

    case 'updateMember': {
      const m = findMember_(req.userId);
      if (!m) throw new Error('找不到成員');
      const patch = {};
      ['name', 'phone', 'role', 'group'].forEach(k => { if (req[k] !== undefined) patch[k] = clean_(req[k], 30); });
      if (req.status === 'active' || req.status === 'disabled') patch.status = req.status;
      update_('成員', m._row, patch);
      return { ok: true };
    }

    case 'announcements':
      return readAll_('公告').reverse().map(stripRow_);

    case 'broadcast': {
      const title = requireText_(req.title, '標題', 50);
      const content = requireText_(req.content, '內容', 4000);
      const target = String(req.target || '');
      const message = text_(`【${title}】\n${content}`);
      let recipients;
      if (target === 'all') {
        broadcastAll_([message]);
        recipients = '所有好友';
      } else if (target === 'members' || target.indexOf('group:') === 0) {
        const group = target.slice(6);
        const ids = readAll_('成員')
          .filter(m => isActiveMember_(m) && (target === 'members' || m.group === group))
          .map(m => m.userId);
        if (!ids.length) throw new Error('沒有符合的收件人');
        multicast_(ids, [message]);
        recipients = ids.length;
      } else {
        throw new Error('未知的推播對象');
      }
      append_('公告', { id: newId_('A'), createdAt: now_(), target, title, content, recipients });
      return { recipients };
    }

    case 'slots': {
      const counts = countBy_(readAll_('諮詢預約').filter(b => b.status !== '已取消'), 'slotId');
      return readAll_('律師時段').reverse().map(s => Object.assign(stripRow_(s), { booked: counts[s.slotId] || 0 }));
    }

    case 'addSlot': {
      const slot = {
        slotId: newId_('S'),
        date: requireMatch_(req.date, /^\d{4}-\d{2}-\d{2}$/, '日期'),
        start: requireMatch_(req.start, /^\d{2}:\d{2}$/, '開始時間'),
        end: requireMatch_(req.end, /^\d{2}:\d{2}$/, '結束時間'),
        lawyer: clean_(req.lawyer, 30),
        capacity: Math.max(1, Math.min(50, parseInt(req.capacity, 10) || 1)),
        note: clean_(req.note, 200),
      };
      append_('律師時段', slot);
      return slot;
    }

    case 'deleteSlot': {
      const s = readAll_('律師時段').find(x => x.slotId === req.slotId);
      if (!s) throw new Error('找不到時段');
      if (readAll_('諮詢預約').some(b => b.slotId === s.slotId && b.status === '已預約')) {
        throw new Error('此時段已有預約，請先取消預約');
      }
      sheet_('律師時段').deleteRow(s._row);
      return { ok: true };
    }

    case 'bookings': {
      const slots = indexBy_(readAll_('律師時段'), 'slotId');
      return readAll_('諮詢預約').reverse().map(b => {
        const s = slots[b.slotId] || {};
        return Object.assign(stripRow_(b), { date: s.date, start: s.start, end: s.end, lawyer: s.lawyer });
      });
    }

    case 'updateBooking': {
      const b = readAll_('諮詢預約').find(x => x.id === req.id);
      if (!b) throw new Error('找不到預約');
      if (BOOKING_STATUSES.indexOf(req.status) < 0) throw new Error('狀態不正確');
      update_('諮詢預約', b._row, { status: req.status });
      return { ok: true };
    }
  }
  throw new Error('未知的管理動作：' + action);
}

// ───────────────────────── 業務邏輯 ─────────────────────────

function findMember_(userId) {
  return readAll_('成員').find(m => m.userId === userId) || null;
}

function isActiveMember_(m) {
  return !!m && m.status === 'active';
}

function upsertMember_(userId, fields) {
  const m = findMember_(userId);
  if (m) {
    update_('成員', m._row, Object.assign({}, fields, m.status === 'disabled' ? {} : { status: 'active' }));
  } else {
    append_('成員', Object.assign({ userId, role: '成員', group: '', status: 'active', joinedAt: now_() }, fields));
  }
}

function createReport_(member, r) {
  const row = Object.assign({
    createdAt: now_(), userId: member.userId, name: member.name,
    location: '', photoUrl: '', status: '待處理', handler: '', note: '', updatedAt: '',
  }, r);
  append_('回報', row);
  notifyAdmins_(`新回報 ${row.id}｜${row.category}\n回報人：${row.name}\n${row.content.slice(0, 200)}${row.location ? '\n地點：' + row.location : ''}`);
}

function notifyAdmins_(message) {
  if (prop_('NOTIFY_ADMINS') === 'false') return;
  const ids = readAll_('成員').filter(m => isActiveMember_(m) && ADMIN_ROLES.indexOf(m.role) >= 0).map(m => m.userId);
  if (ids.length) multicast_(ids, [text_(message)]);
}

function openSlots_() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const counts = countBy_(readAll_('諮詢預約').filter(b => b.status === '已預約' || b.status === '已完成'), 'slotId');
  return readAll_('律師時段')
    .filter(s => s.date >= today)
    .map(s => ({
      slotId: s.slotId, date: s.date, start: s.start, end: s.end, lawyer: s.lawyer, note: s.note,
      remaining: Number(s.capacity || 1) - (counts[s.slotId] || 0),
    }))
    .filter(s => s.remaining > 0)
    .sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
}

function savePhoto_(id, photo) {
  const m = String(photo).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('照片格式不支援');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('照片太大');
  const blob = Utilities.newBlob(bytes, m[1], id + '.' + m[1].split('/')[1]);
  // 照片維持私人權限，只有 Drive 擁有者（里辦帳號）可開啟連結。
  return DriveApp.getFolderById(prop_('PHOTO_FOLDER_ID')).createFile(blob).getUrl();
}

/** 每日由時間觸發器執行：提醒明天的律師諮詢預約。 */
function sendBookingReminders() {
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 864e5), TZ, 'yyyy-MM-dd');
  const slots = indexBy_(readAll_('律師時段'), 'slotId');
  readAll_('諮詢預約').forEach(b => {
    const s = slots[b.slotId];
    if (!s || s.date !== tomorrow || b.status !== '已預約' || b.reminded) return;
    push_(b.userId, [text_(`提醒您：明天 ${s.date} ${s.start}-${s.end} 有律師諮詢預約（${b.topic}）。\n如需取消請輸入「律師諮詢」進入頁面取消。`)]);
    update_('諮詢預約', b._row, { reminded: now_() });
  });
}

// ───────────────────────── LINE API ─────────────────────────

function lineApi_(path, payload, method) {
  const options = {
    method: method || 'post',
    headers: { Authorization: 'Bearer ' + prop_('LINE_CHANNEL_ACCESS_TOKEN') },
    muteHttpExceptions: true,
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, options);
  const code = res.getResponseCode();
  if (code >= 300) console.error('LINE API ' + path + ' ' + code + ' ' + res.getContentText());
  return { ok: code < 300, code, body: res.getContentText() };
}

/** 回傳是否成功；失敗代表 reply token 無效（可能是偽造的 webhook）。 */
function reply_(replyToken, messages) {
  return lineApi_('message/reply', { replyToken, messages }).ok;
}

function push_(to, messages) {
  return lineApi_('message/push', { to, messages }).ok;
}

function multicast_(ids, messages) {
  for (let i = 0; i < ids.length; i += 500) {
    const res = lineApi_('message/multicast', { to: ids.slice(i, i + 500), messages });
    if (!res.ok) throw new Error('推播失敗：' + res.body);
  }
}

function broadcastAll_(messages) {
  const res = lineApi_('message/broadcast', { messages });
  if (!res.ok) throw new Error('推播失敗：' + res.body);
}

function messageQuota_() {
  try {
    const q = lineApi_('message/quota', null, 'get');
    const c = lineApi_('message/quota/consumption', null, 'get');
    if (!q.ok || !c.ok) return null;
    const quota = JSON.parse(q.body);
    return { type: quota.type, limit: quota.value, used: JSON.parse(c.body).totalUsage };
  } catch (err) {
    return null;
  }
}

/** 驗證 LIFF 的 ID token，回傳 LINE userId 與顯示名稱。 */
function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('請從 LINE 開啟此頁面');
  const cache = CacheService.getScriptCache();
  const key = 'idt_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: String(prop_('LIFF_ID')).split('-')[0] },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('登入已過期，請重新開啟頁面');
  const p = JSON.parse(res.getContentText());
  const user = { userId: p.sub, name: p.name || '' };
  const ttl = Math.min(600, Math.floor(p.exp - Date.now() / 1000));
  if (ttl > 0) cache.put(key, JSON.stringify(user), ttl);
  return user;
}

function text_(text) {
  return { type: 'text', text: String(text).slice(0, 5000) };
}

function linkButton_(text, label, uri) {
  return {
    type: 'template',
    altText: text,
    template: { type: 'buttons', text: text.slice(0, 160), actions: [{ type: 'uri', label: label.slice(0, 20), uri }] },
  };
}

function liffUrl_(page) {
  return 'https://liff.line.me/' + prop_('LIFF_ID') + (page ? '?page=' + page : '');
}

// ───────────────────────── 試算表存取 ─────────────────────────

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 全部以純文字儲存，避免日期、電話號碼被自動轉換。
    sh.getRange(1, 1, sh.getMaxRows(), SHEETS[name].length).setNumberFormat('@');
  }
  return sh;
}

function readAll_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const header = values.shift() || [];
  return values
    .map((row, i) => {
      const o = { _row: i + 2 };
      header.forEach((k, j) => { o[k] = normalize_(k, row[j]); });
      return o;
    })
    .filter(o => header.some(k => o[k] !== ''));
}

function normalize_(key, v) {
  if (v instanceof Date) {
    if (key === 'date') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    if (key === 'start' || key === 'end') return Utilities.formatDate(v, TZ, 'HH:mm');
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  }
  return v === null || v === undefined ? '' : String(v);
}

function append_(name, obj) {
  sheet_(name).appendRow(SHEETS[name].map(k => (obj[k] === undefined || obj[k] === null ? '' : String(obj[k]))));
}

function update_(name, row, patch) {
  const sh = sheet_(name);
  Object.keys(patch).forEach(k => {
    const col = SHEETS[name].indexOf(k);
    if (col >= 0) sh.getRange(row, col + 1).setValue(String(patch[k]));
  });
}

// ───────────────────────── 工具函式 ─────────────────────────

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
}

function randomToken_(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

function clean_(v, max) {
  return String(v === undefined || v === null ? '' : v).trim().slice(0, max);
}

function requireText_(v, label, max) {
  const s = clean_(v, max);
  if (!s) throw new Error('請填寫' + label);
  return s;
}

function requireMatch_(v, re, label) {
  const s = clean_(v, 20);
  if (!re.test(s)) throw new Error(label + '格式不正確');
  return s;
}

function pick_(o, keys) {
  const r = {};
  keys.forEach(k => { r[k] = o[k]; });
  return r;
}

function stripRow_(o) {
  const r = Object.assign({}, o);
  delete r._row;
  return r;
}

function indexBy_(list, key) {
  const r = {};
  list.forEach(o => { r[o[key]] = o; });
  return r;
}

function countBy_(list, key) {
  const r = {};
  list.forEach(o => { r[o[key]] = (r[o[key]] || 0) + 1; });
  return r;
}
