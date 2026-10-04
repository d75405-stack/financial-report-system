/**
 * 里辦公處 LINE 官方帳號後端（Google Apps Script）
 *
 * 一個 Web App 同時處理兩種請求：
 *   1. LINE Messaging API 的 webhook（body 含 events 陣列）
 *   2. LIFF 頁面與管理後台的 API 呼叫（body 含 action）
 *
 * 設定方式：在試算表選單「LINE 系統 → 一鍵設定」依序貼上設定值即可，
 * 會寫入以下指令碼屬性（也可在「專案設定 → 指令碼屬性」直接修改）。
 *
 * 必要：
 *   LINE_CHANNEL_ACCESS_TOKEN  Messaging API 的 Channel access token
 *   LIFF_ID                    LIFF App ID，例如 1651234567-AbCdEfGh
 * 由 setup() 自動產生（也可自行修改）：
 *   ADMIN_TOKEN                管理後台登入密碼
 *   INVITE_CODE                工作人員綁定用邀請碼
 * 選用：
 *   WEBAPP_URL                 這個網頁應用程式的 /exec 網址，用來自動設定 LINE webhook
 *   BACKUP_FOLDER_ID           備份用的 Google Drive 資料夾 ID（或資料夾網址）。
 *                              設定後每天自動備份並分類，回報照片也會存到這裡。
 *   NOTIFY_ADMINS              'false' 可關閉新回報／新預約時推播給里長
 * 系統自動維護：
 *   PHOTO_FOLDER_ID            未設定 BACKUP_FOLDER_ID 時，回報照片存放的資料夾
 *   LAST_BACKUP_AT／LAST_BACKUP_ERROR／BACKUP_RUNNING  備份狀態
 *   RICH_MENU_ID               目前使用中的圖文選單
 *   SETUP_OWNER                執行設定的帳號，只有這個帳號能重新設定
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

/**
 * 建立工作表、密碼、排程，並自動完成 LINE 與雲端硬碟的設定。
 * 平常從試算表選單「LINE 系統 → 一鍵設定」執行；可重複執行，不會覆蓋已有的資料。
 */
function setup() {
  checkOwner_();
  Object.keys(SHEETS).forEach(sheet_);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_TOKEN')) props.setProperty('ADMIN_TOKEN', randomToken_(24));
  if (!props.getProperty('INVITE_CODE')) props.setProperty('INVITE_CODE', randomToken_(6).toUpperCase());
  ensureDailyTrigger_('sendBookingReminders', 18);
  ensureDailyTrigger_('backupToDrive', 2);

  const lines = ['✅ 已建立工作表與每日排程（18:00 預約提醒、02:00 雲端備份）'];
  const hasToken = !!prop_('LINE_CHANNEL_ACCESS_TOKEN');
  if (hasToken && prop_('WEBAPP_URL')) lines.push(setWebhookEndpoint_());
  if (hasToken && prop_('LIFF_ID')) {
    try {
      lines.push(setupRichMenu());
    } catch (err) {
      lines.push('⚠️ 圖文選單：' + err.message);
    }
  }
  if (prop_('BACKUP_FOLDER_ID')) {
    try {
      lines.push('✅ 第一次備份完成：' + backupToDrive().snapshot);
    } catch (err) {
      lines.push('❌ 備份沒有執行：' + err.message);
    }
  }
  lines.push('', '── 目前狀態 ──');
  lines.push.apply(lines, statusLines_());
  lines.push('', '管理後台密碼與工作人員邀請碼：請點「LINE 系統 → 查看設定狀態」。');
  console.log(lines.join('\n'));
  return lines;
}

/**
 * 排程和備份都以執行設定的帳號身分運作，所以只讓第一次設定的帳號重新設定，
 * 避免其他共用試算表的人多建一組排程、把備份寫到別人的權限下。
 */
function checkOwner_() {
  const email = Session.getEffectiveUser().getEmail();
  const owner = prop_('SETUP_OWNER');
  if (!email) return;
  if (!owner) PropertiesService.getScriptProperties().setProperty('SETUP_OWNER', email);
  else if (owner !== email) throw new Error('只有第一次設定的帳號（' + owner + '）可以執行設定');
}

function ensureDailyTrigger_(handler, hour) {
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === handler)) return;
  ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(hour).inTimezone(TZ).create();
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
      reply_(ev.replyToken, [
        text_(myBookingsText_(uid)),
        linkButton_('要預約、取消或改時段，請開啟預約頁面。', '開啟預約頁面', liffUrl_('booking')),
      ]);
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
    // 備份可能要跑數十秒，不佔用全域鎖，避免 LIFF 與後台其他操作等候逾時。
    if (action === 'admin.backup') {
      backupRoot_();
      return backupToDrive();
    }
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
      const photo = req.photo ? decodePhoto_(id, req.photo) : null;
      let photoUrl = '';
      let photoError = '';
      if (photo) {
        // 雲端硬碟出問題時照片存不了，但文字回報仍要寫入，避免整筆遺失。
        try {
          photoUrl = photoFolder_(category).createFile(photo).getUrl();
        } catch (err) {
          photoError = String(err.message || err);
          console.error('照片儲存失敗：' + photoError);
        }
      }
      createReport_(member, { id, category, content, location: clean_(req.location, 300), photoUrl },
        photoError ? '（照片儲存失敗：' + photoError + '）' : '');
      return { id, photoSaved: !photo || !!photoUrl };
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
        backup: backupStatus_(),
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

function createReport_(member, r, extraNotice) {
  const row = Object.assign({
    createdAt: now_(), userId: member.userId, name: member.name,
    location: '', photoUrl: '', status: '待處理', handler: '', note: '', updatedAt: '',
  }, r);
  append_('回報', row);
  notifyAdmins_(`新回報 ${row.id}｜${row.category}\n回報人：${row.name}\n${row.content.slice(0, 200)}${row.location ? '\n地點：' + row.location : ''}${extraNotice ? '\n' + extraNotice : ''}`);
}

/** always=true 時不受 NOTIFY_ADMINS 影響（例如備份失敗）。 */
function notifyAdmins_(message, always) {
  if (!always && prop_('NOTIFY_ADMINS') === 'false') return;
  const ids = readAll_('成員').filter(m => isActiveMember_(m) && ADMIN_ROLES.indexOf(m.role) >= 0).map(m => m.userId);
  // 通知失敗不影響主流程，否則回報已寫入卻回傳錯誤，使用者會重複送出。
  try {
    if (ids.length) multicast_(ids, [text_(message)]);
  } catch (err) {
    console.error('通知里長失敗：' + err.message);
  }
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

/** 檢查並解碼 LIFF 傳來的照片；格式或大小不對時直接回報錯誤給使用者。 */
function decodePhoto_(id, photo) {
  const m = String(photo).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('照片格式不支援');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('照片太大');
  return Utilities.newBlob(bytes, m[1], id + '.' + m[1].split('/')[1]);
}

/**
 * 回報照片依「類別／月份」分資料夾，存在備份資料夾的「回報照片」底下。
 * 備份資料夾不能用時（未設定、沒權限、被公開分享），改存到私人資料夾。
 * 照片不另外開放分享，權限跟著所在資料夾。
 */
function photoFolder_(category) {
  const path = [safeName_(category), Utilities.formatDate(new Date(), TZ, 'yyyy-MM')];
  if (prop_('BACKUP_FOLDER_ID')) {
    try {
      return folderPath_(backupRoot_(), ['回報照片'].concat(path));
    } catch (err) {
      console.error('照片改存到私人資料夾：' + err.message);
    }
  }
  return folderPath_(privatePhotoRoot_(), path);
}

function privatePhotoRoot_() {
  const id = prop_('PHOTO_FOLDER_ID');
  if (id) {
    const folder = DriveApp.getFolderById(id);
    if (!folder.isTrashed()) return folder;
  }
  const folder = DriveApp.createFolder('LINE 回報照片');
  PropertiesService.getScriptProperties().setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
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

// ───────────────────────── 雲端硬碟備份 ─────────────────────────

const BACKUP_LABELS = {
  userId: 'LINE ID', name: '姓名', phone: '電話', role: '職務', group: '組別', status: '狀態', joinedAt: '加入時間',
  id: '編號', createdAt: '建立時間', category: '類別', content: '內容', location: '地點', photoUrl: '照片',
  handler: '處理人', note: '備註', updatedAt: '更新時間', target: '對象', title: '標題', recipients: '收件人數',
  slotId: '時段編號', date: '日期', start: '開始', end: '結束', lawyer: '律師', capacity: '名額',
  topic: '諮詢類別', detail: '問題簡述', reminded: '提醒時間',
};
const BACKUP_STALE_MS = 10 * 60 * 1000;

/**
 * 每天凌晨由觸發器執行，也可從試算表選單或管理後台手動執行。
 * 備份資料夾結構：
 *   每日備份/年/年-月/里辦LINE資料_日期_時間.xlsx      整份試算表的歷史快照
 *   最新資料/工作人員、回報、推播公告、律師諮詢/…csv     依類別整理的最新資料（每次覆蓋）
 *   回報照片/類別/年-月/                                 回報時就直接存到這裡
 */
function backupToDrive() {
  if (!prop_('BACKUP_FOLDER_ID')) {
    console.warn('尚未設定 BACKUP_FOLDER_ID，略過備份');
    return null;
  }
  if (!startBackup_()) throw new Error('備份正在進行中，請稍後再試');
  const props = PropertiesService.getScriptProperties();
  try {
    const result = runBackup_();
    props.setProperty('LAST_BACKUP_AT', now_());
    props.deleteProperty('LAST_BACKUP_ERROR');
    return result;
  } catch (err) {
    const message = String(err.message || err);
    props.setProperty('LAST_BACKUP_ERROR', now_() + ' ' + message);
    notifyAdmins_('雲端硬碟備份失敗：' + message, true);
    throw err;
  } finally {
    props.deleteProperty('BACKUP_RUNNING');
  }
}

/**
 * 用指令碼屬性當作「備份中」旗標。網頁、排程、選單三種執行環境拿到的
 * LockService 鎖不同，只有全域鎖共用，所以只在檢查旗標的瞬間借用全域鎖。
 * 超過 10 分鐘的旗標視為上次執行被中斷（Apps Script 單次最多 6 分鐘）。
 */
function startBackup_() {
  return withLock_(() => {
    const props = PropertiesService.getScriptProperties();
    const running = Number(props.getProperty('BACKUP_RUNNING') || 0);
    if (running && Date.now() - running < BACKUP_STALE_MS) return false;
    if (running) props.setProperty('LAST_BACKUP_ERROR', now_() + ' 上次備份沒有完成（可能超過執行時間上限）');
    props.setProperty('BACKUP_RUNNING', String(Date.now()));
    return true;
  });
}

function runBackup_() {
  const root = backupRoot_();
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();

  // 1. 整份試算表匯出成 Excel，依年／月分資料夾保存。
  const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('匯出 Excel 失敗（HTTP ' + res.getResponseCode() + '）');
  const snapshot = folderPath_(root, ['每日備份', stamp.slice(0, 4), stamp.slice(0, 7)])
    .createFile(res.getBlob().setName('里辦LINE資料_' + stamp + '.xlsx'));

  // 2. 依類別整理的最新 CSV（Excel 可直接開啟）。
  const latest = folderPath_(root, ['最新資料']);
  let count = 0;
  const write = (path, name, keys, rows) => {
    upsertCsv_(folderPath_(latest, path), name, toCsv_(keys, rows));
    count++;
  };
  // 分組資料夾：每組寫一個檔，並把本次沒有的舊檔移到垃圾桶，避免留下過時資料或已刪除的個資。
  const writeGroups = (path, keys, groups) => {
    const folder = folderPath_(latest, path);
    const names = Object.keys(groups).map(g => {
      upsertCsv_(folder, g + '.csv', toCsv_(keys, groups[g]));
      return g + '.csv';
    });
    count += names.length;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (!f.isTrashed() && /\.csv$/i.test(f.getName()) && names.indexOf(f.getName()) < 0) f.setTrashed(true);
    }
  };

  const members = readAll_('成員').map(m => Object.assign({}, m, { status: m.status === 'active' ? '啟用' : '停用' }));
  write(['工作人員'], '工作人員名單.csv', SHEETS.成員, members);

  const reports = readAll_('回報').reverse();
  write(['回報'], '全部回報.csv', SHEETS.回報, reports);
  writeGroups(['回報', '依類別'], SHEETS.回報, groupBy_(reports, r => r.category, REPORT_CATEGORIES));
  writeGroups(['回報', '依狀態'], SHEETS.回報, groupBy_(reports, r => r.status, REPORT_STATUSES));

  write(['推播公告'], '推播公告.csv', SHEETS.公告, readAll_('公告').reverse());

  const slots = readAll_('律師時段');
  const slotIndex = indexBy_(slots, 'slotId');
  const bookingKeys = ['id', 'date', 'start', 'end', 'lawyer', 'name', 'phone', 'topic', 'detail', 'status', 'createdAt'];
  const bookings = readAll_('諮詢預約').map(b => {
    const s = slotIndex[b.slotId] || {};
    return Object.assign({}, b, { date: s.date || '', start: s.start || '', end: s.end || '', lawyer: s.lawyer || '' });
  }).sort((a, b) => String(b.date + b.start).localeCompare(String(a.date + a.start)));
  write(['律師諮詢'], '律師時段.csv', SHEETS.律師時段, slots.slice().sort((a, b) => String(b.date).localeCompare(String(a.date))));
  write(['律師諮詢'], '全部預約.csv', bookingKeys, bookings);
  writeGroups(['律師諮詢', '依月份'], bookingKeys, groupBy_(bookings, b => String(b.date).slice(0, 7) || '時段已刪除', []));

  console.log('備份完成：' + snapshot.getName() + '，CSV ' + count + ' 個');
  return { at: now_(), snapshot: snapshot.getName(), csvFiles: count, folderUrl: root.getUrl() };
}

/** 依檔名分組；base 裡的組別即使沒有資料也會產生空檔。 */
function groupBy_(rows, keyFn, base) {
  const groups = {};
  base.forEach(k => { groups[safeName_(k)] = []; });
  rows.forEach(r => {
    const k = safeName_(keyFn(r));
    (groups[k] = groups[k] || []).push(r);
  });
  return groups;
}

/** 取得備份資料夾；資料夾不存在、在垃圾桶或開放公開連結時拒絕寫入。 */
function backupRoot_() {
  const id = parseFolderId_(prop_('BACKUP_FOLDER_ID'));
  if (!id) throw new Error('尚未設定備份資料夾');
  let folder;
  try {
    folder = DriveApp.getFolderById(id);
  } catch (_) {
    throw new Error('無法開啟備份資料夾，請確認網址正確，且這個 Google 帳號有「編輯」權限');
  }
  if (folder.isTrashed()) throw new Error('備份資料夾在垃圾桶裡，請先還原或換一個資料夾');
  if (isPublicFolder_(folder)) {
    throw new Error('備份資料夾開放「知道連結的任何人」存取，為了保護個資已暫停寫入。請到雲端硬碟把這個資料夾的一般存取改成「限制」');
  }
  return folder;
}

function parseFolderId_(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/(?:folders\/|[?&]id=)([-\w]{10,})/);
  if (m) return m[1];
  return /^[-\w]{10,}$/.test(s) ? s : '';
}

function isPublicFolder_(folder) {
  try {
    const access = folder.getSharingAccess();
    return access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK;
  } catch (_) {
    return false;
  }
}

function backupStatus_() {
  const status = { configured: !!prop_('BACKUP_FOLDER_ID'), lastAt: prop_('LAST_BACKUP_AT'), lastError: prop_('LAST_BACKUP_ERROR') };
  const running = Number(prop_('BACKUP_RUNNING') || 0);
  if (running && Date.now() - running >= BACKUP_STALE_MS && !status.lastError) {
    status.lastError = '上次備份沒有完成（可能超過執行時間上限）';
  }
  if (!status.configured) return status;
  try {
    const folder = backupRoot_();
    status.folderUrl = folder.getUrl();
    status.folderName = folder.getName();
  } catch (err) {
    status.warning = err.message;
  }
  return status;
}

/** 依名稱逐層取得資料夾（略過垃圾桶裡的），不存在就建立。同一次執行內會快取。 */
const folderCache_ = {};
function folderPath_(root, names) {
  return names.reduce((parent, name) => {
    const key = parent.getId() + '/' + name;
    if (folderCache_[key]) return folderCache_[key];
    let found = null;
    const it = parent.getFoldersByName(name);
    while (!found && it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed()) found = f;
    }
    return (folderCache_[key] = found || parent.createFolder(name));
  }, root);
}

function upsertCsv_(folder, name, csv) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f.setContent(csv);
  }
  return folder.createFile(name, csv, MimeType.CSV);
}

/** 產生 Excel 可正確顯示中文的 CSV（開頭加 UTF-8 BOM），並擋掉公式注入。 */
function toCsv_(keys, rows) {
  const cell = (key, v) => {
    let s = String(v === undefined || v === null ? '' : v);
    // 電話用 ="0912…" 保留開頭的 0；只在內容全是電話字元時才這樣做。
    if (key === 'phone' && /^[\d+\-() ]+$/.test(s)) return '"=""' + s + '"""';
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [keys.map(k => cell('', BACKUP_LABELS[k] || k))]
    .concat(rows.map(r => keys.map(k => cell(k, r[k]))));
  return '\uFEFF' + lines.map(l => l.join(',')).join('\r\n') + '\r\n';
}

function safeName_(s) {
  return String(s || '').replace(/[\\/:*?"<>|]/g, '_').trim().slice(0, 50) || '未分類';
}

// ───────────────────────── 試算表選單（一鍵設定） ─────────────────────────

const RICH_MENU_IMAGE_URL = 'https://d75405-stack.github.io/financial-report-system/line-bot/web/richmenu.png';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('LINE 系統')
    .addItem('一鍵設定', 'setupWizard')
    .addItem('建立圖文選單', 'setupRichMenuFromMenu')
    .addItem('立即備份到雲端硬碟', 'backupFromMenu')
    .addItem('查看設定狀態', 'showStatus')
    .addToUi();
}

/**
 * 用對話框依序詢問 4 個設定值，每個都當場檢查；全部填完才儲存並自動完成其餘設定。
 * 對話框等待的時間也算在 Apps Script 單次 6 分鐘的上限內，所以請先把 4 個值準備好。
 */
function setupWizard() {
  const ui = SpreadsheetApp.getUi();
  // 授權畫面若只勾了部分權限，webhook 與備份會失敗；這裡要求一次給齊。
  if (typeof ScriptApp.requireAllScopes === 'function') ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  try {
    checkOwner_();
  } catch (err) {
    return ui.alert(err.message);
  }

  const steps = [
    ['LINE_CHANNEL_ACCESS_TOKEN', '1/4  LINE Channel access token',
      'LINE Developers → Messaging API channel → Messaging API 頁籤最下方的「Channel access token (long-lived)」。\n若是空的，先按「Issue」產生再複製。',
      validateToken_],
    ['WEBAPP_URL', '2/4  網頁應用程式網址',
      'Apps Script 右上角「部署 → 管理部署作業」裡的網頁應用程式網址，開頭是 https://script.google.com/macros/s/，結尾是 /exec。',
      validateWebAppUrl_],
    ['LIFF_ID', '3/4  LIFF ID',
      'LINE Developers → LINE Login channel → LIFF 頁籤的 LIFF ID，例如 2001234567-AbCdEfGh。',
      validateLiffId_],
    ['BACKUP_FOLDER_ID', '4/4  備份資料夾網址',
      '在雲端硬碟打開備份資料夾，複製網址列的網址貼上。\n資料夾的一般存取請設成「限制」，裡面會有個資。',
      validateFolder_],
  ];
  const values = {};
  for (const [key, title, hint, validate] of steps) {
    const current = prop_(key);
    let error = '';
    for (;;) {
      const tail = current ? '\n\n已經設定過；留空按「確定」會保留原本的值。' : '\n\n還沒有的話可以先留空，之後再執行一次「一鍵設定」。';
      const r = ui.prompt(title, (error ? '❌ ' + error + '\n\n' : '') + hint + tail, ui.ButtonSet.OK_CANCEL);
      if (r.getSelectedButton() !== ui.Button.OK) return ui.alert('已取消，設定沒有變更。');
      const input = r.getResponseText().trim();
      if (!input) break;
      const checked = validate(input);
      if (checked.error) {
        error = checked.error;
        continue;
      }
      values[key] = checked.value;
      break;
    }
  }
  PropertiesService.getScriptProperties().setProperties(values);
  ui.alert('設定結果', setup().join('\n'), ui.ButtonSet.OK);
}

function validateToken_(v) {
  if (/^[0-9a-f]{32}$/.test(v)) return { error: '這是 Channel secret，請改貼 Channel access token（比較長的那一串）。' };
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', { headers: { Authorization: 'Bearer ' + v }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { error: 'LINE 不接受這個 token，請確認複製完整。' };
  return { value: v };
}

function validateWebAppUrl_(v) {
  if (/\/dev$/.test(v)) return { error: '這是測試用網址（/dev），請到「管理部署作業」複製結尾是 /exec 的網址。' };
  if (!/^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[-\w]+\/exec$/.test(v)) {
    return { error: '網址格式不對，要以 https://script.google.com/macros/s/ 開頭、/exec 結尾。' };
  }
  return { value: v };
}

function validateLiffId_(v) {
  const id = v.replace(/^https:\/\/liff\.line\.me\//, '').replace(/[/?#].*$/, '');
  if (!/^\d{6,}-[A-Za-z0-9]{4,}$/.test(id)) return { error: 'LIFF ID 格式不對，應該像 2001234567-AbCdEfGh。' };
  return { value: id };
}

function validateFolder_(v) {
  const id = parseFolderId_(v);
  if (!id) return { error: '這不是資料夾網址。請在雲端硬碟打開資料夾後，複製網址列（含 /folders/）。' };
  let folder;
  try {
    folder = DriveApp.getFolderById(id);
  } catch (_) {
    return { error: '打不開這個資料夾，請確認網址正確，且目前登入的 Google 帳號有「編輯」權限。' };
  }
  if (folder.isTrashed()) return { error: '這個資料夾在垃圾桶裡，請先還原或換一個資料夾。' };
  if (isPublicFolder_(folder)) {
    return { error: '這個資料夾開放「知道連結的任何人」存取。請先在雲端硬碟按「共用」，把一般存取改成「限制」，再貼一次網址。' };
  }
  return { value: id };
}

function setupRichMenuFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(setupRichMenu());
  } catch (err) {
    ui.alert('建立圖文選單失敗：' + err.message);
  }
}

function backupFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = backupToDrive();
    ui.alert(r ? `備份完成：${r.snapshot}，另整理 ${r.csvFiles} 個分類檔。\n${r.folderUrl}` : '尚未設定備份資料夾，請先執行「一鍵設定」。');
  } catch (err) {
    ui.alert('備份失敗：' + err.message);
  }
}

function showStatus() {
  const lines = statusLines_().concat(['', '管理後台密碼：' + prop_('ADMIN_TOKEN'), '工作人員邀請碼：' + prop_('INVITE_CODE'),
    '', '管理後台密碼可以讀取所有資料、發送推播，請勿外流。']);
  SpreadsheetApp.getUi().alert('設定狀態', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function statusLines_() {
  const lines = [];
  const bot = prop_('LINE_CHANNEL_ACCESS_TOKEN') ? lineApi_('info', null, 'get') : { ok: false };
  lines.push(bot.ok ? '✅ LINE 官方帳號：' + JSON.parse(bot.body).displayName : '❌ LINE Channel access token 無效或未設定');
  lines.push(prop_('WEBAPP_URL') ? '✅ 網頁應用程式網址已設定' : '⚠️ 尚未設定網頁應用程式網址');
  lines.push(prop_('LIFF_ID') ? '✅ LIFF ID：' + prop_('LIFF_ID') : '⚠️ 尚未設定 LIFF ID（回報、預約表單無法開啟）');
  const backup = backupStatus_();
  if (!backup.configured) lines.push('⚠️ 尚未設定備份資料夾');
  else if (backup.folderName) lines.push('✅ 備份資料夾：' + backup.folderName + (backup.lastAt ? '（上次備份 ' + backup.lastAt + '）' : ''));
  if (backup.warning) lines.push('❌ ' + backup.warning);
  if (backup.lastError) lines.push('❌ 上次備份失敗：' + backup.lastError);
  return lines;
}

/** 把 LINE 的 webhook 網址設成這個 Apps Script。 */
function setWebhookEndpoint_() {
  const url = prop_('WEBAPP_URL');
  if (validateWebAppUrl_(url).error) return '⚠️ 網頁應用程式網址格式不對（要以 /exec 結尾），沒有設定 webhook';
  const res = lineApi_('channel/webhook/endpoint', { endpoint: url }, 'put');
  return res.ok ? '✅ 已自動把 LINE webhook 網址設成這個程式' : '❌ 設定 webhook 失敗：' + res.body;
}

/** 建立 6 格圖文選單並設為所有好友的預設選單。重複執行會換掉舊的。 */
function setupRichMenu() {
  if (!prop_('LIFF_ID')) throw new Error('請先在「一鍵設定」填入 LIFF ID');
  const image = UrlFetchApp.fetch(RICH_MENU_IMAGE_URL, { muteHttpExceptions: true });
  if (image.getResponseCode() !== 200) throw new Error('圖文選單圖片還沒上線，等網頁部署完成後點「LINE 系統 → 建立圖文選單」');

  const W = 2500, H = 1686, w = W / 3, h = H / 2;
  const cell = (col, row, action) => ({ bounds: { x: Math.round(col * w), y: row * h, width: Math.round(w), height: h }, action });
  const msg = (label, text) => ({ type: 'message', label, text });
  const menu = lineApi_('richmenu', {
    size: { width: W, height: H },
    selected: true,
    name: '里辦服務選單',
    chatBarText: '服務選單',
    areas: [
      cell(0, 0, msg('回報', '回報')),
      cell(1, 0, msg('律師諮詢', '律師諮詢')),
      cell(2, 0, msg('最新公告', '公告')),
      cell(0, 1, msg('我的預約', '我的預約')),
      cell(1, 1, { type: 'uri', label: '工作人員綁定', uri: liffUrl_('register') }),
      cell(2, 1, msg('使用說明', '說明')),
    ],
  });
  if (!menu.ok) throw new Error(menu.body);
  const richMenuId = JSON.parse(menu.body).richMenuId;

  const upload = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + richMenuId + '/content', {
    method: 'post',
    contentType: 'image/png',
    headers: { Authorization: 'Bearer ' + prop_('LINE_CHANNEL_ACCESS_TOKEN') },
    payload: image.getBlob().getBytes(),
    muteHttpExceptions: true,
  });
  if (upload.getResponseCode() >= 300) throw new Error('上傳圖片失敗：' + upload.getContentText());
  const setDefault = lineApi_('user/all/richmenu/' + richMenuId, null, 'post');
  if (!setDefault.ok) throw new Error(setDefault.body);

  const props = PropertiesService.getScriptProperties();
  const old = props.getProperty('RICH_MENU_ID');
  if (old && old !== richMenuId) lineApi_('richmenu/' + old, null, 'delete');
  props.setProperty('RICH_MENU_ID', richMenuId);
  return '✅ 已建立圖文選單，所有好友都會看到';
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
