async function api(action, data) {
  // 用 text/plain 避免 CORS 預檢，Apps Script 不支援 OPTIONS 請求。
  const res = await fetch(window.APP_CONFIG.API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ action }, data || {})),
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || '發生錯誤');
  return json.data;
}

function esc(v) {
  return String(v === undefined || v === null ? '' : v).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function $(sel, root) {
  return (root || document).querySelector(sel);
}

function toast(msg, isError) {
  let el = $('#toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'toast';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.className = 'show' + (isError ? ' error' : '');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { el.className = ''; }, 3000);
}

function formData(form) {
  return Object.fromEntries(new FormData(form).entries());
}

/** 送出表單時鎖住按鈕，避免重複送出。 */
async function withBusy(button, fn) {
  const label = button.textContent;
  button.disabled = true;
  button.textContent = '處理中…';
  try {
    return await fn();
  } catch (err) {
    toast(err.message, true);
  } finally {
    button.disabled = false;
    button.textContent = label;
  }
}

/** 縮小照片再上傳，長邊最多 1280px。 */
function resizeImage(file, maxSize) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxSize / Math.max(img.width, img.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(img.src);
      resolve(canvas.toDataURL('image/jpeg', 0.8));
    };
    img.onerror = () => reject(new Error('無法讀取照片'));
    img.src = URL.createObjectURL(file);
  });
}
