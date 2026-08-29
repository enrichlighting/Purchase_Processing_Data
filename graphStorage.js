// ============================================================
// graphStorage.js — เก็บ/อ่านข้อมูลผ่าน SharePoint List แทน localStorage
// ต้องโหลดไฟล์นี้ "หลัง" auth.js เสมอ เพราะใช้ window.GraphAuth
// SharePoint List: PackingPOData (site: PackingPO) — คอลัมน์ Title / Value
// ============================================================

const SITE_ID = 'richestsupply.sharepoint.com,f80946f0-2dc7-4992-8ce7-2725202fd09e,ce74b3bd-5357-4e38-9461-980f164deddd';
const LIST_ID = 'd01c34b8-1908-4cbb-87ad-8b5ad5097c9d';
const GRAPH_BASE = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}`;

async function graphFetch(url, options = {}) {
  const token = await window.GraphAuth.getGraphToken();
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Graph API error (${response.status}): ${errText}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

/**
 * หา Item ID ของแถวที่มี Title ตรงกับ key ที่ต้องการ
 * คืนค่า null ถ้ายังไม่เคยมีการบันทึก key นี้มาก่อน
 */
async function findItemIdByKey(key) {
  const url = `${GRAPH_BASE}/items?expand=fields(select=Title,Value)&$filter=fields/Title eq '${encodeURIComponent(key)}'`;
  const data = await graphFetch(url, {
    headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
  });
  if (data && data.value && data.value.length > 0) {
    return data.value[0].id;
  }
  return null;
}

/**
 * บันทึกข้อมูล (key/value)
 * ถ้ามี key นี้อยู่แล้วในระบบ จะอัปเดตทับค่าเดิม ถ้ายังไม่มีจะสร้างแถวใหม่ให้อัตโนมัติ
 * value เป็น object/array ได้เลย จะถูกแปลงเป็นข้อความ JSON ให้อัตโนมัติ
 */
async function saveData(key, value) {
  const stringValue = typeof value === 'string' ? value : JSON.stringify(value);
  const existingId = await findItemIdByKey(key);

  if (existingId) {
    await graphFetch(`${GRAPH_BASE}/items/${existingId}/fields`, {
      method: 'PATCH',
      body: JSON.stringify({ Value: stringValue }),
    });
  } else {
    await graphFetch(`${GRAPH_BASE}/items`, {
      method: 'POST',
      body: JSON.stringify({ fields: { Title: key, Value: stringValue } }),
    });
  }
}

/**
 * อ่านข้อมูลจาก key ที่กำหนด
 * คืนค่าเป็นข้อความดิบ (string) — ถ้าตอน saveData เก็บเป็น JSON ให้ใช้ JSON.parse() แปลงกลับเอง
 * คืนค่า null ถ้ายังไม่เคยมีการบันทึก key นี้
 */
async function loadData(key) {
  const url = `${GRAPH_BASE}/items?expand=fields(select=Title,Value)&$filter=fields/Title eq '${encodeURIComponent(key)}'`;
  const data = await graphFetch(url, {
    headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
  });
  if (data && data.value && data.value.length > 0) {
    return data.value[0].fields.Value;
  }
  return null;
}

// ============================================================
// Presence (ใครออนไลน์บ้าง) — ใช้ token แบบเงียบ (ไม่เด้ง popup) · แต่ละเครื่องเขียน key ของตัวเอง (presence-<clientId>)
// ============================================================
async function _presToken() {
  return (window.GraphAuth && window.GraphAuth.getGraphTokenSilent) ? await window.GraphAuth.getGraphTokenSilent() : null;
}
async function _presFetch(token, url, options = {}) {
  const r = await fetch(url, { ...options, headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(options.headers || {}) } });
  if (!r.ok) throw new Error('presence graph ' + r.status);
  if (r.status === 204) return null;
  return r.json();
}
// เขียน/อัปเดต heartbeat ของเครื่องนี้ (คืน false ถ้ายังไม่ได้ล็อกอิน)
async function presenceSet(key, value) {
  const token = await _presToken(); if (!token) return false;
  const stringValue = typeof value === 'string' ? value : JSON.stringify(value);
  const findUrl = `${GRAPH_BASE}/items?expand=fields(select=Title,Value)&$filter=fields/Title eq '${encodeURIComponent(key)}'`;
  const fd = await _presFetch(token, findUrl, { headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  const id = (fd && fd.value && fd.value[0]) ? fd.value[0].id : null;
  if (id) await _presFetch(token, `${GRAPH_BASE}/items/${id}/fields`, { method: 'PATCH', body: JSON.stringify({ Value: stringValue }) });
  else await _presFetch(token, `${GRAPH_BASE}/items`, { method: 'POST', body: JSON.stringify({ fields: { Title: key, Value: stringValue } }) });
  return true;
}
// อ่านรายการ presence ทั้งหมด (คืน null ถ้ายังไม่ได้ล็อกอิน) · กรอง prefix ฝั่ง client
async function presenceList(prefix) {
  const token = await _presToken(); if (!token) return null;
  const url = `${GRAPH_BASE}/items?expand=fields(select=Title,Value)&$top=500`;
  const d = await _presFetch(token, url);
  return (d && d.value ? d.value : [])
    .map(it => ({ id: it.id, key: (it.fields && it.fields.Title) || '', value: (it.fields && it.fields.Value) || '' }))
    .filter(x => x.key.indexOf(prefix) === 0);
}
// ลบ heartbeat ของเครื่องนี้ (ตอนปิดแท็บ · best-effort)
async function presenceDelete(key) {
  const token = await _presToken(); if (!token) return;
  const findUrl = `${GRAPH_BASE}/items?expand=fields(select=Title)&$filter=fields/Title eq '${encodeURIComponent(key)}'`;
  const fd = await _presFetch(token, findUrl, { headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  const id = (fd && fd.value && fd.value[0]) ? fd.value[0].id : null;
  if (id) await _presFetch(token, `${GRAPH_BASE}/items/${id}`, { method: 'DELETE' });
}

// ============================================================
// อัปโหลดไฟล์ (Excel/รูปภาพ) ตรงเข้า SharePoint Document Library แบบอัตโนมัติ
// ทุกคนที่ login เข้าแอปจะเซฟตรงที่เดียวกันได้เลย ไม่ต้องตั้งโฟลเดอร์ในเครื่อง
// ============================================================
const UPLOAD_FOLDER = 'Packing list log'; // ชื่อโฟลเดอร์ปลายทางใน SharePoint (Documents)
async function uploadFile(filename, blob) {
  const token = await window.GraphAuth.getGraphToken();
  const path = `${encodeURIComponent(UPLOAD_FOLDER)}/${encodeURIComponent(filename)}`;
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:/${path}:/content`;
  const buf = await blob.arrayBuffer();
  const response = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}` },
    body: buf,
  });
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Graph upload error (${response.status}): ${errText}`);
  }
  return response.json();
}
// ============================================================
// Login ระบบ username/password ของแอป (แยกจาก Microsoft sign-in ที่มีอยู่แล้ว)
// เก็บใน SharePoint List: AppUsers (site: PackingPO) — คอลัมน์ Title(=username) / PasswordHash / Role / DisplayName
// รหัสผ่านไม่ได้เก็บตรงๆ — แปลงเป็น hash (SHA-256) ก่อนเทียบ/บันทึกเสมอ แปลงกลับเป็นรหัสผ่านจริงไม่ได้
// ============================================================
const APPUSERS_LIST_ID = '52ed1658-b58f-40d9-90b4-db880a52fe78';
const APPUSERS_BASE = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${APPUSERS_LIST_ID}`;

async function _sha256Hex(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

// ตรวจสอบ username/password กับ AppUsers List
// สำเร็จ -> คืน {username, role, displayName} · ไม่สำเร็จ -> คืน null
async function verifyLogin(username, password) {
  const token = await window.GraphAuth.getGraphToken();
  const url = `${APPUSERS_BASE}/items?expand=fields(select=Title,PasswordHash,Role,DisplayName)&$filter=fields/Title eq '${encodeURIComponent(username)}'`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
  });
  if (!response.ok) throw new Error(`Graph API error (${response.status})`);
  const data = await response.json();
  if (!data.value || !data.value.length) return null;
  const item = data.value[0];
  const f = item.fields;
  const hash = await _sha256Hex(password);
  if (hash !== f.PasswordHash) return null;
  // บันทึกเวลาใช้งานล่าสุด (ยิงแบบไม่รอผล กัน login ช้าถ้าเน็ตช้า)
  fetch(`${APPUSERS_BASE}/items/${item.id}/fields`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ LastLogin: new Date().toISOString() }),
  }).catch(()=>{});
  return { username: f.Title, role: f.Role, displayName: f.DisplayName || f.Title };
}

// เพิ่ม user ใหม่เข้า AppUsers List (ใช้ตอนสร้าง user ครั้งแรก หรือให้ Admin เพิ่ม user ในแอปทีหลัง)
async function createUser(username, password, role, displayName) {
  const token = await window.GraphAuth.getGraphToken();
  const hash = await _sha256Hex(password);
  const response = await fetch(`${APPUSERS_BASE}/items`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: { Title: username, PasswordHash: hash, Role: role, DisplayName: displayName || username } }),
  });
  if (!response.ok) { const t = await response.text(); throw new Error(`Graph API error (${response.status}): ${t}`); }
  return response.json();
}
window.GraphStorage = {
  saveData,
  loadData,
  findItemIdByKey,
  presenceSet,
  presenceList,
  presenceDelete,
  uploadFile,
  verifyLogin,
  createUser,
};
