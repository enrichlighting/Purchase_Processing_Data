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

window.GraphStorage = {
  saveData,
  loadData,
  findItemIdByKey,
};
