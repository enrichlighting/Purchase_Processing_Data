// ============================================================
// auth.js — จัดการ Login เข้าองค์กรผ่าน MSAL.js (Microsoft Authentication Library)
// ใช้คู่กับ Entra ID App Registration: Richest_Purchase_Processing
// ============================================================

const msalConfig = {
  auth: {
    clientId: '5bf2fbde-fbfa-4469-b7d6-7a035a01ef1b', // Application (client) ID
    authority: 'https://login.microsoftonline.com/77ec4080-3602-4d2e-b2fb-eee9ed878b1c', // Tenant ID
    //redirectUri: 'https://enrichlighting.github.io/Purchase_Processing_Data/', // ต้องตรงกับ Redirect URI ที่ลงทะเบียนไว้ใน Entra ID เป๊ะๆ
       redirectUri: window.location.hostname.includes('dynamics.com')
     ? 'https://org747e4176.crm5.dynamics.com/main.aspx?appid=f0d7fc49-839f-f111-b8de-70a8a502a6cf&pagetype=webresource&webresourceName=rich_index.html'
     : 'https://enrichlighting.github.io/Purchase_Processing_Data/', // เลือก Redirect URI อัตโนมัติตามโดเมนที่รันอยู่ (GitHub Pages หรือ Dataverse)
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
  scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All'],
};

let msalInitialized = false;
async function ensureMsalInitialized() {
  if (!msalInitialized) {
    await msalInstance.initialize();
    msalInitialized = true;
  }
}

/**
 * เรียกฟังก์ชันนี้ตอนเปิดหน้าเว็บ (หรือกดปุ่ม Login) เพื่อให้ผู้ใช้ล็อกอินด้วยบัญชีองค์กร
 * ถ้าล็อกอินค้างอยู่แล้วในเบราว์เซอร์ จะข้ามการเปิด popup ไปเลย
 */
async function signIn() {
  await ensureMsalInitialized();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  const loginResponse = await msalInstance.loginPopup(loginRequest);
  msalInstance.setActiveAccount(loginResponse.account);
  return loginResponse.account;
}

/**
 * เรียกฟังก์ชันนี้ทุกครั้งก่อนยิง Graph API เพื่อขอ Access Token ล่าสุด
 * (จะพยายามขอแบบเงียบๆ ก่อน ถ้าไม่ได้ค่อยเปิด popup ให้ล็อกอินใหม่)
 */
async function getGraphToken() {
  await ensureMsalInitialized();

  let account = msalInstance.getActiveAccount();
  if (!account) {
    account = await signIn();
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return response.accessToken;
  } catch (error) {
    const response = await msalInstance.acquireTokenPopup(loginRequest);
    return response.accessToken;
  }
}

/**
    * สำหรับขอ Access Token เพื่อเรียก Power Automate HTTP trigger (ปุ่ม "รีเฟรชข้อมูลทั้งหมด")
    * ใช้ resource คนละตัวกับ Microsoft Graph ด้านบน
    */
   const flowLoginRequest = {
     scopes: ['https://service.flow.microsoft.com//.default'],
   };

   async function getFlowToken() {
     await ensureMsalInitialized();

     let account = msalInstance.getActiveAccount();
     if (!account) {
       account = await signIn();
     }

     try {
       const response = await msalInstance.acquireTokenSilent({
         ...flowLoginRequest,
         account,
       });
       return response.accessToken;
     } catch (error) {
       const response = await msalInstance.acquireTokenPopup(flowLoginRequest);
       return response.accessToken;
     }
   }

   // ขอ token แบบเงียบเท่านั้น (ไม่เด้ง popup) — คืน null ถ้ายังไม่ได้ล็อกอิน · ใช้กับฟีเจอร์ presence ที่ไม่ควรรบกวนผู้ใช้
   async function getGraphTokenSilent() {
     try {
       await ensureMsalInitialized();
       const account = msalInstance.getActiveAccount() || (msalInstance.getAllAccounts()[0]);
       if (!account) return null;
       const response = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
       return response.accessToken;
     } catch (e) { return null; }
   }

   // เปิดให้ไฟล์อื่น (เช่น graphStorage.js หรือ bc-*.js) เรียกใช้งานผ่าน window.GraphAuth
   window.GraphAuth = {
     signIn,
     getGraphToken,
     getGraphTokenSilent,
     getFlowToken,
     msalInstance,
   };

