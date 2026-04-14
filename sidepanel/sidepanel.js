const iframe = document.getElementById("underside-iframe-container");
const statusEl = document.getElementById("status-msg");
const btnPageContext = document.getElementById("btn-page-context");
const btnReload = document.getElementById("btn-reload");
const loginBanner = document.getElementById("login-banner");
const btnLogin = document.getElementById("btn-login");

const COPILOT_URL = "https://copilot.microsoft.com/";

// Cho phép cả 2 domain Microsoft — tài khoản M365 sẽ redirect sang m365.cloud.microsoft
const ALLOWED_ORIGINS = new Set([
  "https://copilot.microsoft.com",
  "https://m365.cloud.microsoft",
]);

// Domain login của Microsoft — MSAL không cho login trong iframe,
// phải mở tab mới thay thế
const LOGIN_DOMAINS = [
  "login.microsoftonline.com",
  "login.live.com",
  "login.microsoft.com",
];

// Giới hạn nội dung trang gửi cho Copilot (~8000 ký tự là đủ để tóm tắt)
const PAGE_TEXT_LIMIT = 8000;

// --- Khởi tạo iframe ---
iframe.src = COPILOT_URL;

// --- Helpers ---
function setStatus(msg, type = "", duration = 4000) {
  statusEl.textContent = msg;
  statusEl.className = type;
  if (duration > 0) {
    setTimeout(() => { statusEl.textContent = ""; statusEl.className = ""; }, duration);
  }
}

function showLoginBanner(loginUrl) {
  loginBanner.classList.add("visible");
  btnLogin.onclick = () => {
    chrome.tabs.create({ url: loginUrl || COPILOT_URL });
  };
}

function hideLoginBanner() {
  loginBanner.classList.remove("visible");
}

function sendEventToIframe(name, args) {
  // Gửi đến cả 2 origin — browser tự bỏ qua nếu iframe đang ở origin khác
  for (const origin of ALLOWED_ORIGINS) {
    iframe.contentWindow.postMessage({ eventName: name, eventArgs: args }, origin);
  }
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs[0];
}

function buildActiveTabInfo(tab) {
  return {
    serpQuery: "",
    isActive: true,
    tabId: "",
    windowId: "",
    groupId: "",
    windowType: "",
    isLoading: false,
    title: tab.title,
    url: tab.url,
    pageLanguage: "",
  };
}

// --- Phát hiện iframe chuyển sang trang login ---
// MSAL chặn redirect trong iframe, cần mở tab mới
chrome.webNavigation.onCommitted.addListener((details) => {
  if (details.frameType !== "sub_frame" && details.frameId === 0) return;
  const url = new URL(details.url);
  if (LOGIN_DOMAINS.some((d) => url.hostname.endsWith(d))) {
    // Đưa iframe về trang chính, hiện banner hướng dẫn login
    iframe.src = COPILOT_URL;
    showLoginBanner(details.url);
  }
});

// --- Nút Tải lại ---
btnReload.addEventListener("click", () => {
  hideLoginBanner();
  iframe.src = COPILOT_URL;
  chatPageInitialized = false;
});

// --- Nút "Đọc trang này" ---
btnPageContext.addEventListener("click", async () => {
  const tab = await getActiveTab();
  if (!tab) {
    setStatus("Không tìm thấy tab hiện tại", "err");
    return;
  }

  btnPageContext.disabled = true;
  setStatus("Đang đọc trang...", "", 0);

  try {
    const response = await chrome.tabs.sendMessage(tab.id, { action: "getPageData" });
    if (!response?.text) {
      setStatus("Không đọc được nội dung trang", "err");
      return;
    }

    const pageText = response.text.trim().slice(0, PAGE_TEXT_LIMIT);
    const prompt = `Trang: ${tab.title}\nURL: ${tab.url}\n\nNội dung:\n${pageText}`;

    // Thử gửi qua postMessage (hoạt động trên copilot.microsoft.com)
    sendEventToIframe("Discover.Chat.Page", { text: response.text });

    // Copy vào clipboard — fallback đáng tin cậy cho m365.cloud.microsoft
    await navigator.clipboard.writeText(prompt);
    setStatus("✓ Đã copy! Paste vào chat (Ctrl+V)", "ok");
  } catch (err) {
    setStatus("Lỗi: " + (err instanceof Error ? err.message : "Không rõ"), "err");
  } finally {
    btnPageContext.disabled = false;
  }
});

// --- Xử lý postMessage từ iframe (copilot.microsoft.com) ---
let chatPageInitialized = false;

async function postMessageListner(event) {
  if (!ALLOWED_ORIGINS.has(event.origin)) return;
  const eventName = event.data.eventName;

  if (eventName === "Discover.Chat.Interact.Req") {
    sendEventToIframe("Discover.Chat.Interact.Rep", { status: true });
  } else if (eventName === "Discover.Chat.Consent.Req") {
    sendEventToIframe("Discover.Chat.Consent.Rep", { text: "Accepted" });
  } else if (eventName === "Discover.Chat.Page.GetData") {
    const tab = await getActiveTab();
    const response = await chrome.tabs.sendMessage(tab.id, { action: "getPageData" });
    sendEventToIframe("Discover.Chat.Page", { text: response.text });
  } else if (eventName === "Discover.Ready" && !chatPageInitialized) {
    sendEventToIframe("Discover.VisibilityState", { isShow: true, timeStamp: Date.now() });
    sendEventToIframe("Discover.Tab.Click", { tabName: "chat", clientLevel: "window" });
    chatPageInitialized = true;
    const tab = await getActiveTab();
    if (tab) {
      sendEventToIframe("Discover.Client.TabStripModelChange", {
        eventType: "Activate",
        tabInfo: buildActiveTabInfo(tab),
      });
    }
  }
}

window.addEventListener("message", postMessageListner, false);

chrome.tabs.onActivated.addListener(async (activeInfo) => {
  const tab = await chrome.tabs.get(activeInfo.tabId);
  sendEventToIframe("Discover.Client.TabStripModelChange", {
    eventType: "Activate",
    tabInfo: buildActiveTabInfo(tab),
  });
});

chrome.tabs.onUpdated.addListener((_tabId, changeInfo, tab) => {
  if (tab.active && changeInfo.status === "complete") {
    sendEventToIframe("Discover.Client.TabStripModelChange", {
      eventType: "Activate",
      tabInfo: buildActiveTabInfo(tab),
    });
  }
});
