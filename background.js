if (chrome.sidePanel) {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
} else {
  chrome.action.onClicked.addListener(() => {
    chrome.notifications.create({
      type: "basic",
      title: "Unsupported",
      iconUrl: chrome.runtime.getURL("icon.png"),
      message: "Please upgrade your Chrome browser to version 114+",
    });
  });
}

// Chỉ giả mạo User-Agent thành Edge cho đúng các domain cần thiết,
// KHÔNG xóa CSP hay X-Frame-Options (rules/bing.json đã xử lý riêng cho iframe copilot).
chrome.declarativeNetRequest.updateDynamicRules({
  removeRuleIds: [1],
  addRules: [
    {
      id: 1,
      priority: 1,
      action: {
        type: "modifyHeaders",
        requestHeaders: [
          {
            header: "user-agent",
            operation: "set",
            value:
              "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36 Edg/110.0.1587.41",
          },
          {
            header: "sec-ch-ua",
            operation: "set",
            value: '"Microsoft Edge";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
          },
        ],
      },
      condition: {
        requestDomains: ["bing.com", "copilot.microsoft.com", "m365.cloud.microsoft"],
        resourceTypes: ["main_frame", "sub_frame", "xmlhttprequest", "websocket"],
      },
    },
  ],
});
