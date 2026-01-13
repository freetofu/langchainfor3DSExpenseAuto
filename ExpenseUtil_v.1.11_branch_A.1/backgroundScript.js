chrome.runtime.onMessage.addListener(function (msg) {
    if (msg.action === 'page_done') {
        chrome.windows.update(-2, {
            drawAttention: true
        });
    }
});