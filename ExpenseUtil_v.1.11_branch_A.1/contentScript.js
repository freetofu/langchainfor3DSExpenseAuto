const scripts = ['scripts/jquery.min.js', 'scripts/xlsx.full.min.js', 'pageScript.js'];

scripts.forEach(src => {
    const script = document.createElement('script');
    script.src = chrome.runtime.getURL(src);
    (document.head || document.documentElement).appendChild(script);
});

chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
    if (message.action === 'popup_start_input') {
        let event = new CustomEvent('content_start_input', {
            detail: message
        });
        window.dispatchEvent(event);
    }
});

window.addEventListener('message', function (event) {
    if (event.data.action === 'page_done') {
        chrome.runtime.sendMessage(event.data);
    }
}, false);
