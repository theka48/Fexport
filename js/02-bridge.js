function postToAHK(obj) {
    if (window.chrome && window.chrome.webview)
        window.chrome.webview.postMessage(JSON.stringify(obj));
}