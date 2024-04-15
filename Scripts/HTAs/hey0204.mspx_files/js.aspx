
function adjustIFrameSize (iframeWindow) {

    ratingHeight = document.getElementById("footerId");
    if(ratingHeight)ratingHeight.height = null;

if (iframeWindow.document.height) {var iframeElement = document.getElementById(iframeWindow.name);iframeElement.style.height = iframeWindow.document.height + 17 + 'px';}else if (document.all) {var iframeElement = document.all[iframeWindow.name];if (iframeWindow.document.compatMode && iframeWindow.document.compatMode != 'BackCompat') {iframeElement.style.height = iframeWindow.document.documentElement.scrollHeight + 5 + 'px';} else {iframeElement.style.height = iframeWindow.document.body.scrollHeight + 5 + 'px';}}

}
