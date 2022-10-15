"use strict";

function setElementText(elementid, text) {
    const statuspanel = document.getElementById(elementid)
    statuspanel.innerHTML = ""
    statuspanel.append(document.createTextNode(text))
}

chrome.runtime.onMessage.addListener(
    function (request, sender, sendResponse) {
        console.log(sender.tab ?
            "from a content script:" + sender.tab.url :
            "from the extension:" + JSON.stringify(request));
        //   if (request.greeting === "hello")
        //     sendResponse({farewell: "goodbye"});
        
        if(!request.status) {
            setElementText("statuspanel","Error: I don't know what goes onabort.")
            return
        }

        if(request.status == "Error") {
            setElementText("statuspanel", "Error:" + request.error)
            return
        }

        if(request.status == "Success") {
            setElementText("diagram",request.diagramText)
            mermaid.init({ startOnLoad: false },".dynmermaid");
        }

        console.log("Finished stuff")
    }
);
