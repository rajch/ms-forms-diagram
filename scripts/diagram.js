"use strict";

function setElementText(elementid, text) {
    const statuspanel = document.getElementById(elementid)
    statuspanel.innerHTML = ""
    statuspanel.append(document.createTextNode(text))
}

function toggleopenbrancheditorpanel() {
    const openbrancheditorpanel = document.getElementById("openbrancheditor")
    const diagram = document.getElementById("diagram")
    openbrancheditorpanel.classList.toggle("hidden")
    diagram.classList.toggle("hidden")
}

chrome.runtime.onMessage.addListener(
    function (request, sender) {
        console.log(sender.tab ?
            "from a content script:" + sender.tab.url :
            "from the extension:" + JSON.stringify(request));

        if(!request.status) {
            setElementText("statuspanel","Error: Wrong message sent.")
            return
        }

        if(request.status == "Error") {
            setElementText("statuspanel", "Error:" + request.error)
            toggleopenbrancheditorpanel()
            return
        }

        if(request.status == "Success") {
            setElementText("statuspanel", "")
            setElementText("diagram",request.diagramText)
            mermaid.init({ startOnLoad: false },".dynmermaid");
        }

        console.log("Finished stuff")
    }
);
