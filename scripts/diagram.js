"use strict";

function setElementText(elementid, text) {
    const statuspanel = document.getElementById(elementid)
    statuspanel.innerHTML = ""
    statuspanel.append(document.createTextNode(text))
}

function showopenbrancheditorpanel() {
    const openbrancheditorpanel = document.getElementById("openbrancheditor")
    const diagram = document.getElementById("diagram")
    openbrancheditorpanel.classList.toggle("hidden")
    diagram.classList.toggle("hidden")
}

function showsavepanel() {
    const savepanel = document.getElementById("savepanel")
    savepanel.classList.toggle("hidden")
}

function savebuttonClicked(e) {
    const diagramsvg = document.querySelector("pre.dynmermaid svg")
    if (!diagramsvg) {
        window.alert("Diagram not found.")
        return
    }

    // First, convert the mermaid-generated svg into a data: url
    // with base64 encoding
    const svgdataUrl = "data:image/svg+xml;base64," +
        btoa(diagramsvg.outerHTML)

    // Then, create a new image object, set up a load handler to
    // process it, and point it to the data url
    const img = new Image();
    img.addEventListener('load', () => {
        // draw the image on an ad-hoc canvas
        const bbox = diagramsvg.getBBox()

        const canvas = document.createElement('canvas')
        canvas.width = bbox.width
        canvas.height = bbox.height

        const context = canvas.getContext('2d')
        context.drawImage(img, 0, 0, bbox.width, bbox.height)

        // Create a temporary link, configure it for download,
        // point it to the canvas-generated data url, and 
        // 'click' it to trigger a download.
        const a = document.createElement('a')
        const diagramtitle = document.getElementById("diagramtitle")
        const downloadFileName = diagramtitle ?
                                    `${diagramtitle.innerText}.png` :
                                    'image.png'               
        a.download = downloadFileName
        document.body.appendChild(a)
        a.href = canvas.toDataURL()
        a.click()
        a.remove()
    })

    img.src = svgdataUrl;
}

function chromeMessageReceived(request, sender) {
    console.log(sender.tab ?
        "from a content script:" + sender.tab.url :
        "from the extension:" + JSON.stringify(request));

    if (!request.status) {
        setElementText("statuspanel", "Error: Wrong message sent.")
        return
    }

    if (request.status == "Error") {
        setElementText("statuspanel", "Error:" + request.error)
        showopenbrancheditorpanel()
        return
    }

    if (request.status == "Success") {
        setElementText("statuspanel", "")
        setElementText("diagramtitle", request.diagramTitle)
        setElementText("diagram", request.diagramText)

        mermaid.init(undefined, ".dynmermaid");

        showsavepanel()
    }

    console.log("Finished stuff")
}

const savebutton = document.getElementById("savebutton")
savebutton.addEventListener("click", savebuttonClicked)

chrome.runtime.onMessage.addListener(chromeMessageReceived);
