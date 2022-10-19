'use strict'

const diagramPage = {
  self: undefined,
  downloadUrl: undefined,
  svgBlobUrl: undefined,
  diagramTitle: '',
  diagramText: '',
  diagramMode: 'default',
  setTitle (diagramtitle) {
    this.diagramTitle = diagramtitle
    this.setElementText('diagramtitle', diagramtitle)
  },
  hideElement (elementid) {
    const element = document.getElementById(elementid)
    element.classList.add('hidden')
  },
  showElement (elementid, hide) {
    const element = document.getElementById(elementid)
    element.classList.remove('hidden')
  },
  setElementText (elementid, text) {
    const element = document.getElementById(elementid)
    element.innerHTML = ''
    element.append(document.createTextNode(text))
  },
  setStatus (text) {
    this.setElementText('statuspanel', text)
  },
  setErrorStatus (text) {
    this.setTitle('Error')
    this.setStatus(text)
  },
  setDiagramText (text) {
    this.diagramText = text
    this.setElementText('diagram', text)
  },
  showDiagram (diagramtitle, diagramtext) {
    this.setTitle(diagramtitle)
    this.setDiagramText(diagramtext)

    mermaid.init(undefined, '.dynmermaid')
  },
  clearDiagram () {
    const diagram = document.getElementById('diagram')
    diagram.innerHTML = ''
    diagram.removeAttribute('data-processed')
  },
  showOpenBranchEditorPanel () {
    this.showElement('openbrancheditor')
    this.hideElement('diagram')
  },
  downloadImage (imageUrl) {
    // Create a temporary link, configure it for download,
    // point it to the canvas-generated data url, and
    // 'click' it to trigger a download.
    const a = document.createElement('a')

    const downloadFileName = this.diagramTitle
      ? `${this.diagramTitle}.png`
      : 'image.png'
    a.download = downloadFileName
    document.body.appendChild(a)
    a.href = imageUrl
    a.click()
    a.remove()
  },
  createDownloadImage (sourceurl, diagramsvg, callback) {
    const img = new Image()
    img.addEventListener('load', () => {
      // Create an ad-hoc canvas, same size as the SVG
      const bbox = diagramsvg.getBBox()

      const canvas = document.createElement('canvas')
      canvas.width = bbox.width
      canvas.height = bbox.height

      const context = canvas.getContext('2d')

      // Solid white background for non-transparent PNG
      context.fillStyle = 'white'
      context.fillRect(0, 0, canvas.width, canvas.height)

      // Draw the image on the canvas
      context.drawImage(img, 0, 0, bbox.width, bbox.height)

      canvas.toBlob((blob) => {
        if (this.downloadUrl) {
          try {
            URL.revokeObjectURL(this.downloadUrl)
            this.downloadUrl = undefined
          } catch { }
        }
        this.downloadUrl = URL.createObjectURL(blob)
        this.downloadImage(this.downloadUrl)
        if (callback) {
          callback(this.downloadUrl)
        }
      })
    })
    img.src = sourceurl
  },
  setFallbackMode () {
    this.diagramMode = 'fallback'
    // Clean out diagram
    this.clearDiagram()

    // Re-render mermaid diagram without
    // foreignObject elements
    mermaid.initialize({
      flowchart: {
        htmlLabels: false
      }
    })

    // Reset diagram element
    this.showDiagram(this.diagramTitle, this.diagramText)

    // Wait a bit, and save it
    const self = this
    setTimeout(() => {
      self.saveFallbackDiagram()
    }, 2000)
  },
  saveFallbackDiagram () {
    const diagram = document.getElementById('diagram')
    const fallbacksvg = diagram.querySelector('svg')
    const svgdata = (new XMLSerializer()).serializeToString(fallbacksvg)
    const blob = new Blob([svgdata], { type: 'image/svg+xml' })
    this.svgBlobUrl = URL.createObjectURL(blob)

    console.log('Downloading from blob url')
    this.createDownloadImage(this.svgBlobUrl, fallbacksvg)
  },
  chromeMessageReceived (request, sender) {
    if (!request || !request.status) {
      this.setErrorStatus(
        'Error: Wrong message sent. Please contact the developers.'
      )
      return
    }

    if (request.status === 'Error') {
      this.setErrorStatus('Error: ' + request.error)
      this.showOpenBranchEditorPanel()
      return
    }

    if (request.status === 'Success') {
      this.setStatus('')

      this.showDiagram(request.diagramTitle, request.diagramText)

      const self = this
      const saveButton = document.getElementById('savebutton')
      saveButton.addEventListener('click', (e) => {
        self.saveButtonClicked()
      })

      this.showElement('savepanel')
    }
  },
  saveButtonClicked (e) {
    if (this.downloadUrl) {
      console.log('Downloading from cached blob URL')
      this.downloadImage(this.downloadUrl)
      return
    }

    if (this.diagramMode === 'default') {
      const diagramsvg = document.querySelector('pre.dynmermaid svg')
      if (!diagramsvg) {
        window.alert('Diagram not found.')
        return
      }

      // First, convert the mermaid-generated svg into a data: url
      // with base64 encoding
      const svgdataUrl = 'data:image/svg+xml;base64,' +
        btoa(diagramsvg.outerHTML)

      console.log(`SIZE OF DATAURL:${svgdataUrl.length}`)
      // Chrome silently fails if the data url size goes beyond a
      // limit. The number 45000 is guesswork, based on tests.
      if (svgdataUrl.length <= 45000) {
        console.log('Downloading from data url')
        this.createDownloadImage(svgdataUrl, diagramsvg)
      } else {
        this.setFallbackMode()
      }
    } else {
      this.saveFallbackDiagram()
    }
  },
  visibilityChanged (e) {
    if (document.visibilityState === 'hidden') {
      if (this.downloadUrl) {
        URL.revokeObjectURL(this.downloadUrl)
        this.downloadUrl = ''
      }
      if (this.svgBlobUrl) {
        URL.revokeObjectURL(this.svgBlobUrl)
        this.svgBlobUrl = ''
      }
      console.log('Blobs cleared')
    }
  },
  init () {
    const self = this

    chrome.runtime.onMessage.addListener((request, sender) => {
      self.chromeMessageReceived(request, sender)
    })

    document.addEventListener('visibilitychange', (e) => {
      self.visibilityChanged(e)
    })
  }
}

diagramPage.init()
