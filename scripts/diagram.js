'use strict'

/* global mermaid, chrome */

const diagramPage = {
  downloadUrl: undefined,
  svgBlobUrl: undefined,
  diagramTitle: '',
  diagramText: '',
  diagramPrimaryColor: '#ececff',
  diagramStyle: 'basis',
  hideElement(elementid) {
    const element = document.getElementById(elementid)
    element.classList.add('hidden')
  },
  showElement(elementid, hide) {
    const element = document.getElementById(elementid)
    element.classList.remove('hidden')
  },
  setElementText(elementid, text) {
    const element = document.getElementById(elementid)

    element.innerHTML = ''
    element.append(document.createTextNode(text))

    return element
  },
  setDiagramTitle(diagramtitle) {
    this.diagramTitle = diagramtitle
    this.setElementText('diagramtitle', diagramtitle)
  },
  setStatus(text) {
    this.setElementText('statuspanel', text)
  },
  setErrorStatus(text) {
    this.setDiagramTitle('Error')
    this.setStatus(text)
  },
  setDiagramText(text) {
    this.diagramText = text

    this.refreshDiagram()
  },
  setDiagramStyle(stylename) {
    this.diagramStyle = stylename

    this.refreshDiagram()
  },
  diagramSource(noHTMLFlag) {
    return [
      '---',
      'config:',
      '  theme: base',
      '  themeVariables:',
      '    primaryColor: "#ececff"',
      '    secondaryColor: "#e8e8e8"',
      `    primaryBorderColor: "${this.diagramPrimaryColor}"`,
      '  flowchart:',
      `    curve: ${this.diagramStyle}`,
      noHTMLFlag && '    htmlLabels: false' || '',
      '---',
      `${this.diagramText}`
    ].join('\n')
  },
  refreshDiagram() {
    const self = this

    const diagramElement = document.getElementById('diagram')
    diagramElement.removeAttribute('data-processed')

    // The setTimeout is needed to let mermaid queue the render
    setTimeout(() => {
      mermaid.render('diagramSvg', self.diagramSource()).then((result) => {
        diagramElement.innerHTML = result.svg
      }).catch((error) => {
        self.setErrorStatus(error)
      })
    }, 0)

  },
  showOpenBranchEditorPanel() {
    this.showElement('openbrancheditor')
    this.hideElement('diagram')
  },
  createPngBlob(callback) {
    const self = this

    mermaid.render('diagramPng', this.diagramSource(true))
      .then((result) => {
        const b = new Blob([result.svg], { type: 'image/svg+xml' })

        if (self.svgBlobUrl) {
          try {
            URL.revokeObjectURL(self.svgBlobUrl)
            self.svgBlobUrl = undefined
          } catch { }
        }
        self.svgBlobUrl = URL.createObjectURL(b)

        const img = new Image()
        img.addEventListener('load', () => {
          const diagramsvg = document.getElementById('diagramSvg')
          if (!diagramsvg) {
            return
          }

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
            callback(blob)
          })
        })
        img.src = self.svgBlobUrl
      }).catch(() => { })
  },
  downloadImage(imageUrl) {
    // Create a temporary link, configure it for download,
    // point it to the canvas-generated data url, and
    // 'click' it to trigger a download.
    const a = document.createElement('a')

    const downloadFileName = this.diagramTitle
      ? `${this.diagramTitle}.png`
      : 'image.png'
    a.download = downloadFileName
    a.href = imageUrl
    a.click()
    a.remove()
  },
  styleDropdownChanged(e) {
    const selectedstyle = e.target.value
    this.setDiagramStyle(selectedstyle)
  },
  copyButtonClicked(e) {
    const self = this

    self.createPngBlob((blob) => {
      try {
        navigator.clipboard.write([
          new ClipboardItem({
            'image/png': blob
          })
        ])
        window.alert('Diagram copied to clipboard.')
      } catch (error) {
        self.setErrorStatus(error)
      }
    })
  },
  copySourceButtonClicked(e) {
    navigator.clipboard.writeText(this.diagramSource())
    window.alert('Mermaid source copied to clipboard.')
  },
  saveButtonClicked(e) {
    const self = this

    if (self.downloadUrl) {
      console.log('Downloading from cached blob URL')
      self.downloadImage(self.downloadUrl)
      return
    }

    self.createPngBlob((blob) => {
      self.downloadUrl = URL.createObjectURL(blob)
      self.downloadImage(self.downloadUrl)
    })
  },
  visibilityChanged(e) {
    if (document.visibilityState === 'hidden') {
      if (this.downloadUrl) {
        URL.revokeObjectURL(this.downloadUrl)
        this.downloadUrl = undefined
        console.log('Download png cleared')
      }
      if (this.svgBlobUrl) {
        URL.revokeObjectURL(this.svgBlobUrl)
        this.svgBlobUrl = undefined
        console.log('Copy svg cleared')
      }
      console.log('Blobs cleared')
    }
  },
  chromeMessageReceived(request, sender) {
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

      // Set colors per theme
      if (request.themePrimaryColor) {
        document.body.style.setProperty('--highlight-color', request.themePrimaryColor)
        this.diagramPrimaryColor = request.themePrimaryColor.toLowerCase()
      }

      this.setDiagramTitle(request.diagramTitle)
      this.setDiagramText(request.diagramText)

      const self = this

      const saveButton = document.getElementById('savebutton')
      saveButton.addEventListener('click', (e) => {
        self.saveButtonClicked(e)
      })

      const copybutton = document.getElementById('copybutton')
      copybutton?.addEventListener('click', (e) => {
        self.copyButtonClicked(e)
      })

      const copysourcebutton = document.getElementById('copysourcebutton')
      copysourcebutton.addEventListener('click', (e) => {
        self.copySourceButtonClicked(e)
      })

      const styledropdown = document.getElementById('styledropdown')
      styledropdown.addEventListener('change', (e) => {
        self.styleDropdownChanged(e)
      })

      this.showElement('savepanel')
    }
  },
  init() {
    const self = this

    mermaid.initialize({ startOnLoad: false })

    chrome.runtime.onMessage.addListener((request, sender) => {
      self.chromeMessageReceived(request, sender)
    })

    document.addEventListener('visibilitychange', (e) => {
      self.visibilityChanged(e)
    })
  }
}

diagramPage.init()
