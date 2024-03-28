'use strict'

/* global mermaid, chrome */

const diagramPage = {
  downloadUrl: undefined,
  svgBlobUrl: undefined,
  diagramTitle: '',
  diagramText: '',
  diagramStyle: 'basis',
  diagramPrimaryColor: '#ececff',
  diagramThemeName: 'base',
  diagramThemePrimaryColor: '#ececff',
  diagramThemeSecondaryColor: '',
  diagramThemePrimaryBorderColor: undefined,

  hexToHsl (hexColor) {
    // Remove the hash symbol if present
    const cleanedHex = hexColor.replace(/^#/, '')

    // Convert the hex value to RGB
    const r = parseInt(cleanedHex.substring(0, 2), 16) / 255
    const g = parseInt(cleanedHex.substring(2, 4), 16) / 255
    const b = parseInt(cleanedHex.substring(4, 6), 16) / 255

    // Find the maximum and minimum RGB values
    const max = Math.max(r, g, b)
    const min = Math.min(r, g, b)

    // Calculate the luminosity (lightness)
    const l = (max + min) / 2

    // Calculate the saturation
    let s
    if (max === min) {
      s = 0 // No saturation (gray)
    } else {
      s = l > 0.5 ? (max - min) / (2 - max - min) : (max - min) / (max + min)
    }

    // Calculate the hue
    let h
    if (max === min) {
      h = 0 // No hue (gray)
    } else if (max === r) {
      h = (g - b) / (max - min)
    } else if (max === g) {
      h = 2 + (b - r) / (max - min)
    } else {
      h = 4 + (r - g) / (max - min)
    }
    h = (h * 60 + 360) % 360

    // Round values and return the HSL object
    return {
      h: Math.round(h),
      s: Math.round(s * 100),
      l: Math.round(l * 100)
    }
  },
  hslToHex (hslObject) {
    const { h, s, l } = hslObject

    // Convert hue to the range [0, 360)
    const hue = ((h % 360) + 360) % 360

    // Normalize saturation and luminosity to [0, 1]
    const saturation = Math.max(0, Math.min(1, s / 100))
    const luminosity = Math.max(0, Math.min(1, l / 100))

    // Calculate chroma (colorfulness)
    const chroma = (1 - Math.abs(2 * luminosity - 1)) * saturation

    // Calculate intermediate values
    const x = chroma * (1 - Math.abs(((hue / 60) % 2) - 1))
    const m = luminosity - chroma / 2

    // Convert to RGB
    let r, g, b
    if (hue >= 0 && hue < 60) {
      r = chroma
      g = x
      b = 0
    } else if (hue >= 60 && hue < 120) {
      r = x
      g = chroma
      b = 0
    } else if (hue >= 120 && hue < 180) {
      r = 0
      g = chroma
      b = x
    } else if (hue >= 180 && hue < 240) {
      r = 0
      g = x
      b = chroma
    } else if (hue >= 240 && hue < 300) {
      r = x
      g = 0
      b = chroma
    } else {
      r = chroma
      g = 0
      b = x
    }

    // Convert RGB to 8-bit integers
    const red = Math.round((r + m) * 255)
    const green = Math.round((g + m) * 255)
    const blue = Math.round((b + m) * 255)

    // Convert to hex format
    const hexColor = `#${red.toString(16).padStart(2, '0')}${green.toString(16).padStart(2, '0')}${blue.toString(16).padStart(2, '0')}`
    return hexColor
  },
  lighten (hexColor, luminosity) {
    const hslv = this.hexToHsl(hexColor)
    return this.hslToHex({
      h: hslv.h,
      s: hslv.s,
      l: luminosity
    })
  },

  themeFuncs: {
    default (self) {
      self.diagramThemeName = 'base'
      self.diagramThemePrimaryColor = '#ececff'
      self.diagramThemeSecondaryColor = '#e8e8e8'
      self.diagramThemePrimaryBorderColor = self.diagramPrimaryColor
    },
    themed (self) {
      self.diagramThemeName = 'base'
      self.diagramThemePrimaryColor = self.lighten(self.diagramPrimaryColor, 90)
      self.diagramThemeSecondaryColor = '#ffffff'
      self.diagramThemePrimaryBorderColor = self.diagramPrimaryColor
    },
    mono (self) {
      self.diagramThemeName = 'neutral'
      self.diagramThemePrimaryColor = '#ececff'
      self.diagramThemeSecondaryColor = '#e8e8e8'
      self.diagramThemePrimaryBorderColor = undefined
    }
  },
  getLocalizedMessage (messagename) {
    return chrome.i18n.getMessage(messagename)
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

    return element
  },
  setDiagramTitle (diagramtitle) {
    this.diagramTitle = diagramtitle
    this.setElementText('diagramtitle', diagramtitle)
  },
  setStatus (text) {
    this.setElementText('statuspanel', text)
  },
  setErrorStatus (text) {
    this.setDiagramTitle('Error')
    this.setStatus(text)
  },
  setDiagramText (text) {
    this.diagramText = text

    this.refreshDiagram()
  },
  setDiagramStyle (stylename) {
    this.diagramStyle = stylename

    this.refreshDiagram()
  },
  setDiagramTheme (themename) {
    this.themeFuncs[themename](this)

    this.refreshDiagram()
  },
  diagramSource (noHTMLFlag) {
    return [
      '---',
      'config:',
      `  theme: ${this.diagramThemeName}`,
      '  themeVariables:',
      `    primaryColor: "${this.diagramThemePrimaryColor}"`,
      `    secondaryColor: "${this.diagramThemeSecondaryColor}"`,
      this.diagramPrimaryColor ? `    primaryBorderColor: "${this.diagramPrimaryColor}"` : '',
      '  flowchart:',
      `    curve: ${this.diagramStyle}`,
      noHTMLFlag ? '    htmlLabels: false' : '',
      '---',
      `${this.diagramText}`
    ].join('\n')
  },
  refreshDiagram () {
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
  showOpenBranchEditorPanel () {
    this.showElement('openbrancheditor')
    this.hideElement('diagram')
  },
  sanitiseMermaidSvg (text) {
    return text.replaceAll('&nbsp;', ' ')
  },
  createPngBlob (callback) {
    const self = this

    mermaid.render('diagramPng', this.diagramSource(true))
      .then((result) => {
        const cleansvg = this.sanitiseMermaidSvg(result.svg)
        const b = new Blob(
          [cleansvg],
          { type: 'image/svg+xml;charset=utf-8' }
        )

        if (self.svgBlobUrl) {
          try {
            URL.revokeObjectURL(self.svgBlobUrl)
            self.svgBlobUrl = undefined
          } catch (error) {
            console.debug('Error in createPngBlob self.svgBlobUrl revoke:', error)
            window.alert(this.getLocalizedMessage('errCouldNotCompleteOperation'))
          }
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
      }).catch((error) => {
        console.debug('Error in createPngBlob mermaid.render:', error)
        window.alert(this.getLocalizedMessage('errCouldNotCompleteOperation'))
      })
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
    a.href = imageUrl
    a.click()
    a.remove()
  },
  colorDropdownChanged (e) {
    const selectedThemeSetting = e.target.value
    this.setDiagramTheme(selectedThemeSetting)
  },
  styleDropdownChanged (e) {
    const selectedstyle = e.target.value
    this.setDiagramStyle(selectedstyle)
  },
  copyButtonClicked (e) {
    const self = this

    self.createPngBlob((blob) => {
      try {
        navigator.clipboard.write([
          new ClipboardItem({
            'image/png': blob
          })
        ])
        window.alert(this.getLocalizedMessage('msgDiagramCopied'))
      } catch (error) {
        console.debug('Error in copyButtonClicked:', error)
        window.alert(this.getLocalizedMessage('errCouldNotCompleteOperation'))
      }
    })
  },
  copySourceButtonClicked (e) {
    navigator.clipboard.writeText(this.diagramSource())
    window.alert(this.getLocalizedMessage('msgSourceCopied'))
  },
  saveButtonClicked (e) {
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
  visibilityChanged (e) {
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
  chromeMessageReceived (request, sender) {
    if (!request || !request.status) {
      this.setErrorStatus(
        this.getLocalizedMessage('errWrongMessage')
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
      saveButton?.addEventListener('click', (e) => {
        self.saveButtonClicked(e)
      })

      const copybutton = document.getElementById('copybutton')
      copybutton?.addEventListener('click', (e) => {
        self.copyButtonClicked(e)
      })

      const copysourcebutton = document.getElementById('copysourcebutton')
      copysourcebutton?.addEventListener('click', (e) => {
        self.copySourceButtonClicked(e)
      })

      const styledropdown = document.getElementById('styledropdown')
      styledropdown?.addEventListener('change', (e) => {
        self.styleDropdownChanged(e)
      })

      const colordropdown = document.getElementById('colordropdown')
      colordropdown?.addEventListener('change', (e) => {
        self.colorDropdownChanged(e)
      })

      this.showElement('savepanel')
    }
  },
  init () {
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
