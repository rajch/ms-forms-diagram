'use strict'

/* global chrome */

const FORMSHOST = 'forms.office.com'
const FORMSHOST2 = 'forms.microsoft.com'
const PATHPREFIX1 = '/pages/designpagev2.aspx'
const PATHPREFIX2 = '/Pages/DesignPageV2.aspx'
const MSFORMSDESIGNPAGE = `https://${FORMSHOST}${PATHPREFIX1}`
const MSFORMSDESIGNPAGE2 = `https://${FORMSHOST}${PATHPREFIX2}`
const MSFORMSDESIGNPAGE3 = `https://${FORMSHOST2}${PATHPREFIX1}`
const MSFORMSDESIGNPAGE4 = `https://${FORMSHOST2}${PATHPREFIX2}`

chrome.runtime.onInstalled.addListener(
  function extensionInstalled (details) {
    chrome.action.disable()

    // Ensure that the extension icon only "lights up" on Microsoft Forms
    // diagram edit pages. As on 22-May-2023, that means the following 4
    // urls, with at least the query string option shown:
    //
    // https://forms.office.com/pages/designpagev2.aspx?subpage=design
    // https://forms.office.com/pages/DesignPageV2.aspx?subpage=design
    // https://forms.microsoft.com/pages/designpagev2.aspx?subpage=design
    // https://forms.microsoft.com/pages/DesignPageV2.aspx?subpage=design
    //
    // Yes, those are four different urls. URLs are case-sensitive.
    //
    // "Lighting up" is achieved by changing the icon from the default
    // gray to full-color. This requires ImageData values, not paths.
    chrome.declarativeContent.onPageChanged.removeRules(
      undefined,
      async function rulesAdding () {
        const formsrule = {
          conditions: [
            new chrome.declarativeContent.PageStateMatcher({
              pageUrl: { hostSuffix: FORMSHOST, pathPrefix: PATHPREFIX1, queryContains: 'subpage=design', schemes: ['https'] }
            }),
            new chrome.declarativeContent.PageStateMatcher({
              pageUrl: { hostSuffix: FORMSHOST, pathPrefix: PATHPREFIX2, queryContains: 'subpage=design', schemes: ['https'] }
            }),
            new chrome.declarativeContent.PageStateMatcher({
              pageUrl: { hostSuffix: FORMSHOST2, pathPrefix: PATHPREFIX1, queryContains: 'subpage=design', schemes: ['https'] }
            }),
            new chrome.declarativeContent.PageStateMatcher({
              pageUrl: { hostSuffix: FORMSHOST2, pathPrefix: PATHPREFIX2, queryContains: 'subpage=design', schemes: ['https'] }
            })
          ],
          actions: [
            new chrome.declarativeContent.SetIcon({
              imageData: {
                16: await loadImageData('images/icon16.png'),
                32: await loadImageData('images/icon32.png'),
                48: await loadImageData('images/icon48.png'),
                128: await loadImageData('images/icon128.png')
              }
            }),
            new chrome.declarativeContent.ShowAction()
          ]
        }

        chrome.declarativeContent.onPageChanged.addRules([formsrule])
      }
    )
  }
)

async function loadImageData (url) {
  const img = await createImageBitmap(await (await fetch(chrome.runtime.getURL(url))).blob())
  const { width: w, height: h } = img
  const canvas = new OffscreenCanvas(w, h)
  const ctx = canvas.getContext('2d')
  ctx.drawImage(img, 0, 0, w, h)
  return ctx.getImageData(0, 0, w, h)
}

chrome.action.onClicked.addListener(
  function actionClicked (tab) {
    if (
      tab.url.includes(MSFORMSDESIGNPAGE) || tab.url.includes(MSFORMSDESIGNPAGE2) ||
      tab.url.includes(MSFORMSDESIGNPAGE3) || tab.url.includes(MSFORMSDESIGNPAGE4)
    ) {
      chrome.scripting.executeScript(
        {
          target: { tabId: tab.id },
          files: [
            'scripts/content.js'
          ]
        },
        function executeScriptCallback (injectionresults) {
          for (const injectionresult of injectionresults) {
            console.log(injectionresult.result)

            chrome.tabs.create(
              {
                url: chrome.runtime.getURL('diagram.html'),
                active: true
              },
              function tabCreated (tab) {
                // Give the new tab about a second to load, and then
                // send it a message.
                setTimeout(
                  () => chrome.tabs.sendMessage(tab.id, injectionresult.result),
                  1000
                )
              }
            )
          }
        }
      )
    } else {
      // Wrong page
      chrome.action.setBadgeText({ tabId: tab.id, text: 'NO' })
    }
  }
)
