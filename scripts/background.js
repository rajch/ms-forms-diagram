"use strict";

const MSFORMSDESIGNPAGE = "https://forms.office.com/pages/designpagev2.aspx"

chrome.action.onClicked.addListener(
    function actionClicked(tab) {
        if (tab.url.includes(MSFORMSDESIGNPAGE)) {
            chrome.scripting.executeScript(
                {
                    target: { tabId: tab.id },
                    files: [
                        "scripts/content.js"
                    ]
                },
                function executeScriptCallback(injectionresults) {
                    for (const injectionresult of injectionresults) {
                        console.log(injectionresult.result)

                        chrome.tabs.create(
                            {
                                url: chrome.runtime.getURL('diagram.html'),
                                active: true
                            },
                            function tabCreated(tab) {
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
            );
        } else {
            // Wrong page
            chrome.action.setBadgeText({tabId: tab.id, text:"NO"})
        }
    }
);