(function () {
    "use strict";


    const finalresult = {
        status: "Error",
        error: "",
        diagramTitle: "",
        diagramText: ""
    }

    // This will work only on the Branching Options page while
    // editing a form in Microsoft Forms.

    const headingSpan = document.querySelector("span[role='heading']")
    if (!headingSpan) {
        finalresult.error = "Not on Branching Options screen (check 1)"
        return finalresult
    }

    if (headingSpan.innerText !== "Branching options") {
        finalresult.error = "Not on Branching Options screen (check 2)"
        return finalresult
    }

    // Begin a mermaid flowchart
    let result = "graph TD\nStart([Start])\nEnd([End])\nStart --> 1\n"

    // Questions are contained in elements which have the CSS class "office-form-question"
    const q = document.querySelectorAll(".office-form-question")

    const qcount = q.length

    q.forEach(function processQuestions(i) {
        // The question text, type and requireness are all contained as a single string
        // in an attribute called "aria-label" of the question element. 
        const question = i.getAttribute("aria-label")

        // The format of the question string is:
        // <number>. <text> <type> <required>
        const qnumregexp = /^(\d{1,5})\.(.*)$/

        const fromNum = question.match(qnumregexp)[1]

        // Just the question text can be found in a span element with 
        // the class .text-format-content, which is inside a div which
        // has a 'data-automation-id' attribute with the value
        // 'questionTitle', which is under the question container 
        // element.
        const titlespan = i.querySelector("div[data-automation-id='questionTitle'] span.text-format-content")
        let fromText = titlespan ? titlespan.innerText : "NO TITLE"

        // Mermaid cannot handle parenthesis, semicolons or angle brackets
        // in text
        fromText = fromText.replace(/[\(\/\);<>:]/g, '')

        // We want the number to be shown in flowchart boxes
        fromText = `${fromNum}. ${fromText}`


        // If a question is not currently selected for editing, it is contained
        // in a <button> element with a role attribute with value 'button'.
        const selectedforediting = !(i.getAttribute("role") == "button")

        // If a question can have multiple branches, the container element
        // will have a child div with a 'role' attribute with value 
        // 'radiogroup'. The branches will be organized below that. Each
        // branch is identified by a value in a span which has the class
        // '.text-format-content'.
        const branchgroup = i.querySelector('div[role="radiogroup"]')
        const firstbranchcondition = branchgroup ?
            branchgroup.querySelector("div[role='radio'] span.text-format-content") :
            undefined


        // If a question has just one branch, the destination question text
        // can be found in an element with the class '.dropdown-placeholder-text'
        // This element will not exist if the destination question is the 
        // next one in order
        const gotoElement = i.querySelector(".dropdown-placeholder-text")

        // A 'ratings' type question also has a div with the 'radiogroup' role.
        // It doesn't have branch identification values under it. And finally,
        // if a multi-branch question is selected for editing, the branch
        // destinations are found in spans with the class 
        // .dropdown-placeholder-text, which is the same marker as one-branch
        // questions. So, a question has mutiple branches only if:
        // 1. The 'radiogroup' role exists, AND... 
        // 2. EITHER
        //      a) span.text-format-content can be detected AND 
        //      b) .dropdown-placeholder-text cannot
        // 3. OR
        //      a) question is currently selected for editing AND
        //      b) .dropdown-placeholder-text can be detected
        const multiplebranches = (
            branchgroup &&
            (
                (firstbranchcondition && (!gotoElement)) ||
                (selectedforediting && gotoElement)
            )
        )

        if (multiplebranches) {
            // The question has multiple branches
            // Each destination is contained in a div with the 'role' attribute
            // having value 'radio'
            const choices = branchgroup.querySelectorAll("div[role='radio']")



            choices.forEach(function processChoices(choice) {
                // A span with the class '.text-format-content' will have the
                // choice value for the current destination
                const ifValue = choice.querySelector("span.text-format-content").innerText

                let toNum = ""

                let gotoElement = undefined

                if (selectedforediting) {
                    // If the question is current selected for editing, it is contained
                    // in a <div> tag. In this case, each choice's destination
                    // question string is available in a span with the class
                    // .dropdown-placeholder-text inside the choice element.
                    gotoElement = choice.querySelector("span.dropdown-placeholder-text")
                } else {
                    // If a question is not currently selected for editing, it is contained
                    // in a <button> element with a role attribute with value 'button'.
                    // If this case,
                    // the destination question string is contained in a structure like
                    // this inside the choice element:
                    // <div></div>
                    // <div>
                    //    <div>
                    //       <div></div>
                    //       <div>DESTINATION QUESTION HERE</div>
                    //    </div>
                    // </div>
                    gotoElement = choice.children[1]?.children[0]?.children[1]
                }

                toNum = getToNum(gotoElement, fromNum, qcount, qnumregexp)

                result += `${fromNum}[${fromText}] -->|${ifValue}|${toNum}\n`
            })
        } else {
            // Question has just one branch
            let toNum = getToNum(gotoElement, fromNum, qcount, qnumregexp)

            result += `${fromNum}[${fromText}] --> ${toNum}\n`
        }

    })

    // Write out result
    console.log("LOGGING FROM content.js:\n" + result)

    finalresult.status = "Success"
    finalresult.diagramTitle = document.title
    finalresult.diagramText = result
    return finalresult


    function getNextNum(fromNum, qcount) {
        const toNum = parseInt(fromNum) + 1
        return toNum > qcount ? "End" : toNum.toString()
    }

    function getToNum(gotoElement, fromNum, qcount, qnumregexp) {
        let toNum = ""
        if (gotoElement) {
            const gotoValue = gotoElement.innerText
            if (gotoValue == "Next") {
                toNum = getNextNum(fromNum, qcount)
            } else if (gotoValue == "End of the form") {
                toNum = "End";
            } else {
                toNum = gotoValue.match(qnumregexp)[1];
            }
        } else {
            toNum = getNextNum(fromNum, qcount)
        }
        return toNum
    }
})()