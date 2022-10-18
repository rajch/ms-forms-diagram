(function () {
    "use strict";

    // The format of a question string is:
    // <number>. <text> <type> <required>
    // The format of a section string is:
    // <number>. <section title>
    // This regexp matches both
    const qnumregexp = /^(\d{1,5})\.(.*)$/;

    function createChoice(choiceElement, questionObject) {
        // A span with the class '.text-format-content' will have the
        // choice value for the current destination
        const ifValue =
            choiceElement.querySelector(
                'span.text-format-content'
            ).innerText

        let choicegotoElement = undefined;

        if (questionObject.selectedForEditing()) {
            // If the question is current selected for editing, it is
            // contained in a <div> tag. In this case, each choice's
            // destination string is available in a span with the class
            // .dropdown-placeholder-text inside the choice element.
            choicegotoElement =
                choiceElement.querySelector('span.dropdown-placeholder-text')
        } else {
            // If a question is not currently selected for editing, it is
            // contained in a <button> element with a role attribute with
            // value 'button'. In this case, the destination string is
            // available in a structure like this inside the choice
            // element:
            // <div></div>
            // <div>
            //    <div>
            //       <div></div>
            //       <div>DESTINATION QUESTION HERE</div>
            //    </div>
            // </div>
            choicegotoElement = choiceElement.children[1]?.children[0]?.children[1];
        }

        const destinationstring = choicegotoElement ?
            choicegotoElement.innerText :
            ''

        return {
            ifValue() {
                return ifValue
            },
            destionationString() {
                return destinationstring
            }
        }
    }

    function createQuestion(questionElement, sectionObject) {
        // The question text, type and requireness are all contained 
        // as a single string in the "aria-label" attribute of the
        // question element. 
        const title = questionElement.getAttribute('aria-label')

        const id = title.match(qnumregexp)[1]

        // Just the question text can be found in a span element with 
        // the class .text-format-content, which is inside a div which
        // has a 'data-automation-id' attribute with the value
        // 'questionTitle', which is under the question container 
        // element.
        const titletextspan = questionElement.querySelector(
            'div[data-automation-id="questionTitle"] span.text-format-content'
        )
        let titletext = titletextspan ? titletextspan.innerText : "NO TITLE";

        // Mermaid cannot handle parenthesis, semicolons or angle brackets
        // in text
        titletext = titletext.replace(/[\(\/\);<>:]/g, '');

        // We want the number to be shown in flowchart boxes
        titletext = `${id}. ${titletext}`;


        // If a question is not currently selected for editing, the element
        // is a <button> with a role attribute with value 'button'.
        // Otherwise, it is a <div> without that attribute.
        const selectedforediting =
            (questionElement.getAttribute('role') !== 'button')

        // If a question can have multiple branches, the container element
        // will have a child div with a 'role' attribute with value 
        // 'radiogroup'. The branches will be organized below that. Each
        // branch is identified by a value in a span which has the class
        // '.text-format-content'.
        const branchgroup = questionElement.querySelector(
            'div[role="radiogroup"]'
        )
        const firstbranchcondition = branchgroup ?
            branchgroup.querySelector("div[role='radio'] span.text-format-content") :
            undefined

        // If a question has just one branch, the destination question text
        // can be found in an element with the class '.dropdown-placeholder-text'
        // This element will not exist if the destination question is the 
        // next one in order
        const gotoElement = questionElement.querySelector(
            ".dropdown-placeholder-text"
        )

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

        const destination = gotoElement && !multiplebranches ?
            gotoElement.innerText :
            ''

        const questionObject = {
            choices: [],
            id() {
                return id
            },
            titleText() {
                return titletext
            },
            title() {
                return `${id}. ${titletext}`
            },
            selectedForEditing() {
                return selectedforediting
            },
            multipleBranches() {
                return multiplebranches
            },
            destionationString() {
                return destination
            },
            section() {
                return sectionObject
            },
            isLastInSection() {
                return sectionObject ?
                    id == sectionObject.lastQuestionId() :
                    false
            }
        }

        // If the question has multiple branches
        // Each choice is contained in a div with the 'role' attribute
        // having value 'radio' under the questiom element
        if (multiplebranches) {
            const choiceElements =
                branchgroup.querySelectorAll("div[role='radio']")

            for (let i = 0, l = choiceElements.length; i < l; i++) {
                questionObject.choices.push(
                    createChoice(choiceElements[i], questionObject)
                )
            }
        }

        return questionObject
    }

    function createSection(sectionOrdinal, sectionMarkerElement) {
        // Sections are contained in structures like this:
        // <div>
        //      <div>
        //          <div aria-label="{SECTION TITLE}">
        //              <div></div>
        //              <div>
        //                  <div role="heading" data-automation-id="SectionTitle"></div>
        //              </div<
        //          </div>
        //      <div>
        //      { QUESTION ELEMENTS }
        //      <div>
        //          {SEVERAL LEVELS DOWN}
        //          <span class="dropdown-placeholder-text">{{ SECTION NEXT }}</span>
        //      </di>
        // </div>

        const secLabel = sectionMarkerElement.closest('[aria-label]')
        const secContainer = secLabel.parentNode.parentNode
        const destinationSpan = secContainer.querySelector(
            ':scope >div:last-child span.dropdown-placeholder-text'
        )
        const destination = destinationSpan ?
            destinationSpan.innerText :
            "Next"
        const titletext = secLabel.getAttribute('aria-label')

        const section = {
            questions: [],
            ordinal() {
                return sectionOrdinal
            },
            id() {
                return `Section${this.ordinal()}`
            },
            titleText() {
                return titletext
            },
            title() {
                let titleValue = titletext == 'Section title' ?
                    '' :
                    titletext

                return `${sectionOrdinal}. ${titleValue}`.trimEnd()
            },
            firstQuestionId() {
                if (this.questions && this.questions.length) {
                    return this.questions[0].id()
                } else {
                    return ""
                }
            },
            lastQuestionId() {
                const qs = this.questions
                if (qs && qs.length) {
                    return qs[qs.length - 1].id()
                } else {
                    return ""
                }
            },
            destionationString() {
                return destination
            }
        }

        const questionElements = section.container.querySelectorAll(
            ".office-form-question"
        )
        if (questionElements) {
            for (let i = 0, l = questionElements.length; i < l; i++) {
                const childQuestion = createQuestion(
                    questionElements[i], section
                )
                section.questions.push(childQuestion)
            }
        }

        return section
    }

    function createForm() {
        let sections = undefined
        let questions = undefined
        let qcount = 0
        let scount = 0

        // See createSection() above for the structure of section
        // elements.
        // The most reliable marker is data-automation-id="SectionTitle"
        const sectionElements = document.querySelectorAll('[data-automation-id="SectionTitle"]')

        const hassections = (sectionElements && sectionElements.length)
        if (hassections) {
            sections = []
            for (let i = 0; i < sectionElements.length; i++) {
                let sec = createSection(i + 1, sectionElements[i])
                sections.push(sec)
                qcount += sec.questions.length
            }
            scount = sections.length
        } else {
            questions = []
            const questionElements = document.querySelectorAll(".office-form-question")
            for (let i = 0; i < questionElements.length; i++) {
                questions.push(
                    createQuestion(questionElements[i], undefined)
                )
            }
            qcount = questions.length
        }

        return {
            getSectionId(sectionTitle) {
                let result = ""
                if (scount) {
                    let sec = sections.find((s) => sectionTitle == s.title())
                    result = sec ? sec.id() : ""
                }
                return result
            },
            getSectionDestination(question) {
                if (!question.section()) {
                    return this.getNextQuestion(question)
                }

                const sec = question.section()
                const secdest = sec.sectionDestination()
                if (secdest == "Next") {
                    const destsecordinal = sec.ordinal() + 1
                    if (destsecordinal > scount) {
                        return "End"
                    } else {
                        return `Section${destsecordinal}`
                    }
                }

                return this.getSectionId(secdest)
            },
            getNextQuestion(question) {
                const toNum = parseInt(question.id()) + 1
                return toNum > qcount ? "End" : toNum.toString()
            },
            getNextDestination(question) {
                if (question.isLastInSection()) {
                    return this.getSectionDestination(question)
                } else {
                    return this.getNextQuestion(question)
                }
            },
            getDestinationId(gotoValue, question) {
                // Empty gotoValue implies next
                if (!gotoValue) {
                    return this.getNextDestination(question)
                }

                if (gotoValue == 'Next') {
                    return this.getNextDestination(question)
                }

                if (gotoValue == 'End of the form') {
                    return "End"
                }

                // gotoValue is a title. Could be section or
                // question. Check section first.
                const toSection = this.getSectionId(gotoValue)
                if (toSection) {
                    return toSection
                }

                // Now it must be a question title. Extract id
                return gotoValue.match(qnumregexp)[1];
            },
            processQuestion(question) {
                const thisForm = this

                let result = ""

                const qId = question.id()
                const qTitle = question.titleText()

                if (question.multipleBranches()) {
                    result += `${qId}[${qTitle}]\n`

                    const choices = question.choices

                    choices.forEach(function processChoices(choice) {
                        const ifValue = choice.ifValue()
                        const destId = thisForm.getDestinationId(
                            choice.destionationString(), question
                        )

                        result += `${qId} -->|${ifValue}|${destId}\n`;
                    })
                } else {
                    const destId = thisForm.getDestinationId(
                        question.destionationString(), question
                    )

                    result += `${qId}[${qTitle}] --> ${destId}\n`;
                }

                return result
            },
            processSection(section) {
                let result = ""

                const sec = section
                result += `${sec.id()}\{\{${sec.titleText()}\}\}\n`
                result += `${sec.id()} --> ${sec.firstQuestionId()}\n`

                sec.questions.forEach((q) => {
                    result += this.processQuestion(q)
                })

                return result
            },
            process() {
                // Begin a mermaid flowchart
                let result = "graph TD\nStart([Start])\nEnd([End])\n"

                // Check for sections
                if (hassections) {
                    result += "Start --> Section1\n"

                    sections.forEach((s) => {
                        result += this.processSection(s)
                    })
                } else {
                    result += "Start --> 1\n"

                    questions.forEach((q) => {
                        result += this.processQuestion(q)
                    })
                }

                return result
            }
        }
    }

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

    const f = createForm()
    const result = f.process()

    // Write out result
    console.log("LOGGING FROM content.js:\n" + result)

    finalresult.status = "Success"
    finalresult.diagramTitle = document.title
    finalresult.diagramText = result
    return finalresult
})()