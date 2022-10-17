(function () {
    "use strict";

    // The format of a question string is:
    // <number>. <text> <type> <required>
    // The format of a section string is:
    // <number>. <section title>
    // This regexp matches both
    const qnumregexp = /^(\d{1,5})\.(.*)$/;

    function createQuestion(questionElement, sectionObject) {
        // The question text, type and requireness are all contained as a single string
        // in an attribute called "aria-label" of the question element. 
        const question = questionElement.getAttribute("aria-label");

        const fromNum = question.match(qnumregexp)[1];

        // Just the question text can be found in a span element with 
        // the class .text-format-content, which is inside a div which
        // has a 'data-automation-id' attribute with the value
        // 'questionTitle', which is under the question container 
        // element.
        const titlespan = questionElement.querySelector("div[data-automation-id='questionTitle'] span.text-format-content");
        let fromText = titlespan ? titlespan.innerText : "NO TITLE";

        // Mermaid cannot handle parenthesis, semicolons or angle brackets
        // in text
        fromText = fromText.replace(/[\(\/\);<>:]/g, '');

        // We want the number to be shown in flowchart boxes
        fromText = `${fromNum}. ${fromText}`;


        // If a question is not currently selected for editing, it is contained
        // in a <button> element with a role attribute with value 'button'.
        const selectedforediting = !(questionElement.getAttribute("role") == "button");

        // If a question can have multiple branches, the container element
        // will have a child div with a 'role' attribute with value 
        // 'radiogroup'. The branches will be organized below that. Each
        // branch is identified by a value in a span which has the class
        // '.text-format-content'.
        const branchgroup = questionElement.querySelector('div[role="radiogroup"]');
        const firstbranchcondition = branchgroup ?
            branchgroup.querySelector("div[role='radio'] span.text-format-content") :
            undefined;

        // If a question has just one branch, the destination question text
        // can be found in an element with the class '.dropdown-placeholder-text'
        // This element will not exist if the destination question is the 
        // next one in order
        const gotoElement = questionElement.querySelector(".dropdown-placeholder-text");

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
        );

        return {
            element: questionElement,
            id() {
                return fromNum
            },
            titleText() {
                return fromText
            },
            title() {
                return `${this.id()}. ${this.titleText()}`
            },
            selectedForEditing() {
                return selectedforediting
            },
            multipleBranches() {
                return multiplebranches
            },
            branchGroup() {
                return branchgroup
            },
            singleGotoElement() {
                return gotoElement
            },
            section() {
                return sectionObject
            },
            isLastInSection() {
                return sectionObject ?
                    this.id() == sectionObject.lastQuestionId() :
                    false
            }
        }
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
        const destination = destinationSpan ? destinationSpan.innerText : "Next"

        const section = {
            label: secLabel,
            container: secContainer,
            questions: [],
            ordinal() {
                return sectionOrdinal
            },
            id() {
                return `Section${this.ordinal()}`
            },
            titleText() {
                return this.label.getAttribute('aria-label')
            },
            title() {
                let titletext = this.titleText() == "Section title" ?
                    "" : this.titleText()
                return `${this.ordinal()}. ${titletext}`.trimEnd()
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
            sectionDestination() {
                return destination
            }
        }

        const childQuestionElements = section.container.querySelectorAll(
            ".office-form-question"
        )
        if (childQuestionElements) {
            childQuestionElements.forEach((q) => {
                const childQuestion = createQuestion(q, section)
                section.questions.push(childQuestion)
            })
        }

        return section
    }

    function createForm() {
        let sections = []
        let questions = undefined
        let qcount = 0
        let scount = 0

        // See createSection() above for the structure of section
        // elements.
        // The most reliable marker is data-automation-id="SectionTitle"
        const sectionElements = document.querySelectorAll('[data-automation-id="SectionTitle"]')
        console.log("Sections:", sectionElements)

        const hassections = (sectionElements && sectionElements.length)
        if (hassections) {
            for (let i = 0; i < sectionElements.length; i++) {
                let sec = createSection(i + 1, sectionElements[i])
                sections.push(sec)
                qcount += sec.questions.length
            }
            scount = sections.length
        } else {
            questions = []
            const questionElements = document.querySelectorAll(".office-form-question")
            questionElements.forEach((q) => {
                questions.push(createQuestion(q, undefined))
            })
            qcount = questions.length
        }

        return {
            hasSections() {
                return hassections
            },
            sections() {
                return sections
            },
            questions() {
                return questions
            },
            qCount() {
                return qcount
            },
            sCount() {
                return scount
            },
            getSectionId(sectionTitle) {
                let result = ""
                if (scount) {
                    for (let i = 0; i < scount; i++) {
                        let sec = sections[i]

                        if (sectionTitle == sec.title()) {
                            result = sec.id()
                            break;
                        }
                    }
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
                    if (destsecordinal > this.sCount()) {
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
                    console.log(`QUESTION ${question.title()} is the last in its section`)
                    return this.getSectionDestination(question)
                } else {
                    console.log(`QUESTION ${question.title()} is NOT the last in its section`)
                    return this.getNextQuestion(question)
                }
            },
            getDestination(gotoElement, question) {
                let dest = ""

                if (gotoElement) {
                    const gotoValue = gotoElement.innerText
                    if (gotoValue == "Next") {
                        dest = this.getNextDestination(question)
                    } else if (gotoValue == "End of the form") {
                        dest = "End";
                    } else {
                        //const toSection = getToSection(gotoValue, this.sections())
                        const toSection = this.getSectionId(gotoValue)
                        if (toSection) {
                            dest = toSection
                        } else {
                            dest = gotoValue.match(qnumregexp)[1];
                        }
                    }
                } else {
                    // Implied next
                    dest = this.getNextDestination(question)
                }
                return dest
            },
            processQuestion(questionObject) {
                const thisForm = this

                let result = ""
                const sections = this.sections()

                const fromNum = questionObject.id()
                const fromText = questionObject.titleText()

                if (questionObject.multipleBranches()) {
                    // The question has multiple branches
                    // Each destination is contained in a div with the 'role' attribute
                    // having value 'radio'
                    const choices = questionObject.branchGroup().querySelectorAll("div[role='radio']")

                    choices.forEach(function processChoices(choice) {
                        // A span with the class '.text-format-content' will have the
                        // choice value for the current destination
                        const ifValue = choice.querySelector("span.text-format-content").innerText

                        let choicegotoElement = undefined;

                        if (questionObject.selectedForEditing()) {
                            // If the question is current selected for editing, it is contained
                            // in a <div> tag. In this case, each choice's destination
                            // question string is available in a span with the class
                            // .dropdown-placeholder-text inside the choice element.
                            choicegotoElement = choice.querySelector("span.dropdown-placeholder-text")
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
                            choicegotoElement = choice.children[1]?.children[0]?.children[1];
                        }

                        const toNum = thisForm.getDestination(choicegotoElement, questionObject)

                        result += `${fromNum}[${fromText}] -->|${ifValue}|${toNum}\n`;
                    })
                } else {

                    // Question has just one branch
                    const toNum = thisForm.getDestination(questionObject.singleGotoElement(), questionObject);

                    result += `${fromNum}[${fromText}] --> ${toNum}\n`;
                }

                return result
            },
            process() {
                // Begin a mermaid flowchart
                let result = "graph TD\nStart([Start])\nEnd([End])\n"

                // Check for sections
                if (this.hasSections()) {
                    const sections = this.sections()
                    console.log("Sections:", sections)

                    result += "Start --> Section1\n"

                    for (let i = 0; i < sections.length; i++) {
                        const sec = sections[i]

                        result += `${sec.id()}\{\{${sec.titleText()}\}\}\n`
                        result += `${sec.id()} --> ${sec.firstQuestionId()}\n`

                        sec.questions.forEach((q) => {
                            result += this.processQuestion(q)
                        })
                    }
                } else {
                    result += "Start --> 1\n"

                    this.questions().forEach((q) => {
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