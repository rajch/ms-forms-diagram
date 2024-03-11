(function () {
  'use strict'

  // The form can be reduced to a data structure, from which a mermaid
  // diagram can be generated. A form is made up of multiple questions
  // or multiple sections, each section having multiple questions. A
  // question or a section will have a destination string, which shows
  // the next question or section in the branching sequence.

  // The format of a question's destination string is:
  // <number>. <question text>
  // The format of a section's destination string is:
  // <number>. <section title>
  // or just
  // <number>.
  // if the section does not have a title.
  // This regexp matches both cases.
  const qnumregexp = /^(\d{1,5})\.(.*)$/

  // Strings in mermaid, if they contain special characters, should be
  // enclosed in double quotes, with any actual double quotes encoded.
  function sanitiseForMermaid (text) {
    return `"${text.replaceAll('"', '#quot;')}"`
  }

  // Titles do not need surrounding quotes
  const sanitiseTitle = (text) => `${text.replaceAll('"', '#quot;')}`

  // A form consists of multiple questions. Each question is contained
  // in an HTML element, which has a specific structure. The structure
  // contains the question's ordinal (id) and text (titleText). It
  // also contains a pointer to the next question (destinationString),
  // or multiple pointers if it is multiple-choice (choices).
  // A form may optionally be divided into sections. Each section is
  // contained in an HTML element, which contains a set of question
  // elements, and a pointer to the next section (destinationString).
  function parseForm () {
    let sections
    let questions
    let qcount = 0
    let scount = 0

    // See parseSection() to understand the structure of section
    // elements.
    // The most reliable marker is an element, inside the section
    // container element, which has an attribute called
    // 'data-automation-id' with the value "SectionTitle".
    const sectionElements = document.querySelectorAll(
      '[data-automation-id="SectionTitle"]'
    )

    const hassections = (sectionElements && sectionElements.length)
    if (hassections) {
      // Each section has its own question elements
      sections = []
      for (let i = 0, l = sectionElements.length; i < l; i++) {
        const sec = parseSection(i + 1, sectionElements[i])
        sections.push(sec)
        qcount += sec.questions.length
      }
      scount = sections.length
    } else {
      // All question elements are directly availble in the
      // form body.
      // See parseQuestion() to understand the structure of
      // question elements.
      // Every question element has the css class
      // '.office-form-question'
      questions = []
      const questionElements = document.querySelectorAll(
        '.office-form-question'
      )
      for (let i = 0, l = questionElements.length; i < l; i++) {
        questions.push(
          parseQuestion(questionElements[i], undefined)
        )
      }
      qcount = questions.length
    }

    return {
      getSectionId (sectionTitle) {
        let result = ''
        if (scount) {
          const sec = sections.find((s) => sectionTitle === s.title())
          result = sec ? sec.id() : ''
        }
        return result
      },
      getSectionDestination (question) {
        if (!question.section()) {
          return this.getNextQuestion(question)
        }

        const sec = question.section()
        const secdest = sec.destinationString()
        if (secdest === 'Next') {
          const destsecordinal = sec.ordinal() + 1
          if (destsecordinal > scount) {
            return 'End'
          } else {
            return `Section${destsecordinal}`
          }
        }

        return this.getSectionId(secdest)
      },
      getNextQuestion (question) {
        const toNum = parseInt(question.id()) + 1
        return toNum > qcount ? 'End' : toNum.toString()
      },
      getNextDestination (question) {
        if (question.isLastInSection()) {
          return this.getSectionDestination(question)
        } else {
          return this.getNextQuestion(question)
        }
      },
      getDestinationId (gotoValue, question) {
        // Empty gotoValue implies next
        if (!gotoValue) {
          return this.getNextDestination(question)
        }

        if (gotoValue === 'Next') {
          return this.getNextDestination(question)
        }

        if (gotoValue === 'End of the form') {
          return 'End'
        }

        // gotoValue is a title. Could be section or
        // question. Check section first.
        const toSection = this.getSectionId(gotoValue)
        if (toSection) {
          return toSection
        }

        // Now it must be a question title. Extract id
        return gotoValue.match(qnumregexp)[1]
      },
      processQuestion (question) {
        const thisForm = this

        let result = ''

        const qId = question.id()
        const qTitle = question.titleText()

        if (question.multipleBranches()) {
          result += `${qId}[${qTitle}]\n`

          const choices = question.choices

          choices.forEach(function processChoices (choice) {
            const choiceTitle = choice.title()
            const destId = thisForm.getDestinationId(
              choice.destinationString(), question
            )

            result += `${qId} -->|${choiceTitle}|${destId}\n`
          })
        } else {
          const destId = thisForm.getDestinationId(
            question.destinationString(), question
          )

          result += `${qId}[${qTitle}] --> ${destId}\n`
        }

        return result
      },
      processSection (section) {
        let result = ''

        const sec = section
        result += `${sec.id()}{{${sec.titleText()}}}\n`
        result += `${sec.id()} --> ${sec.firstQuestionId()}\n`

        sec.questions.forEach((q) => {
          result += this.processQuestion(q)
        })

        return result
      },
      process () {
        // Begin a mermaid flowchart
        let result = 'graph TD\nStart([Start])\nEnd([End])\n'

        // Check for sections
        if (hassections) {
          result += 'Start --> Section1\n'

          sections.forEach((s) => {
            result += this.processSection(s)
          })
        } else {
          result += 'Start --> 1\n'

          questions.forEach((q) => {
            result += this.processQuestion(q)
          })
        }

        return result
      }
    }
  }

  function parseSection (sectionOrdinal, sectionMarkerElement) {
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
    const destination = destinationSpan
      ? destinationSpan.innerText
      : 'Next'

    const titletext = sanitiseTitle(secLabel.getAttribute('aria-label'))
    // If no specific title is given to a section, a default
    // value of 'Section title' is shown. This value is NOT
    // included in destination strings.
    const titleValue = titletext === 'Section title'
      ? ''
      : titletext

    const title = `${sectionOrdinal}. ${titleValue}`.trimEnd()

    const section = {
      questions: [],
      ordinal () {
        return sectionOrdinal
      },
      id () {
        return `Section${sectionOrdinal}`
      },
      titleText () {
        return titletext
      },
      title () {
        return title
      },
      firstQuestionId () {
        if (this.questions && this.questions.length) {
          return this.questions[0].id()
        } else {
          return ''
        }
      },
      lastQuestionId () {
        const qs = this.questions
        if (qs && qs.length) {
          return qs[qs.length - 1].id()
        } else {
          return ''
        }
      },
      destinationString () {
        return destination
      }
    }

    // Question elements inside a section container all have
    // the css class '.office-form-question'
    const questionElements = secContainer.querySelectorAll(
      '.office-form-question'
    )
    if (questionElements) {
      for (let i = 0, l = questionElements.length; i < l; i++) {
        const childQuestion = parseQuestion(
          questionElements[i], section
        )
        section.questions.push(childQuestion)
      }
    }

    return section
  }

  function parseQuestion (questionElement, section) {
    // The question text, type and required-ness are all contained
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
    let titletext = titletextspan ? titletextspan.innerText : 'NO TITLE'

    // We want the number to be shown in flowchart boxes
    titletext = `${id}. ${titletext}`

    // Mermaid cannot handle parenthesis, semicolons, slashes or
    // angle brackets in text
    titletext = sanitiseForMermaid(titletext)

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
    const firstbranchcondition = branchgroup
      ? branchgroup.querySelector("div[role='radio'] span.text-format-content")
      : undefined

    // If a question has just one branch, the destination question text
    // can be found in an element with the class '.dropdown-placeholder-text'
    // This element will not exist if the destination question is the
    // next one in order
    const gotoElement = questionElement.querySelector(
      '.dropdown-placeholder-text'
    )

    // A 'ratings' type question also has a div with the 'radiogroup' role.
    // It doesn't have branch identification values under it. And finally,
    // if a multi-branch question is selected for editing, the branch
    // destinations are found in spans with the class
    // .dropdown-placeholder-text, which is the same marker as one-branch
    // questions. So, a question has multiple branches only if:
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

    const destination = gotoElement && !multiplebranches
      ? gotoElement.innerText
      : ''

    const question = {
      choices: [],
      id () {
        return id
      },
      titleText () {
        return titletext
      },
      title () {
        return `${id}. ${titletext}`
      },
      selectedForEditing () {
        return selectedforediting
      },
      multipleBranches () {
        return multiplebranches
      },
      destinationString () {
        return destination
      },
      section () {
        return section
      },
      isLastInSection () {
        return section
          ? id === section.lastQuestionId()
          : false
      }
    }

    // If the question has multiple branches
    // Each choice is contained in a div with the 'role' attribute
    // having value 'radio' under the question element
    if (multiplebranches) {
      const choiceElements =
        branchgroup.querySelectorAll("div[role='radio']")

      for (let i = 0, l = choiceElements.length; i < l; i++) {
        question.choices.push(
          parseChoice(choiceElements[i], question)
        )
      }
    }

    return question
  }

  function parseChoice (choiceElement, question) {
    // A span with the class '.text-format-content' will have the
    // choice title, or text.
    const choiceTitle =
      sanitiseForMermaid(
        choiceElement.querySelector(
          'span.text-format-content'
        ).innerText
      )

    let choiceGotoElement

    if (question.selectedForEditing()) {
      // If the choice's question is selected for editing, it is
      // contained in a <div> tag. In this case, each choice's
      // destination string is available in a span with the class
      // .dropdown-placeholder-text inside the choice element.
      choiceGotoElement =
        choiceElement.querySelector('span.dropdown-placeholder-text')
    } else {
      // If the choice's question is not selected for editing, it is
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
      choiceGotoElement = choiceElement.children[1]?.children[0]?.children[1]
    }

    const destinationstring = choiceGotoElement
      ? choiceGotoElement.innerText
      : ''

    return {
      title () {
        return choiceTitle
      },
      destinationString () {
        return destinationstring
      }
    }
  }

  // Main code begins

  // The background process expects a result in this shape.
  const finalresult = {
    status: 'Error',
    error: '',
    diagramTitle: '',
    diagramText: ''
  }

  // This will work only on the Branching Options page while
  // editing a form in Microsoft Forms.

  const headingSpan = document.querySelector("span[role='heading']")
  if (!headingSpan) {
    finalresult.error = 'Not on Branching Options screen (check 1)'
    return finalresult
  }

  if (headingSpan.innerText !== 'Branching options') {
    finalresult.error = 'Not on Branching Options screen (check 2)'
    return finalresult
  }

  const f = parseForm()
  const result = f.process()

  // Log result for debugging
  console.log('LOGGING FROM content.js:\n' + result)

  finalresult.status = 'Success'
  finalresult.diagramTitle = document.title
  finalresult.diagramText = result

  // Fetch and pass theme information
  const bodyStyles = getComputedStyle(document.body)
  finalresult.themePrimaryColor = bodyStyles.getPropertyValue('--palette-form-primary')

  return finalresult
})()
