# ms-forms-diagram

Show branch logic in a Microsoft Forms form as a diagram

![Forms Diagram Icon](images/icon48.png)

## Rationale

A form created using [Microsoft Forms](https://www.microsoft.com/en-gb/microsoft-365/online-surveys-polls-quizzes) can have complex branching logic. This Chrome (and Edge, I suppose) Extension allows that logic to be visualized using a [mermaid](https://mermaid-js.github.io/)-generated diagram.

## Installation

For now:

1. Clone this repository, or download as a zip file and unzip, to a local directory.
2. Put Chrome (or Edge) into [developer mode](https://developer.chrome.com/docs/extensions/mv3/faq/#faq-dev-01).
3. Load the local directory created in step 1 using the "Load Unpacked" button.

Then, pin the extension icon to your browser toolbar. It will "light up" when you start editing a form in Microsoft Forms. Click it while editing branches to see a diagram of your branching logic.

## Disclaimer

This extension was created in October 2022, and uses HTML scraping to parse the branching structure of forms. If the HTML structure or URLs of Microsoft Forms changes, it will probably stop working and need an update.

## Acknowledgements

This extension was created because of prodding by [Dr. Nitin Paranjape](https://efficiency365.com/) ([@drnitinp](https://github.com/drnitinp) on GitHub). He came up with the basic requirement, and tested the solution through multiple iterations and refinements.
