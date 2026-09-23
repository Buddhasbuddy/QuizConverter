# Brightspace Quiz Converter

This is a dependency-free browser converter for turning a structured `.txt` or `.docx` quiz into a Brightspace content package (`.zip`). It is intentionally client-side: files are read in the browser and are not sent to a server.

## Run it

Open `index.html` in a current browser. If the browser blocks local file features, serve this folder with any static file server, for example:

```text
npx serve .
```

No build step or package installation is required.

## Recommended quiz format

```text
Title: Workplace Safety Quiz
Description: Select the best answer.

1. What does PPE stand for?
A. Personal protective equipment
B. Public policy evaluation
C. Project planning estimate
Answer: A

2. Safety training is required before using the equipment.
Type: True/False
Answer: True

3. Name the colour used for a fire exit sign.
Type: Fill in the Blank
Answer: green
```

The parser also recognizes multiple choice, multi-select, matching (`Term -> Definition`), ordering, and long answer questions. Always review the detected questions and test the generated package in a sandbox Brightspace course before using it in a live course. Brightspace package details can vary by tenant and version, so this project treats the supplied export as its compatibility reference.
