# Brightspace Quiz Converter

This is a dependency-free browser converter for turning a Respondus Standard Format `.txt` or `.docx` quiz into a Brightspace content package (`.zip`). It also accepts the guide's tab/comma-delimited `.csv` format. It is intentionally client-side: files are read in the browser and are not sent to a server.

## Run it

Open `index.html` in a current browser. If the browser blocks local file features, serve this folder with any static file server, for example:

```text
npx serve .
```

No build step or package installation is required.

## Supported Respondus Standard Format

The converter follows the [University of Wyoming Respondus formatting guide](https://www.uwyo.edu/lms/wyocourses/improved/respondustabsfolder/exam_formatting.pdf). It supports question numbers using a period or parenthesis, lettered choices, starred correct answers, trailing answer lists, question types, titles, points, feedback, and matching questions.

```text
Type: MR
Title: Speed of Light
Points: 2
3) Which people helped determine the speed of light?
a. Albert Einstein
*b. Albert Michelson
*c. Edward Morley
d. Thomas Edison

Type: E
Title: Relativity
4) Explain the Michelson-Morley experiment.

Type: F
5) Who is known as the father of television?
a. Vladimir Zworykin

Type: MT
6) Match each person to the discovery.
a. Michelson-Morley = Speed of light
b. Einstein = Theory of Relativity

Answers:
3. B, C
4. A suggested essay answer
5. Vladimir Zworykin
```

For text-file images, select the quiz file and its image files together. Reference images with `[ img: "filename.jpg" "Alternative text" ]`. Embedded DOCX images are included automatically.

## Simple format

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
