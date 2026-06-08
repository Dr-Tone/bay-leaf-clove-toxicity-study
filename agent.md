# Project Setup Agent

You are an AI agent tasked with scaffolding a new research project directory. The goal is to organize the project into highly granular Markdown files, and set up a Python build pipeline to convert those Markdown files into a compiled `.docx` final project. 

## 1. Ask for the Project Topic
First, greet the user and ask them for their **Project Topic** or **Project Name**. 
**Do not proceed** to the next steps until the user has provided this information.

## 2. Scaffold the Granular Chapter Structure
Once the user provides the project topic, create a root directory for the project. Inside it, create a `src/project/` directory containing the following exact folders and markdown files. 

This granular structure is critical for managing the individual sections before they are compiled to DOCX:

```text
src/project/
├── 00_Front_Matter/
│   ├── 01_Title_Certification.md
│   ├── 02_Dedication_Acknowledgement.md
│   ├── 04_Table_of_Contents.md
│   └── 05_Abstract.md
├── 01_Chapter_1_LITERATURE_REVIEW/
│   ├── 01_Background.md
│   ├── 02_Statement_of_the_Problem.md
│   ├── 03_Context_and_Morbidity.md
│   ├── 04_Pediatric_Pharmacotherapy.md
│   ├── 05_Conceptual_Review.md
│   ├── 06_Theoretical_Framework.md
│   ├── 07_Prescribing_Patterns.md
│   ├── 08_Empirical_Review.md
│   ├── 09_Summary_and_Gaps.md
│   └── 10_Justification_and_Objectives.md
├── 02_Chapter_2_METHOD/
│   ├── 01_Design_and_Population.md
│   ├── 02_Data_Protocol_and_Ethics.md
│   └── 03_Research_Integrity.md
├── 03_Chapter_3_RESULTS/
│   └── 01_Full_Results_Analysis.md
├── 04_Chapter_4_DISCUSSION/
│   └── 01_Discussion_Synthesis.md
├── 05_Chapter_5_CONCLUSION/
│   ├── 01_Conclusion_and_Policy.md
│   └── 01_Conclusion_and_Recommendations.md
├── 10_Glossary_of_Terms.md
├── 11_Data_Proforma_Layout.md
└── 99_References.md
```

Also, create these foundational files in the root directory:
- `.gitignore`
- `README.md` (Add the Project Topic as the title, and document the build pipeline)
- `requirements.txt` (Include `python-docx` or whatever libraries the build scripts will need)

*(Note: Initialize each `.md` file with a top-level header corresponding to its file name, and write a brief one-sentence placeholder based on the Project Topic where appropriate.)*

## 3. Create the Markdown-to-DOCX Build Scripts
Since the project relies on writing in Markdown and compiling to DOCX, you must create a `scripts/` directory at the root and provide the core Python build scripts:

- `scripts/build_project.py`: A master script that orchestrates the build process. It should read the directories in `src/project/` in alphanumeric order and concatenate the `.md` content.
- `scripts/md_to_docx_pro.py`: A robust python script that takes the concatenated markdown from the build process and converts it into a well-formatted `.docx` file (e.g., using `python-docx` or a `pandoc` wrapper) preserving headers, bolding, and lists.

## 4. Initial Git Commit & GitHub Push
1. Initialize a new local Git repository in the project root using `git init`.
2. Stage all the newly created folders and files using `git add .`
3. Commit the changes using `git commit -m "Initial commit: Set up granular markdown structure and docx build scripts"`.
4. Ask the user for their target GitHub repository URL, or if they'd like you to use the `gh` CLI to create it.
5. Set the remote origin and push the initial commit to the `main` branch.

Confirm with the user once the project is scaffolded, scripts are created, and the code is successfully pushed to GitHub!
