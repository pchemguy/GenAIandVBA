# AI-Driven VBA Development for Enhanced Word/Zotero/AI Manuscript Authoring Workflows

## Summary

This repository provides a practical solution and case study for improving academic authoring workflows involving Microsoft Word, Zotero, and AI assistants.
- **Problem Solved:** Addresses common issues like:
    - Loss of Zotero citation metadata (Word fields) when editing text with AI tools.
    - Loss of internal Word hyperlinks during external editing.
    - Lack of direct navigation links from in-text citations to bibliography entries in Word/PDF documents.
- **Solution:** Offers a suite of **VBA macros** for Word 2002/2003 (potentially compatible with newer versions, untested) to:
    - Automatically create hyperlinks between Zotero citations and bibliography items (`modBibliographyHyperlinker`).
    - Recover lost Zotero field information after external plain-text editing (`modZoteroFieldRecovery`).
    - Manage internal document bookmarks and hyperlinks using a plain-text markup that survives copy-pasting (`MarkupProcessor`).
- **Methodology:** Demonstrates **Prompt-Driven Development (PDD)** using Google Gemini Advanced 2.5 Pro.
    - Includes insights into using **meta-prompting** and In-Context Learning (ICL) to generate and refine the VBA code.
    - Investigates the trade-offs between prompt engineering effort and debugging time.
    - Suggests that investing in more detailed prompts upfront (guiding the AI's logic) can significantly streamline subsequent AI-driven development of the generated VBA code, especially for complex tasks.
- **Contents:** Includes:
    - Ready-to-use VBA module source code (`.bas` files).
    - The exact prompts used for AI generation.
    - Links to the AI conversations demonstrating the development and debugging process.

---

Academic manuscript authoring often involves integrating tools like Microsoft Word for writing, Zotero for reference management, and increasingly, Generative AI (GAI) for drafting and revision assistance. However, a common workflow challenge arises when transferring text between Word and plain-text AI interfaces: critical metadata, such as Zotero citation fields and internal document hyperlinks, is often lost. Furthermore, standard Zotero integration in Word does not automatically create navigable hyperlinks between in-text citations and their corresponding entries in the bibliography, hindering reader navigation, especially in PDF outputs.

To address these workflow limitations, this project explores the use of **Prompt-Driven Development (PDD)**, leveraging Large Language Models (LLMs), specifically Google Gemini Advanced 2.5 Pro, to create a suite of Visual Basic for Applications (VBA) macros for Word. The choice of VBA allows these tools to function directly within Word without external software installation. This document details the PDD methodology, focusing on **meta-prompting** techniques used to generate and refine the prompts that guided the AI in creating the VBA code. I share the developed VBA modules, the prompts used, AI conversations illustrating development workflows, and insights gained during the AI-assisted development process.

## The Workflow Problem: Data Loss and Navigation Gaps

1. **Metadata Loss:** Copying text sections from Word (with Zotero fields) to a plain-text AI chat for revision and pasting back results in the loss of embedded citation data (Word fields). Similarly, internal document hyperlinks created in Word are lost in this transfer.
2. **Lack of Citation-to-Bibliography Hyperlinks:** Zotero efficiently manages references but doesn't automatically create hyperlinks from citations (e.g., `[7]`) to the bibliography entry (e.g., `[7] Bibliography item...`). This makes navigating long reference lists cumbersome, especially in formats like PDF where field functionality is absent.
3. **Manual Recovery is Inefficient:** Manually recreating lost fields or adding hyperlinks is tedious and error-prone, negating the efficiency gains sought from using AI assistance.

## The Solution: AI-Generated VBA Macros

I developed three VBA macro modules as a proof-of-concept solution:

1. **`modBibliographyHyperlinker`:** Creates hyperlinks from in-text citations to bibliography entries.
2. **`modZoteroFieldRecovery`:** Facilitates the recovery of lost Zotero citation fields after text has been revised externally.
3. **`MarkupProcessor`:** Enables encoding/decoding internal bookmarks and hyperlinks using a plain-text markup, preserving them during external editing.

These modules were developed using a prompt-driven approach, where detailed instructions (prompts) were given to an LLM (Gemini Advanced 2.5 Pro) to generate the VBA code. Subsequently, any error messages and log outputs were also provided to LLM AI-driven debugging of the generated code.

## Methodology: Prompt-Driven Development via Meta-Prompting

Developing complex code like VBA macros benefits from detailed, well-structured prompts. Instead of starting from scratch or relying solely on basic prompts, I employed **meta-prompting** – using the LLM itself to help refine and generate the prompts that would ultimately be used to create the VBA code.

### Core Prompting Strategy

When creating an initial prompt draft (before meta-prompting), the focus is on describing the task as if explaining it to a human expert. This typically includes:

1. **Role/Expertise:** Defining the required skills (e.g., expert VBA programmer familiar with the Word Object Model, Regular Expressions, etc.).
2. **Problem Description:** Outlining the workflow issue and the desired functionality of the macro.
3. **Requirements/Constraints:** Specifying target Word version (Word 2002/2003 in this case), coding standards (adapting Python style guides), input/output formats, etc.
4. **Desired Solution Format:** Requesting a self-contained VBA module (`.bas` file content) operating on the `ActiveDocument`.

### Meta-Prompting for Prompt Refinement and Generation

Complexity in prompts requires careful management. I use Markdown formatting (via Obsidian.md) to structure prompts. To develop these structured prompts, I used meta-prompts.

- **Basic Meta-Prompt:** The simplest form is instructing the LLM to improve a given prompt, implicitly encouraging the LLM to analyze clarity, ambiguity, and completeness:

```
Help me improve the following prompt:

---

{PROMPT_TO_BE_ANALYZED}
```

- **Elaborated Meta-Prompts:** For more complex tasks, the meta-prompt itself can be more detailed, guiding the LLM on how to analyze and improve the target prompt (e.g., focusing on technical accuracy, positive language, logical flow). For more advanced meta-prompts, meta-meta-prompt (`Help me improve the following meta-prompt`) may help refine the prompt-improvement instructions themselves. An example is documented in [MarkupProcessorPromptMeta.md][] and the resulting chat [MarkupProcessorChat][].
- **Templated Meta-Prompts for Workflow Generation:** I also used meta-prompts to instruct the LLM to fill in specific parts of a prompt template, such as devising the detailed algorithm steps for a macro. The LLM analyzes the provided context and suggests the workflow logic, which can then be refined manually or interactively (see full example in conversation: [Meta-Prompting with Templated Prompt - VBA Citation Recovery Workflow Design][TemplatedMetaPrompting]).

```
Analyze the following prompt and consider if instructions are clear and unambiguous. Provide feedback/questions on any potential issues. Devise a workflow to be placed in place of "{To be suggested by AI}".

---

# Prompt: Recovery of Citation Fields
## Persona:
... (Expert VBA developer profile) ...
## Task:
Create a self-contained VBA6 macro module... for recovery of field-based in-text citations...
... (Detailed requirements) ...
### Macro Processing Steps:
{To be suggested by AI} 
```

- **Meta-Prompting with In-Context Learning (ICL):** When developing subsequent prompts for similar tasks, previously successful prompts can be provided as examples (ICL) within the meta-prompt. This helps the LLM understand the desired style, structure, and level of detail. For instance, when creating the prompt for `modBibliographyHyperlinker`, the prompt for `modZoteroFieldRecovery` was referenced. (See conversation: Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging, use this link https://gemini.google.com/share/57062c5d202c#:~:text=use%20previous%20prompts%20as%20a%20reference).

> **Important Note on Prompts:** The shared prompts ([MarkupProcessorPrompt.md][], [modBibliographyHyperlinkerPrompt.md][], [modZoteroFieldRecoveryPrompt.md][]) use **Markdown formatting**. When using these prompts with an AI, ensure you copy the **raw Markdown source**, not the rendered HTML, to preserve formatting crucial for the AI's interpretation. On GitHub, use the "Raw" button to view the source.

## Notes on AI-Assisted VBA Development

Generating VBA code via LLMs presented unique challenges compared to more modern languages like Python. My experience highlighted several points:

- **Varying Control Strategies:** No single prompting strategy was universally optimal.
    - For `modZoteroFieldRecovery` and `modBibliographyHyperlinker`, letting the AI generate the initial workflow ("Macro Processing Steps") based on a high-level description, followed by iterative debugging, worked reasonably well. Using the first prompt as ICL for the second was beneficial.
    - For `MarkupProcessor`, an initial attempt using AI-generated logic became difficult to debug when adding features. A second attempt, where I provided a detailed workflow logic within the prompt, resulted in cleaner, more maintainable code with fewer bugs. This suggests that for complex logic, investing time in outlining the algorithm upfront can lead to better AI-generated results.
- **Iterative Debugging:** AI-generated code often requires debugging. Providing clear error messages and context back to the LLM was effective for fixing issues. Often, LLM extended debugging functionality based on provided error logs without further instructions. Other times, additional specific requests facilitated debugging process.
- **Style Guidance:** Referencing Python style guides and asking the LLM to adapt them for VBA might have improved code quality, although the specific impact wasn't formally evaluated.
- **LLM Capabilities:** Frontier models like Gemini Advanced 2.5 Pro seem capable of handling complex VBA generation tasks, especially when guided by well-structured prompts and iterative refinement.

## VBA Modules: Installation and Usage

The VBA code is provided as standard module files (`.bas`).

**Prerequisites:**

1. **Microsoft Word:** Developed on Word 2002/2003. Should theoretically work on later versions, but **this has not been tested.**
2. **VBA References:** Some modules require specific references to be enabled in the VBA editor (Tools -> References...):
    `modBibliographyHyperlinker`: Requires "Microsoft VBScript Regular Expressions 5.5" and "Microsoft Scripting Runtime". Check the module's comments for specific requirements.

**Installation:**

1. Open your Word document or template (e.g., `Normal.dotm` if you want the macros available globally).
2. Open the VBA Editor (Alt + F11).
3. In the VBA Editor, go to File -> Import File...
4. Navigate to and select the `.bas` file you want to import (e.g., `modBibliographyHyperlinker.bas`).
5. Repeat for each module you want to use. The imported modules will appear in the "Modules" folder of your VBA project.

**Running the Macros:**

Currently, there is no custom GUI. To run a macro:

1. Open the VBA Editor (Alt + F11).
2. Find the desired module (e.g., `modBibliographyHyperlinker`) in the Project Explorer.
3. Double-click the module to open its code.
4. Click anywhere inside the main public subroutine you want to run (the "Entry Point" listed below, e.g., `CreateBibliographyHyperlinks`).
5. Press F5 or click the Run button on the toolbar. Alternatively, you can run macros via Word's Macro dialog (Alt + F8), selecting the macro name (e.g., `CreateBibliographyHyperlinks`) and clicking "Run".

**Important:** Always **save your document** before running macros, especially during testing. These macros modify the `ActiveDocument`.

### Module Details

#### 1. Zotero Bibliographic Hyperlinks (`modBibliographyHyperlinker`)

- **Entry Point:** `CreateBibliographyHyperlinks`
- **Purpose:** Creates hyperlinks from numeric in-text citations (e.g., `[23]`, `[25,26]`, `[17-19]`) to the corresponding bibliography items.
- **Requirements:**
    - Bibliography must be generated by Zotero before running.
    - Bibliography items must start with `[#]` (e.g., `[7]`).
    - In-text citations must use Zotero fields.
    - VBA References: Microsoft VBScript Regular Expressions 5.5, Microsoft Scripting Runtime.
- **Action:** Scans the bibliography, creates bookmarks (`BIB_#`) for each item. Scans the document, finds citations, and links the numbers to the corresponding `BIB_#` bookmarks. Deletes pre-existing `BIB_#` bookmarks and associated hyperlinks before running.

#### 2. Zotero Field Recovery (`modZoteroFieldRecovery`)

- **Entry Point:** `RecoverZoteroFields`
- **Purpose:** Restores Zotero field codes to plain-text citations (e.g., `[23]`) after they were lost during external editing.
- **Requirement:** The original text section containing the intact Zotero fields must still be present somewhere in the document (e.g., temporarily pasted below the revised section) for the macro to find and copy the field information.
- **Action:** Matches plain-text citations in the revised text with corresponding field-based citations in the original text (based on the displayed number) and copies the field information over.

#### 3. Internal Navigation Markup (`MarkupProcessor`)

- **Entry Point:** `AutoMarkup`
- **Purpose:** Allows defining internal bookmarks and hyperlinks using plain-text markup that survives copy-pasting to external editors.
- **Markup Format:**
    - Bookmarks: `{{Displayed Text}}{{BMK: BookmarkName}}`
    - Hyperlinks: `{{Displayed Text}}{{LNK: BookmarkName}}`
- **Action:** Finds the markup. Sets the metadata part (`{{BMK:...}}` or `{{LNK:...}}` and surrounding braces `{{}}`) as hidden text. Creates a Word bookmark named `BookmarkName` around the `Displayed Text` for `BMK` tags, or creates a hyperlink from `Displayed Text` to the bookmark named `BookmarkName` for `LNK` tags.

## Summary of Shared Artifacts

### modBibliographyHyperlinker

|                          |                                                                                                                                                                          |
| :----------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Module                   | [modBibliographyHyperlinker.bas][]                                                                                                                                       |
| Prompt                   | [modBibliographyHyperlinkerPrompt.md][]                                                                                                                                  |
| Development Conversation | [Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging][ICLMetaPromptingDebugging] ([Private URL](https://gemini.google.com/app/59e84d4879cebb1c)) |
| Language                 | VBA Version 6                                                                                                                                                            |
| Host Platform            | MS Word 2002/2003, newer versions (untested)                                                                                                                             |
| Entry Point              | CreateBibliographyHyperlinks                                                                                                                                             |
| Development Mode         | Prompt-driven (via Google Gemini Advanced 2.5 Pro)                                                                                                                       |

### modZoteroFieldRecovery

|                          |                                                                                                                                                                          |
| :----------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Module                   | [modZoteroFieldRecovery.bas][]                                                                                                                                           |
| Prompt                   | [modZoteroFieldRecoveryPrompt.md][]                                                                                                                                      |
| Development Conversation | [Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging][ICLMetaPromptingDebugging] ([Private URL](https://gemini.google.com/app/59e84d4879cebb1c)) |
| Language                 | VBA Version 6                                                                                                                                                            |
| Host Platform            | MS Word 2002/2003, newer versions (untested)                                                                                                                             |
| Entry Point              | RecoverZoteroFields                                                                                                                                                      |
| Development Mode         | Prompt-driven (via Google Gemini Advanced 2.5 Pro)                                                                                                                       |

### MarkupProcessor

|                          |                                                                                                                                        |
| :----------------------- | :------------------------------------------------------------------------------------------------------------------------------------- |
| Module                   | [MarkupProcessor.bas][]                                                                                                                |
| Prompt                   | [MarkupProcessorPrompt.md][]                                                                                                           |
| Meta-Prompt              | [MarkupProcessorPromptMeta.md][]                                                                                                       |
| Development Conversation | [VBA-Based Navigation Markup Workflow in MS Word][MarkupProcessorChat] ([Private URL](https://gemini.google.com/app/1571d7a44e0e6355)) |
| Language                 | VBA Version 6                                                                                                                          |
| Host Platform            | MS Word 2002/2003, newer versions (untested)                                                                                           |
| Entry Point              | AutoMarkup                                                                                                                             |
| Development Mode         | Prompt-driven (via Google Gemini Advanced 2.5 Pro)                                                                                     |

### Meta-Prompting Demos

- [Meta-Prompting with Templated Prompt - VBA Citation Recovery Workflow Design][TemplatedMetaPrompting] ([Private URL](https://gemini.google.com/app/efb8c56cfe127897))
- [Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging][ICLMetaPromptingDebugging] ([Private URL](https://gemini.google.com/app/59e84d4879cebb1c))


---

[Zotero]: https://zotero.org
[MarkupProcessorPrompt.md]: MarkupProcessorPrompt.md
[MarkupProcessorPromptMeta.md]: MarkupProcessorPromptMeta.md
[MarkupProcessor.bas]: MarkupProcessor.bas
[modBibliographyHyperlinkerPrompt.md]: modBibliographyHyperlinkerPrompt.md
[modBibliographyHyperlinker.bas]: modBibliographyHyperlinker.bas
[modZoteroFieldRecoveryPrompt.md]: modZoteroFieldRecoveryPrompt.md
[modZoteroFieldRecovery.bas]: modZoteroFieldRecovery.bas
[MarkupProcessorChat]: https://g.co/gemini/share/50e01f6b36be
[TemplatedMetaPrompting]: https://g.co/gemini/share/3239df438507
[ICLMetaPromptingDebugging]: https://gemini.google.com/share/57062c5d202c#:~:text=use%20previous%20prompts%20as%20a%20reference
