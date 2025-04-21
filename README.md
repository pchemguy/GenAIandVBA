# Prompt-Driven Development of VBA Macros for Management of Navigation and Field Information in Word/Zotero/AI Manuscript Authoring Workflow

One basic approach to using generative AI assistance for manuscript development involves periodic transfer of sections of text to an AI conversation for revisions and copying revised text back into manuscript document file. If the text is developed using plain text markup, such as Markdown, metadata, such as cited references and hyperlinks are typically persevered (assuming text is not changed drastically). On the other hand, if the text is developed using MS Word, and the references are managed with a reference manager, such as [Zotero][], reference information is typically encoded via Word fields, and this information is lost when copying a section of text to a plain-text AI chat. Further, hyperlink information also used for local navigation is also lost for the same reason. Additionally, there was another issue that I wanted to address. While Zotero manages references and creates citations automatically, it does not create hyperlinks leading from citations to corresponding bibliographic items. When a paper with Zotero-based references is converted to the PDF format, all references become plain text without convenient hyperlinks that make it easy to find the cited item in the bibliography list, from where the sources can often than be accessed via cited DOI-based or publisher-based URLs.

In an attempt to address these issues, I decided to developed a set of VBA macro modules as a proof-of-concept solution. The choice of VBA was dictated by its availability directly within MS Word without the need to install any additional software. And since I am working on text focused on generative AI (GAI) prompting and also want to use GAI for revising the text, it was a natural decision to use GAI for macro development.

I am sharing generated macro modules, prompts and some related prompt engineering information, and one Google Gemini Advanced 2.5 Pro conversation used for revising one of the prompts and subsequent generation of the associated modules. The two other modules were generated in a separate conversation; however, that conversation has become somewhat messy due to my early experiments, and I do not see much value in sharing that conversation as well.

> [!Warning]
> 
> **IMPORTANT:** Markdown formatting is an integral constituent of the shared prompts. Do not copy html rendered prompt texts to chat bots. Instead, use the raw plain-text Markdown formatted text, which is accessible online via the "Raw" button on the right side of the file toolbar above the text window.

The VBA code is developed as standard macro modules and is shared as exported plain-text VBA source code files which can be imported into the target document project for further use. All shared code is designed to operate on the `ActiveDocument` object, which means they can be imported into a Word document, which can then be saved as a Word add-in, loaded into Word and used like any other add-in provided tools (though I have not tested this approach yet). The code has only been tested in Word 2002/2003, though it should in theory work as-is with later versions as well. Presently, there is no GUI available: to activate associated functionality, the main entry (implemented as a public procedure) of the corresponding module must be executed either from the VBA editor or via the Word's standard macro running mechanism.

## Prompt-Driven Workflow: Prompt -> Meta-Prompt -> Meta-Meta-Prompt

While there are a variety of resources featuring prompt libraries, usually, I do not use them. When I do not have a suitable prompt to start with and need to start from scratch, I think of how I would describe the task to find an expert. Such a description might include
1. List of expert qualifications and skills I consider important for solving the problem, e.g., what I could include in a job posting (these characteristics can be projected on to model using the role prompting techniques).
2. Description of the problem (e.g., it might include the description of a particular workflow I need to enable or improve; present specific limitations, such as unreliable, limited, or missing implementation for certain operations leading to inefficient or broken pipelines).
3. Specification of any requirements and limitations (This section may include, for example, previously developed snippets used for similar problems or texts adapted from common guidelines. For example, if I need a Python script, I might incorporate, adapt or reference one of the common Python Style Guides and some more general programming practices )
4. Information on solution I need, how I want to use it (perhaps sample workflows).

I mostly focus on complex tasks that benefit from correspondingly complex prompt. Complexity, however, needs to be properly managed. Two common ways to structure complex prompts are Markdown-based and xml-based formatting. Personally, I use Markdown with Obsidian.md editor.

After having a preliminary draft, I might test it to see the initial model response, but more often I would use an abstract meta-prompt to have the model to improve the prompt itself first. One of the simplest In the simplest meta-prompts is `Help me improve the following prompt` used like this:  

**Basic Meta-Prompt**

```
Help me improve the following prompt

---

{PROMPT TO BE ANALYZED}
```
 
This meta-prompt is universal and may work reasonably well with frontier reasoning models (probably, non-reasoning models as well) and moderately complex tasks/prompts. For more complex prompts, it might be beneficial introducing an intermediate abstraction layer by elaborating on the meta-prompt.

While the prompt being developed focuses on the final problem or task, the task/objective of the meta-prompt is the prompting process. In other words, the goal of a prompt is to generate a solution to an actual problem; the goal of a meta-prompt is to generated a prompt that, in turn, will be used to generate a solution to an actual problem. The meta-prompt should, in general, focus on linguistic characteristics of the target prompt and its efficiency as a tool. For example, meta-prompt may instruct the LLM to analyze the prompt as a piece of technical writing and list detailed criteria commonly used for revising technical texts or emphasize some specific points, such as clarity and positive actionable language. Different task may also benefit from somewhat different emphasis. Meta-prompt may also request the model to analyze the prompt, provide constructive feedback, including suggestions for improvements. When performing prompt analysis, LLM will have access to all the same background data that will be used for executing the prompt. For this reason, prompt analysis feedback may include not only linguistic suggestions, but also semantic suggestions. And sometimes meta-prompt abstract nature may be broken to some extent depending on specific task. For these reasons, even meta-prompt may benefit from certain task specific adaptation. In such a case, the baseline meta-prompt above can be used as meta-meta-prompt (*Help me improve the following **META**-prompt*) to improve meta-prompt. Examples of a meta-prompt draft and its improved versions are included in [MarkupProcessorPromptMeta.md][] (the final version of the associated generated prompt is in [MarkupProcessorPrompt.md][]) and can be also seen in the shared conversations ([VBA-Based Navigation Markup Workflow in MS Word][MarkupProcessorChat]).

## Meta-Prompting with Templated Prompts and In-Context Learning (ICL)

An important variation of meta-prompting technique where abstraction is intentionally broken, is the use of meta-prompts not just for prompt revision, but for extending and/or generating prompts. This approach is useful for developing detailed prompts for complex tasks. The idea is that concise generic prompts are often suboptimal for complex tasks, yielding generic and varying solutions. Developing detailed specific prompts from scratch, on the other hand, is a time consuming process. One possible compromise would be to start with a relatively simple prompt or perhaps using for initial prompt sections of previously developed prompts for similar tasks. Such a prompt would include a concise description of the desired task, and the meta-prompt would instruct the model to generate specific detailed steps.

Here is an example of a meta-prompt used with a templated prompt. The ultimate task of the prompt being generated / extended is generation of a VBA macro based on textual description of the desired functionality. The job of the meta-prompting stage is generation of detailed structured description of macro algorithm that should implement desired functionality (see the template placeholder `{To be suggested by AI}` at the end of the prompt). After initial algorithm description is generated, macro workflow may be manually edited or interactively adjusted via successive requests. Once the the algorithm details are refined, the final prompt may be generated and executed in a subsequent request.

**Templated Meta-Prompt**

```
Analyze the following prompt and consider if instructions are clear and unambiguous. Provide feedback/questions
on any potential issues. Devise a workflow to be placed in place of "{To be suggested by AI}".

---

# Prompt: Recovery of Citation Fields

## Persona:

============================== CONTENT REDUCED FOR BREVITY. ==============================

## Task:

Create a self-contained VBA6 macro module (`.bas` file content) for Microsoft Word (2002/XP) for recovery of field-based
in-text citations after revision of edited text. The text will be pasted next to original text containing all references. 

============================== CONTENT REDUCED FOR BREVITY. ==============================

### Macro Processing Steps:

{To be suggested by AI}
```

Note: the prompt example above is shown with reduced content for brevity. The full content, as we as an example of an iterative interactive refinement process is available from the shared conversation ([Meta-Prompting with Templated Prompt - VBA Citation Recovery Workflow Design][TemplatedMetaPrompting]). 

Another sample conversation ([Meta-Prompting with ICL and Refinement -  BMK - Generated VBA Code Debugging][ICLMetaPromptingDebugging]) used to develop two macro VBA modules starts with a basic meta-prompt

```
Analyze the following prompt and consider if instructions are clear and
unambiguous.Provide feedback/questions on any potential issues.

{PROMPT TO BE ANALYZED}
```

and demonstrates interactive improvements of the prompt by providing answers to LLM's feedback

```
Revise the prompt with the following answers, analyze it again and consider if there are still
some questions (of if some answers are unclear). Provide additional feedback, if necessary, or
generated a revised prompt with clear and well organized structure and language.

# Answers

{ANSWERS TO LLM FEEDBACK}
```

Then the generated prompt is executed yielding the first draft of the module and the subsequent prompts are used for iterative debugging. After debugging is complete, a prompt for a second macro VBA module is generated via meta-prompting with ICL, using the first developed prompt as a reference (search the conversation text for `use previous prompts as a reference`):

**modBibliographyHyperlinker Meta-Prompt (ICL-Based Generation)**

```
Help me create a prompt for generating a VBA6 / MS Word macro (use previous prompts as a reference). The macro would need to
1. Delete all bookmarks with `AUTO_` prefix.
2. Search for patterns `{{Supporting Information}}{{BMK: SI}}`.
3. Verify that
    1. The text inside the first pair {{}} is visible AND the rest is hidden.
    2. The content in the second pair starts with "BMK:".
    3. The trimmed part after "BMK:" contains only alphanumeric characters and underscores.
    4. The trimmed part after "BMK:" starts with a letter.
    5. The trimmed part after "BMK:" is no longer than 35 chars.
4. Create bookmark around the visible part.

Make sure to ask me for clarification, if necessary, before starting the prompt generation process.
```

## Notes on AI-Assisted VBA Development

I am under the impression that for a number of reasons, VBA code generation is somewhat more challenging for AI than, for example, Python. In fact, the three presented modules, employ somewhat different (meta-)prompting strategies. With the first module [modZoteroFieldRecovery.bas][], I described the desired functionality to AI and asked it to generate the workflow (the "Macro Processing Steps" in the [prompt][modZoteroFieldRecoveryPrompt.md]), following by module generation and interactive iterative debugging. With the second module, [modBibliographyHyperlinker.bas][] ([prompt][modBibliographyHyperlinkerPrompt.md]), I followed a similar strategy, except I could also use the first generated prompt for ICL during the meta-prompting stage. With the last module, [MarkupProcessor.bas][], the same strategy initially seemed to work, but when I tried to add a feature the process got messed up, so I had to start from scratch. The second time I amended my strategy: instead of letting the AI to develop processing logic, I drafted a fairly detailed workflow myself (see "Bookmarks and Internal Hyperlinks" in [prompt][MarkupProcessorPrompt.md]). I spent somewhat more time on the processing logic, though I mostly sorted it out on my way home, I think I ended up with a significantly better organized code. I think there was just one bug, which was actually my fault, and Gemini readily sorted it out I after I provided the error information. Another trick I used with the hope to get better code quality is that I referenced Python style guides in the prompts and asked Gemini to adapt it to VBA, where possible (I have not tried to evaluate, whether this specific trick made a significant difference though).

## VBA Modules
### Zotero Bibliographic Hyperlinks - modBibliographyHyperlinker

[modBibliographyHyperlinker.bas][] module (entry point: `CreateBibliographyHyperlinks`) creates local hyperlinks that make it possible to navigate from in-text citation to associated bibliographic record in the bibliography list. Bibliography must be created before executing this module. Requirements:
- In-text citations: `[23]`, `[25,26,30]`, `[17-19]`, `[17-19, 23, 25]`.
- Bibliography list items: Must begin with `[#]` (e.g., `[7]`, `[23]`).
- Field-based citations.
Required VBA project references:
- Microsoft VBScript Regular Expressions 5.5
- Microsoft Scripting Runtime

The code looks for the bibliography field and creates for each item a bookmark using template `BIB_{#}` (e.g., `BIB_7`, `BIB_23`). The bookmarks cover the bracketed number (e.g., `[7]`, `[23]`). The code looks for all in-text citations and creates a hyperlink for each number leading to corresponding bibliographic item. For ranges, the entire range points to the first citation in the range. Before creating bookmarks and links, the code deletes all bookmarks that match the name template and all hyperlinks targeting such bookmarks.

### Zotero Field Recovery - modZoteroFieldRecovery

[modZoteroFieldRecovery.bas][] module (entry point: `RecoverZoteroFields`) facilitates recovery of lost in-text citation fields (for bracketed in-text citation only, e.g., `[23]`), after the text has been copied outside Word as plain text with all in-text citation losing associated field metadata. It is important that the old text section is not deleted until this module is executed. This module looks for plain-text in-text citations and it assumes that the same citations with proper field metadata are still in the document. When the module is executed, the old citation are matched based on the displayed text and replace plain-text citations.

### Internal Navigation - Bookmarks / Hyperlinks - MarkupProcessor

[MarkupProcessor.bas][] module (entry point: `AutoMarkup`) enables encoding local hyperlinks and bookmarks via plain text, so that the associated metadata can be recovered automatically if the text section is copied outside Word, for example, for proofreading, and copied back. The module assumes that such bookmarks and hyperlinks follow templates: `{{Displayed Text}}{{BMK: #BookmarkName}}` and `{{Displayed Text}}{{LNK: #BookmarkName}}`. The code set everything, except for the `Displayed Text` as `Hidden` text, which is not exported to PDF.

## Summary of Shared Artifacts

### modBibliographyHyperlinker

|                       |                                                                             |
| --------------------- | --------------------------------------------------------------------------- |
| Module                | [modBibliographyHyperlinker.bas][]                                          |
| Language              | VBA Version 6                                                               |
| Primary host platform | MS Word 2002/2003                                                           |
| Other host platforms  | MS Word, newer versions (not tested)                                        |
| Entry Point           | `CreateBibliographyHyperlinks`                                              |
| Development mode      | Prompt-driven                                                               |
| Prompt source         | [modBibliographyHyperlinkerPrompt.md][]                                     |
| AI conversation title | Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging |
| Generative AI model   | Google Gemini Advanced 2.5 Pro                                              |
| Public URL            | https://g.co/gemini/share/57062c5d202c                                      |
| Private URL           | https://gemini.google.com/app/59e84d4879cebb1c                              |

### modZoteroFieldRecovery

|                       |                                                                             |
| --------------------- | --------------------------------------------------------------------------- |
| Module                | [modZoteroFieldRecovery.bas][]                                              |
| Language              | VBA Version 6                                                               |
| Primary host platform | MS Word 2002/2003                                                           |
| Other host platforms  | MS Word, newer versions (not tested)                                        |
| Entry Point           | `RecoverZoteroFields`                                                       |
| Development mode      | Prompt-driven                                                               |
| Prompt source         | [modZoteroFieldRecoveryPrompt.md][]                                         |
| AI conversation title | Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging |
| Generative AI model   | Google Gemini Advanced 2.5 Pro                                              |
| Public URL            | https://g.co/gemini/share/57062c5d202c                                      |
| Private URL           | https://gemini.google.com/app/59e84d4879cebb1c                              |

### MarkupProcessor

|                       |                                                 |
| --------------------- | ----------------------------------------------- |
| Module                | [MarkupProcessor.bas][]                         |
| Language              | VBA Version 6                                   |
| Primary host platform | MS Word 2002/2003                               |
| Other host platforms  | MS Word, newer versions (not tested)            |
| Entry Point           | `AutoMarkup`                                    |
| Development mode      | Prompt-driven                                   |
| Prompt source         | [MarkupProcessorPrompt.md][]                    |
| Meta-prompt source    | [MarkupProcessorPromptMeta.md][]                |
| AI conversation title | VBA-Based Navigation Markup Workflow in MS Word |
| Generative AI model   | Google Gemini Advanced 2.5 Pro                  |
| Public URL            | https://g.co/gemini/share/50e01f6b36be          |
| Private URL           | https://gemini.google.com/app/1571d7a44e0e6355  |

### Meta-Prompting and VBA Macro Module Development Demos

|                       |                                                                              |
| --------------------- | ---------------------------------------------------------------------------- |
| AI conversation title | Meta-Prompting with Templated Prompt - VBA Citation Recovery Workflow Design |
| Generative AI model   | Google Gemini Advanced 2.5 Pro                                               |
| Public URL            | https://g.co/gemini/share/3239df438507                                       |
| Private URL           | https://gemini.google.com/app/efb8c56cfe127897                               |

|                       |                                                                             |
| --------------------- | --------------------------------------------------------------------------- |
| AI conversation title | Meta-Prompting with ICL and Refinement - BMK - Generated VBA Code Debugging |
| Generative AI model   | Google Gemini Advanced 2.5 Pro                                              |
| Public URL            | https://g.co/gemini/share/57062c5d202c                                      |
| Private URL           | https://gemini.google.com/app/59e84d4879cebb1c                              |















 
 
 <!-- References -->

[Zotero]: https://zotero.org
[MarkupProcessorPrompt.md]: MarkupProcessorPrompt.md
[MarkupProcessorPromptMeta.md]: MarkupProcessorPromptMeta.md
[MarkupProcessor.bas]: MarkupProcessor.bas
[modBibliographyHyperlinkerPrompt]: modBibliographyHyperlinkerPrompt.md
[modBibliographyHyperlinker]: modBibliographyHyperlinker.bas
[modZoteroFieldRecoveryPrompt.md]: modZoteroFieldRecoveryPrompt.md
[modZoteroFieldRecovery]: modZoteroFieldRecovery.bas
[MarkupProcessorChat]: https://g.co/gemini/share/50e01f6b36be
[TemplatedMetaPrompting]: https://g.co/gemini/share/3239df438507
[ICLMetaPromptingDebugging]: https://g.co/gemini/share/57062c5d202c
