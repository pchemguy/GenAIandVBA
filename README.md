# Prompt-Driven Development of VBA Macros for Management of Navigation and Field Information in Word/Zotero/AI Manuscript Authoring Workflow

One basic approach to using generative AI assistance for manuscript development involves periodic transfer of sections of text to an AI conversation for revisions and copying revised text back into manuscript document file. If the text is developed using plain text markup, such as Markdown, metadata, such as cited references and hyperlinks are typically persevered (assuming text is not changed drastically). On the other hand, if the text is developed using MS Word, and the references are managed with a reference manager, such as [Zotero][], reference information is typically encoded via Word fields, and this information is lost when copying a section of text to a plain-text AI chat. Further, hyperlink information also used for local navigation is also lost for the same reason. Additionally, there was another issue that I wanted to address. While Zotero manages references and creates citations automatically, it does not create hyperlinks leading from citations to corresponding bibliographic items. When a paper with Zotero-based references is converted to the PDF format, all references become plain text without convenient hyperlinks that make it easy to find the cited item in the bibliography list, from where the sources can often than be accessed via cited DOI-based or publisher-based URLs. In an attempt to address these issues, I decided to developed a set of VBA macro modules as a proof-of-concept solution. The choice of VBA was dictated by its availability directly within MS Word without the need to install any additional software. And since I am working on text focused on generative AI (GenAI) prompting and also want to use GenAI for revising the text, it was a natural decision to use GenAI for macro development.

I am sharing generated macro modules, prompts and some related prompt engineering information, and one Google Gemini Advanced Pro 2.5 conversation used for revising one of the prompts and subsequent generation of the associated modules. The two other modules were generated in a separate conversation; however, that conversation has become somewhat messy due to my early experiments, and I do not see much value in sharing that conversation as well.

> [!Warning]
> 
> **IMPORTANT:** Markdown formatting is an integral constituent of the shared prompts. Do not copy html rendered prompt texts to chat bots. Instead, use the raw plain-text Markdown formatted text, which is accessible online via the "Raw" button on the right side of the file toolbar above the text window.

The VBA code is developed as standard macro modules and is shared as exported plain-text VBA source code files which can be imported into the target document project for further use. All shared code is designed to operate on the `ActiveDocument` object, which means they can be imported into a Word document, which can then be saved as a Word add-in, loaded into Word and used like any other add-in provided tools (though I have not tested this approach yet). The code has only been tested in Word 2002/2003, though it should in theory work as-is with later versions as well. Presently, there is no GUI available: to activate associated functionality, the main entry (implemented as a public procedure) of the corresponding module must be executed either from the VBA editor or via the Word's standard macro running mechanism.

## 1. Development Workflow
## 2. Creation of Bibliographic Hyperlinks

The first generated module 
modBibliographyHyperlinker.bas






https://g.co/gemini/share/bcff2aa6b15f  
VBA for Bookmarks and Hyperlinks Information Management in MS Word  




### MarkupProcessor


|                       |                                                 |
| --------------------- | ----------------------------------------------- |
| Module                | MarkupProcessor.bas                             |
| Language              | VBA Version 6                                   |
| Primary host platform | MS Word 2002/2003                               |
| Other host platforms  | MS Word, newer versions (not tested)            |
| Development mode      | Prompt-driven                                   |
| Generative AI model   | Google Gemini Advanced Pro 2.5                  |
| Prompt source         | MarkupProcessorPrompt.md                        |
| Meta-prompt source    | MarkupProcessorPromptMeta.md                    |
| AI conversation title | VBA-Based Navigation Markup Workflow in MS Word |
| Public URL            | https://g.co/gemini/share/50e01f6b36be          |
| Private URL           | https://gemini.google.com/app/1571d7a44e0e6355  |




****

```
Meta-Meta-Prompt

Help me improve the following meta-prompt

---

{META-PROMPT TO BE ANALYZED}

```

 
 
 <!-- References -->

[Zotero]: https://zotero.org
[MarkupProcessorPrompt]: MarkupProcessorPrompt.md
[MarkupProcessor]: MarkupProcessor.bas
[modBibliographyHyperlinkerPrompt]: modBibliographyHyperlinkerPrompt.md
[modBibliographyHyperlinker]: modBibliographyHyperlinker.bas
[modZoteroFieldRecovery]: modZoteroFieldRecovery.bas
