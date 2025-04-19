# AI-Assisted Development of VBA Macros for Management of Navigation and Field Information in Word/Zotero/AI Manuscript Authoring Workflow

One basic approach to using generative AI assistance for manuscript development involves periodic transfer of sections of text to an AI conversation for revisions and copying revised text back into manuscript document file. If the text is developed using plain text markup, such as Markdown, metadata, such as cited references and hyperlinks are typically persevered (assuming text is not changed drastically). On the other hand, if the text is developed using MS Word, and the references are managed with a reference manager, such as [Zotero][], reference information is typically encoded via Word fields, and this information is lost when copying a section of text to a plain-text AI chat. Further, hyperlink information also used for local navigation is also lost for the same reason. Additionally, there was another issue that I wanted to address. While Zotero manages references and creates citations automatically, it does not create hyperlinks leading from citations to corresponding bibliographic items. When a paper with Zotero-based references is converted to the PDF format, all references become plain text without convenient hyperlinks that make it easy to find the cited item in the bibliography list, from where the sources can often than be accessed via cited DOI-based or publisher-based URLs. In an attempt to address these issues, I decided to developed a set of VBA macro modules as a proof-of-concept solution.

The choice of VBA was dictated by its availability directly within MS Word without the need to install any additional software. And since I am working on text focused on generative AI (GenAI) prompting and also want to use GenAI for revising the text, it was a natural decision to use GenAI for macro development. I am sharing generated macro modules, prompts and approaches used for these prompt development, and one Google Gemini Advanced Pro 2.5 conversation used for revising one of the prompts and subsequent generation of the associated modules. The two other modules were generated in a separate conversation; however, that conversation has become somewhat messy due to my early experiments, and I do not see much value in sharing that conversation as well. 

## 1. Creation of Bibliographic Hyperlinks

The first generated module 
modBibliographyHyperlinker.bas






https://g.co/gemini/share/bcff2aa6b15f  
VBA for Bookmarks and Hyperlinks Information Management in MS Word  

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
[modBibliographyHyperlinker]: modBibliographyHyperlinker.bas
[modZoteroFieldRecovery]: modZoteroFieldRecovery.bas
