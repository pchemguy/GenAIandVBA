# Template-Based

```
Analyze the following prompt and consider if instructions are clear and unambiguous.
Provide feedback/questions on any potential issues.
Devise a workflow to be placed in place of "{TO BE SUGGESTED BY AI}".

---

# Prompt: ... (Title) ...

## Persona:
... (Description of a suitable role) ...

## Task:
... (Description of the task) ...
... (Detailed requirements) ...

## Processing Steps:
{TO BE SUGGESTED BY AI} 
```

# ICL-Facilitated

```
Help me create a prompt, using previous prompts as a reference.
The generated prompt should acco

The macro would need to
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