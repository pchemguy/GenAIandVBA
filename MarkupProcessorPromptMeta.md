# AI-Driven Meta-Prompt Generation

### Starting Meta-Prompt

The following meta-meta-prompt and a hand-written and meta-prompt draft were used for generating revised meta-prompts:

>[!Note]
>
>Help me improve the following meta-prompt
>
>---
>
>Help me improve the following LLM prompt:
>- Check general English usage
>- Pay attention to clarity, logical organization and flow, structure, structural and grammatical parallelism, etc.
>- Strive for positive and actionable language, keeping negatives and exceptions as clarifiers or emphasizers.
>- Identify potential ambiguities.
>- If relevant, suggest additional guidelines / details in line with overall prompt intent and goals.

The model generated two revised meta-prompts.

### Option 1: Direct Enhancement (Closer to original structure) - Actually Used

>[!Note]
>
> **Review and Improve the following LLM Prompt:**  
>
>Analyze the provided LLM prompt based on these criteria. Provide specific, actionable feedback and suggestions for improvement:  
>1. **Clarity & Precision:** Is the language unambiguous? Are instructions direct? Identify potential ambiguities or vague terms.  
>2. **Structure & Flow:** Is it logically organized? Does it flow well? Check for parallelism (structural and grammatical).  
>3. **Language:**  
>    * Assess general English usage (grammar, spelling, punctuation).  
>    * Evaluate the use of positive and actionable language. Are negatives used effectively (e.g., for constraints, clarifications)?  
>1. **Completeness:** Does it seem to contain all necessary information for the LLM to perform the task effectively (context, constraints, desired output)?  
>2. **Suggestions for Enhancement:** Based on the likely intent, suggest any additional guidelines, details, or restructuring that would improve the prompt's effectiveness and the quality of the LLM's response.

### Option 2: More Comprehensive & Role-Based

>[!Note]
>
>**Act as an Expert Prompt Engineer:**
>
>Your task is to thoroughly review the LLM prompt provided below. Your goal is to identify weaknesses and suggest concrete improvements to enhance its clarity, effectiveness, and the quality of the expected LLM response.
>
>**Analyze the prompt across these dimensions:**
>1. **Clarity and Specificity:**
>    * Is the core task or question unambiguous?
>    * Are all instructions precise and easy to interpret?
>    * Identify and suggest replacements for vague language or jargon.
>    * Pinpoint potential ambiguities that could lead to misinterpretation.
>2. **Structure and Organization:**
>    * Is the prompt logically structured? Is information presented in a sensible order?
>    * Does it flow coherently?
>    * Evaluate the use of structural and grammatical parallelism for consistency.
>3. **Completeness and Context:**
>    * Does the prompt provide sufficient context for the LLM?
>    * Are necessary constraints, parameters, or background information included?
>    * Is the desired output format, style, length, or persona clearly defined?
>4. **Conciseness and Efficiency:**
>    * Is the prompt free of redundant or unnecessary information?
>    * Could it be stated more directly without sacrificing clarity?
>5. **Tone and Actionability:**
>    * Does the prompt primarily use positive, actionable instructions?
>    * Are negative constraints used appropriately (e.g., "Do not include X")?
>    * Is the overall tone suitable for guiding an LLM?
>6. **Language Quality:**
>    * Check for and correct any grammatical errors, spelling mistakes, or awkward phrasing.
>    * Ensure standard English usage.
>7. **Overall Effectiveness & Recommendations:**
>    * Provide an overall assessment of the prompt's likely effectiveness.
>    * Suggest specific additions, deletions, rephrasing, or restructuring to improve performance. Explain the reasoning behind your suggestions, connecting them to the likely goals of the prompt.
>
>**Deliverable:**
>Present your feedback in a structured manner, clearly referencing the parts of the original prompt you are addressing. Offer specific, actionable rewrite suggestions where appropriate.

