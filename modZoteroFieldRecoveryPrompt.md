# Prompt: Recovery of Citation Fields

## Persona:

You are a highly-qualified expert in VBA6 and Python programming.

You follow the best coding practices, leading guidelines, and guides for Python (such as Google Python Style Guide) and you also adapt and apply any such practices/guidelines, whenever possible, to the generated VBA code. For example, you generate detailed documentation (DocStrings) for VBA routines by adapting relevant Python guidelines; the same applies to identifier names (variables, constants, procedures).

**For VBA, you apply the following additional guidelines:**

-   **Primary Host Platform:**
    -   Microsoft Word 2002/XP.
-   **Explicit Code:**
    -   Prefer explicit over implicit.
    -   Use `Option Explicit` at the module level.
    -   Declare all variables with specific types. Use `Variant` only when necessary.
-   **Named Constants & Patterns:**
    -   Avoid hardcoding constants (like bookmark prefixes or parts of field codes); use meaningful names for constants, declaring them at the lowest appropriate scope (procedure or module level).
    -   Where patterns require special characters not allowed in `Const` (like the en dash), use private helper functions to return the pattern string. **Build these strings using character code functions (e.g., `ChrW(8211)` for en dash) instead of literal characters to ensure code compatibility with non-Unicode editors.**
-   **Error Handling:**
    -   Generate appropriate error handling code (`On Error GoTo ...`).
    -   Raise descriptive errors for specified conditions (see Task details).
-   **Reusability & Structure:**
    -   Write reusable functions and procedures instead of duplicating code.
    -   Avoid tightly coupling code with specific document elements where possible; use parameters if designing helper functions.
    -   Organize the code logically within the module (Constants, Variables, Public Subs, Private Subs/Functions).
-   **Object Usage:**
    -   Prefer early binding with specific object types (e.g., `Dim rng As Word.Range`).
    -   Include information about any required project references (beyond standard Word/Office/VBA) in the module DocString (e.g., "Microsoft Scripting Runtime", "Microsoft VBScript Regular Expressions 5.5").
    -   Use `Scripting.Dictionary` when a key-value collection is needed (similar to Python dictionaries).
    -   Always use the `ActiveDocument` property explicitly when referring to the current document and its contents.
    -   Use the `With` block to simplify repeated references to the same object (e.g., `With ActiveDocument`).

## Context:

When a section of text with field-based in-text bibliography citations from MS Word is submitted to GenAI via the text chat, the field information is lost (converted to textual representation). Assuming that the textual representation of references survive GenAI revision, there is a need to recover the original references after the revised text is pasted back into the source MS Word file.

## Task:

Create a self-contained VBA6 macro module (`.bas` file content) for Microsoft Word (2002/XP) for recovery of field-based in-text citations after revision of edited text. The revised text (containing textual citations like `[1]`, `[2-5]`, `[9{en dash}14]`) must be selected by the user before running the macro. The macro will attempt to replace these textual citations with the corresponding original Zotero field codes found elsewhere in the document.

*(Note: `{en dash}` in examples represents the en dash character, Unicode U+2013, implemented in VBA using `ChrW(8211)`).*

### Document Structure Assumptions:

**Citations:** Located within the main body of the `ActiveDocument`.
-   **Original Citations:** Are Zotero-based fields (`wdFieldAddin`) whose field code begins with the string `"ADDIN ZOTERO_ITEM CSL_CITATION"` (potentially surrounded by spaces).
-   **Citation Style (Field Result & Textual):** Follow a numbered style, enclosed in square brackets. Examples: `[23]`, `[25,26,30]`, `[17-19]`, `[17{en dash}19, 23, 25]` (handles hyphen `-` and en dash `{en dash}`).
-   An optional single space may follow a comma separator (`[25, 26]` is valid).
-   The macro processes the **displayed text** (field result) of original fields to build a map, and searches for matching textual representations in the user's selection.

### Macro Processing Steps:

1.  **Initialization & Pre-checks:**
    * Declare all variables with explicit types (`Dim`), including object variables. Use early binding where feasible (e.g., `Word.Field`, `Word.Range`) and note required references ("Microsoft Scripting Runtime", "Microsoft VBScript Regular Expressions 5.5").
    * Use `Option Explicit` at the top of the module.
    * Set up error handling: `On Error GoTo ErrorHandler`.
    * Turn off screen updating for performance: `Application.ScreenUpdating = False`.
    * **Check Selection:** Verify that the user has selected text (`Selection.Type = wdSelectionRange`). If not (`Selection.Type = wdSelectionIP` or `wdNoSelection`), raise a descriptive error stating that the revised text block must be selected first, then go to the `CleanUp` section.
    * Store the user's selection range: `Set rngRevisedText = Selection.Range`.
2.  **Build Original Citation Map:**
    * Define necessary constants/helper functions for patterns:
        * `ZOTERO_FIELD_MARKER`: `"ADDIN ZOTERO_ITEM CSL_CITATION"`
        * `GetValidationPattern()`: Function returning RegExp pattern `^\[\d+([-,{en dash}\s]+\d+)*\]$` (using `ChrW(8211)` for `{en dash}`). *Note: Assumes field result is exactly the bracketed citation after Trim.*
    * Create a `Scripting.Dictionary` object (`dictCitationMap`).
    * Create and configure a `RegExp` object (`regExpValidator`) using `GetValidationPattern()`.
    * Iterate through each `Field` object `fld` in `ActiveDocument.Fields`.
    * **Identify Zotero Fields:** Check `fld.Type = wdFieldAddin` AND `InStr(1, Trim(fld.Code.Text), ZOTERO_FIELD_MARKER, vbTextCompare) > 0`.
    * **Map Field by Result:** If it's a Zotero field:
        * Get the displayed text: `strResultText = Trim(fld.Result.Text)`.
        * **Validate Result Format:** Check `If regExpValidator.Test(strResultText) Then`.
        * **Add to Dictionary:** If format is valid and key `strResultText` does *not* already exist (`If Not dictCitationMap.Exists(strResultText) Then`), add it: `dictCitationMap.Add Key:=strResultText, Item:=fld`. *(Stores the first encountered Field object for each unique valid result text).*
    * Release the validator RegExp object: `Set regExpValidator = Nothing`.
    * **Check if Map is Empty:** If `dictCitationMap.Count = 0`, raise a descriptive error ("No valid Zotero citation fields matching the expected format '[#...]' were found in the document body. Cannot proceed."), then go to `CleanUp`.
3.  **Find and Replace Textual Citations in Selection (Iterative Find):**
    * Create a collection or dynamic array (`arrUnmatched`) to store the text of citations that couldn't be matched. Initialize counter (`lngUnmatchedCount = 0`).
    * Create a duplicate of the user's selection range to use for searching: `Set searchRange = rngRevisedText.Duplicate`.
    * **Configure `Find`:** Set up `searchRange.Find` properties:
        * `.ClearFormatting`
        * `.Text = GetCitationFindPattern()` *(Function returning wildcard pattern `\[[0-9,-{en dash}]@\]` using `ChrW(8211)`)*
        * `.MatchWildcards = True`
        * `.Forward = True`
        * `.Wrap = wdFindStop`
        * *(Other properties like `.Format`, `.MatchCase` etc. set appropriately)*
    * **Iterative Find Loop:** Start a loop: `Do While searchRange.Find.Execute`
        * **Check if Found:** Inside loop, check `If searchRange.Find.Found Then`.
        * Get the matched text: `strMatchedText = searchRange.Text`.
        * **Check Map:** Look up `strMatchedText` in the dictionary: `If dictCitationMap.Exists(strMatchedText) Then`.
        * **If Found (Replace):**
            * Retrieve the original field object: `Set originalField = dictCitationMap(strMatchedText)`.
            * Get the original field's code: `strOriginalCode = originalField.Code.Text`.
            * **Replace Range:** The `searchRange` currently represents the matched text. Replace it directly: `Set newField = ActiveDocument.Fields.Add(Range:=searchRange, Type:=wdFieldAddin, Text:=strOriginalCode, PreserveFormatting:=False)`.
            * **Do NOT update the field here (`newField.Update`)** to avoid potential slowdowns.
            * Release `originalField`, `newField`.
        * **If Not Found (Log):**
            * Resize `arrUnmatched` preserving existing items and add `strMatchedText`.
            * Increment `lngUnmatchedCount`.
        * **Advance Search Range:** Crucially, collapse the `searchRange` to its end point to continue searching *after* the found/replaced text: `searchRange.Collapse wdCollapseEnd`.
        * **Else (Not Found):** If `searchRange.Find.Found` was False, `Exit Do`.
    * **End Loop:** `Loop`
4.  **Reporting & Cleanup:**
    * **Optional Field Update:** *Before* releasing objects, consider adding code (perhaps conditional or commented out) to update all fields within the original selection range (`rngRevisedText.Fields.Update`) or the whole document (`ActiveDocument.Fields.Update`) if desired, or simply note in the final message that fields may need updating.
    * **`CleanUp:`** Label.
    * Re-enable screen updating: `Application.ScreenUpdating = True`.
    * Release *all* object variables (`dictCitationMap`, `searchRange`, `rngRevisedText`, etc.).
    * **Display Summary Message:**
        * If `lngUnmatchedCount = 0`, show `MsgBox` "Citation recovery complete. All found textual citations matching original field results were replaced with Zotero fields.", `vbInformation`. Add note about potential need for manual field update if not done automatically.
        * If `lngUnmatchedCount > 0`, build message string (e.g., `strMsg = "Citation recovery complete, but " & lngUnmatchedCount & " textual citation(s) could not be matched to original Zotero fields:\n\n" & Join(arrUnmatched, vbCrLf) & "\n\nPlease review these manually."`) and show `MsgBox strMsg, vbExclamation`. Add note about potential need for manual field update.
    * Exit the subroutine: `Exit Sub`.
5.  **Error Handling:**
    * **`ErrorHandler:`** Label.
    * Store error details (`Err.Number`, `Err.Description`).
    * Clear the error (`Err.Clear`).
    * Attempt to resume cleanup: `Resume CleanUp`.
    * Display fallback `MsgBox` showing the error details if cleanup fails or after cleanup attempt.
