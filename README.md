# UpdatePilcrows Word Visual Basic Application Script

This repository contains a Visual Basic for Applications (VBA) script designed to automatically manage pilcrow symbols (`¶`) in Microsoft Word documents. The script is executed before each document save operation, ensuring that every paragraph starts and ends with a pilcrow symbol.

## How It Works

The script is composed of a few key components:

1. **Application Event Handling**: The script utilizes the `WithEvents` keyword to handle Word application events, specifically the `DocumentBeforeSave` event. This event is triggered automatically every time the document is saved.

2. **Pilcrow Management**: The `UpdatePilcrows` subroutine is responsible for adding a pilcrow (`¶`) at the start and end of each paragraph in the document. It first removes any existing pilcrows at the beginning or end of paragraphs to avoid duplication, then ensures that a single pilcrow is present in the correct locations.

## Implementation Steps

To use this script in your Word document:

1. **Enable Macros**: Before using this script, make sure that macros are enabled in your Microsoft Word settings. Without enabling macros, the script will not run automatically.

2. **Add the Script to ThisDocument**:
   - Open your Word document.
   - Press `Alt + F11` to open the VBA editor.
   - In the Project Explorer, locate `ThisDocument` under `Microsoft Word Objects`.
   - Copy the provided script into the `ThisDocument` module. This ensures that the script runs for this specific document.

3. **Save the Document**: Save the document as a macro-enabled Word document (`.docm`). This format supports the execution of VBA scripts.

4. **Automatic Execution**: The script will now run automatically each time the document is saved. It will update the paragraphs in the document by adding pilcrows at the start and end of each one.

## Important Notes

- **Macros Must Be Enabled**: Ensure that macros are enabled in your Word environment. The script will not execute if macros are disabled.
- **Script Location**: For the automation to work across documents, the script must be placed in the `ThisDocument` module within `Microsoft Word Objects`. This is crucial as the script relies on document-specific events that are only accessible in this module.

## Example

Here’s a simplified version of what the script does:

- Before Save:
  - Removes any existing pilcrows at the start and end of each paragraph.
  - Adds a single pilcrow at the start and end of each paragraph.

- After Save:
  - Your document will have each paragraph neatly enclosed with a pilcrow symbol.
