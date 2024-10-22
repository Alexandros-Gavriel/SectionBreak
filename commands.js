/* global Word, Office */

// Office onReady event
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Initialization code if needed
  }
});

// Function to insert a section break at the end of the document
async function insertSectionBreakDocumentEnd(event) {
  try {
    await Word.run(async (context) => {
      // Insert a section break at the end of the document
      context.document.body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
      await context.sync();
    });
    event.completed(); // Mark the event as completed
  } catch (error) {
    console.log("Error inserting section break at document end: " + error);
    event.completed(); // Ensure the event is completed even if there's an error
  }
}

// Function to insert a section break at the cursor position
async function insertSectionBreakAtCursor(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
      await context.sync();
    });
    event.completed(); // Mark the event as completed
  } catch (error) {
    console.log("Error inserting section break at cursor: " + error);
    event.completed(); // Ensure the event is completed even if there's an error
  }
}

// Remove section break at cursor
async function removeSectionBreakAtCursor(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // Expand selection to include possible section break
      selection.expand("Character");
      await context.sync();

      // Search for section break
      const searchResults = selection.search("^b", { matchWildcards: true });
      context.load(searchResults, "items");
      await context.sync();

      if (searchResults.items.length > 0) {
        // Delete the section break
        searchResults.items[0].delete();
        await context.sync();
      } else {
        console.log("No section break found at the cursor position.");
      }
    });
    event.completed();
  } catch (error) {
    console.error("Error removing section break at cursor: " + error);
    event.completed();
  }
}

// Remove all section breaks in document
async function removeAllSectionBreaks(event) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // Search for all section breaks
      const searchResults = body.search("^b", { matchWildcards: true });
      context.load(searchResults, "items");
      await context.sync();

      // Delete all section breaks found
      for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].delete();
      }
      await context.sync();
    });
    event.completed();
  } catch (error) {
    console.error("Error removing all section breaks: " + error);
    event.completed();
  }
}

// Associate the functions with the names used in the manifest
Office.actions.associate("button1Function", insertSectionBreakAtCursor);
Office.actions.associate("button2Function", insertSectionBreakDocumentEnd);
Office.actions.associate("button3Function", removeSectionBreakAtCursor);
Office.actions.associate("button4Function", removeAllSectionBreaks);