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
      const body = context.document.body;
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

      // Move the selection to the end to force UI refresh
      const lastParagraph = body.paragraphs.getLast();
      lastParagraph.select();
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

      // Insert the section break after the current selection
      selection.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);

      // Move the selection to the new location to force UI refresh
      const newRange = selection.getRange(Word.RangeLocation.after);
      newRange.select();
      await context.sync();
    });
    event.completed(); // Mark the event as completed
  } catch (error) {
    console.log("Error inserting section break at cursor: " + error);
    event.completed(); // Ensure the event is completed even if there's an error
  }
}

// Associate the functions with the names used in the manifest
Office.actions.associate("button1Function", insertSectionBreakAtCursor);
Office.actions.associate("button2Function", insertSectionBreakDocumentEnd);
