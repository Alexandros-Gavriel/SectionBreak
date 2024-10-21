/* global Word, Office */

// Office onReady event
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Initialization code if needed
  }
});

// Function to insert a section break at the cursor position
async function insertSectionBreakAtCursor(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // Insert the section break after the current selection
      selection.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
      await context.sync();

      // Insert an empty content control to force UI refresh
      const contentControl = selection.insertContentControl();
      contentControl.tag = "tempContentControl";
      contentControl.appearance = "Hidden"; // Make it invisible
      await context.sync();

      // Remove the content control
      contentControl.delete();
      await context.sync();
    });
    event.completed();
  } catch (error) {
    console.log("Error inserting section break at cursor: " + error);
    event.completed();
  }
}


// Function to insert a section break at the end of the document
async function insertSectionBreakDocumentEnd(event) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // Insert a section break at the end of the document
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
      await context.sync();

      // Move the selection to the end to force UI refresh
      const lastParagraph = body.paragraphs.getLast();
      lastParagraph.select();
      await context.sync();

      // Optionally, move the selection back to the original position
      // Depending on your use case
    });
    event.completed(); // Mark the event as completed
  } catch (error) {
    console.log("Error inserting section break at document end: " + error);
    event.completed();
  }
}

// Associate the functions with the names used in the manifest
Office.actions.associate("button1Function", insertSectionBreakAtCursor);
Office.actions.associate("button2Function", insertSectionBreakDocumentEnd);
