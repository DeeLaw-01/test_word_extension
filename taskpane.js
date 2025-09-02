/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    console.log("Hello World Add-in loaded in Excel!");
  } else if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    console.log("Hello World Add-in loaded in Word!");
    
    // Insert "hello world" into the Word document when the extension loads
    insertHelloWorldIntoDocument();
  } else if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    console.log("Hello World Add-in loaded in PowerPoint!");
  } else {
    console.log("Hello World Add-in loaded in unknown host!");
  }
});

// Simple function to demonstrate Office.js functionality
function sayHello() {
  console.log("Hello from Office.js Add-in!");
  
  // You can add more Office.js specific functionality here
  // For example, interacting with the document, workbook, or presentation
}

// Function to insert "hello world" into the Word document
function insertHelloWorldIntoDocument() {
  return Word.run(async (context) => {
    try {
      // Get the current selection or insert at the beginning of the document
      const selection = context.document.getSelection();
      
      // Insert "hello world" text
      selection.insertText("hello world", Word.InsertLocation.replace);
      
      // Sync the context to execute the queued commands
      await context.sync();
      
      console.log("Successfully inserted 'hello world' into the document");
    } catch (error) {
      console.error("Error inserting text into document:", error);
      
      // Fallback: try inserting at the beginning of the document body
      try {
        const body = context.document.body;
        body.insertText("hello world", Word.InsertLocation.start);
        await context.sync();
        console.log("Successfully inserted 'hello world' at document start");
      } catch (fallbackError) {
        console.error("Fallback insertion also failed:", fallbackError);
      }
    }
  });
}
