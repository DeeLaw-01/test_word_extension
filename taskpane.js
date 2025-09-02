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
