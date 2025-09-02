# Hello World Office.js Add-in

A simple "Hello World" Office.js extension that displays a basic page with a greeting message.

## Files Structure

- `manifest.xml` - The add-in manifest file that defines the add-in configuration
- `taskpane.html` - The main HTML page that displays "Hello World"
- `taskpane.js` - JavaScript file that initializes Office.js
- `taskpane.css` - CSS styles for the task pane
- `commands.html` - HTML file for function commands (required by manifest)
- `assets/` - Directory for images and other assets

## How to Use

1. Host these files on a web server (e.g., localhost:3000)
2. Update the URLs in `manifest.xml` to match your server location
3. Sideload the add-in in Excel, Word, or PowerPoint using the manifest file
4. The add-in will appear in the ribbon and display "Hello World" when opened

## Requirements

- Microsoft Office (Excel, Word, or PowerPoint)
- Web server to host the files
- Modern web browser support

## Notes

- The manifest uses localhost:3000 as the default URL - change this to your actual server URL
- You'll need to replace the placeholder logo file with actual PNG images
- This is a basic example - you can extend it with more Office.js functionality
