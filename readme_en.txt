================================================================================
UC Case Registry
================================================================================

A clinical data collection web application designed for Ulcerative Colitis (UC) 
studies, specifically comparing JAK inhibitors and S1P modulators.

[Please see readme_ja.txt for the Japanese version.]

--------------------------------------------------------------------------------
Features
--------------------------------------------------------------------------------

* No Server Required
  Runs entirely in the web browser. No backend server or database setup is needed.

* Local File Storage
  Uses the modern File System Access API to save and load data directly to/from 
  a local .json file on your computer. Falls back to localStorage if not supported.

* Bilingual Voice Input
  Enter laboratory data quickly using voice commands. Supports both English and 
  Japanese medical terminology (e.g., "CRP 5.2", "White blood cell 6500").

* Auto-Calculations
  Automatically calculates Age, BMI, Disease Duration, PRO-2 scores, Absolute 
  Neutrophil Count (ANC), and Absolute Lymphocyte Count (ALC) based on input data.

* Data Export
  Export your collected cases to a CSV file or copy the data formatted for direct 
  pasting into Excel.

* Theme Support
  Includes a premium medical dark theme and a light theme, with a toggle button.

--------------------------------------------------------------------------------
How to Use
--------------------------------------------------------------------------------

1. Open the App
   Simply open index.html in a modern web browser (Google Chrome or Microsoft 
   Edge is recommended for full feature support, including File System Access 
   and Speech Recognition).

2. Link a Storage File
   - Click the "Link File" (folder icon) button in the top navigation bar.
   - Choose a location on your computer to save the cases.json file.
   - If you select an existing JSON file created by this app, it will load the 
     data. If you create a new one, your current data will be saved to it.
   - Note: For security reasons, browsers require you to link the file each time 
     you open the app or refresh the page.

3. Register a Case
   - Fill out the form in the "+ New Case" tab.
   - Use the microphone button next to laboratory data sections to use voice 
     input. You can toggle between English (EN) and Japanese (JA) recognition.
   - Click "Register Case" at the end of the form to save.

4. Manage Cases
   - Switch to the "Case List" tab to view, search, edit, or delete registered 
     cases.

5. Export Data
   - Use "Export CSV" to download the data.
   - Use "Copy for Excel" to copy the data and paste it directly into an Excel 
     spreadsheet.

--------------------------------------------------------------------------------
File Structure
--------------------------------------------------------------------------------

* index.html  : The main user interface.
* style.css   : The stylesheet containing the design system and themes.
* app.js      : The core application logic.
* fields.js   : Configuration file defining the form steps and fields.
* readme_en.txt  : English documentation.
* readme_ja.txt  : Japanese documentation.

--------------------------------------------------------------------------------
Requirements
--------------------------------------------------------------------------------

* A modern web browser.
* Google Chrome or Microsoft Edge is strongly recommended for:
  - File System Access API: For saving data directly to a local file.
  - Web Speech API: For voice input functionality.