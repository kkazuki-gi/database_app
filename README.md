# UC Case Registry

A clinical data collection web application designed for Ulcerative Colitis (UC) studies, specifically comparing JAK inhibitors and S1P modulators.

## Features

- **No Server Required**: Runs entirely in the web browser. No backend server or database setup is needed.
- **Local File Storage**: Uses the modern File System Access API to save and load data directly to/from a local .json file on your computer. Falls back to localStorage if not supported.
- **Bilingual Voice Input**: Enter laboratory data quickly using voice commands. Supports both English and Japanese medical terminology.
- **Auto-Calculations**: Automatically calculates Age, BMI, Disease Duration, PRO-2 scores, ANC, and ALC.
- **Data Export**: Export your collected cases to a CSV file or copy data formatted for Excel.
- **Theme Support**: Includes premium medical dark and light themes.

## How to Use

1. **Open the App**: Open `index.html` in a modern web browser (Chrome or Edge recommended).
2. **Link a Storage File**: Click "Link File" button to save data to a local JSON file.
3. **Register a Case**: Fill out the multi-step form with patient information.
4. **Manage Cases**: View, search, edit, or delete cases in the Case List tab.
5. **Export Data**: Use "Export CSV" or "Copy for Excel" buttons to export your data.

## File Structure

- `index.html` - Main user interface
- `style.css` - Stylesheet with dark/light themes
- `app.js` - Core application logic
- `fields.js` - Form configuration (127 columns)
- `README.md` - Documentation

## Requirements

- Modern web browser (Chrome/Edge recommended)
- No installation or backend required
- For full features: support for File System Access API and Web Speech API
