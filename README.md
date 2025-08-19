# AutoFlow PDF – PDF Automation Tool

## 📌 Overview
AutoFlow PDF is a Windows Forms application built in C# for automating repetitive PDF-related workflows.  
It allows you to record **action points** (clicks, text inputs, scrolling, hotkeys, file dialogs, etc.) and replay them across multiple PDF files.

---

## ✨ Features
- 📂 Upload and process multiple PDF files
- 🎯 Define automation steps:
  - Tap / Click
  - Scroll
  - Text Input
  - Hotkey
  - Right Click
  - File Dialog handling
- ⏱ Add delays before/after actions
- 🎨 Conditional clicks (based on screen pixel color)
- 📋 Save and load automation profiles as JSON
- 🖱 Drag-and-drop to reorder steps
- ⏸ Pause/Resume automation with `Ctrl + Shift + Z`
- 📝 Logs of executed actions
- 🔍 Highlight current PDF being processed

---

## 🚀 How It Works
1. **Upload PDFs** – Select one or more PDF files.
2. **Add Action Points** – Define where and what actions should occur on the screen.
3. **Save Profile** – Store your sequence of steps for reuse.
4. **Run Automation** – The app executes steps sequentially on each PDF.
5. **Resume/Continue** – If paused, you can resume from a selected step.

---

## 🛠 Tech Stack
- **C#** (.NET Windows Forms)
- **Win32 API** (for mouse/keyboard control)
- **JSON** (for profiles)
- **Regex & System.IO** (for filename parsing and file handling)

---

## 📦 Project Structure
/PdfAutomationApp
├── Form1.cs # Main automation logic and UI
├── BrowserForm.cs # WebView2-based browser window
├── DelayPopup.cs # Delay configuration dialog
├── HotkeyCaptureForm.cs # For capturing hotkeys
├── Profiles/ # Saved profiles in JSON
├── PdfAutomationApp.csproj
└── README.md

yaml
Copy
Edit

---

## 🎮 Usage
- Run the app (`PdfAutomationApp.exe`).
- Upload PDFs using the **Upload PDF** button.
- Add automation steps with the left-side buttons.
- Start automation with **Start**.
- Use **Ctrl + Shift + Z** to pause automation.
- Use **Continue** to resume from the selected step.

---

## ⚠️ Notes
- Ensure correct screen resolution and UI layout for consistent automation.
- File dialogs may need manual intervention if they don’t open/close as expected.
- Use color-matching taps for reliable automation on dynamic UIs.

---

## 📜 License
[MIT License](LICENSE)
