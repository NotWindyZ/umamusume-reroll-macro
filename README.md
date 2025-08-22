# 🐴 Uma Musume Reroll Macro

A GUI-based automation macro for **Uma Musume Pretty Derby**

## 📥 Download

👉 [**Download the latest release here**](https://github.com/NotWindyZ/umamusume-reroll-macro/releases)

Simply extract the folder, then run `start_macro.bat`. No Python installation required! For further instruction please ahead to "🚀 How to Use" section!

---

## ✨ Features

-  Automated rerolling & account registration  
-  Easy region/click assignment via GUI  
-  OCR-based rarity detection using EasyOCR  
-  Discord webhook notifications (with screenshots)  
-  Macro failsafe check 

### ⌨️ Hotkeys
- **F1** → Start macro  
- **F3** → Stop macro
  
---

## 🖥️ Requirements

- Windows 10/11  
- Python 3.12 (**embedded, already included in release folder**)
---

## 🚀 How to Use

1. **Extract the release folder** anywhere you like.
2. **Start the macro:**
   - Double-click `start_macro.bat` (recommended)
   - _Tip: Create a desktop shortcut to `start_macro.bat` for quick access!_
3. The macro GUI will open. Assign all required regions and clicks using the GUI tabs:
   - **Card regions** (for rarity detection)
   - **Macro reroll clicks**
   - **Register account clicks**
   - **In-game menu clicks**
   - **Safe check regions**
   - **Link account clicks**
   - Use the “Assign” buttons to snip or set coordinates for each action.
4. **Set up Discord webhook (optional):**
   - Enter your Discord webhook URL and (optionally) a user ID to ping in the Webhook tab.
   - Click “Send Test Notification” to verify.
5. **Start/Stop the Macro:**
   - Press `F1` to start, `F3` to stop, or use the GUI buttons.

---

## 📝 Notes

- On first run, a default `config.ini` will be created if missing. _If you lose or delete it, just restart the macro!_
- All configuration is saved in `config.ini` and can be edited via the GUI.
- Screenshots for Discord notifications are saved temporarily and deleted after sending.
- For best OCR results, assign card regions as tightly as possible around the card rarity icons.
- The macro will not run unless all required regions/clicks are assigned.
- All logs are saved to `macro_log.txt` (rotates at 5MB) and successful reroll accounts to `met_required_accounts.txt` (rotates at 10MB).

---

## ⚡ Optional: Enable GPU Support for Faster OCR

By default, the macro uses CPU for EasyOCR. This is fine for most users, but it can be a little slow.  
If you have an **NVIDIA GPU with CUDA support**, you can install **PyTorch with CUDA** inside the embedded Python environment to make OCR much faster.

⚠️ **Warning:** Installing GPU support will download about **7 GB** of data.
Just locate the bat file: **`install-ocr-gpu-support.bat`**

---
## 🛠️ Troubleshooting & Common Problems

- **`config.ini` missing or corrupted?**  
  The macro will auto-create a default `config.ini` on startup. Just restart the macro.

- **OCR not working or inaccurate?**  
  - Make sure your screen is scaled (100% recommended).
  - Nvidia GPU for faster card detection
  - If OCR fails to load, try reinstalling EasyOCR or ensure your graphics drivers are up to date.
  - For best results, assign card regions closely around the rarity icons.

- **Macro doesn't detect the game window?**  
  - The game must be running and visible!
  - Fullscreen on the game also!

- **Discord notifications not sending?**  
  - Double-check your Discord webhook URL.

- **Can't find `start_macro.bat`?**  
  - If missing, create a `.bat` file with the following content in the macro folder:
    ```bat
    @echo off
    REM === Always go into MACRO_HERE_OPEN_IT folder first ===
    cd /d "%~dp0MACRO_HERE_OPEN_IT"
    
    REM === Run main.py with the embedded python.exe ===
    "..\python.exe" "main.py"
    
    echo.
    echo Macro finished running.
    pause
    ```
  - Then create a shortcut to this `.bat` on your desktop for easy access.

---

## 📦 Included Python Modules (for reference)

The macro uses the following Python modules (already included in the embedded Python):

- autoit
- opencv-python
- easyocr
- keyboard
- numpy
- pygetwindow
- pyperclip
- pillow
- requests
- ttkbootstrap

---

Enjoy rerolling! For issues or suggestions, open an issue on GitHub or contact [@notwindybee_](https://discord.com/users/youridhere) on Discord.
