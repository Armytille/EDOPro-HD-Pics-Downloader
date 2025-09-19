# <img src="https://projectignis.github.io/assets/img/ignis_logo.png" width="80"/> Armytille's EDOPro HD Pics Downloader

[![PowerShell Version](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/en-us/powershell/scripting/overview)  
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)  

**Armytille’s EDOPro HD Pics Downloader** is a **PowerShell GUI application** that allows you to easily download **HD images of Yu-Gi-Oh! cards** for **EDOPro**.

---

## 🌟 Features

- Download **all Yu-Gi-Oh! cards in HD**.
- Special handling for **Field Spells cropped images**.
- Automatic support for **alternate arts**.
- Option to **force overwrite** existing images.
- **Real-time progress bar** and **log display**.
- Simple **Windows Forms GUI**.
- Automatic **error handling and retries**.

---

## ⚡ Installation & Usage

1. **Download the script** `EDOPro-HD-Pics-Downloader.ps1`.
2. **Place the file** at the **root of your EDOPro folder**, in the same location as the `pics` folder.  
3. **Run the script**: **Right-click → Run with PowerShell**.

4. **Using the GUI**:

- Click **“Download All Cards”** to start downloading.
- Check **“Force Overwrite Existing”** to overwrite existing images.
- Click **“Cancel”** to stop the download at any time.
- The **progress bar** and **log window** display real-time status.
  
<img width="616" height="261" alt="image" src="https://github.com/user-attachments/assets/69c0684e-5961-4e64-a503-192aede20b93" />
---

*(Optional)* Enable unrestricted script execution in PowerShell if needed:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

