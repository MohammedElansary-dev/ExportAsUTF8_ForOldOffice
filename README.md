# Excel 2010 UTF-8 CSV Exporter (VBA)
The Problem
Microsoft Excel 2010 does not natively support exporting to **CSV UTF-8**
The Solution
This repository provides a VBA macro that bypasses the built-in "Save As" function. It uses the `ADODB.Stream` object to manually construct a CSV file and force **UTF-8 encoding**, ensuring all international characters are preserved perfectly.
### The Problem
Microsoft Excel 2010 does not natively support exporting to **CSV UTF-8**. When saving as a standard CSV, special characters (like Arabic, Persian, or emojis) often turn into gibberish or question marks (????) because the software defaults to ANSI encoding.

### The Solution
This repository provides a VBA macro that bypasses the built-in "Save As" function. It uses the `ADODB.Stream` object to manually construct a CSV file and force **UTF-8 encoding**, ensuring all international characters are preserved perfectly.

---

## ✨ Features
* **True UTF-8 Support:** Works perfectly for Arabic and other non-Latin scripts.
* **Save Dialog:** Opens a standard Windows window to choose your file destination.
* **Data Sanitization:** Automatically handles commas inside cells to prevent column breaking.
* **High Performance:** Disables screen updating and calculations during export for speed.
* **Row Counter:** Confirms the exact number of rows exported upon completion.

---

## 🚀 How to Use

1. **Open Excel:** Open the workbook containing the data you want to export.
2. **Open VBA Editor:** Press `Alt + F11`.
3. **Insert Module:** Go to `Insert > Module`.
4. **Paste Code:** Copy the code from `ExportAsUTF8_ForOldOffice_Cgroup.vba` (found in this repo) and paste it into the module.
5. **Enable Reference:** * In the VBA window, go to `Tools > References`.
   * Check the box for **Microsoft ActiveX Data Objects 6.1 Library** (or the highest version you see should be on).
6. **Run:** Press `Alt + F8`, select `ExportAsUTF8_Professional`, and click **Run**.

---

## 🛠️ Technical Explanation
Since Excel 2010 lacks the `xlCSVUTF8` constant, this script utilizes:
- **ADODB.Stream:** To handle the character encoding at a system level.
- **UsedRange Loop:** To identify only cells containing data.
- **Error Handling:** To prevent crashes if the file is locked or the path is invalid.

## ⚠️ Important Note on Opening CSVs in Excel
Even though the file is correctly saved in UTF-8, double-clicking a CSV to open it in Excel 2010 may still display characters incorrectly. 

**To view correctly:**
Open Excel > Go to **Data** Tab > **Get External Data (From Text)** > Choose your file > Set File Origin to **65001: Unicode (UTF-8)**.
or just open it with a text editor or using specific programs like [modern csv (free)][https://www.moderncsv.com/] or [CSView (free and open source)][https://kothar.net/csview] 
---

*Created to solve legacy software limitations.*
