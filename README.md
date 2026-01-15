# ExportAsUTF8_ForOldOffice
The Problem Microsoft Excel 2010 does not natively support exporting to **CSV UTF-8** The Solution This repository provides a VBA macro that bypasses the built-in "Save As" function. It uses the `ADODB.Stream` object to manually construct a CSV file and force **UTF-8 encoding**, ensuring all international characters are preserved perfectly.
