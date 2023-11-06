TranslateWinForm library
========================

## Overview
This library provides utility functions to translate the text in VB.net WinForm objects (e.g. menu items, forms and controls) to a different natural language (e.g. to French). 

The library contains utilities to:
- Automatically and recusively traverse all WinForm controls (menus, dialogs, labels, buttons etc.) in an application, and build a database of each control's text.
- Generate AI translations into any language (Spanish, Russian, Swahili etc.) of all the application's text using [CrowdIn](https://crowdin.com).
- If needed, use human translaters to improve the AI translations using [CrowdIn](https://crowdin.com).
- Generate a database of all the application's translations.
- Use this database to allow the user to dynamically flip between languages during runtime.

This class uses an SQLite database to translate text items to a new language. The database must contain the following tables:
```
CREATE TABLE "form_controls" (
"form_name"	TEXT,
"control_name"	TEXT,
"id_text"	TEXT,
PRIMARY KEY("form_name", "control_name")
)

CREATE TABLE "translations" (
"id_text"	TEXT,
"language_code"	TEXT,
"translation"	TEXT,
PRIMARY KEY("id_text", "language_code")
)
```
For example, if the 'form_controls' table contains a row with the values `{'frmMain', 'mnuFile', 'File'}`, 
then the 'translations' table should have a row for each supported language, e.g. `{'File', 'en', 'File'}, {'File', 'fr', 'Fichier'}`.

For an example of this library in use (including an extensive database), please see [R-Instat](https://github.com/africanmathsinitiative/R-Instat).