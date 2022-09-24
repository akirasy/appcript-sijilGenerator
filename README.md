# appcript-sijilGenerator

## Introduction

Generate sijil for any occasion using Google Spreadsheet. This app uses Google Appscript.
Just provide your own template or use our template available here.

## How to use

### 1. Prepare Google Drive Folder and Google Spreadsheet

1. Open `Google Drive` and create empty folder inside it (name the folder anything as you like).
1. Create an empty `Google Spreadsheet` inside the folder.
1. Open Appscript in-browser editor available at topbar-menu.
    ```
    Extensions --> App Script
    ```
1. Copy all code in [Code.gs](Code.gs) and paste it to the editor. Save and close it.
1. Reload your `Google Spreadsheet`.

### 2. Set up Google Appscript permission and Script Setup

1. There will be a new topbar-menu named `Sijil Generator`.
1. Click `Set Google permission` available at:
    ```
    Sijil Generator --> Setup --> Set Google Permission
    ```
1. Allow the script to run.
1. Click `Setup Spreadsheet` to configure your fresh Google Spreadsheet available at:
    ```
    Sijil Generator -- Setup --> Setup spreadsheet
    ```
1. Get your `sijil template` GoogleDoc ID and paste it to `Template file ID`.

### 3. Use the generator

1. This generator accepts up to 5 parameter. Please use `<<VAR1>>` to `<<VAR5>>` inside your
`sijil template` to replace it with your parameter (eg. name, id, date).
1. Prepare your candidate data at sheet `Data`.
1. Click `Generate Gdoc File` and let the script runs.
1. If you want it in PDF format, click `Convert Gdoc to PDF File`.
1. Enjoy!!

## Sijil Template

This is our template. Feel free to use and copy it to your folder.
1. [Hari Kemerdekaan](https://docs.google.com/document/d/1m-ehl05mF9zTI_EDGkkCUDnA-9xEZiVHgBhQ2i85KQ8/)
1. [Hari Tanpa Tembakau](https://docs.google.com/document/d/1Sc2qg3a0Yq533EyxR0f7ULTMD_8oWJWNEclYAgFBP4A/)

## License

This app is licensed under [GNU GPLv3](LICENSE).<br>Feel free to use under the terms of this license.
