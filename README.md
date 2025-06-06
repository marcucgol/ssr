# SSR Demo

This project processes `.gge` files into Excel and provides a web interface for browsing and editing data.

## Usage

1. Install dependencies
   ```bash
   npm install
   ```
2. Prepare the input data:
   - Place your `.gge` files under `Объекты/` (use the upload page to create folders if needed)
   - Put `NLSR.xlsx` and `TEP.xlsx` in the project root
3. Start the server
   ```bash
   node app_s.js
   ```
   The web interface will be available at <http://localhost:2500>.
  Buttons on the main page let you upload new objects, edit `NLSR.xlsx` or `TEP.xlsx`, download the combined Excel file, or view its contents directly in the browser. The tables use sticky headers so column titles stay visible while scrolling. When viewing the combined Excel output you can zoom the table with a slider fixed at the top, and you can resize columns by dragging their headers. The chosen widths are stored in your browser.
  The upload page opens a file explorer for the `Объекты` folder. Use the toggle to switch between a grid of icons and a detailed list view (name, date, type, size) just like in Windows Explorer. Folder icons are embedded directly in the stylesheet so no binary images are required. There is also a button to open `C:\Users\User\4` in your system file manager.
  The NLSR and TEP editors include “Добавить строку” buttons above and below the table for quick entry. From those pages you can download the current Excel or upload a replacement file. When нажимаете кнопку «Обновить данные», вверху страницы появляется сообщение о ходе обработки и после завершения предлагается перезагрузить страницу.

## Generating Data

Use the "Обновить данные" button on the main page or send a GET request to `/generate` to reprocess the `.gge` files and refresh `combined_output.xlsx`.
