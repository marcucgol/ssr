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
   Buttons on the main page let you upload new objects, edit `NLSR.xlsx` or `TEP.xlsx`, and download the combined Excel file.
   The upload page shows a collapsible tree of the `Объекты` folder so you can create or remove directories and files visually.

4. **Optional SSH access**

   Run `./ssh_connect.sh` to open a shell (user: `user`, pass: `pass`). The server listens on port `2222`.

## Generating Data

Use the "Обновить данные" button on the main page or send a GET request to `/generate` to reprocess the `.gge` files and refresh `combined_output.xlsx`.
