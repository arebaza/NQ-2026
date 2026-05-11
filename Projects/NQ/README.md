# NQ Recon Dashboard

This project converts your Excel NQ report into a web dashboard similar to your Power BI page.

## Stack
- HTML
- CSS
- JavaScript
- Chart.js
- SheetJS

## Project structure
- `index.html` → page layout
- `styles.css` → dashboard styling
- `app.js` → Excel loader, filters, tables, charts
- `data/Team Recon Report NQ 2.xlsx` → default Excel source

## Run locally in VS Code
1. Open the folder in Visual Studio Code.
2. Install the **Live Server** extension.
3. Right-click `index.html`.
4. Select **Open with Live Server**.

## Publish in GitHub Pages
1. Create a new GitHub repository.
2. Upload all files from this project.
3. Go to **Settings > Pages**.
4. In **Build and deployment**, choose **Deploy from a branch**.
5. Select branch `main` and folder `/root`.
6. Save.

## Update with a new Excel export
- Replace the file inside `/data` with the same name, or
- Use the upload button in the browser

## Suggested next features
- Add a Summary Outstanding page
- Add Comments page
- Add month filter
- Add export buttons
- Add role access by office
