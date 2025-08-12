
# 📦 Product Images Downloader from Excel to ZIP

A lightweight **web app** built with **HTML, CSS, and JavaScript** that allows you to:  
- Upload an Excel file containing **product codes** and **image URLs** (multiple links per product separated by commas, semicolons, or pipes).  
- Automatically **download all images** and export them as a **ZIP file**.  
- Optionally **organize images into folders by product code**.  
- Supports **English and Arabic column names** for automatic column detection.  

---

## ✨ Features
- **Smart column detection**: Finds product code and image URL columns automatically (supports Arabic & English headers or infers from data).  
- **Multiple link delimiters** supported: `,` `;` `|` `،` or new lines.  
- **Dashboard statistics**:
  - Total products in file.
  - Products with image links.
  - Total image links.
  - Estimated total image size before download.
  - Estimated ZIP file size after compression.
  - Number of ZIP files if split.
- **Progress tracking** with live updates.  
- **Arabic-friendly design** with [Tajawal font](https://fonts.google.com/specimen/Tajawal).  
- **Folder per product** option.  
- **Responsive UI** with styled upload button and clean charts.  

---

## 📂 How to Use
1. Open the web app in your browser.  
2. Upload your Excel file containing product codes and image URLs.  
3. Adjust settings if needed (concurrency, ZIP split size, folder per product).  
4. Click **Start Download** and wait for the process to complete.  
5. Save your ZIP file(s).  

---

## 🛠 Technologies Used
- **HTML5**, **CSS3**, **JavaScript (Vanilla)**  
- [JSZip](https://stuk.github.io/jszip/) – for ZIP creation.  
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) – for saving files.  
- [Chart.js](https://www.chartjs.org/) – for dashboard charts.  
- [XLSX.js](https://sheetjs.com/) – for Excel parsing.  
- Google Fonts – [Tajawal](https://fonts.google.com/specimen/Tajawal) for Arabic styling.

---

## 📸 Screenshots

### Upload & Dashboard
![Upload Screenshot](screenshots/upload-dashboard.png)

### Progress Tracking
![Progress Screenshot](screenshots/progress-tracking.png)

### Settings Panel
![Settings Screenshot](screenshots/settings.png)

---

## 🚀 Pro Features
- Robust networking: **Retry + Timeout + Backoff**, and **per-host rate limiting**.
- Controls: **Pause / Resume / Cancel** during download.
- Smart parsing: **Manual column selection** + preview top 50 rows.
- Naming & org: **Custom filename patterns** and **folder-per-product**.
- Safety: **Whitelist/Blacklist** by domains.
- Insights: ETA & download speed, **extension pie**, **top host bar**, **size histogram**.
- Estimation: HEAD-based size estimation; shows known/unknown, min/avg/max, total, and expected ZIP parts.
- Reporting: Export **full CSV report** and **failures-only CSV** after run.


## 📜 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🤝 Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

---


## 💡 Author
Developed by **[Ahmed Elsawi]**
