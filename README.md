<div align="center">
  <img src="assets/icon-80.png" alt="Active Cell Tracker Logo" width="80" height="80">
  <h1>Active Cell Tracker for Excel</h1>
  <p><strong>A lightweight and easy Office Add-in to dynamically track and export active cell coordinates in real-time.</strong></p>

  <p>
    <a href="https://github.com/komt0/Active_Cell_Excel/blob/main/LICENSE">
      <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License">
    </a>
    <a href="https://komt0.github.io/Active_Cell_Excel/taskpane.html">
      <img src="https://img.shields.io/badge/Hosted%20On-GitHub%20Pages-blue.svg" alt="GitHub Pages">
    </a>
    <a href="https://developer.microsoft.com/en-us/office">
      <img src="https://img.shields.io/badge/Office.js-v1.1-0f7b41.svg" alt="Office JS">
    </a>
  </p>
</div>

---

## Features

*  **Real-time Tracking:** Automatically detects active cell selection changes across your sheet.
*  **Custom Destination:** Set a specific target cell like `D4` or `Sheet2!A1`, where coordinates are sent instantly.
*  **Flexible Output Options:** Toggle whether to include the sheet name (e.g., `Sheet1!D4` vs. `D4`).
*  **Persistent Settings:** Stores your target preferences locally using `localStorage`.
*  **Cloud Hosted:** Fully deployed via GitHub Pages over secure HTTPS.

---

##  How to Install / Sideload

To use this add-in in **Excel for Web**:

1. Download or copy the [`manifest.xml`](manifest.xml) file from this repository.
2. Open **Excel for Web** in your browser.
3. Go to **Insert** > **Add-ins** > **Upload My Add-in**.
4. Select the `manifest.xml` file.
5. Click **Upload**. The **Active Cell** button will appear under the **Home** tab!

---

##  Built With

* [Office.js API](https://learn.microsoft.com/en-us/office/dev/add-ins/) - Official JavaScript API for Microsoft Office Add-ins.
* [Node.js](https://nodejs.org/) & [Webpack](https://webpack.js.org/) - Module bundler and build process.
* [GitHub Pages](https://pages.github.com/) - Live HTTPS hosting service.

---
