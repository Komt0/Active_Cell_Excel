<div align="center">
  <img src="assets/icon-80.png" alt="Active Cell Tracker Logo" width="80" height="80">
  <h1>Active Cell Tracker for Excel</h1>
  <p><strong>A lightweight and powerful Office Add-in to dynamically track, capture, and export active cell data in real-time.</strong></p>

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

 **Real-time Tracking:** Automatically detects active cell selection changes across your worksheets.
* **Multiple Export Modes:** Choose what to export to your destination cell:
  * **Cell Address** (e.g., `D4`)
  * **Cell Value** (Content/Text)
  * **Cell Formula**
  * **Combined Value & Address** (e.g., `$1000 / D4`)
* **Target Sheet & Cell Selection:** Select any specific worksheet and target cell (e.g., `Sheet2!V1436`) to receive output.
* **Jump to Target Cell:** Quickly navigate and focus directly on your defined target cell with a single click.
* **Play / Pause Control:** Toggle cell tracking on or off anytime without closing the add-in panel.
* **Auto-Copy to Clipboard:** Automatically copy captured cell coordinates or data straight to your OS clipboard.
* **Live Activity Card:** View the last registered cell value along with an exact timestamp.
* **Persistent Settings:** Saves all your custom preferences locally via `localStorage`.

---

## How to Install / Sideload

To use this add-in in **Excel for Web**:

1. Download or copy the [`manifest.xml`](manifest.xml) file from this repository.
2. Open **Excel for Web** in your browser.
3. Go to **Insert** > **Add-ins** > **Upload My Add-in**.
4. Select the `manifest.xml` file.
5. Click **Upload**. The **Active Cell** button will appear under the **Home** tab!

---

## Built With

* [Office.js API](https://learn.microsoft.com/en-us/office/dev/add-ins/) - Official JavaScript API for Microsoft Office Add-ins.
* [Node.js](https://nodejs.org/) & [Webpack](https://webpack.js.org/) - Module bundler and build process.
* [GitHub Pages](https://pages.github.com/) - Live HTTPS hosting service.

---
