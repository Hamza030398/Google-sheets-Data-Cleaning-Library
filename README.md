# Google-sheets-Data-Cleaning-Library

🧹 Google Sheets Data Cleaning Library

A modular Google Apps Script library that automates post-ingestion data cleanup, normalization, and regex-based analysis directly inside Google Sheets.

🚀 Features

🔄 Auto Clean — Trim spaces, remove duplicates, normalize dates.

🗂️ Smart Backup System — Auto backup sheet versions before transformations.

📅 Date & Time Formatter — Automatically formats messy timestamps.

🔍 Regex Pattern Tool — Sidebar for pattern-based search, extract, replace, or test.

🧱 Preview Mode — Review regex changes before applying.

🪶 UI Integrated Menu — Appears directly in Google Sheets toolbar.

🧠 Who It Helps

Built for data analysts, researchers, and automation engineers working with:

CSV imports from SaaS platforms.

Data ingested from APIs or ETL workflows into Sheets

Teams performing quick cleaning before BI / Dashboard modeling

🧰 How to Use

Add this library to your Google Apps Script project:

Script ID: (add after you publish it as a library)

In your bound Google Sheet script editor, include this proxy example:

function onOpen() {
  DataCleaningTools.onOpen();
}


From the Sheets toolbar, open
→ 🧾 Receipt Tools → 🧩 Data Tools → Regex Sidebar

Run regex or cleanup actions, preview results, and log actions.

🧪 Example Regex

Find dimensional strings like "14 7/8\" x 11\"":

\d+\s?\d*(?:\/\d+)?"\s*x\s*\d+\s?\d*(?:\/\d+)?"
