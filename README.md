<!--
  README for: Outlook Email Extractor
-->

<h1 align="center">📧 Outlook Email Extractor</h1>

<p align="center">
  A fast, modern, Windows desktop app to search Outlook, export to Excel, and optionally save attachments — all with a slick GUI.
</p>

<p align="center">
  <a href="https://github.com/griffgriff5000/Spotlight-on-Outlook/actions">
    <img alt="Build" src="https://img.shields.io/github/actions/workflow/status/griffgriff5000/Spotlight-on-Outlook/build.yml?branch=main&label=Build&logo=githubactions">
  </a>
  <a href="https://github.com/griffgriff5000/Spotlight-on-Outlook/releases/latest">
    <img alt="Download" src="https://img.shields.io/github/v/release/griffgriff5000/Spotlight-on-Outlook?display_name=release&sort=semver&label=Latest%20Release&logo=github">
  </a>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white">
  <img alt="OS" src="https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows&logoColor=white">
</p>

---

## ⚡ TL;DR

1. **Download** the ZIP from **[Releases → Latest](https://github.com/griffgriff5000/Spotlight-on-Outlook/releases/latest)**  
2. **Unzip** it anywhere (e.g. Desktop)  
3. **Run** `Outlook Email Extractor.exe`  
4. Click **Connect / Load Accounts** → set filters → **Export to Excel** ✅

> 🛡️ **Privacy-first**: Everything runs locally. No data leaves your machine.

---

## ✨ Highlights

- 🌓 **Modern UI** with Light/Dark theme (`sv-ttk` if present)
- 📅 **UK date pickers** (`tkcalendar` if present) or plain text fallback (DD-MM-YYYY)
- 🧵 **Fast scanning** (MAPI/COM), sortable, subfolder recursion
- 🧲 **Smart filters**: date range, unread, has/no attachments, subject/from contains, max items
- 🧰 **Attachment filter** by type (PDF/Images/Excel/Docs/PPT/Archives/Custom)
- 🧽 **Exclude inline images** (signature clutter) toggle
- 💾 **Exports to Excel** with **Filters** & **Emails** sheets (+ **Attachments** sheet if saved)
- 🗂️ **Per-email attachment folders** + auto hyperlinks back from Excel
- 🧠 **Auto-named outputs**:  
  - Excel → `Emails DD-MM-YYYY - DD-MM-YYYY.xlsx`  
  - Attachments → `Attachments DD-MM-YYYY - DD-MM-YYYY`

---

## 🖥️ How it works (in 30 seconds)

1. **Connect** → Loads your Outlook “stores” (mailboxes).  
2. **Scope** → Scan entire account or a specific `Folder/Subfolder`.  
3. **Filter** → Dates, unread/read, attachments, subject/from, max items.  
4. **Types** → If “Only with attachments”, restrict file types and hide inline images.  
5. **Save location** → Pick a base folder; names auto-update as dates change.  
6. **Go** → **Preview Count** (dry-run) or **Export to Excel** (writes file + optional attachments).

---

## 📦 What you get

### Excel workbook

| Sheet       | What’s inside                                                                                           |
|-------------|----------------------------------------------------------------------------------------------------------|
| **Emails**  | `ReceivedTime`, `Subject`, `SenderName`, `SenderEmail`, `To/CC/BCC`, `Categories`, `Unread`, `HasAttachments`, counts, `FolderPath`, `ConversationID`, `EntryID`, optional `BodyPreview`, plus attachment columns when enabled |
| **Filters** | Human-readable snapshot of the exact run config (store, folder, dates, flags, file types, timestamp)     |
| **Attachments** *(optional)* | `ReceivedTime`, `Subject`, `SenderEmail`, `AttachmentName`, `AttachmentPath`, Excel `Link` to open file |

### Attachments on disk (optional)


