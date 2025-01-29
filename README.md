---

# WarrantyBMW

## Project Overview

**WarrantyBMW** is a large-scale project aimed at automating the processing of BMW warranty claims and technical campaigns. The project consists of multiple modules, including dealer warranty processing workstations and technical campaign automation tools. Проект разрабатывался для внутреннего использования с русскоязычным интерфейсом. The repository will be populated progressively as the codebase is cleaned and translated.

---

## Repository Structure

- **/bmwActionTbl/** – Automation module for handling BMW **Technical Campaigns**.

  - `bmwActionTbl.mdb` – Supporting database for processing technical campaigns.
  - `/Modules/` – VBA modules exported from the database.
  - `/Forms/` – Screenshots of key forms.
  - `README.md` – Description of the database structure, key features, and list of included modules.

- **/WarrantyWorkPlaceN/** – Workstation for specialists handling dealer warranty claims (Ukraine).

  - `WarrantyWkstN.mdb` – Primary database file.
  - `/Modules/` – Exported VBA modules (only completed ones, e.g., `dgPathGlgl`, `frFileSearch`).
  - `/Forms/` – Screenshots of relevant forms.
  - `README.md` – Description of features, forms, and included modules.

- **/WarrantyWorkPlaceM/** – Workstation for specialists handling dealer warranty claims (Kyiv).

  - `WarrantyWkstM.mdb` – Primary database file.
  - `/Modules/` – Exported VBA modules (only completed ones, e.g., `dgPathGlgl`, `frFileSearch`).
  - `/Forms/` – Screenshots of relevant forms.
  - `README.md` – Description of features, forms, and included modules.

- **/WarrantyWorkPlaceA/** – Workstation for the warranty engineer - database administrator.

  - `WarrantyWkstA.mdb` – Primary database file.
  - `/Modules/` – Exported VBA modules (only completed ones, e.g., `dgPathGlgl`, `frFileSearch`).
  - `/Forms/` – Screenshots of relevant forms.
  - `README.md` – Description of features, forms, and included modules.

- **/sharedLibrary/** – Common library used across multiple projects.

  - `README.md` – Documentation explaining the purpose of `libOB.mda`. The library is still in development and will be included once ready.

- **README.md (Main)** – This document, providing an overview of all projects, their structure, and setup instructions.

---

## Current Status

At present, only the **TechnicalCampaigns** module has been uploaded, and partial infrastructure has been created for **WarrantyWorkPlaceN**. The repository will be updated progressively as more components are prepared.

### **TechnicalCampaigns Module**

The **TechnicalCampaigns** module automates BMW **Technical Campaigns**, handling:

- **Customer notifications.**
- **Parts ordering and delivery tracking.**
- **Campaign initiation and processing.**

For detailed documentation, refer to `TechnicalCampaigns/README.md`.

### **Warranty Processing Workstations**

The main warranty processing database is split into three workstations:

- **WarrantyWorkPlaceN** – Handling dealer warranty claims (Ukraine).
- **WarrantyWorkPlaceM** – Handling dealer warranty claims (Kyiv).
- **WarrantyWorkPlaceA** – Lead administrator workstation for managing the database.

These workstations are still in progress and will be added once modules and forms are ready.

---

## Setup Instructions

### **Requirements**

- **Microsoft Access** (MS Office XP or later recommended)
- **Microsoft Office VBA support enabled**
- Ensure necessary **VBA references** are enabled (Alt+F11 > Tools > References):
  - Microsoft DAO 3.6 Object Library
  - Microsoft ActiveX Data Objects (ADO)
  - Microsoft Scripting Runtime
  - Microsoft Office Object Library
  - Microsoft Word Object Library
  - Custom Library: `libOB.mda` (when available)

### **How to Open the Project**

1. Download the required **.mdb** file from the repository.
2. Place it in a **trusted location** in MS Access settings.
3. Open the database in **MS Access**.
4. Run the main form **frTA** (for Technical Campaigns) or relevant forms for Warranty Processing.

For module-specific instructions, refer to the README files inside each subfolder.

---

## Future Plans

- **Complete the translation of all modules and UI elements.**
- **Gradually upload more components** of the warranty processing system.
- **Provide detailed documentation** for each module.

This repository is a **work in progress**, and updates will be made as components become ready.

For more details, refer to the respective `README.md` files in each subdirectory.
