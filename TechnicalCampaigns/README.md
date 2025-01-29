# TechnicalCampaigns

## Project Description

This module is part of the **WarrantyBMW** repository and is responsible for processing **BMW Technical Campaigns**. It includes various modules and data structures that automate the processing of Technical Campaigns, such as customer notifications, part orders, and deliveries.

---

## Repository Structure

- **TechnicalCampaigns/**  

  The main folder containing database files and modules responsible for processing **BMW Technical Campaigns**.

  Includes:

  - **PickTbl.doc** – A file with images of BMW structural groups.
  - **acbMailMerge.doc** – A merge file template for generating customer letters.
  - **bmwActionTbl.mdb** – A supporting database that automates technical campaign processing (not the main database).
  - **qsLetterSource.doc** – An auxiliary file for mail merge operations.
  - **tmpNewTA.dot** – A template letter for notifying customers about a technical campaign.

  - **Modules/**  

    Contains VBA programming modules. Each file represents an individual module exported from the database.

    **File Descriptions:**

    - **Form_frEntryNewTA.cls** – Automates the processing of a new technical campaign (campaign description, input of involved vehicles).
    - **Form_frInfScan.cls** – Controls customer notifications about technical campaigns.
    - **Form_frTA.cls** – Automates the main form `frTA`.
    - **Form_sf2OrdItem.cls** – Automates the processing of part orders required for completing the technical campaign.
    - **Form_sfDlv.cls** – Automates the processing of part deliveries required for completing the technical campaign.
    - **Form_sfLetters.cls** – Automates the creation of customer notification letters about the technical campaign.
    - **Form_sfOrders.cls** – Automates the subform of part orders required for technical campaigns.
    - **Form_sfPartSum.cls** – Processes the calculation of part quantities (for different variations).
    - **Form_sfParts_Veh.cls** – Processes navigation through records of vehicles involved in the technical campaign.
    - **Form_sqParts_Parts.cls** – Automates the input of required parts for vehicles involved in the technical campaign.
    - **MyFunktons.bas** – Global procedures and functions of the project.
    - **basWord.bas** – Automates the conversion of database query results into a Word table.
    - **BSI_Class.cls** – Checks whether BSI (BMW Service Inclusive) reports need to be generated.
    - **CClipboard_Class.cls** – Manages clipboard operations.
    - **DivQ_Class.cls** – Handles discrepancies in approved part quantities.
    - **NewCredit_Class.cls** – Imports and processes new credit data.

---

## Key Features

- **Automation of technical campaign initiation.**
- **Automated customer letter generation.**
- **Management of part ordering and deliveries.**
- **Customer notification tracking and control.**
- **Calculation of required part quantities.**
- **Discrepancy handling for approved part quantities.**
- **Processing new credit data related to technical campaigns.**

---

## Installation and Launch

### **Requirements**
Ensure that the Word files are located in the same folder as `bmwActionTbl.mdb`.

### **Referenced Libraries**
To run the modules properly, enable the following libraries in MS Access (Alt+F11 > Tools > References):

- Microsoft DAO 3.6 Object Library
- Microsoft ActiveX Data Objects (ADO)
- Microsoft Scripting Runtime
- Microsoft Office Object Library
- Microsoft Word Object Library
- Custom Library: `libOB.mda` (if applicable)

### **Steps to Launch**
1. Open `bmwActionTbl.mdb` in **MS Access**.
2. Ensure that all referenced libraries are enabled.
3. Run the main form `frTA` to initiate the processing of technical campaigns.
4. Use supporting forms and modules to manage customer notifications, part orders, and deliveries.
5. Verify generated reports and ensure data consistency across different technical campaign processes.


