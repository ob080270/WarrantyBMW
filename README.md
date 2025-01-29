---

# WarrantyBMW

## Project Description

This repository demonstrates database functionalities for managing BMW warranty claims. The project includes various modules and data for automating the processing of BMW warranty claims and technical campaigns. The warranty claims processing module is not yet uploaded but will be added later.

---

## Repository Structure

- **TechnicalCampaigns/**  

  The main folder containing database files and modules responsible for processing BMW Technical Campaigns.

  Includes:

  - **PickTbl.doc** – A file with images of BMW structural groups.
  - **acbMailMerge.doc** – A merge file template for generating customer letters.
  - **bmwActionTbl.mdb** – A database with a suite of modules automating the technical campaigns' processing.
  - **qsLetterSource.doc** – An auxiliary file for mail merge operations.
  - **tmpNewTA.dot** – A template letter for customers.

  - **Modules/**  

    Contains VBA programming modules. Each file represents an individual module exported from the database.  

    **File Descriptions:**

    - **Form_frEntryNewTA.cls** – A module for automating the processing of a new technical campaign (campaign description, input of involved vehicles).
    - **Form_frInfScan.cls** – A module for controlling customer notifications about technical campaigns.
    - **Form_frTA.cls** – A module automating the main form `frTA`.
    - **Form_sf2OrdItem.cls** – A module for automating the processing of part orders required for completing the technical campaign.
    - **Form_sfDlv.cls** – A module for automating the processing of part deliveries required for completing the technical campaign.
    - **Form_sfLetters.cls** – A module for automating the creation of customer notification letters about the technical campaign.
    - **Form_sfOrders.cls** – A module for automating the subform of part orders required for technical campaigns.
    - **Form_sfPartSum.cls** – A module for processing the calculation of part quantities (for different variations).
    - **Form_sfParts_Veh.cls** – A module for processing navigation through records of vehicles involved in the technical campaign.
    - **Form_sqParts_Parts.cls** – A module for automating the input of required parts for vehicles involved in the technical campaign.
    - **MyFunktons.bas** – Global procedures and functions of the project.
    - **basWord.bas** – A module for automating the conversion of database query results into a Word table.

---

## Key Features

- Automation of procedures for launching a new Technical Campaign.
- Automation of customer letter creation.
- Automation of part ordering and deliveries required for the technical campaign.
- Control of customer notifications about technical campaigns.

---

## Installation and Launch

Ensure that the Word files are located in the same folder as `bmwActionTbl.mdb`.

List of referenced libraries (Alt+F11 > Tools > References):

---

This file is now ready for use on GitHub. Let me know if further changes or additions are needed!
