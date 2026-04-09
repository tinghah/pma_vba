# Enterprise HR Data Integration Suite (Excel to PostgreSQL)

A robust, enterprise-grade solution for automating the Extraction, Transformation, and Loading (ETL) of HR Attendance data from disparate Excel workbooks into a centralized PostgreSQL database.

## 🚀 Overview

This suite consists of three modular VBA engines designed to handle the full data lifecycle: from the initial import of messy vendor/attendance files to a structured database upload with transactional integrity.

### 🛠 Core Modules

#### 1. Data Intake Engine (`import.vba`)
*   **Purpose**: Strategic ingestion of raw attendance workbooks.
*   **Key Features**:
    *   **Automated Discovery**: Scans selected workbooks for relevant sheets (Leave, OT, Late) using keyword-based heuristics.
    *   **Data Integrity**: Implements "Leading Zero Protection" for critical ID and Code columns to prevent truncation.
    *   **Pre-Cleaning**: Automatically removes non-essential columns and standardizes naming conventions upon import.
    *   **Auto-Route**: Seamlessly returns the user to the MAIN dashboard upon completion.

#### 2. Transformation & Normalization Engine (`fix.vba`)
*   **Purpose**: Structural alignment of imported data for relational database compatibility.
*   **Key Features**:
    *   **Dynamic Mapping**: Handles varied source formats for Overtime and Late records using fuzzy header matching.
    *   **Strict Logic (Leave)**: Utilizes color-code mapping to distribute complex leave types into granular database fields.
    *   **Postgres Preparation**: Formats dates to `YYYY-MM-DD` and structures data into a flattened "result" format optimized for bulk insertion.

#### 3. PostgreSQL Database Manager (`upload.vba`)
*   **Purpose**: Secure and efficient data transmission to the PostgreSQL backend.
*   **Key Features**:
    *   **Transactional UPSERT**: Utilizes `ON CONFLICT (id_no, date)` logic, allowing for idempotent daily updates. This ensures that re-running an upload overwrites existing records rather than creating duplicates.
    *   **SQL Sanitization**: Automatically handles character escaping and decimal formatting to prevent SQL errors.
    *   **Schema Awareness**: Dynamically maps Excel headers to a pre-defined database schema (`leave`, `ot`, `late` tables).
    *   **Connectivity**: Leverages high-performance ADODB connections via the PostgreSQL Unicode driver.

---

## 📂 Implementation & Security

To maintain enterprise security standards, the implementation scripts and database configuration files are stored in the `/scripts/` directory, which is ignored by version control to protect sensitive connection strings and intellectual property.

### How to use in other projects:
1.  **Modularization**: Each script can be imported into any Excel `.xlsm` project via the VBA Editor (VBE) as a separate module.
2.  **Configuration**: Update the constant variables (`DB_HOST`, `DB_NAME`, etc.) in `upload.vba` to point to your specific environment.
3.  **Extensibility**: The `GetDBColumnName` mapping function in the Upload Manager can be easily extended to support additional database tables or custom fields.

## ⚙️ Requirements
*   **Microsoft Excel**: 2016 or newer.
*   **PostgreSQL**: 12+ (compatible with standard relational schema).
- **ODBC Driver**: [PostgreSQL Unicode Driver](https://www.postgresql.org/ftp/odbc/versions/msi/) must be installed on the host machine.
- **Reference Libraries**: `Microsoft ActiveX Data Objects 6.1 Library` (optional, uses late binding).

---
*Created for automated HR workflow optimization.*
