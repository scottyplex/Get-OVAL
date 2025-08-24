# Red Hat Security Advisory Reporter

This PowerShell script automates the process of gathering and organizing detailed security data from the Red Hat security advisory API. By simply providing a CVE number, the script will query Red Hat's database and compile all relevant information into a single, organized Excel file.

### Features

* **API Querying:** Automatically fetches data for a specific CVE from Red Hat's OVAL API.

* **Data Consolidation:** Combines multiple data points—such as notes, mitigation, affected packages, and errata—into a single, multi-tab Excel workbook for easy review.

* **Structured Output:** Organizes data into separate, clearly labeled tabs within the Excel file, making the information easy to navigate and analyze.

* **Automated Cleanup:** Cleans up intermediate CSV files after the final Excel report is generated, keeping your workspace tidy.

### Prerequisites

* A Windows host with **PowerShell 5.1 or newer**.

* **Microsoft Excel** must be installed on the host to enable the script to merge CSV files into a multi-tab workbook.

### Usage

The script is designed to be run from the command line and takes a single, mandatory parameter for the CVE you want to look up.

```
.\get-redhat-cve.ps1 -cve CVE-2023-44487
```

### Output

The script creates a directory named `logs` on your desktop. Inside, you will find a single Excel file named after the specific advisory ID (e.g., `RHSA-2024_1234.xlsx`).

The Excel file contains the following tabs:

* **notes**: All notes related to the advisory, including the description, CVSS base score, and security fixes.

* **mitigation**: Red Hat's assessment and recommended mitigation for the vulnerability.

* **packageState**: A list of applicable packages, their fix state, and associated CPEs.

* **affectedRelease**: Details on the packages being pushed, their associated CPEs, and the release date.

* **relatedCVE**: A list of other CVEs being resolved in the same advisory.

* **errata**: Information on the advisory errata, its URL, type, and description.

### Notes

* **Version:** 1.0

* **Author:** Scott B Lichty

* **Creation Date:** 11/19/2021

* **Purpose:** Initial script development to quickly gather information on RHEL vulnerabilities.
