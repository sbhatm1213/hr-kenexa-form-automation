### hr-kenexa-form-automation
### HR Operations
### Offer Form Automation (IBM Kenexa)

### Overview

Python + Selenium automation to create and submit **Offer Forms** in the IBM Kenexa recruitment system using data from Excel.

Built to eliminate manual data entry and speed up offer creation for multiple candidates.

---

### What it does

* Reads candidate and offer details from Excel (.xlsm)
* Opens HR Links Portal → IBM Kenexa
* Searches by **Requisition ID**
* Updates candidate **HR Status** to *Prepare Offer* (if required)
* Creates **Offer Form** (if not already completed)
* Auto-fills form fields:

  * Job details, recruiter, workspace
  * Hire/Transfer type, shift, virtual status
  * Offer dates and start date
  * Compensation, CTC, currency, frequency
  * Benefits, relocation, allowances, bonuses
  * Versant scores and other required fields
* Marks status in Excel (IN PROGRESS / Completed handling)

---

### Tech Stack

Python, Selenium, pandas, xlwings

---

### How to run

1. Update:

   * ChromeDriver path
   * Excel file path
   * Portal URL (internal access required)
2. Install dependencies
3. Run:

```bash
python script.py
```

---

### Notes

* Designed for internal enterprise HR/IBM Kenexa system
* Requires portal access and network connectivity
* UI element selectors may need updates if the application changes
* Handles existing-form checks to avoid duplicates

---

**Year:** ~2018–2020
Reference project demonstrating large-scale HR offer processing automation.
