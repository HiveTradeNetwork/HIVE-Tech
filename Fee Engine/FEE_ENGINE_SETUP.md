# HIVE Fee Engine — Setup & Run Guide

**Who this is for:** Jake. Run this once per month after all supplier invoices for the period are entered in Airtable.

---

## One-time setup (do this once)

### Step 1 — Install Python

1. Go to https://python.org/downloads
2. Download the latest Python 3.x installer for Windows
3. Run the installer — **tick "Add Python to PATH"** before clicking Install
4. Verify: open Command Prompt and run `python --version` — should show a version number

### Step 2 — Install dependencies

Open Command Prompt in the Fee Engine folder:

```
cd "G:\My Drive\HIVE CLAUDE ACCESS\Outputs\HIVE-Tech\Fee Engine"
pip install -r requirements.txt
```

### Step 3 — Set your Airtable API key

Your Airtable personal access token needs to be set as an environment variable. This only needs to be done once per machine (or once per session if you use the temporary method).

**Permanent (recommended):**

1. Search Windows for "Edit the system environment variables"
2. Click Environment Variables
3. Under User variables, click New
4. Variable name: `AIRTABLE_API_KEY`
5. Variable value: your Airtable personal access token
6. Click OK, then restart Command Prompt

**To get your Airtable token:**
- Go to airtable.com → click your profile icon → Developer Hub → Personal access tokens
- Create a token with `data.records:read`, `data.records:write`, `schema.bases:read` scopes
- Add HIVE Fee Engine base as the resource

### Step 4 — Add bank details to the script

Before you ever send an invoice, open `fee_engine.py` in Notepad and update the HIVE dict at the top:

```python
HIVE = {
    ...
    "bank_name":     "TBC — add before sending invoices",   # <-- replace
    "bsb":           "TBC",                                  # <-- replace
    "account":       "TBC",                                  # <-- replace
    ...
}
```

---

## Running the engine (monthly)

### Before you run

- Confirm all supplier invoices for the period are entered in Airtable Transactions table
- Confirm each transaction record has the Fee Period field set (e.g. `April 2026`)
- Confirm each Supply Partner linked to those transactions is set to Active status

### Run command

Open Command Prompt in the Fee Engine folder and run:

```
python fee_engine.py "April 2026"
```

Replace `April 2026` with the actual fee period. Must match exactly what is in Airtable.

### What it does

1. Pulls all transactions for that fee period from Airtable
2. Groups by Supply Partner
3. Calculates 2% service fee on total Amount Ex GST, then 10% GST on that fee
4. Creates a Fee Calculation record in Airtable (status: Draft)
5. Generates one branded Excel invoice per Supply Partner
6. Saves invoices to: `Fee Engine\invoices\HIVE_Invoice_HIVEYYMM-001_PartnerName.xlsx`

### Example output

```
HIVE Fee Engine
Fee period : April 2026
Run date   : 01/05/2026
Due date   : 15/05/2026
==================================================
Fetching transactions for 'April 2026'...
  Found 12 transaction(s).
Fetching active Supply Partners...
  Found 3 active partner(s).
Calculating fees...

Results for April 2026:
--------------------------------------------------

  Acme Hardware Pty Ltd
    Transactions :  5
    Total ex GST : $    18,500.00
    Fee (2%)     : $       370.00
    GST on fee   : $        37.00
    Total due    : $       407.00
    Invoice no.  : HIVE-202604-001
    Saved        : invoices/HIVE_Invoice_HIVE-202604-001_Acme_Hardware_Pty_Ltd.xlsx

==================================================
Done. 3 invoice(s) generated in /invoices/
ACTION: Add bank details to HIVE dict in fee_engine.py before sending.
```

### After running

1. Open each invoice in Excel and check figures look right
2. Update the Fee Calculation records in Airtable from Draft to Sent once invoices are emailed
3. File invoices in Google Drive under the relevant period folder

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| `AIRTABLE_API_KEY environment variable not set` | Set the env variable (Step 3 above) and restart Command Prompt |
| `No transactions found for this period` | Check Fee Period field in Airtable matches the command exactly (case sensitive) |
| `ModuleNotFoundError: No module named 'pyairtable'` | Run `pip install -r requirements.txt` again |
| Invoice has $0.00 fee | Amount Ex GST is blank on one or more transaction records in Airtable — fill it in and re-run |

---

*HIVE Trade Network Pty Ltd | fee_engine.py v1.0 | April 2026*
