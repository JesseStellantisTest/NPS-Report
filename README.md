# NPS-Report
NPS Report Python Shell

## Run the report

Use ISO 8601 dates (`YYYY-MM-DD`) on the command line:

```powershell
python manpssw.py --start-date 2026-11-03 --end-date 2026-11-30
```

You may omit either or both arguments. The script will prompt for missing dates and accept Enter to use the default shown in the prompt.

## Email options

The default recipient is `anthony.malgapo@stellantis.com`. Override it and optionally add CC recipients:

```powershell
python manpssw.py --recipient anthony.malgapo@stellantis.com --cc first@stellantis.com;second@stellantis.com
```

Use `--draft` to open the email in Outlook without sending it:

```powershell
python manpssw.py --start-date 2026-11-03 --end-date 2026-11-30 --recipient anthony.malgapo@stellantis.com --draft
```

## Python setup

A workspace virtual environment is stored in `.venv`. Run the report with that interpreter:

```powershell
& .\.venv\Scripts\python.exe manpssw.py
```

When the script starts, enter your Snowflake user when prompted, for example:

```text
Snowflake user (for example, TA48135@stellantis.com): TA48135@stellantis.com
```

You can also provide the user and reporting dates on the command line:

```powershell
& .\.venv\Scripts\python.exe manpssw.py --snowflake-user TA48135@stellantis.com --start-date 2026-11-03 --end-date 2026-11-30
```

To recreate or update the dependencies:

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```
