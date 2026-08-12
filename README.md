
# DORA Compliance Tracker & Risk Management

This repository contains tools to process EU DORA compliance data and generate risk trackers.

> **Developer Note:** 
> Looking for a raw, lightweight, and 100% plain Python/Pandas implementation without any reporting bloat or heavy formatting? 
> Check out `dora_tracker_fast.py`. It focuses purely on speed, vectorization, and clean data extraction with zero external styling dependencies.


Just a quick Python script I put together to help track compliance for the EU Digital Operational Resilience Act (DORA). 

Instead of doing manual formatting in Excel every time a new assessment comes in, this takes a structured dataset of DORA controls, calculates the key metrics, and builds an executive dashboard automatically.

## What it does

- Builds a summary tab with KPIs (Overall Compliance Rate, gaps vs compliant count)
- Generates native Excel charts (pie chart for status breakdown and column chart by domain)
- Outputs a master list with status badges (green/yellow/red) for all 38 controls
- Filters out urgent remediation gaps into a separate tab so team leads can focus on Critical/High risk failures first

## DORA Domains Covered

1. ICT Risk Management Framework
2. ICT Incident Management & Reporting
3. Digital Operational Resilience Testing (TLPT, vuln scanning)
4. Third-Party Risk Management (TPRM & vendor risks)
5. Information & Threat Intelligence Sharing

## How to run it

First install the requirements:

`pip install pandas xlsxwriter`

Then run the script:

`python dora_compliance_tracker.py`

It will generate `DORA_Compliance_Executive_Tracker.xlsx` in the same directory.

## Tech stack

- Python 3
- pandas for data manipulation
- xlsxwriter for Excel formatting & charts
