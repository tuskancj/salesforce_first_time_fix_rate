# Background

In field service organizations, it is important to understand repair metrics to make strategic management decisions.  First-Time-Fix (FTF) Rate is often used to help gauge the quality level of a field service engineer (FSE).  Mean Time Between Failure (MTBF) is often used to understand how frequently particular Assets fail.  Together, they provide an overview of field operations and allow for quantitative measures of performance for both machine and engineer.  Calculating such measures often can be complicated depending on an organization's record management system and budget.  

# Overview

This GUI attempts to merge various Salesforce field service reports in order to obtain 3 things:

1. Time until next repair for every closed Work Order - from this a 45-day first-time-fix rate can be determined
2. Mean Time Between Failure for Assets
3. Mean Time Between Failure for Asset Service Contract Periods

Salesforce for field service is often laid out in a Case -> Work Order -> Service Appointment hierarchy.  Cases are handled by a technical support department, leaving Work Orders and Service Appointments to be handled by the field engineering team and/or dispatch team.  Repairs are found on the the Work Order level, however, sometimes Case information is needed to validate that a Repair qualify for FTF Rate.

### Input reports:
1. assets.csv
2. contracts.csv
3. cases.csv
4. WOs.csv
5. timesheets.csv

### Output files (within 'OUTPUT' folder):
1. TBR_raw_FSE.csv -> FTF information
2. MTBF_raw_Asset.csv -> MTBF per Asset
3. MTBF_raw_Contracts.csv -> MTBF per Asset Service Contract Period
4. log_main
5. log_same_case -> situations where a Repair should not go against an FSE's FTF rate.  (more information below)

# Further Notes About the Script
*  The GUI will provide general guidelines regarding Salesforce report configuration
*  Links to reports will need to be configured for specific needs
*  Reports should be exported as '.csv' file, encoding = 'UTF-8'
*  Data is read from the '.csv' file and stored into dictionaries of Classes
  * The script will need to be referenced in order to match up column headers with Class variables

### Validate your Organization's Timestamps
* Garbage in, Garbage out.  Ensure timestamps are accurate if no stop-gaps are in-place
* While pulling data, the script gathers the range of data by storing the earliest timestamp and the most recent timestamp.  Timeframes will reference the most recent timestamp as 'current day'

### Time Between Repair (TBR)
* Represented in calendar days
* TBR is generated for every Work Order, regardless of type or FSE owner.  Filter for Service/Repair Work Order (WO) to get true FTF Rate.  Filter by PM or Installation WO to understand how often a return visit is necessary for a repair after these procedures.
  * TBR can be negative if one WO engulfs another via onsite timestamps.  WOs of this nature are given a FTF rate of ‘NA’.  They typically reflect a problem with FSE timestamp/documentation practices.
* 45-day FTF rate is labeled True/False/NA and determined from the TBR
  * NA if the WO has not been given a 45-day opportunity yet or for other obvious reasons
  * True if there has been at least 45 days without a Repair WO
  * False if there has been less than 45 days before a Repair WO
* Asset WOs are sorted by FINAL onsite timestamp before TBR is calculated.  This is due to the potential for auto-Case/WO/Service Appointment creation such as PM's that give an FSE a large window to schedule & complete
* Both multi-WO Cases and Parent Cases are taken into account when a Repair should not negatively affect an FSE's FTF Rate
  * e.g. a repair is necessary during a PM and logged in a different Repair Case that references the PM Case
  * e.g. a repair is necessary during a PM and logged in a 2nd WO under the PM Case

### Mean Time Between Failure (MTBF) for Assets
* represented in calendar days
* If Asset Installation date pre-dates the earliest timestamp within the Salesforce report data:
  * MTBF = [Most Recent Timestamp - Oldest Timestamp]/(# Repairs)
* Otherwise:
  * MTBF = [Most Recent Timestamp - Install Date]/(# Repairs)

### Mean Time Between Failure (MTBF) for Service Contracts
* MTBF = [Contract End Date - Contract Start Date]/(# Repairs)

### Service Territories
* Throughout this script, there is currently no good way to understand exact Service Territory for each WO.  Therefore, it is sometimes extrapolated based on context information.  It should be considered an estimate.
