# Background

In field service organizations, it is important to understand repair metrics to make strategic management decisions.  First-Time-Fix (FTF) Rate is often used to help gauge the quality level of a field service engineer (FSE).  Mean Time Between Failure (MTBF) is often used to understand how frequently particular Assets fail.  Together, they provide an overview of field operations and allow for quantitative measures of performance for both machine and engineer.  Calculating such measures often can be complicated depending on an organization's record management system and budget.  

# Example of Insights

'Best-of-best' FSEs complete tasks quicker than peers, while maintaining lower than average callback ratio.  The charts below plot 45-day Effective Preventive Maintenance (PM) rate and Hours per Work Order of an example best-of-best 'FSE-1' on Cytometer 'A' over time.  For reference, 45-day Effective PMs are completed without a repair callback within a 45-day window post-PM.  This particular 'FSE-1' has 10% higher than average Effective PM rate while completing PMs ~2 hours faster than the average in the Central US district which is consistent over time.
![](/images/fse1.jpg)

Compare this to 'FSE-2' with similar workload on Cytometer 'A' who has 10% lower than average Effective PM rate while taking ~2 hours longer than average, again consistent over time.
![](/images/fse2.jpg)

These are strictly quantitative insights that generate questions:

* what is FSE-1 doing different than FSE-2?
* is FSE-1 following procedures? 
  * if yes, can we pair FSE-1 with FSE-2 to mentor/transfer knowledge/increase speed?
  * if no, are the current procedures necessary to maintain instrument uptime?  
  * can we pair FSE-1 with sustaining engineering department to modify procedures?

**These are valuable insights with potential to help both the individual engineer and organization continuously improve.**  This study looks at not only PMs, but also Repair and Installation Work Orders.  It can be generated on a quarterly basis to observe trends.

# Overview

The script utilizes a GUI that attempts to merge various Salesforce field service reports.  The idea behind the GUI was to allow a more Windows-friendly interface for other non-technical personnel to utilize.  

![](/images/gui.jpg)

The script attempts to obtain 3 things:

1. Time until next repair for every closed Work Order - from this a 45-day first-time-fix rate can be determined
2. Mean Time Between Failure for Assets
3. Mean Time Between Failure for Asset Service Contract Periods

Salesforce for field service is often laid out in a Case -> Work Order -> Service Appointment hierarchy.  Cases are handled by a technical support department, leaving Work Orders and Service Appointments to be handled by the field engineering team and/or dispatch team.  Repairs are found on the the Work Order level, however, sometimes Case information is needed to validate that a Repair qualifies for FTF Rate.

### Input reports:

* assets.csv
* contracts.csv
* cases.csv
* WOs.csv
* timesheets.csv

### Output files (within 'OUTPUT' folder):

* TBR_raw_FSE.csv -> FTF information
* MTBF_raw_Asset.csv -> MTBF per Asset
* MTBF_raw_Contracts.csv -> MTBF per Asset Service Contract Period
* log_main.csv
* log_same_case.csv -> situations where a Repair should not go against an FSE's FTF rate.  (more information below)

# Further Notes About the Script

*  The GUI will provide general guidelines regarding Salesforce report configuration
*  Links to reports will need to be configured for specific needs
*  Reports should be exported as '.csv' file, encoding = 'UTF-8'
*  Data is read from the '.csv' file and stored into dictionaries of various Class values
  * The script will need to be referenced in order to match up column headers with Class variables

### Validate your Organization's Timestamps

* Garbage in, Garbage out.  Ensure timestamps are accurate if no stop-gaps are in-place
* While pulling data, the script interprets the range of Salesforce data by storing the earliest timestamp and the most recent timestamp.  Timeframes will reference the most recent timestamp as 'current day' so it's important to ensure erroneous timesheets are not in the data (e.g. timestamps in the future)

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
