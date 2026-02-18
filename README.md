**Hanshow Integration Service
**


Excel → Hanshow Integration API v2.0

Automated Windows service for retail price updates.



**Overview**



This project implements a production-grade integration service that:



Monitors a folder for Excel interface files



Converts Excel rows into Hanshow API payloads



Sends data in batches of 1000 items



Logs all activity with daily rotation



Retains logs for 30 days



Deletes processed files older than 30 days



Runs automatically at Windows startup



The service is designed to operate unattended.



**Architecture**

Excel File Drop

&nbsp;       ↓

Folder Monitor (5 min cycle)

&nbsp;       ↓

Excel → JSON Mapping

&nbsp;       ↓

1000 Item Batch Push

&nbsp;       ↓

Move to /done

&nbsp;       ↓

Retention Cleanup (30 days)



**Folder Structure
**


Watched folder:



C:\\wamp64\\www\\csv\\INTERFACE\_FILE





Processed files:



C:\\wamp64\\www\\csv\\INTERFACE\_FILE\\done





Log file:



hanshow\_integration.log



**Excel Requirements
**


The Excel file must contain the following columns:



ItemID



ItemName



ItemNumber



PrimaryUpc



UnitPrice



**Mapping rules:
**


sku = ItemID



itemName = ItemName



ean = ItemNumber + PrimaryUpc (comma separated if both exist)



price1 = UnitPrice



IIS\_COMMAND = COMPLETE\_UPDATE



Batching



1000 items per API request



Each batch uses a unique batch number



File is moved only if all batches succeed



**Logging**



Daily rotating logs



30-day retention



Console + file logging



Automatic deletion of old logs



Processed File Retention



Files in /done older than 30 days are automatically deleted



Retention based on file modified time



Requirements



Python 3.10+



Dependencies:



requests

pandas

openpyxl





Install with:



pip install -r requirements.txt



Configuration



Sensitive credentials must be stored as environment variables:



HANSHOW\_BASIC\_USER

HANSHOW\_BASIC\_PASSWORD





Set in PowerShell:



setx HANSHOW\_BASIC\_USER "your\_user"

setx HANSHOW\_BASIC\_PASSWORD "your\_password"





Restart shell after setting.



Running Manually

python excel\_to\_hanshow\_v5.py



Windows Service Installation



Run the PowerShell installer as Administrator:



scripts\\install\_service.ps1





The service:



Runs at system startup



Runs as SYSTEM



Restarts on failure



Operates without user login



Failure Handling



If any batch fails, file is not moved



Errors are logged



Service continues monitoring next cycle



Versioning



Current Version: v5

Features included:



1000 batch support



5-minute polling



30-day log retention



30-day processed file cleanup



Windows startup automation



Security Notes



No credentials should be committed to Git



.gitignore excludes logs and runtime Excel files



Environment variables must be configured on deployment machine



Maintenance



Recommended periodic checks:



Verify log rotation



Confirm cleanup of old files



Confirm service status via Task Scheduler



License



Internal Smart Label Solutions integration project.

