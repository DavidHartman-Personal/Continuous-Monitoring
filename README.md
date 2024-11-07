Continuous Monitoring git repo.

* [ ]  Add gitignore
* [ ]  Update readme.md

## cmdb

System Inventory related scripts, etc.


| Name                                                                                                  | Inputs/Outputs                                                                                            | Description                                                                                                                                                                                                                   |
| :---------------------------------------------------------------------------------------------------- | :-------------------------------------------------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [create-cmdb-masters-from-fedramp-inventory.py](./cmdb/create-cmdb-masters-from-fedramp-inventory.py) | **++Inputs++**: Excel CMDB master file<br>**++Outputs++**: JSON file containing system inventory results. | This module will load the current OPEN CMDBS into a dictionary object and then create a JSON file containing system inventory information.<br>This JSON file is read by the [current_cmdb.py](./cmdb/current_cmdb.py) module. |
| [create-cmdb-output-frm-inventory.py](./cmdb/create-cmdb-output-frm-inventory.py)                     | **++Inputs++**: Excel CMDB master file<br>**++Outputs++**: JSON file containing system inventory results. | This module will load the current OPEN CMDBS into a dictionary object and then create a JSON file containing system inventory information.<br>This JSON file is read by the [current_cmdb.py](cmdb\current_cmdb.py) module.   |
| [create-cmdb-masters-from-fedramp-inventory.py](cmdb/create-cmdb-masters-from-fedramp-inventory.py)   | Inputs: Excel CMDB master file<br>Outputs: JSON file containing system inventory results.                 | This module will load the current OPEN CMDBS into a dictionary object and then create a JSON file containing system inventory information.<br>This JSON file is read by the [current_cmdb.py](cmdb\current_cmdb.py) module.   |
| [create-cmdb-masters-from-fedramp-inventory.py](cmdb/create-cmdb-masters-from-fedramp-inventory.py)   | Inputs: Excel CMDB master file<br>Outputs: JSON file containing system inventory results.                 | This module will load the current OPEN CMDBS into a dictionary object and then create a JSON file containing system inventory information.<br>This JSON file is read by the [current_cmdb.py](cmdb\current_cmdb.py) module.   |

## conf

Configuration files.

## docs

Documentation

## model

objecct model definitions

## poam

POAM related utilities

## util

General utilities

## vdrf

Deviation requests

## vulnerabilitiy_scans

Scripts for processing nessus/tenable vulnerability scan data.

[current_vuln_scan.py](vulnerability_scans/current_vuln_scan.py): Module
that can be imported to accesss CMDB JSON data/file as well as functions
for accessing the CMDB data.

* Inputs: CMDB JSON File, set as a variable/constant
* Variables/Contants
  * ENVIRONMENT: L5 or FRM
  * VULN_SCAN_FILE_JSON: JSON File containing vuln scan results data.
* Outputs/Results: None

[create_vuln_scan_results_from_spreadsheet.py](vulnerability_scans/create_vuln_scan_results_from_spreadsheet.py):
Reads the "All Items" sheet from the excel spreadsheet containing
vulnerability scan data.

* Inputs: Excel spreadsheet with All Items worksheet.
* Variables/Contants
  * SCAN_RESULTS_SHEET_NAME: "All Items"
  * VULN_SCAN_OUTPUT_FILE_JSON: Output JSON file name from Vuln Scan
    data.
* Outputs/Results: JSON file containing vulnerability scan data.

[process_csv_nessus_results.py](vulnerability_scans/process_csv_nessus_results.py):
Reads a .csv file containing nessus scan data and prints/extracts
details as needed.

* Inputs: CSV file containing nessus scan results
* Outputs/Results: N/A

[process_nessue_xml_file.py](vulnerability_scans/process_nessue_xml_file.py):
process raw .nesssus file and extract information related to the scan.
This includes items such as plugins, targets, policy used, etc.

* Inputs: Raw .nessus file
* Outputs/Results: Prints summary information about the scan result.
