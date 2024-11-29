"""This script will read or write to a spreadsheet that follows the current FedRAMP template for reporting POAMS.

TODO: Add ability to write to a poam file based on XXX

Inputs/Arguments/Settings:
- poam Spreadsheet file
- First Data Row

Outputs
- Outputs results to JSON files as well as CSV file

NOTES: The spreadsheet should not have any extra rows at the bottom

HISTORY:
 - 6/25/2020 - This script is the current version for generating the POAM in JSON Format.

"""
import logging
import os
import re
import openpyxl
import datetime
import json
from util.extract_field_info import split_server_port, get_plugin_id
import cmdb.current_cmdb as current_cmdb


# Default first data row for FedRAMP poam Template
FIRST_DATA_ROW = 6

ENVIRONMENT = "L5"
YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")
# 3 character month abbreviation
MONTH_ABBREV = datetime.datetime.now().strftime("%b")
# Month number
MONTH_NUM = datetime.datetime.now().strftime("%m")
YEAR = datetime.datetime.now().strftime("%Y")
MONTH_FOLDER = "{month_abbrev} - {month_number}".format(month_abbrev=MONTH_ABBREV,
                                                        month_number=MONTH_NUM)
SCAN_RESULTS_GENERATION_DAY = "20200705"

FEDRAMP_POAM_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\{environment}\\{month}\\".format(environment=ENVIRONMENT,
                                                                                                                           month=MONTH_FOLDER)
DEFAULT_PROTOCOL = 'TCP'

# C:\Users\dhartman\Documents\FedRAMP\Continuous Monitoring\poam\L5\04 - APR\PTC-CS -L5-APR-2020-CONMON-poam.xlsx
CURRENT_L5_POAM_SPREADSHEET_FILE = "L5-POAM-May-2020-Internal.xlsm"
POAM_RESULTS_SPREADSHEET_FULL_PATH = os.path.join(FEDRAMP_POAM_DIR, CURRENT_L5_POAM_SPREADSHEET_FILE)
# One worksheet has POAMS generated from Nessus Vulnerability Scans and the other sheet has POAMS not generated
# directly from a Nessus Vulnerability Scan
POAM_DATA_WORKSHEET = "VULN-SCAN-POAMS"
NON_VULN_SCAN_POAMS_WORKSHEET = "NON-VULN-SCAN-POAMS"

# Output Files
POAM_OUTPUT_FILE_JSON = "{environment}\\PTC-{environment}-poam-{year_month_day}.json".format(
        environment=ENVIRONMENT,
        year_month_day=YEAR_MONTH_DAY)
POAM_RESULTS_JSON_FULL_PATH = os.path.join(FEDRAMP_POAM_DIR, POAM_OUTPUT_FILE_JSON)


SERVER_IP_RE = r"(\S+)\s*Ports:\s*((.*)+)"
STARTS_WITH_AFFECTS_RE = r"\s*Affect.*"
PLUGIN_ID_RE = r"Plugin ID: (\d+)"

server_ip_regex = re.compile(SERVER_IP_RE, re.MULTILINE)
starts_with_affects_regex = re.compile(STARTS_WITH_AFFECTS_RE, re.MULTILINE)
plugin_id_regex = re.compile(PLUGIN_ID_RE, re.MULTILINE)

POAM_FIELDS = ['POAM_ID',
               'CONTROLS',
               'WEAKNESS_NAME',
               'WEAKNESS_DESCRIPTION',
               'WEAKNESS_DETECTOR_SOURCE',
               'WEAKNESS_SOURCE_IDENTIFIER',
               'ASSET_IDENTIFIER',
               'POINT_OF_CONTACT',
               'RESOURCES_REQUIRED',
               'OVERALL_REMEDIATION_PLAN',
               'ORIGINAL_DETECTION_DATE',
               'SCHEDULED_COMPLETION_DATE',
               'PLANNED_MILESTONES',
               'MILESTONE_CHANGES',
               'STATUS_DATE',
               'VENDOR_DEPENDENCY',
               'LAST_VENDOR_CHECK_IN_DATE',
               'VENDOR_DEPENDENT_PRODUCT_NAME',
               'ORIGINAL_RISK_RATING',
               'ADJUSTED_RISK_RATING',
               'RISK_ADJUSTMENT',
               'FALSE_POSITIVE',
               'OPERATIONAL_REQUIREMENT',
               'DEVIATION_RATIONALE',
               'SUPPORTING_DOCUMENTS',
               'COMMENTS',
               'AUTO_APPROVE',
               'IP_PORT'
               ]


def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()


def read_poam_excel(in_worksheet):
    poam_results_out = {}
    line_number = 0
    for row in in_worksheet.iter_rows(min_row=FIRST_DATA_ROW, max_col=100, values_only=True):
        if row[POAM_FIELDS.index('POAM_ID')] is None:
            continue;
        line_number += 1
        poam_id = row[POAM_FIELDS.index('POAM_ID')].upper()
        print(str(poam_id))
        plugin_id = get_plugin_id(str(row[POAM_FIELDS.index('WEAKNESS_SOURCE_IDENTIFIER')]))
        if row[POAM_FIELDS.index('WEAKNESS_DESCRIPTION')] is not None:
            weakness_description = row[POAM_FIELDS.index('WEAKNESS_DESCRIPTION')].strip()
        else:
            weakness_description = ""
        poam_results_out[poam_id] = dict(POAM_ID=poam_id,
                                         CONTROLS=row[POAM_FIELDS.index('CONTROLS')],
                                         WEAKNESS_NAME=row[POAM_FIELDS.index('WEAKNESS_NAME')].strip(),
                                         WEAKNESS_DESCRIPTION=weakness_description,
                                         WEAKNESS_DETECTOR_SOURCE=row[POAM_FIELDS.index('WEAKNESS_DETECTOR_SOURCE')].strip(),
                                         WEAKNESS_SOURCE_IDENTIFIER=str(row[POAM_FIELDS.index(
                                                 'WEAKNESS_SOURCE_IDENTIFIER')]).strip(),
                                         ASSET_IDENTIFIER=row[POAM_FIELDS.index('ASSET_IDENTIFIER')].strip(),
                                         POINT_OF_CONTACT=row[POAM_FIELDS.index('POINT_OF_CONTACT')],
                                         RESOURCES_REQUIRED=row[POAM_FIELDS.index('RESOURCES_REQUIRED')],
                                         OVERALL_REMEDIATION_PLAN=row[POAM_FIELDS.index('OVERALL_REMEDIATION_PLAN')],
                                         ORIGINAL_DETECTION_DATE=row[POAM_FIELDS.index('ORIGINAL_DETECTION_DATE')],
                                         SCHEDULED_COMPLETION_DATE=row[POAM_FIELDS.index('SCHEDULED_COMPLETION_DATE')],
                                         PLANNED_MILESTONES=row[POAM_FIELDS.index('PLANNED_MILESTONES')],
                                         MILESTONE_CHANGES=row[POAM_FIELDS.index('MILESTONE_CHANGES')],
                                         STATUS_DATE=row[POAM_FIELDS.index('STATUS_DATE')],
                                         VENDOR_DEPENDENCY=row[POAM_FIELDS.index('VENDOR_DEPENDENCY')],
                                         LAST_VENDOR_CHECK_IN_DATE=row[POAM_FIELDS.index('LAST_VENDOR_CHECK_IN_DATE')],
                                         VENDOR_DEPENDENT_PRODUCT_NAME=row[
                                             POAM_FIELDS.index('VENDOR_DEPENDENT_PRODUCT_NAME')],
                                         ORIGINAL_RISK_RATING=row[POAM_FIELDS.index('ORIGINAL_RISK_RATING')],
                                         ADJUSTED_RISK_RATING=row[POAM_FIELDS.index('ADJUSTED_RISK_RATING')],
                                         RISK_ADJUSTMENT=row[POAM_FIELDS.index('RISK_ADJUSTMENT')],
                                         FALSE_POSITIVE=row[POAM_FIELDS.index('FALSE_POSITIVE')],
                                         OPERATIONAL_REQUIREMENT=row[POAM_FIELDS.index('OPERATIONAL_REQUIREMENT')],
                                         DEVIATION_RATIONALE=row[POAM_FIELDS.index('DEVIATION_RATIONALE')],
                                         SUPPORTING_DOCUMENTS=row[POAM_FIELDS.index('SUPPORTING_DOCUMENTS')],
                                         COMMENTS=row[POAM_FIELDS.index('COMMENTS')],
                                         AUTO_APPROVE=row[POAM_FIELDS.index('AUTO_APPROVE')],
                                         IP_PORT=row[POAM_FIELDS.index('IP_PORT')],
                                         AFFECTED_CMDB_RESOURCES=[],
                                         STATUS="Open",
                                         PLUGIN_ID=plugin_id
                                         )
    return poam_results_out

#
# def get_name_from_poam_asset_identifier(asset_identifier):



if __name__ == "__main__":
    poam_results_spreadsheet_file = os.path.join(POAM_RESULTS_SPREADSHEET_FULL_PATH)
    poam_results_json_out = POAM_RESULTS_JSON_FULL_PATH
    print("Reading poam details from {}".format(poam_results_spreadsheet_file))
    poam_results_wb = openpyxl.load_workbook(filename=poam_results_spreadsheet_file, data_only=True)
    poam_ws = poam_results_wb[POAM_DATA_WORKSHEET]
    poam_results = read_poam_excel(poam_ws)
    non_vuln_poam_ws = poam_results_wb[NON_VULN_SCAN_POAMS_WORKSHEET]
    non_vuln_poam_results = read_poam_excel(non_vuln_poam_ws)
    poam_results.update(non_vuln_poam_results)

    for key in poam_results.keys():
        server_port_lines = split_server_port(poam_results[key]['ASSET_IDENTIFIER'])
        poam_results[key]['AFFECTED_CMDB_RESOURCES'].append(server_port_lines)

    print("Writing poam details in JSON format to {}".format(poam_results_json_out))
    with open(poam_results_json_out, 'w') as json_out_handle:
        json.dump(poam_results, json_out_handle, default=datetime_default)