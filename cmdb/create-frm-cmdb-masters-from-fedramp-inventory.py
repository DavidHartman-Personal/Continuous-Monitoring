"""This module will load the current OPEN CMDBS into a dictionary object and make available various functions for
retrieving cmdb results.

"""
import logging
import os

import coloredlogs
import dateutil
from dateutil.parser import parse
import datetime
import re
import csv
import openpyxl
import json

SPREADSHEET_FIRST_DATA_ROW = 3
# HEADER_ROW_FLAG = True

# This effectively defines the root of the project and so adding ..\, etc is not needed
# in config files,etc
# PROJECT_ROOT_DIR = os.path.dirname(os.path.dirname(__file__))

# Add script directory to the path to allow searching for modules
# sys.path.insert(0, PROJECT_ROOT_DIR)

YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")
ENVIRONMENT = "FRM"

FEDRAMP_ROOT_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\CMDB\\"
CMDB_SPREADSHEET_FILENAME = "{environment}\\CMDB-{environment}-Master.xlsx".format(
        environment=ENVIRONMENT)
CMDB_RESULTS_SPREADSHEET_FULL_PATH = os.path.join(FEDRAMP_ROOT_DIR, CMDB_SPREADSHEET_FILENAME)
CMDB_WORKSHEET = "Inventory"

# Output Files
CMDB_OUTPUT_FILE_JSON = "{environment}\\PTC-{environment}-cmdb.json".format(
        environment=ENVIRONMENT,
        year_month_day=YEAR_MONTH_DAY)
CMDB_RESULTS_JSON_FULL_PATH = os.path.join(FEDRAMP_ROOT_DIR, CMDB_OUTPUT_FILE_JSON)

CMDB_FIELDS = ['UNIQUE_ASSET_IDENTIFIER',
               'IPV4_OR_IPV6_ADDRESS',
               'VIRTUAL',
               'PUBLIC',
               'DNS_NAME_OR_URL',
               'NETBIOS_NAME',
               'MAC_ADDRESS',
               'AUTHENTICATED_SCAN',
               'BASELINE_CONFIGURATION_NAME',
               'OS_NAME_AND_VERSION',
               'LOCATION',
               'ASSET_TYPE',
               'HARDWARE_MAKE_MODEL',
               'IN_LATEST_SCAN',
               'SOFTWARE_DATABASE_VENDOR',
               'SOFTWARE_DATABASE_NAME_VERSION',
               'PATCH_LEVEL',
               'FUNCTION',
               'COMMENTS',
               'SERIAL_NUMBER_ASSET_TAG_NUMBER',
               'VLAN_NETWORK_ID',
               'SYSTEM_ADMINISTRATOR_OWNER',
               'APPLICATION_ADMINISTRATOR_OWNER'
               ]

logging_level = 'INFO'
coloredlogs.install(level=logging_level,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')


def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()


def decompose_server_list(assets_impacted_in):
    server_ip_re = r"(\S+)\n"
    # remove any leading/trailing double quotes
    re.sub(r'^"|"$', '', assets_impacted_in)
    server_list = assets_impacted_in.split('\n')
    server_out_result = [server.strip() for server in server_list]

    return server_out_result


def read_cmdb_excel(in_worksheet):
    cmdb_results_out = {}
    line_number = 0
    for row in in_worksheet.iter_rows(min_row=SPREADSHEET_FIRST_DATA_ROW, max_col=100, values_only=True):
        if row[CMDB_FIELDS.index('UNIQUE_ASSET_IDENTIFIER')] is None:
            continue;
        line_number += 1
        print("row:" + str(row))
        cmdb_id = row[CMDB_FIELDS.index('UNIQUE_ASSET_IDENTIFIER')].upper()
        cmdb_results_out[cmdb_id] = dict(UNIQUE_ASSET_IDENTIFIER=cmdb_id,
                                         IPV4_OR_IPV6_ADDRESS=row[CMDB_FIELDS.index('IPV4_OR_IPV6_ADDRESS')],
                                         VIRTUAL=row[CMDB_FIELDS.index('VIRTUAL')],
                                         PUBLIC=row[CMDB_FIELDS.index('PUBLIC')],
                                         DNS_NAME_OR_URL=row[CMDB_FIELDS.index('DNS_NAME_OR_URL')],
                                         NETBIOS_NAME=row[CMDB_FIELDS.index('NETBIOS_NAME')],
                                         MAC_ADDRESS=row[CMDB_FIELDS.index('MAC_ADDRESS')],
                                         AUTHENTICATED_SCAN=row[CMDB_FIELDS.index('AUTHENTICATED_SCAN')],
                                         BASELINE_CONFIGURATION_NAME=row[
                                             CMDB_FIELDS.index('BASELINE_CONFIGURATION_NAME')],
                                         OS_NAME_AND_VERSION=row[CMDB_FIELDS.index('OS_NAME_AND_VERSION')],
                                         LOCATION=row[CMDB_FIELDS.index('LOCATION')],
                                         ASSET_TYPE=row[CMDB_FIELDS.index('ASSET_TYPE')],
                                         HARDWARE_MAKE_MODEL=row[CMDB_FIELDS.index('HARDWARE_MAKE_MODEL')],
                                         IN_LATEST_SCAN=row[CMDB_FIELDS.index('IN_LATEST_SCAN')],
                                         SOFTWARE_DATABASE_VENDOR=row[CMDB_FIELDS.index('SOFTWARE_DATABASE_VENDOR')],
                                         SOFTWARE_DATABASE_NAME_VERSION=row[
                                             CMDB_FIELDS.index('SOFTWARE_DATABASE_NAME_VERSION')],
                                         PATCH_LEVEL=row[CMDB_FIELDS.index('PATCH_LEVEL')],
                                         FUNCTION=row[CMDB_FIELDS.index('FUNCTION')],
                                         COMMENTS=row[CMDB_FIELDS.index('COMMENTS')],
                                         SERIAL_NUMBER_ASSET_TAG_NUMBER=row[
                                             CMDB_FIELDS.index('SERIAL_NUMBER_ASSET_TAG_NUMBER')],
                                         VLAN_NETWORK_ID=row[CMDB_FIELDS.index('VLAN_NETWORK_ID')],
                                         SYSTEM_ADMINISTRATOR_OWNER=row[
                                             CMDB_FIELDS.index('SYSTEM_ADMINISTRATOR_OWNER')],
                                         APPLICATION_ADMINISTRATOR_OWNER=row[
                                             CMDB_FIELDS.index('APPLICATION_ADMINISTRATOR_OWNER')],
                                         IP_ADDRESSES=[]
                                         )
    return cmdb_results_out


if __name__ == "__main__":
    cmdb_results_spreadsheet_file = os.path.join(CMDB_RESULTS_SPREADSHEET_FULL_PATH)
    cmdb_results_json_out = CMDB_RESULTS_JSON_FULL_PATH
    print("Reading cmdb details from {}".format(cmdb_results_spreadsheet_file))
    cmdb_results_wb = openpyxl.load_workbook(filename=cmdb_results_spreadsheet_file, data_only=True)
    cmdb_ws = cmdb_results_wb[CMDB_WORKSHEET]
    cmdb_results = read_cmdb_excel(cmdb_ws)

    for key in cmdb_results.keys():
        ip_addresses = cmdb_results[key]['IPV4_OR_IPV6_ADDRESS'].split('\n')
        for ip in ip_addresses:
            cmdb_results[key]['IP_ADDRESSES'].append(ip)

    print("Writing cmdb details in JSON format to {}".format(cmdb_results_json_out))
    with open(cmdb_results_json_out, 'w') as json_out_handle:
        json.dump(cmdb_results, json_out_handle, default=datetime_default)
