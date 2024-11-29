"""This module will load the current OPEN CMDBS into a dictionary object and make available various functions for
retrieving cmdb results.

"""
import logging
import coloredlogs
import os
import dateutil
from dateutil.parser import parse
import datetime
import re
import csv
import openpyxl
import json
import cmdb.current_cmdb as current_cmdb

SPREADSHEET_FIRST_DATA_ROW = 2

YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")
ENVIRONMENT = "L5"

FEDRAMP_ROOT_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\CMDB\\"
SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE = "PTC-System-Inventory-Template.xlsx"
SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE_FULL_PATH = os.path.join(FEDRAMP_ROOT_DIR, SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE)
SYSTEM_INV_WORKSHEET = "Inventory"

SYSTEM_INVENTORY_OUTPUT_FIELDS = ['UNIQUE_ASSET_IDENTIFIER',
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


if __name__ == "__main__":
    logging.info("Writing system inventory data to %s", SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE_FULL_PATH)
    # system_inv_spreadsheet_output = os.path.join(SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE_FULL_PATH)
    system_inv_wb = openpyxl.load_workbook(filename=SYSTEM_INVENTORY_SPREADSHEET_TEMPLATE_FULL_PATH, data_only=True)
    cmdb_ws = system_inv_wb[SYSTEM_INV_WORKSHEET]

    for key in cmdb_results.keys():
        ip_addresses = cmdb_results[key]['IPV4_OR_IPV6_ADDRESS'].split('\n')
        for ip in ip_addresses:
            cmdb_results[key]['IP_ADDRESSES'].append(ip)

    print("Writing cmdb details in JSON format to {}".format(cmdb_results_json_out))
    with open(cmdb_results_json_out, 'w') as json_out_handle:
        json.dump(cmdb_results, json_out_handle, default=datetime_default)
