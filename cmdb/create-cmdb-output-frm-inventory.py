"""This script reads an excel file that contains CMDB information and creates a JSON dictionary export of the CMDB.


"""
import logging
import os
import io
import re
import openpyxl
import pandas
import json
import coloredlogs
import logging

FIRST_DATA_ROW = 2
ENVIRONMENT = "FRM"
FEDRAMP_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\"
CMDB_DIR = "CMDB\\{environment}".format(environment=ENVIRONMENT)
CURRENT_CMDB_SPREADSHEET_FILE = "{environment}-CMDB.xlsx".format(environment=ENVIRONMENT)
DATA_WORKSHEET = "{environment}_CMDB".format(environment=ENVIRONMENT)
OUTPUT_CMDB_JSON = "PTC-{environment}-CMDB.json".format(environment=ENVIRONMENT)
OUTPUT_CMDB_CSV = "CMDB-{environment}.csv".format(environment=ENVIRONMENT)

CMDB_SPREADSHEET_FULL_PATH = os.path.join(FEDRAMP_DIR, CMDB_DIR, CURRENT_CMDB_SPREADSHEET_FILE)
CMDB_OUT_JSON_FULL_PATH = os.path.join(FEDRAMP_DIR, CMDB_DIR, OUTPUT_CMDB_JSON)

# cmdb Column constants for accessing Excel.  These are 0 based.
ID = 0
PRIMARY_IP_ADDRESS = 1
NAME = 2
ADDITIONAL_IP_ADDRESSES = 3
VIRTUAL = 4
PUBLIC = 5
DNS_NAME_URL = 6
NETBIOS_NAME = 7
MAC_ADDRESS = 8
AUTHENTICATED_SCAN = 9
BASELINE_CONFIGURATION_NAME = 10
OS_NAME_VERSION = 11
LOCATION = 12
ASSET_TYPE = 13
HARDWARE_MAKE_MODEL = 14
IN_LATEST_SCAN = 15
SOFTWARE_DATABASE_VENDOR = 16
SOFTWARE_DB_NAME_VERSION = 17
PATCH_LEVEL = 18
FUNCTION = 19
COMMENTS = 20
SERIAL_NUM_ASSET_TAG = 21
VLAN_NETWORK_ID = 22
SYSTEM_ADMIN_OWNER = 23
APP_ADMIN_OWNER = 24
ENVIRONMENT = 25
NAME_IP = 26

logging_level = 'INFO'
coloredlogs.install(level=logging_level,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')

if __name__ == "__main__":
    logging.info("Opening spreadsheet %s to extract CMDB", CMDB_SPREADSHEET_FULL_PATH)
    cmdb_results_wb = openpyxl.load_workbook(filename=CMDB_SPREADSHEET_FULL_PATH, data_only=True)
    ws = cmdb_results_wb[DATA_WORKSHEET]
    cmdb_results = {}
    line_number = 0
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=37, values_only=True):
        line_number += 1
        # Skip first row if this is a header row
        # cmdb_item_primary_ip_address = row[PRIMARY_IP_ADDRESS]
        cmdb_item_name = row[NAME].upper()
        cmdb_item_additional_ip_addresses = row[ADDITIONAL_IP_ADDRESSES]
        cmdb_item_virtual = row[VIRTUAL]
        cmdb_item_public = row[PUBLIC]
        cmdb_item_dns_name_url = row[DNS_NAME_URL]
        cmdb_item_netbios_name = row[NETBIOS_NAME]
        cmdb_item_mac_address = row[MAC_ADDRESS]
        cmdb_item_authenticated_scan = row[AUTHENTICATED_SCAN]
        cmdb_item_baseline_configuration_name = row[BASELINE_CONFIGURATION_NAME]
        cmdb_item_os_name_version = row[OS_NAME_VERSION]
        cmdb_item_location = row[LOCATION]
        cmdb_item_asset_type = row[ASSET_TYPE]
        cmdb_item_hardware_make_model = row[HARDWARE_MAKE_MODEL]
        cmdb_item_in_latest_scan = row[IN_LATEST_SCAN]
        cmdb_item_software_database_vendor = row[SOFTWARE_DATABASE_VENDOR]
        cmdb_item_software_db_name_version = row[SOFTWARE_DB_NAME_VERSION]
        cmdb_item_patch_level = row[PATCH_LEVEL]
        cmdb_item_function = row[FUNCTION]
        cmdb_item_comments = row[COMMENTS]
        cmdb_item_serial_num_asset_tag = row[SERIAL_NUM_ASSET_TAG]
        cmdb_item_vlan_network_id = row[VLAN_NETWORK_ID]
        cmdb_item_system_admin_owner = row[SYSTEM_ADMIN_OWNER]
        cmdb_item_app_admin_owner = row[APP_ADMIN_OWNER]
        cmdb_item_environment = row[ENVIRONMENT]
        cmdb_item_name_ip = row[NAME_IP]

        # This key is used for the dictionary object.  There will be only one entry per cmdb_ID
        # Returns an array of all of the servers that are impacted by the cmdb
        cmdb_servers = cmdb_item_additional_ip_addresses.split(' ')
        cmdb_results[cmdb_item_name] = dict(UNIQUE_ASSET_IDENTIFIER=cmdb_item_name,
                                            VIRTUAL=cmdb_item_virtual,
                                            PUBLIC=cmdb_item_public,
                                            DNS_NAME_OR_URL=cmdb_item_dns_name_url,
                                            NETBIOS_NAME=cmdb_item_netbios_name,
                                            MAC_ADDRESS=cmdb_item_mac_address,
                                            AUTHENTICATED_SCAN=cmdb_item_authenticated_scan,
                                            BASELINE_CONFIGURATION_NAME=cmdb_item_baseline_configuration_name,
                                            OS_NAME_VERSION=cmdb_item_os_name_version,
                                            LOCATION=cmdb_item_location,
                                            ASSET_TYPE=cmdb_item_asset_type,
                                            HARDWARE_MAKE_MODEL=cmdb_item_hardware_make_model,
                                            IN_LATEST_SCAN=cmdb_item_in_latest_scan,
                                            SOFTWARE_DATABASE_VENDOR=cmdb_item_software_database_vendor,
                                            SOFTWARE_DB_NAME_VERSION=cmdb_item_software_db_name_version,
                                            PATCH_LEVEL=cmdb_item_patch_level,
                                            FUNCTION=cmdb_item_function,
                                            COMMENTS=cmdb_item_comments,
                                            SERIAL_NUM_ASSET_TAG=cmdb_item_serial_num_asset_tag,
                                            VLAN_NETWORK_ID=cmdb_item_vlan_network_id,
                                            SYSTEM_ADMIN_OWNER=cmdb_item_system_admin_owner,
                                            APP_ADMIN_OWNER=cmdb_item_app_admin_owner,
                                            ENVIRONMENT=cmdb_item_environment,
                                            NAME_IP=cmdb_item_name_ip,
                                            IP_ADDRESSES=cmdb_servers
                                            )
    logging.info("Writing CMDB to JSON output to %s", CMDB_OUT_JSON_FULL_PATH)
    with open(CMDB_OUT_JSON_FULL_PATH, 'w') as json_out_handle:
        json.dump(cmdb_results, json_out_handle)
