import os
import datetime
import model.CMDBMongoDB as CMDB
import logging
import coloredlogs
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import openpyxl
from mongoengine import *


YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")
ENVIRONMENT = "L5"

FEDRAMP_ROOT_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\CMDB\\"
CMDB_MASTER_SPREADSHEET = "{environment}\\CMDB-{environment}-Master.xlsx".format(environment=ENVIRONMENT)
CMDB_MASTER_SPREADSHEET_FULL_PATH = os.path.join(FEDRAMP_ROOT_DIR, CMDB_MASTER_SPREADSHEET)
CMDB_WORKSHEET = "CMDB-INVENTORY"
CMDB_DATA_START_ROW = 2

SCAN_TARGET_POLICY_SHEET = "all_scans_policies_targets"

MONGODB_DATABASE = "CMDB"
CMDB_FIELDS = ['ID',
               'PRIMARY_IP_ADDRESS',
               'NAME',
               'ADDITIONAL_IP_ADDRESSES',
               'ENVIRONMENT',
               'FUNCTION',
               'SYSTEM_ADMINISTRATOR_OWNER',
               'APPLICATION_ADMINISTRATOR_OWNER',
               'UNIQUE_ASSET_IDENTIFIER',
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
               'COMMENTS',
               'SERIAL_NUMBER_ASSET_TAG_NUMBER',
               'VLAN_NETWORK_ID'
               ]

SCAN_POLICY_TARGET = ["SCAN_NAME",
                      "POLICY_NAME",
                      "CREDENTIALS",
                      "TARGETS"]

def clean_cell(cell):
    if type(cell) == str:
        return cell.strip()
    elif cell is None:
        return ""

    return cell


if __name__ == "__main__":
    # Establishing a Connection
    connect('CMDB', host='localhost', port=27017, username="cmdbuser", password="Helen)))1")

    # open up source workbook for CMDB
    cmdb_wb = openpyxl.load_workbook(CMDB_MASTER_SPREADSHEET_FULL_PATH)
    cmdb_ws = cmdb_wb[CMDB_WORKSHEET]
    line_number = 0
    for raw_row in cmdb_ws.iter_rows(min_row=CMDB_DATA_START_ROW, max_col=100, values_only=True):
        row = [clean_cell(field) for field in raw_row]
        if row[CMDB_FIELDS.index('UNIQUE_ASSET_IDENTIFIER')] is None:
            continue
        line_number += 1
        cmdb_item = CMDB.SystemResource(number=row[CMDB_FIELDS.index('ID')],
                                      primary_ip_address=row[CMDB_FIELDS.index('PRIMARY_IP_ADDRESS')],
                                      name=row[CMDB_FIELDS.index('NAME')],
                                      additional_ip_addresses=row[CMDB_FIELDS.index('ADDITIONAL_IP_ADDRESSES')],
                                      environment=row[CMDB_FIELDS.index('ENVIRONMENT')],
                                      function=row[CMDB_FIELDS.index('FUNCTION')],
                                      system_administrator_owner=row[
                                          CMDB_FIELDS.index('SYSTEM_ADMINISTRATOR_OWNER')],
                                      application_administrator_owner=row[
                                          CMDB_FIELDS.index('APPLICATION_ADMINISTRATOR_OWNER')],
                                      unique_asset_identifier=row[CMDB_FIELDS.index('UNIQUE_ASSET_IDENTIFIER')],
                                      ipv4_or_ipv6_address=row[CMDB_FIELDS.index('IPV4_OR_IPV6_ADDRESS')],
                                      virtual=row[CMDB_FIELDS.index('VIRTUAL')],
                                      public=row[CMDB_FIELDS.index('PUBLIC')],
                                      dns_name_or_url=row[CMDB_FIELDS.index('DNS_NAME_OR_URL')],
                                      netbios_name=row[CMDB_FIELDS.index('NETBIOS_NAME')],
                                      mac_address=row[CMDB_FIELDS.index('MAC_ADDRESS')],
                                      authenticated_scan=row[CMDB_FIELDS.index('AUTHENTICATED_SCAN')],
                                      baseline_configuration_name=row[
                                          CMDB_FIELDS.index('BASELINE_CONFIGURATION_NAME')],
                                      os_name_and_version=row[CMDB_FIELDS.index('OS_NAME_AND_VERSION')],
                                      location=row[CMDB_FIELDS.index('LOCATION')],
                                      asset_type=row[CMDB_FIELDS.index('ASSET_TYPE')],
                                      hardware_make_model=row[CMDB_FIELDS.index('HARDWARE_MAKE_MODEL')],
                                      in_latest_scan=row[CMDB_FIELDS.index('IN_LATEST_SCAN')],
                                      software_database_vendor=row[CMDB_FIELDS.index('SOFTWARE_DATABASE_VENDOR')],
                                      software_database_name_version=row[
                                          CMDB_FIELDS.index('SOFTWARE_DATABASE_NAME_VERSION')],
                                      patch_level=row[CMDB_FIELDS.index('PATCH_LEVEL')],
                                      comments=row[CMDB_FIELDS.index('COMMENTS')],
                                      serial_number_asset_tag_number=row[
                                          CMDB_FIELDS.index('SERIAL_NUMBER_ASSET_TAG_NUMBER')],
                                      vlan_network_id=row[CMDB_FIELDS.index('VLAN_NETWORK_ID')]
                                      )
        cmdb_item.save()



