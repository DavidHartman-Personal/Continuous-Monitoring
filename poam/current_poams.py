"""This module will load the current OPEN CMDBS into a dictionary object and make available various functions for
retrieving cmdb results.

"""
import os
import dateutil
from dateutil.parser import parse
import datetime
import re
import csv
import openpyxl
import pandas
import json

ENVIRONMENT = "L5"
YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")

SCAN_RESULTS_GENERATION_DAY = "20200531"


FEDRAMP_POAM_ROOT_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\"

# JSON Source File
POAM_MASTER_JSON_FILE = "{environment}\\PTC-{environment}-POAM.json".format(environment=ENVIRONMENT)
POAM_MASTER_JSON_FILE_FULL_PATH = os.path.join(FEDRAMP_POAM_ROOT_DIR, POAM_MASTER_JSON_FILE)

ALL_POAMS = {}

try:
    print("Opening poam master JSON file: {}".format(POAM_MASTER_JSON_FILE_FULL_PATH))
    with open(POAM_MASTER_JSON_FILE_FULL_PATH, "r") as poam_json_fh:
        ALL_POAMS = json.load(poam_json_fh)
except Exception as e:
    print("Error opening poam results JSON file: {}".format(POAM_MASTER_JSON_FILE_FULL_PATH))
    raise ValueError("Error Reading poam JSON file [%s]: [%s]", POAM_MASTER_JSON_FILE_FULL_PATH, str(e))

# PLUGIN_IP_MAPPING = []
# def generate_alias_mapping():
#     alias_mapping = []
#     # name_hash = {}
#     for poam_key, values in ALL_POAMS.items():
#         plugin_id = values['PLUGIN_ID']
#         ips = values['IP_ADDRESSES']
#         for key in ips.keys():
#             alias_mapping.append([plugin_id, key])
#     return alias_mapping
#
# PLUGIN_IP_MAPPING = generate_alias_mapping()
# def get_dict_cmdb_name_ip_listing():
#     """Since a resource can have more than one IP, this will return a dictionary that  maps IP and Name
#
#     """
#     return_dict = {}
#     for key, values in ALL_CMDBS.items():
#         name = values['UNIQUE_ASSET_IDENTIFIER']
#         for ip in values['IP_ADDRESSES']:
#             return_dict_key = "{name}-{ip}".format(name=name, ip=ip)
#             return_dict[return_dict_key] = {}
#             return_dict[return_dict_key]['NAME'] = name
#             return_dict[return_dict_key]['IP_ADDRESS'] = ip
#     return return_dict


# def get_array_cmdb_server_ip_listing():
#     """Returns an array of all cmdb, poam, Server combinations
#
#
#     """
#     return_array = []
#     cmdb_server_ip = get_dict_cmdb_name_ip_listing()
#     for cmdb_entry in cmdb_server_ip.values():
#         return_array.append([cmdb_entry['NAME'], cmdb_entry['IP_ADDRESS']])
#     return return_array

# NAME_IP_MAPPING = get_array_cmdb_server_ip_listing()
# # load_cmdb_file()
#
# CMDB_HEADER_FIELDS = ['CMDB_HEADER_ID',
#                       'CMDB_ID',
#                       'CONTROLS',
#                       'WEAKNESS_NAME',
#                       'WEAKNESS_DESCRIPTION',
#                       'WEAKNESS_DETECTOR_SOURCE',
#                       'WEAKNESS_SOURCE_IDENTIFIER',
#                       'POINT_OF_CONTACT',
#                       'RESOURCES_REQUIRED',
#                       'OVERALL_REMEDIATION_PLAN',
#                       'ORIGINAL_DETECTION_DATE',
#                       'SCHEDULED_COMPLETION_DATE',
#                       'PLANNED_MILESTONES',
#                       'MILESTONE_CHANGES',
#                       'STATUS_DATE',
#                       'VENDOR_DEPENDENCY',
#                       'LAST_VENDOR_CHECK_IN_DATE',
#                       'VENDOR_DEPENDENT_PRODUCT_NAME',
#                       'ORIGINAL_RISK_RATING',
#                       'ADJUSTED_RISK_RATING',
#                       'RISK_ADJUSTMENT',
#                       'FALSE_POSITIVE',
#                       'OPERATIONAL_REQUIREMENT',
#                       'DEVIATION_RATIONALE',
#                       'SUPPORTING_DOCUMENTS',
#                       'COMMENTS',
#                       'AUTO_APPROVE'
#                       ]
#
#
def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()


def get_poam_for_plugin(plugin_id_in):
    """Returns POAM_IDs objects where the plugin matches the input plugin id

    """
    # checks if there is a current POAM for a plugin
    return_poams = []
    for key, values in ALL_POAMS.items():
        if values['PLUGIN_ID'] == plugin_id_in:
            return_poams.append(key)
    return return_poams


def plugin_ip_poam_check(plugin_id_in, ip_in):
    # checks if there is a current POAM for a plugin IP combination
    return_has_poam = False
    for key, values in ALL_POAMS.items():
        if values['PLUGIN_ID'] == plugin_id_in:
            print("PLUGIN {plugin} and IP {ip} has an open POAM [{key}]".format(plugin=plugin_id_in,
                                                                                ip=ip_in,
                                                                                key=key)
                  )
            return_has_poam = True
    return return_has_poam

if __name__ == "__main__":
    # print_name_ip_listing()
    # out_name = get_name_from_ip("10.192.192.15")
    # print("name: " + out_name)
    # out_file = os.path.join(FEDRAMP_CMDBS_ROOT_DIR, "test.csv")
    # print_cmdb_to_csv(out_file)
    # create_fedramp_inv_and_mapping_workbook()
    output_string = "Scan Result key [{key}]: for POAM_ID [{poam_id}], [{weakness_name}]"
    poam_id = 99589

    # for key in ALL_POAMS.keys():
    #     print(output_string.format(key=key,
    #                                poam_id=ALL_POAMS[key]['POAM_ID'],
    #                                weakness_name=ALL_POAMS[key]['WEAKNESS_NAME'].strip())
    #           )
    # clipboard_copy_cmdb_server_ip_listing()
    # load_cmdb_file()
    # cmdb_poam_server_results = get_cmdb_poam_server_listing()
    # print_cmdb_poam_server_listing()
    # clipboard_copy_cmdb_poam_server_listing()

#
#     for value in cmdb_poam_server_results.values():
#         print("{dr_number}\t{poam_id}\t{server}".format(dr_number=value['DR_NUMBER'],
#                                                         poam_id=value['POAM_ID'],
#                                                         server=value['SERVER_NAME']))
# #     cmdb_results_wb = openpyxl.load_workbook(filename=cmdb_results_spreadsheet_file, data_only=True)
#     cmdb_header_ws = cmdb_results_wb[CMDB_HEADER_WORKSHEET]
#     cmdb_detail_ws = cmdb_results_wb[CMDB_DETAIL_WORKSHEET]
#
#     detail_cmdb_results = read_excel_detail(cmdb_detail_ws)
#     header_return_results = read_excel_header(cmdb_header_ws)
#     for key in detail_cmdb_results.keys():
#         cmdb_header_id = detail_cmdb_results[key]['CMDB_HEADER_ID']
#         if header_return_results.get(cmdb_header_id):
#             header_return_results[cmdb_header_id]['ASSET_IDENTIFIERS'].append(detail_cmdb_results[key])
#         else:
#             print("Error, no header for: " + str(key))
#     with open(cmdb_results_json_out, 'w') as json_out_handle:
#         json.dump(header_return_results, json_out_handle, default=datetime_default)
