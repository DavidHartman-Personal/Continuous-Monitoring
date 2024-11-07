"""Classes and functions for reading/processing/creating a POAM report.

Core functionality:
1) Read in a Excel file that has Vulnerability scan results
2) DONE: Read in a JSON file of current Scan results
4) DONE: Create JSON dump of scan results report, including report metadata
5) TODO: Support the ability to create a POAM report from a VulnerabilityScanResults object

TODO: Split out POAM and POAM_Report classes

"""
from datetime import datetime
import os
import logging
import coloredlogs
import openpyxl
import re
import pprint
import json
from model import ScanResult

SERVER_IP_RE = r"(\S+)\s*\(?.*\)?\s*Ports:\s*((.*)+)"
STARTS_WITH_AFFECTS_RE = r"\s*Affect.*"

coloredlogs.install(level=logging.INFO,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')
SCAN_RESULT_FIELDS = ['PLUGIN',
                      'PLUGIN_NAME',
                      'FAMILY',
                      'SEVERITY',
                      'IP_ADDRESS',
                      'PROTOCOL',
                      'PORT',
                      'EXPLOIT',
                      'MAC_ADDRESS',
                      'DNS_NAME',
                      'NETBIOS_NAME',
                      'PLUGIN_TEXT',
                      'FIRST_DISCOVERED',
                      'LAST_OBSERVED',
                      'EXPLOIT_FRAMEWORKS',
                      'VDRF',
                      'HOSTNAME',
                      'NAME_IP',
                      'ENVIRONMENT',
                      #'CUSTOMER'
                      'FUNCTION',
                      'LOCATION',
                      'SYSTEM_OWNER',
                      'FIRST_DISCOVERED_DATE',
                      'LAST_OBSERVED_DATE',
                      'DUE_DATE',
                      'DAYS_AGED',
                      'AGEGROUP',
                      'DAYS_TILL_DUE',
                      'REMEDIATION_TYPE',
                      'REMEDIATION_OWNER',
                      'POAM_ID',
                      'PLUGIN_ID_NAME'
                      ]


def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.day)):
        return obj.isoformat()


class ScanResults:
    tab_separator = "\t"
    comma_separator = ","

    def __init__(self, in_scan_date, scan_results_file, scan_results_worksheet):
        """ check what file type we have and then choose the appropriate method to create class
        """
        self.in_scan_date = in_scan_date
        self.scan_results_file = scan_results_file
        # These will be NULL for JSON Files
        self.scan_results_worksheet = scan_results_worksheet
        self.scan_results_processing_dt = datetime.now()
        self.scan_results = []
        self.result_count = 0
        if os.path.isfile(scan_results_file):
            if scan_results_file.endswith(".xlsx") or \
                    scan_results_file.endswith(".xlsm") or \
                    scan_results_file.endswith(".xls"):
                logging.info("processing excel file [%s]", scan_results_file)
                self.process_scan_results_excel_file()
            elif scan_results_file.endswith(".json"):
                logging.info("processing json file [%s]", scan_results_file)
                self.process_scan_results_json_file()
        else:
            logging.error("[%s] is not a valid file or doesn't exist", scan_results_file)
            raise NameError(scan_results_file)

    IPS_IN_SCAN_RESULTS = []

    def ip_in_scan_results(in_ip):
        return in_ip in SCAN_RESULT_IP_SET

    def process_scan_results_excel_file(self):
        logging.debug("in process_poam_excel_file processing excel file [%s]", self.scan_results_file)
        scan_results_wb = openpyxl.load_workbook(self.scan_results_file)
        scan_results_ws = scan_results_wb[self.scan_results_worksheet]
        line_number = 0
        # TODO: make first data row variable
        for row in scan_results_ws.iter_rows(min_row=2, max_col=100, values_only=True):
            if row[SCAN_RESULT_FIELDS.index('PLUGIN')] is None:
                continue
            line_number += 1

            stripped_scan_result_row = [field.strip() if type(field) == str else str(field) for field in row]

            plugin = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PLUGIN')]
            plugin_name = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PLUGIN_NAME')]
            family = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('FAMILY')]
            severity = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('SEVERITY')]
            ip_address = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('IP_ADDRESS')]
            protocol = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PROTOCOL')]
            port = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PORT')]
            exploit = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('EXPLOIT')]
            mac_address = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('MAC_ADDRESS')]
            dns_name = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('DNS_NAME')]
            netbios_name = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('NETBIOS_NAME')]
            plugin_text = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PLUGIN_TEXT')]
            first_discovered = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('FIRST_DISCOVERED')]
            last_observed = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('LAST_OBSERVED')]
            exploit_frameworks = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('EXPLOIT_FRAMEWORKS')]
            vdrf = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('VDRF')]
            hostname = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('HOSTNAME')]
            name_ip = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('NAME_IP')]
            environment = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('ENVIRONMENT')]
            # customer = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('CUSTOMER')]
            function = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('FUNCTION')]
            location = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('LOCATION')]
            system_owner = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('SYSTEM_OWNER')]
            first_discovered_date = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('FIRST_DISCOVERED_DATE')]
            last_observed_date = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('LAST_OBSERVED_DATE')]
            due_date = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('DUE_DATE')]
            days_aged = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('DAYS_AGED')]
            agegroup = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('AGEGROUP')]
            days_till_due = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('DAYS_TILL_DUE')]
            remediation_type = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('REMEDIATION_TYPE')]
            remediation_owner = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('REMEDIATION_OWNER')]
            poam_id = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('POAM_ID')]
            plugin_id_name = stripped_scan_result_row[SCAN_RESULT_FIELDS.index('PLUGIN_ID_NAME')]

            # Key for the dictionary object is plugin, ip, protocol, port and year/month/day
            scan_result_key = "{plugin}-{ip_address}-{protocol}-{port}".format(plugin=plugin,
                                                                               ip_address=ip_address,
                                                                               protocol=protocol,
                                                                               port=port)
            scan_result = ScanResult.ScanResult(scan_result_key=scan_result_key,
                                                   plugin=plugin,
                                                   plugin_name=plugin_name,
                                                   family=family,
                                                   severity=severity,
                                                   ip_address=ip_address,
                                                   protocol=protocol,
                                                   port=port,
                                                   exploit=exploit,
                                                   mac_address=mac_address,
                                                   dns_name=dns_name,
                                                   netbios_name=netbios_name,
                                                   plugin_text=plugin_text,
                                                   first_discovered=first_discovered,
                                                   last_observed=last_observed,
                                                   exploit_frameworks=exploit_frameworks,
                                                   vdrf=vdrf,
                                                   hostname=hostname,
                                                   name_ip=name_ip,
                                                   environment=environment,
                                                   customer="",
                                                   function=function,
                                                   location=location,
                                                   system_owner=system_owner,
                                                   first_discovered_date=first_discovered_date,
                                                   last_observed_date=last_observed_date,
                                                   due_date=due_date,
                                                   days_aged=days_aged,
                                                   agegroup=agegroup,
                                                   days_till_due=days_till_due,
                                                   remediation_type=remediation_type,
                                                   remediation_owner=remediation_owner,
                                                   poam_id=poam_id,
                                                   plugin_id_name=plugin_id_name
                                                   )
            logging.debug("Scan result [%s]", str(scan_result))
            self.scan_results.append(scan_result)
        self.result_count = line_number
        logging.info("Processed [%d] scan results", self.result_count)

    def process_scan_results_json_file(self):
        logging.info("in process_scan_results_json_file processing json file [%s]", self.scan_results_file)
        try:
            logging.info("Opening Vulnerability Scan Results master JSON file: %s", self.scan_results_file)
            with open(self.scan_results_file, "r") as vuln_scan_json_fh:
                json_file_scan_results = json.load(vuln_scan_json_fh)
        except Exception as e:
            logging.error("Error opening cmdb results JSON file: %s", self.scan_results_file)
            raise ValueError("Error Reading cmdb JSON file [%s]: [%s]", self.scan_results_file, str(e))
        # Now pull some of the metadata and then create VulnerabilityScanFinding objects for each scan result
        self.result_count = json_file_scan_results['SCAN_RESULT_COUNT']
        self.in_scan_date = json_file_scan_results['SCAN_RESULTS_DATE']
        self.scan_results_processing_dt = json_file_scan_results['REPORT_GEN_DATETIME']
        scan_results_array = json_file_scan_results['SCAN_RESULTS']
        # Loop through all scan results and create individual scan result finding objects
        for scan_result_dict in scan_results_array:
            scan_result_key = "{plugin}-{ip_address}-{protocol}-{port}".format(plugin=scan_result_dict['PLUGIN'],
                                                                               ip_address=scan_result_dict[
                                                                                   'IP_ADDRESS'],
                                                                               protocol=scan_result_dict['PROTOCOL'],
                                                                               port=scan_result_dict['PORT'])
            scan_result = ScanResult.ScanResult(scan_result_key=scan_result_key,
                                                plugin=scan_result_dict['PLUGIN'],
                                                plugin_name=scan_result_dict['PLUGIN_NAME'],
                                                family=scan_result_dict['FAMILY'],
                                                severity=scan_result_dict['SEVERITY'],
                                                ip_address=scan_result_dict['IP_ADDRESS'],
                                                protocol=scan_result_dict['PROTOCOL'],
                                                port=scan_result_dict['PORT'],
                                                exploit=scan_result_dict['EXPLOIT'],
                                                mac_address=scan_result_dict['MAC_ADDRESS'],
                                                dns_name=scan_result_dict['DNS_NAME'],
                                                netbios_name=scan_result_dict['NETBIOS_NAME'],
                                                plugin_text=scan_result_dict['PLUGIN_TEXT'],
                                                first_discovered=scan_result_dict['FIRST_DISCOVERED'],
                                                last_observed=scan_result_dict['LAST_OBSERVED'],
                                                exploit_frameworks=scan_result_dict['EXPLOIT_FRAMEWORKS'],
                                                vdrf=scan_result_dict['VDRF'],
                                                hostname=scan_result_dict['HOSTNAME'],
                                                name_ip=scan_result_dict['NAME_IP'],
                                                environment=scan_result_dict['ENVIRONMENT'],
                                                customer=scan_result_dict['CUSTOMER'],
                                                function=scan_result_dict['FUNCTION'],
                                                location=scan_result_dict['LOCATION'],
                                                system_owner=scan_result_dict['SYSTEM_OWNER'],
                                                first_discovered_date=scan_result_dict['FIRST_DISCOVERED_DATE'],
                                                last_observed_date=scan_result_dict['LAST_OBSERVED_DATE'],
                                                due_date=scan_result_dict['DUE_DATE'],
                                                days_aged=scan_result_dict['DAYS_AGED'],
                                                agegroup=scan_result_dict['AGEGROUP'],
                                                days_till_due=scan_result_dict['DAYS_TILL_DUE'],
                                                remediation_type=scan_result_dict['REMEDIATION_TYPE'],
                                                remediation_owner=scan_result_dict['REMEDIATION_OWNER'],
                                                poam_id=scan_result_dict['POAM_ID'],
                                                plugin_id_name=scan_result_dict['PLUGIN_ID_NAME']
                                                )
            logging.debug("Scan result [%s]", str(scan_result))
            self.scan_results.append(scan_result)

    def create_json_vulnerability_scan_results_report(self, out_json_file):
        logging.info("create vulnerability scan results JSON output from file [%s]", self.scan_results_file)
        current_date_time = datetime.now().strftime('%B %d, %Y %I:%M:%S EDT')
        out_dictionary = dict(REPORT_NAME=self.scan_results_file,
                              SCAN_RESULT_COUNT=self.result_count,
                              SCAN_RESULTS_DATE=self.in_scan_date,
                              REPORT_GEN_DATETIME=current_date_time,  # self.in_scan_date, # datetime.now(),
                              SCAN_RESULTS=[])
        for scan_result in self.scan_results:
            out_dictionary['SCAN_RESULTS'].append(scan_result.scan_result_dict)
        # pprint.pprint(out_dictionary)
        logging.info("Generating JSON output file [%s]", out_json_file)
        with open(out_json_file, 'w') as json_out_handle:
            json.dump(out_dictionary, json_out_handle, default=datetime_default)


if __name__ == "__main__":
    scan_results_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\Vulnerability-Scanning\\FRM\\07 - JUL\\"
    in_scan_results_file = "PTC_FRM_Weekly_Scan_Review_07072020.xlsx"
    out_json_filename = "PTC_FRM_Weekly_Scan_Review_07072020.json"
    scan_results_worksheet_name = "All Items"
    in_full_path = os.path.join(scan_results_file_path, in_scan_results_file)
    json_file_full_path = os.path.join(scan_results_file_path, out_json_filename)
    # TODO: Make in_scan_date a datetime withonly a day/month/year
    # Read in a previously created JSON file and create a VulnerabilityScanResults object
    current_vulnerability_scan_results = ScanResults(in_scan_date="20200707",
                                                     scan_results_file=json_file_full_path,
                                                     scan_results_worksheet="")
    logging.info("Scan Result count [%d]", current_vulnerability_scan_results.result_count)
