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

SERVER_IP_RE = r"(\S+)\s*\(?.*\)?\s*Ports:\s*((.*)+)"
# SERVER_IP_RE = r"(\S+)\s*Ports:\s*((.*)+)"
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
                      'CUSTOMER',
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


class ScanResult:
    def __init__(self,
                 scan_result_key,
                 plugin,
                 plugin_name,
                 family,
                 severity,
                 ip_address,
                 protocol,
                 port,
                 exploit,
                 mac_address,
                 dns_name,
                 netbios_name,
                 plugin_text,
                 first_discovered,
                 last_observed,
                 exploit_frameworks,
                 vdrf,
                 hostname,
                 name_ip,
                 environment,
                 customer,
                 function,
                 location,
                 system_owner,
                 first_discovered_date,
                 last_observed_date,
                 due_date,
                 days_aged,
                 agegroup,
                 days_till_due,
                 remediation_type,
                 remediation_owner,
                 poam_id,
                 plugin_id_name
                 ):
        self.scan_result_key = scan_result_key
        self.plugin = plugin
        self.plugin_name = plugin_name
        self.family = family
        self.severity = severity
        self.ip_address = ip_address
        self.protocol = protocol
        self.port = port
        self.exploit = exploit
        self.mac_address = mac_address
        self.dns_name = dns_name
        self.netbios_name = netbios_name
        self.plugin_text = plugin_text
        self.first_discovered = first_discovered
        self.last_observed = last_observed
        self.exploit_frameworks = exploit_frameworks
        self.vdrf = vdrf
        self.hostname = hostname
        self.name_ip = name_ip
        self.environment = environment
        self.customer = customer
        self.function = function
        self.location = location
        self.system_owner = system_owner
        self.first_discovered_date = first_discovered_date
        self.last_observed_date = last_observed_date
        self.due_date = due_date
        self.days_aged = days_aged
        self.agegroup = agegroup
        self.days_till_due = days_till_due
        self.remediation_type = remediation_type
        self.remediation_owner = remediation_owner
        self.poam_id = poam_id
        self.plugin_id_name = plugin_id_name
        self.scan_result_dict = dict(PLUGIN=self.plugin,
                                     PLUGIN_NAME=self.plugin_name,
                                     FAMILY=self.family,
                                     SEVERITY=self.severity,
                                     IP_ADDRESS=self.ip_address,
                                     PROTOCOL=self.protocol,
                                     PORT=self.port,
                                     EXPLOIT=self.exploit,
                                     MAC_ADDRESS=self.mac_address,
                                     DNS_NAME=self.dns_name,
                                     NETBIOS_NAME=self.netbios_name,
                                     PLUGIN_TEXT=self.plugin_text,
                                     FIRST_DISCOVERED=self.first_discovered,
                                     LAST_OBSERVED=self.last_observed,
                                     EXPLOIT_FRAMEWORKS=self.exploit_frameworks,
                                     VDRF=self.vdrf,
                                     HOSTNAME=self.hostname,
                                     NAME_IP=self.name_ip,
                                     ENVIRONMENT=self.environment,
                                     CUSTOMER=self.customer,
                                     FUNCTION=self.function,
                                     LOCATION=self.location,
                                     SYSTEM_OWNER=self.system_owner,
                                     FIRST_DISCOVERED_DATE=self.first_discovered_date,
                                     LAST_OBSERVED_DATE=self.last_observed_date,
                                     DUE_DATE=self.due_date,
                                     DAYS_AGED=self.days_aged,
                                     AGEGROUP=self.agegroup,
                                     DAYS_TILL_DUE=self.days_till_due,
                                     REMEDIATION_TYPE=self.remediation_type,
                                     REMEDIATION_OWNER=self.remediation_owner,
                                     POAM_ID=self.poam_id,
                                     PLUGIN_ID_NAME=self.plugin_id_name
                                     )

    def __str__(self):
        return_str = ""
        try:
            return_str = "Plugin {plugin} Server: {server} Port: {port}".format(plugin=self.plugin,
                                                                                server=self.ip_address,
                                                                                port=self.port)
        except Exception as e:
            logging.error("Could not print Vulnerability Scan Finding")
        return return_str

    # def return_dict(self):
    #     logging.debug("Creating dicitonary object for scan result [%s]", self.scan_result_key)
    #     scan_result_dict = dict()

