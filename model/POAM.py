from datetime import datetime
import os
import logging
import coloredlogs
import re
import pprint
import json
import datetime

SERVER_IP_RE = r"(\S+)\s*\(?.*\)?\s*Ports:\s*((.*)+)"
# SERVER_IP_RE = r"(\S+)\s*Ports:\s*((.*)+)"
STARTS_WITH_AFFECTS_RE = r"\s*Affect.*"

coloredlogs.install(level=logging.INFO,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')


def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.day)):
        return obj.isoformat()


POAM_FIELDS = [
        'POAM_ID',
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

class POAM:
    def __init__(self,
                 poam_id,
                 controls,
                 weakness_name,
                 weakness_description,
                 weakness_detector_source,
                 weakness_source_identifier,
                 asset_identifier,
                 point_of_contact,
                 resources_required,
                 overall_remediation_plan,
                 original_detection_date,
                 scheduled_completion_date,
                 planned_milestones,
                 milestone_changes,
                 status_date,
                 vendor_dependency,
                 last_vendor_check_in_date,
                 vendor_dependent_product_name,
                 original_risk_rating,
                 adjusted_risk_rating,
                 risk_adjustment,
                 false_positive,
                 operational_requirement,
                 deviation_rationale,
                 supporting_documents,
                 comments,
                 auto_approve
                 ):
        self.poam_id = poam_id
        self.controls = controls
        self.weakness_name = weakness_name
        self.weakness_description = weakness_description
        self.weakness_detector_source = weakness_detector_source
        self.weakness_source_identifier = weakness_source_identifier
        self.asset_identifier = asset_identifier
        self.point_of_contact = point_of_contact
        self.resources_required = resources_required
        self.overall_remediation_plan = overall_remediation_plan
        self.original_detection_date = original_detection_date
        self.scheduled_completion_date = scheduled_completion_date
        self.planned_milestones = planned_milestones
        self.milestone_changes = milestone_changes
        self.status_date = status_date
        self.vendor_dependency = vendor_dependency
        self.last_vendor_check_in_date = last_vendor_check_in_date
        self.vendor_dependent_product_name = vendor_dependent_product_name
        self.original_risk_rating = original_risk_rating
        self.adjusted_risk_rating = adjusted_risk_rating
        self.risk_adjustment = risk_adjustment
        self.false_positive = false_positive
        self.operational_requirement = operational_requirement
        self.deviation_rationale = deviation_rationale
        self.supporting_documents = supporting_documents
        self.comments = comments
        self.auto_approve = auto_approve
        self.affected_servers_count = 0
        self.affected_assets = []

    def __str__(self):
        return_str = ""
        try:
            return_str = "POAM {poam_id} Name: {weakness_name}".format(poam_id=self.poam_id,
                                                                       weakness_name=self.weakness_name)
        except Exception as e:
            logging.error("Could not print POAM")
        return return_str

    def get_poam_dict(self):
        poam_dict = dict(POAM_ID=self.poam_id,
                         CONTROLS=self.controls,
                         WEAKNESS_NAME=self.weakness_name,
                         WEAKNESS_DESCRIPTION=self.weakness_description,
                         WEAKNESS_DETECTOR_SOURCE=self.weakness_detector_source,
                         WEAKNESS_SOURCE_IDENTIFIER=self.weakness_source_identifier,
                         ASSET_IDENTIFIER=self.asset_identifier,
                         POINT_OF_CONTACT=self.point_of_contact,
                         RESOURCES_REQUIRED=self.resources_required,
                         OVERALL_REMEDIATION_PLAN=self.overall_remediation_plan,
                         ORIGINAL_DETECTION_DATE=self.original_detection_date,
                         SCHEDULED_COMPLETION_DATE=self.scheduled_completion_date,
                         PLANNED_MILESTONES=self.planned_milestones,
                         MILESTONE_CHANGES=self.milestone_changes,
                         STATUS_DATE=self.status_date,
                         VENDOR_DEPENDENCY=self.vendor_dependency,
                         LAST_VENDOR_CHECK_IN_DATE=self.last_vendor_check_in_date,
                         VENDOR_DEPENDENT_PRODUCT_NAME=self.vendor_dependent_product_name,
                         ORIGINAL_RISK_RATING=self.original_risk_rating,
                         ADJUSTED_RISK_RATING=self.adjusted_risk_rating,
                         RISK_ADJUSTMENT=self.risk_adjustment,
                         FALSE_POSITIVE=self.false_positive,
                         OPERATIONAL_REQUIREMENT=self.operational_requirement,
                         DEVIATION_RATIONALE=self.deviation_rationale,
                         SUPPORTING_DOCUMENTS=self.supporting_documents,
                         COMMENTS=self.comments,
                         AUTO_APPROVE=self.auto_approve,
                         AFFECTED_SERVERS_COUNT=len(self.affected_assets),
                         AFFECTED_ASSETS=self.affected_assets
                         )

        return poam_dict

    def create_poam_details(self):
        """ This will create a array of affected servers based on a single Asset Identifier field as defined in the POAM
        """
        SERVER_IP_RE = r"(\S+)\s*\(?(.*?)\)?\s*Ports:\s*((.*)+)"
        STARTS_WITH_AFFECTS_RE = r"\s*Affect.*"
        # SERVER_IP_RE = r"(\S+)\s*Ports:\s*((.*)+)"
        starts_with_affects_regex = re.compile(STARTS_WITH_AFFECTS_RE, re.MULTILINE)
        server_ip_regex = re.compile(SERVER_IP_RE, re.MULTILINE)

        server_ports = []
        server_port_lines = self.asset_identifier.split('\n')
        line_number = 0
        for line in server_port_lines:
            line_number += 1
            # Check if the first line is "Affect.*" and skip if it is.
            if line_number == 1 and re.match(starts_with_affects_regex, line):
                continue
            match = re.search(server_ip_regex, line)
            if match:
                # for match in matches:
                server = match.group(1)
                ip = match.group(2)
                server_ip = str(server).upper()
                ports = [item.strip() for item in match.group(3).split(',')]
                ports = [i for i in ports if i]
                port_array = []
                for port in ports:
                    port_array.append(port)
                server_port = dict(NAME=server_ip,
                                   IP_ADDRESS=ip,
                                   PORTS=port_array)
                self.affected_assets.append(server_port)
            else:
                logging.error("Invalid Asset Identifier for POAM [%s] Asset ID [%s]", self.poam_id,
                              self.asset_identifier)
        self.affected_servers_count = len(self.affected_assets)
        logging.debug("poam [%s] # of servers [%d]", self.poam_id, self.affected_servers_count)

    def add_affected_host(self, server, ip_address, port):
        # if the server already exists as an affected host, add port
        server_exists = False
        for affected_server in self.affected_assets:
            if affected_server['NAME'] == server.upper():
                server_exists = True
                affected_server['PORTS'].append(port)
        if not server_exists:
            new_port = [port]
            server_port = dict(NAME=server,
                               IP_ADDRESS=ip_address,
                               PORTS=new_port)
            self.affected_assets.append(server_port)
