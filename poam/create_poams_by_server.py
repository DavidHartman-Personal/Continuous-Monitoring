"""
Script to create spreadsheet of POAMs split out by Server and Plugin.
This will be used to link back to the remediation owner in the raw scan results as well as to track progress
based on maintenance schedules

"""
import io
import os
import coloredlogs
import logging
import datetime
import pprint
from model.POAMReport import POAMReport
from model.POAM import POAM
from model.ScanResults import ScanResults

# Variable/constant declarations

# Variable/constant declarations

FIRST_DATA_ROW = 2
ENVIRONMENT = "FRM"
YEAR_MONTH_DAY = datetime.datetime.now().strftime("%Y%m%d")
# 3 character month abbreviation
MONTH_ABBREV = datetime.datetime.now().strftime("%b")
# Month number
MONTH_NUM = datetime.datetime.now().strftime("%m")

SCAN_RESULTS_GENERATION_DAY = "20200705"
FEDRAMP_DIR = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous " \
              "Monitoring\\Vulnerability-Scanning\\{environment}\\{month_number} - {month_abb}\\".format(
        environment=ENVIRONMENT,
        month_number=MONTH_NUM,
        month_abb=MONTH_ABBREV)
SCAN_RESULTS_SPREADSHEET = "PTC_{environment}_Weekly_Scan_Review_Results_Only_{year_month_day}.xlsx".format(
        environment=ENVIRONMENT,
        year_month_day=SCAN_RESULTS_GENERATION_DAY)
SCAN_RESULTS_SHEET_NAME = "All Items"

current_open_poam_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\FRM\\06 - JUN\\"
in_poam_file = "PTC-CS-IL5-POAM-June-2020-FINAL_DJH_07152020_CLEAN.xlsm"

out_poam_file = "FedRAMP-FRM-POAM-June-2020.xlsm"
open_poam_worksheet_name = "Open POA&M Items"
out_json_filename = "PTC_FRM_CURRENT_POAMS.json"

coloredlogs.install(level=logging.INFO,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')


def datetime_default(obj):
    if isinstance(obj, (datetime.date, datetime.day)):
        return obj.isoformat()


def create_poam_report_object_from_excel(in_excel_poam_full_path):
    logging.info("Creating and returning a POAM_Report oject created from a populated excel workbook")
    # current_poam_report = POAM_Report("202005", current_poam_full_path, open_poam_worksheet_name, in_current_poam_full_path)
    current_poam_report = POAMReport.poam_from_excel_report(in_year="2020", in_month="5",
                                                            in_poam_excel_file=in_excel_poam_full_path,
                                                            in_open_poam_worksheet_name=open_poam_worksheet_name)
    logging.info("result count [%d]", current_poam_report.results_count)
    # Print out a couple
    # limit = 0
    # for poam_entry in current_poam_report.poams:
    #     limit += 1
    #     if limit < 3:
    #         pprint.pprint(poam_entry.get_poam_dict())
    return current_poam_report


def create_json_output_from_poam_report_object():
    logging.info("Generating a JSON file for a POAM_Report instance")
    poam_json_file_out = os.path.join(current_open_poam_file_path, out_json_filename)


def create_poam_excel_report_output(in_current_poam_report, in_poam_ouput_full_path):
    logging.info("Generate a POAM Report using the Excel FedRAMP Template")
    in_current_poam_report.create_poam_excel_report_output(excel_report_output_file=in_poam_ouput_full_path)


# def create_poam_objects_from_scans_and_open_poams():

if __name__ == "__main__":
    logging.info("Various functions for processing and generating POAM_Reports")
    current_poam_full_path = os.path.join(current_open_poam_file_path, in_poam_file)
    #poam_ouput_full_path = os.path.join(current_open_poam_file_path, out_poam_file)

    # Create POAMReport object from a FedRAMP Template report
    current_poam_report = create_poam_report_object_from_excel(current_poam_full_path)

    # To create a POAM Report using the FedRAMP template for an instance of POAMReport
    # create_poam_excel_report_output(in_current_poam_report=current_poam_report, in_poam_ouput_full_path=poam_ouput_full_path)

    # Read in weekly scan results
    scan_results_spreadsheet_file = os.path.join(FEDRAMP_DIR, SCAN_RESULTS_SPREADSHEET)
    # Normally this would be All Items
    weekly_scan_results = ScanResults(in_scan_date="20200705",
                                      scan_results_file=scan_results_spreadsheet_file,
                                      scan_results_worksheet="ScanResults")

    # create spreadsheet with POAM ID, Server and Plugin

    poam_recs_from_vuln_scan = {}
    point_of_contact = "Tom Wollard"
    for scan_result in weekly_scan_results.scan_results:
        if not poam_recs_from_vuln_scan.get(scan_result.plugin):
            # not seen this plugin previously.  Check if it has a current open POAM and then add new entry to dictionary
            logging.debug("Adding plugin to list to create poam [%s]", scan_result.plugin)
            scan_result_poam_id = scan_result.poam_id
            if scan_result_poam_id != "No POAM":
                logging.debug("Has current open POAM")
                # TODO: Need function to return single POAM object based on passed POAM ID
                for current_poam_entry in current_poam_report.poams:
                    if current_poam_entry[0] == scan_result_poam_id:
                        poam_record = current_poam_entry[1]
                # get information from current POAM regarding Deviation requests, etc
                weakness_source_identifier = poam_record.weakness_source_identifier
                weakness_name = poam_record.weakness_name
                weakness_description = poam_record.weakness_description
                weakness_detector_source = poam_record.weakness_detector_source
                controls = poam_record.controls
                resources_required = poam_record.resources_required
                overall_remediation_plan = poam_record.overall_remediation_plan
                original_detection_date = poam_record.original_detection_date
                # scheduled_completion_date = poam_record.scheduled_completion_date
                planned_milestones = poam_record.planned_milestones
                milestone_changes = poam_record.milestone_changes
                status_date = poam_record.status_date
                vendor_dependency = poam_record.vendor_dependency
                vendor_dependent_product_name = poam_record.vendor_dependent_product_name
                last_vendor_check_in_date = poam_record.last_vendor_check_in_date
                original_risk_rating = poam_record.original_risk_rating
                adjusted_risk_rating = poam_record.adjusted_risk_rating
                risk_adjustment = poam_record.risk_adjustment
                false_positive = poam_record.false_positive
                operational_requirement = poam_record.operational_requirement
                deviation_rationale = poam_record.deviation_rationale
                supporting_documents = poam_record.supporting_documents
                comments = poam_record.comments
            else:
                # new POAM
                weakness_source_identifier = scan_result.plugin
                weakness_name = scan_result.plugin_name
                weakness_description = scan_result.plugin_text
                weakness_detector_source = "Nessus"
                controls = "RA-5"
                overall_remediation_plan = "TBD"
                original_detection_date = scan_result.first_discovered_date
                # Scheduled completion date is a formula
                # scheduled_completion_date = poam_record.scheduled_completion_date
                planned_milestones = "TBD"
                milestone_changes = ""
                status_date = scan_result.last_observed_date
                resources_required = ""
                vendor_dependency = ""
                vendor_dependent_product_name = ""
                last_vendor_check_in_date = ""
                if scan_result.severity == "Medium":
                    original_risk_rating = "Moderate"
                elif scan_result.severity == "Critical" or scan_result.severity == "High":
                    original_risk_rating = "High"
                else:
                    original_risk_rating = "Low"
                adjusted_risk_rating = original_risk_rating
                risk_adjustment = "No"
                false_positive = "No"
                operational_requirement = "No"
                deviation_rationale = ""
                supporting_documents = ""
                comments = ""
            # Create a POAM record
            # new_poam_rec = POAM(poam_id=scan_result_poam_id,
            poam_recs_from_vuln_scan[scan_result.plugin] = POAM(poam_id=scan_result_poam_id,
                                                                controls=controls,
                                                                weakness_name=weakness_name,
                                                                weakness_description=weakness_description,
                                                                weakness_detector_source=weakness_detector_source,
                                                                weakness_source_identifier=weakness_source_identifier,
                                                                asset_identifier="",
                                                                point_of_contact=point_of_contact,
                                                                resources_required=resources_required,
                                                                overall_remediation_plan=overall_remediation_plan,
                                                                original_detection_date=original_detection_date,
                                                                scheduled_completion_date="",
                                                                planned_milestones=planned_milestones,
                                                                milestone_changes=milestone_changes,
                                                                status_date=status_date,
                                                                vendor_dependency=vendor_dependency,
                                                                last_vendor_check_in_date=last_vendor_check_in_date,
                                                                vendor_dependent_product_name=vendor_dependent_product_name,
                                                                original_risk_rating=original_risk_rating,
                                                                adjusted_risk_rating=adjusted_risk_rating,
                                                                risk_adjustment=risk_adjustment,
                                                                false_positive=false_positive,
                                                                operational_requirement=operational_requirement,
                                                                deviation_rationale=deviation_rationale,
                                                                supporting_documents=supporting_documents,
                                                                comments=comments,
                                                                auto_approve=""
                                                                )
            poam_recs_from_vuln_scan[scan_result.plugin].add_affected_host(server=scan_result.hostname.upper(),
                                                                           ip_address=scan_result.ip_address,
                                                                           port=scan_result.port)
        else:
            # POAM record already created, just append the new server/ip/port
            poam_recs_from_vuln_scan[scan_result.plugin].add_affected_host(server=scan_result.hostname.upper(),
                                                                           ip_address=scan_result.ip_address,
                                                                           port=scan_result.port)
    # Create a POAM_Report object and then add POAM objects created
    poam_output_report = POAMReport(in_year="2020", in_month="6", report_name="FRM-2020-05")
    # Now add the POAMS we created
    for key, value in poam_recs_from_vuln_scan.items():
        poam_output_report.poams.append([value.poam_id,value])
    poam_output_report.create_poam_excel_report_output(excel_report_output_file=poam_ouput_full_path)
