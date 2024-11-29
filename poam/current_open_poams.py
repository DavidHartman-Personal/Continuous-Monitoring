"""
This script can be used to process and/or create POAM Reports  Each of these will be it's own function to keep the main section clean

1. Given a populated POAM Excel Report, create an instance of a POAM_Report object via a Class function that takes in the spreadsheet containing POAM Records
2. Given a populated POAM_Report class object, create a JSON export of the POAM_Report oject.
   - Note that this JSON formatted output can be processed to create a POAM_Report instance later.

"""
import io
import os
import coloredlogs
import logging
import datetime
import pprint
from model.POAMReport import POAMReport

# Variable/constant declarations
current_open_poam_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\FRM\\05 - MAY\\"
in_poam_file = "FedRAMP-FRM-POAM-May-2020.xlsm"
out_poam_file = "FedRAMP-FRM-POAM-June-2020.xlsm"
open_poam_worksheet_name = "Open POA&M Items"
closed_poam_worksheet_name = "Closed POA&M Items"
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
    current_poam_report = POAMReport.poam_from_excel_report(in_year="2020", in_month="5", in_poam_excel_file=in_excel_poam_full_path, in_open_poam_worksheet_name=open_poam_worksheet_name)
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


if __name__ == "__main__":
    logging.info("Various functions for processing and generating POAM_Reports")
    current_poam_full_path = os.path.join(current_open_poam_file_path, in_poam_file)
    poam_ouput_full_path = os.path.join(current_open_poam_file_path, out_poam_file)

    # Create POAMReport object from a FedRAMP Template report
    current_poam_report = create_poam_report_object_from_excel(current_poam_full_path)

    # To create a POAM Report using the FedRAMP template for an instance of POAMReport
    # create_poam_excel_report_output(in_current_poam_report=current_poam_report, in_poam_ouput_full_path=poam_ouput_full_path)