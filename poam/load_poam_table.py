from datetime import datetime
import os
import logging
import coloredlogs
import openpyxl
import re
import psycopg2
from model.POAMReport import POAMReport


coloredlogs.install(level=logging.INFO,
                    fmt="%(asctime)s %(hostname)s %(name)s %(filename)s line-%(lineno)d %(levelname)s - %(message)s",
                    datefmt='%H:%M:%S')
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

POAM_HEADER_TABLE = "public.poam_header"
INSERT_STR = "INSERT INTO {table}({columns}) VALUES ({values})"


def insert_poam_header(conn, in_poam_header_record):
    # Create the values insert list for each row

    insert_sql = INSERT_STR.format(table=POAM_HEADER_TABLE,
                                   columns="poam_id, weakness_name, weakness_description, weakness_source_identifier, asset_identifier, month_year_added",
                                   values="%s,%s,%s,%s,%s,%s")

def insert_poam_header_all_recs(cur, in_values_array):
    # [poam_id, weakness_name, weakness_description, weakness_source_identifier, asset_identifier, month_year_added]
    value_str = "(\'{poam_id}\', \'{weakness_name}\',E\'{weakness_description}\')".format(poam_id=in_values_array[0],
                                                                                         weakness_name=in_values_array[1],
                                                                                         weakness_description=in_values_array[2])
    # logging.info("in_values_array [%s]", str(in_values_array))
    # logging.info("value str [%s]", value_str)
    insert_sql = "INSERT INTO public.poam_header_2 (poam_id, weakness_name, weakness_description) VALUES (%s, %s, %s)"
    logging.info("Running the following insert [%s]", insert_sql)
    cur.execute(insert_sql, (in_values_array[0],in_values_array[1],in_values_array[2],))


#    args_str = ','.join(cur.mogrify("(%s,%s,%s,%s,%s,%s)", x).decode('utf-8') for x in in_values_array)
#    cur.execute("INSERT INTO public.poam_header (poam_id, weakness_name, weakness_description, weakness_source_identifier, asset_identifier, month_year_added) VALUES " + args_str)

# Variable/constant declarations

if __name__ == "__main__":
    logging.info("Creating Word Document table for current POAMS")
    current_open_poam_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\L5\\05 - MAY\\"
    in_poam_file = "PTC CS -L5-POAM-Final-May-2020-DB-Test.xlsm"
    open_poam_worksheet_name = "Open POA&M Items"
    current_poam_full_path = os.path.join(current_open_poam_file_path, in_poam_file)
    current_poam_report = POAMReport.poam_from_excel_report(in_year="2020", in_month="5",
                                                            in_poam_excel_file=current_poam_full_path,
                                                            in_open_poam_worksheet_name=open_poam_worksheet_name)
    rows_to_add = []
    for current_poam_entry in current_poam_report.poams:
        logging.info("poam: %s", current_poam_entry[0])
        # For each POAM, create an array containing the values to insert
        # if current_poam_entry[0] == "CM-133619":
        poam_record = current_poam_entry[1]
        poam_id = str(poam_record.poam_id)
        weakness_name = str(poam_record.weakness_name)
        weakness_description = str(poam_record.weakness_description)
        weakness_source_identifier = str(poam_record.weakness_source_identifier)
        asset_identifier = "servers"
        month_year_added = "05-2020"
        rows_to_add.append([poam_id, weakness_name, weakness_description])

    conn = psycopg2.connect(host="localhost",database="poam", user="postgres", password="Helen)))1")
    cur = conn.cursor()
    for insert_row in rows_to_add:
        insert_poam_header_all_recs(cur, insert_row)
    conn.commit()
    cur.close()
    # #insert_sql = "INSERT INTO public.poam_header"
    # insert_columns = " poam_id,vendor_dependency, risk_adjustment, false_positive, operational_requirement, month_year_added"
    # insert_values = " VALUES(%s, %s, %s, %s, %s, %s) RETURNING id;"
    # # INSERT INTO public.poam_header(
    # # 	id, poam_id, controls, weakness_name, weakness_description, weakness_detector_source, weakness_source_identifier, asset_identifier, point_of_contact, resources_required, overall_remediation_plan, original_detection_date, scheduled_completion_date, planned_milestones, milestone_changes, status_date, vendor_dependency, last_vendor_check_in_date, vendor_dependent_product_name, original_risk_rating, adjusted_risk_rating, risk_adjustment, false_positive, operational_requirement, deviation_rationale, supporting_documents, comments, auto_approve, create_date, update_date, month_year_added)
    # # 	VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    #
    # cur.execute(sql, ("1234",False, False, False, False,"072020"))
    # # get the generated id back
    # id = cur.fetchone()[0]
    # # commit the changes to the database
    # conn.commit()
    # # close communication with the database
    # cur.close()
    # # current_poam_report = POAM_Report("202006", current_poam_full_path, open_poam_worksheet_name, closed_poam_worksheet_name)
    # # scan_results_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\Vulnerability-Scanning\\"
    # # in_scan_results_file = "ScanResultsTestFile.xlsx"
    # # logging.debug("in process_poam_excel_file processing excel file [%s]", current_poam_full_path)
    # # poam_wb = openpyxl.load_workbook(current_poam_full_path)
    # # open_poam_ws = poam_wb[open_poam_worksheet_name]
    # # line_number = 0
    # # for row in open_poam_ws.iter_rows(min_row=6, max_col=100, values_only=True):
    # #     if row[POAM_FIELDS.index('POAM_ID')] is None:
    # #         continue
    # #     line_number += 1
    # #     stripped_poam_row = [field.strip() if type(field) == str else str(field) for field in row]
    # #     poam_id = str(stripped_poam_row[POAM_FIELDS.index('POAM_ID')]).upper()
    # #     controls = stripped_poam_row[POAM_FIELDS.index('CONTROLS')]
    # #     weakness_name = stripped_poam_row[POAM_FIELDS.index('WEAKNESS_NAME')]
# controls,
# weakness_name,
# weakness_description,
# weakness_detector_source,
# weakness_source_identifier,
# asset_identifier,
# point_of_contact,
# resources_required,
# overall_remediation_plan,
# original_detection_date,
# scheduled_completion_date,
# planned_milestones,
# milestone_changes,
# status_date,
# vendor_dependency,
# last_vendor_check_in_date,
# vendor_dependent_product_name,
# original_risk_rating,
# adjusted_risk_rating,
# risk_adjustment,
# false_positive,
# operational_requirement,
# deviation_rationale,
# supporting_documents,
# comments, auto_approve,
# create_date,
# update_date,
# month_year_added
