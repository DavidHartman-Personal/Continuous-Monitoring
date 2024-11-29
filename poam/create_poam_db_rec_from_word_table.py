from datetime import datetime
import os
import logging
import coloredlogs
import re
import pprint
import json
import datetime
import docx2python


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


def create_html_table_for_poam(poam_id):
    logging.info("Creating POAM HTML table for POAM %s", poam_id)
    table_string = """<table class="tg">
<thead>
  <tr>
    <th class="tg-ihxd" colspan="4">POAM {poam_id} </th>
  </tr>
</thead>
<tbody>
  <tr>
    <td class="tg-grgk">POAM ID</td>
    <td class="tg-grgk">Weakness Name</td>
    <td class="tg-grgk" colspan="2">Weakness ID</td>
  </tr>
  <tr>
    <td class="tg-0pky">{poam_id}</td>
    <td class="tg-0pky">Security Updates for Internet Explorer<br>(February 2020)</td>
    <td class="tg-0pky" colspan="2">133619</td>
  </tr>
  <tr>
    <td class="tg-phtq">Weakness Description</td>
    <td class="tg-phtq" colspan="3">Plugin Output: <br>  KB : 4537767<br>  - C:\Windows\system32\mshtml.dll has not been patched.<br>    Remote version : 11.0.9600.19597<br>    Should be      : 11.0.9600.19625<br>Note: The fix for this issue is available in either of the following updates:<br>  - KB4537767 : Cumulative Security Update for Internet Explorer<br>  - KB4537821 : Windows 8.1 / Server 2012 R2 Monthly Rollup</td>
  </tr>
  <tr>
    <td class="tg-0lax">Severity</td>
    <td class="tg-0lax">High</td>
    <td class="tg-0lax">Adjusted Risk<br>Rating</td>
    <td class="tg-0lax">High</td>
  </tr>
  <tr>
    <td class="tg-hmp3">False Positive</td>
    <td class="tg-hmp3"></td>
    <td class="tg-hmp3">Operation<br>Requirement</td>
    <td class="tg-hmp3"></td>
  </tr>
  <tr>
    <td class="tg-0pky">Original Detection Date</td>
    <td class="tg-0pky">2/27/2020</td>
    <td class="tg-0pky">Status Date</td>
    <td class="tg-0pky">6/30/2020</td>
  </tr>
  <tr>
    <td class="tg-pvvy" colspan="2">Assets</td>
    <td class="tg-grgk">Protocol</td>
    <td class="tg-grgk">Port</td>
  </tr>
  <tr>
    <td class="tg-9wq8" rowspan="2">Asset Identifier(s)</td>
    <td class="tg-0pky">PHI-MAV-DC02 (10.188.240.111)</td>
    <td class="tg-c3ow">TCP</td>
    <td class="tg-c3ow">445</td>
  </tr>
  <tr>
    <td class="tg-phtq">PHI-MAV-DC02 (10.188.240.111)</td>
    <td class="tg-svo0">TCP</td>
    <td class="tg-svo0">445</td>
  </tr>
  <tr>
    <td class="tg-0pky">Overall Remediation Plan</td>
    <td class="tg-0pky" colspan="3"></td>
  </tr>
  <tr>
    <td class="tg-hmp3">Milestone Changes</td>
    <td class="tg-hmp3" colspan="3"></td>
  </tr>
  <tr>
    <td class="tg-0pky">3</td>
    <td class="tg-0pky">Shooting</td>
    <td class="tg-dvpl">70%</td>
    <td class="tg-dvpl">55%</td>
  </tr>
</tbody>
# </table>
"""
    print(table_string.format(poam_id=poam_id))


def print_style_header():
    style_string = """<style type="text/css">
.tg  {border-collapse:collapse;border-color:#9ABAD9;border-spacing:0;}
.tg td{background-color:#EBF5FF;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#444;
  font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;word-break:normal;}
.tg th{background-color:#409cff;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#fff;
  font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
.tg .tg-phtq{background-color:#D2E4FC;border-color:inherit;text-align:left;vertical-align:top}
.tg .tg-pvvy{background-color:#fe996b;border-color:inherit;font-weight:bold;text-align:center;vertical-align:top}
.tg .tg-9wq8{border-color:inherit;text-align:center;vertical-align:middle}
.tg .tg-hmp3{background-color:#D2E4FC;text-align:left;vertical-align:top}
.tg .tg-c3ow{border-color:inherit;text-align:center;vertical-align:top}
.tg .tg-ihxd{background-color:#3166ff;border-color:inherit;text-align:center;vertical-align:top}
.tg .tg-grgk{background-color:#fe996b;border-color:inherit;font-weight:bold;text-align:left;vertical-align:top}
.tg .tg-0pky{border-color:inherit;text-align:left;vertical-align:top}
.tg .tg-0lax{text-align:left;vertical-align:top}
.tg .tg-svo0{background-color:#D2E4FC;border-color:inherit;text-align:center;vertical-align:top}
.tg .tg-dvpl{border-color:inherit;text-align:right;vertical-align:top}
</style>
"""


if __name__ == "__main__":
    logging.info("Reading POAM Word document")
    # from docx.api import Document
    current_open_poam_file_path = "C:\\Users\\dhartman\\Documents\\FedRAMP\\Continuous Monitoring\\POAM\\L5\\"
    in_poam_file = "poam-test-template.docx"

    # open_poam_worksheet_name = "Open POA&M Items"
    current_poam_full_path = os.path.join(current_open_poam_file_path, in_poam_file)
    doc_result = docx2python.docx2python(current_poam_full_path)
    for items in doc_result.document:
        print("\n")
        print("Item in document: " + str(items))
        poam_row = items[0]
        print("poam row:" + str(poam_row))
        for entry in poam_row:
            print("poam row entry:" + str(entry))
        # for table_row in items[1]:
        #     print("table row:" + str(table_row))
        # poam_id = items[0][1]
        # controls = items[1][1]
        # print(str(poam_id))
        # print(str(controls))

        # for item in items:
        #     print(str(item[1]))

        #print(str(doc_result[0]))
    # document = docx.Document(current_poam_full_path)
    # table = document.tables[0]
    # for i, row in enumerate(table.rows):
    #     text = (cell.text for cell in row.cells)
    #     print(str(text))

# Simplier HTML table
# HTML Table with css styles
# <style type="text/css">
# .tg  {border-collapse:collapse;border-color:#9ABAD9;border-spacing:0;}
# .tg td{background-color:#EBF5FF;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#444;
#   font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;word-break:normal;}
# .tg th{background-color:#409cff;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#fff;
#   font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
# .tg .tg-phtq{background-color:#D2E4FC;border-color:inherit;text-align:left;vertical-align:top}
# .tg .tg-pvvy{background-color:#fe996b;border-color:inherit;font-weight:bold;text-align:center;vertical-align:top}
# .tg .tg-9wq8{border-color:inherit;text-align:center;vertical-align:middle}
# .tg .tg-hmp3{background-color:#D2E4FC;text-align:left;vertical-align:top}
# .tg .tg-c3ow{border-color:inherit;text-align:center;vertical-align:top}
# .tg .tg-ihxd{background-color:#3166ff;border-color:inherit;text-align:center;vertical-align:top}
# .tg .tg-grgk{background-color:#fe996b;border-color:inherit;font-weight:bold;text-align:left;vertical-align:top}
# .tg .tg-0pky{border-color:inherit;text-align:left;vertical-align:top}
# .tg .tg-0lax{text-align:left;vertical-align:top}
# .tg .tg-svo0{background-color:#D2E4FC;border-color:inherit;text-align:center;vertical-align:top}
# .tg .tg-dvpl{border-color:inherit;text-align:right;vertical-align:top}
# </style>
# <table class="tg">
# <thead>
#   <tr>
#     <th class="tg-ihxd" colspan="4">POAM</th>
#   </tr>
# </thead>
# <tbody>
#   <tr>
#     <td class="tg-grgk">POAM ID</td>
#     <td class="tg-grgk">Weakness Name</td>
#     <td class="tg-grgk" colspan="2">Weakness ID</td>
#   </tr>
#   <tr>
#     <td class="tg-0pky">CM-133619</td>
#     <td class="tg-0pky">Security Updates for Internet Explorer<br>(February 2020)</td>
#     <td class="tg-0pky" colspan="2">133619</td>
#   </tr>
#   <tr>
#     <td class="tg-phtq">Weakness Description</td>
#     <td class="tg-phtq" colspan="3">Plugin Output: <br>  KB : 4537767<br>  - C:\Windows\system32\mshtml.dll has not been patched.<br>    Remote version : 11.0.9600.19597<br>    Should be      : 11.0.9600.19625<br>Note: The fix for this issue is available in either of the following updates:<br>  - KB4537767 : Cumulative Security Update for Internet Explorer<br>  - KB4537821 : Windows 8.1 / Server 2012 R2 Monthly Rollup</td>
#   </tr>
#   <tr>
#     <td class="tg-0lax">Severity</td>
#     <td class="tg-0lax">High</td>
#     <td class="tg-0lax">Adjusted Risk<br>Rating</td>
#     <td class="tg-0lax">High</td>
#   </tr>
#   <tr>
#     <td class="tg-hmp3">False Positive</td>
#     <td class="tg-hmp3"></td>
#     <td class="tg-hmp3">Operation<br>Requirement</td>
#     <td class="tg-hmp3"></td>
#   </tr>
#   <tr>
#     <td class="tg-0pky">Original Detection Date</td>
#     <td class="tg-0pky">2/27/2020</td>
#     <td class="tg-0pky">Status Date</td>
#     <td class="tg-0pky">6/30/2020</td>
#   </tr>
#   <tr>
#     <td class="tg-pvvy" colspan="2">Assets</td>
#     <td class="tg-grgk">Protocol</td>
#     <td class="tg-grgk">Port</td>
#   </tr>
#   <tr>
#     <td class="tg-9wq8" rowspan="2">Asset Identifier(s)</td>
#     <td class="tg-0pky">PHI-MAV-DC02 (10.188.240.111)</td>
#     <td class="tg-c3ow">TCP</td>
#     <td class="tg-c3ow">445</td>
#   </tr>
#   <tr>
#     <td class="tg-phtq">PHI-MAV-DC02 (10.188.240.111)</td>
#     <td class="tg-svo0">TCP</td>
#     <td class="tg-svo0">445</td>
#   </tr>
#   <tr>
#     <td class="tg-0pky">Overall Remediation Plan</td>
#     <td class="tg-0pky" colspan="3"></td>
#   </tr>
#   <tr>
#     <td class="tg-hmp3">Milestone Changes</td>
#     <td class="tg-hmp3" colspan="3"></td>
#   </tr>
#   <tr>
#     <td class="tg-0pky">3</td>
#     <td class="tg-0pky">Shooting</td>
#     <td class="tg-dvpl">70%</td>
#     <td class="tg-dvpl">55%</td>
#   </tr>
# </tbody>
# </table>
