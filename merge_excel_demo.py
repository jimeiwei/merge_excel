"""Created by jimw"""
"""demo file"""

import merge_excel_api as merge_api
import merge_excel_comm as merge_comm
import inspect

excel_file_a = "电气部分.xlsx"

workbook_a = merge_api.merge_excel_load_one_excel(excel_file_a)
merge_comm.comm_check_rc(workbook_a)

sheet_names_a = merge_api.merge_excel_load_sheetnames_by_wb(workbook_a)
merge_comm.comm_check_rc(sheet_names_a)

sheet_content_a = merge_api.merge_excel_load_one_sheet(workbook_a, sheet_names_a[0])
merge_comm.comm_check_rc(sheet_content_a)

cell_value_a = merge_api.merge_excel_value_get_by_row_column(sheet_content_a, 2, 1)
merge_comm.comm_check_rc(cell_value_a)

print(sheet_content_a.merged_cells.ranges)

if (1,1) in sheet_content_a.merged_cells.ranges[0].top:
    print("hello")


