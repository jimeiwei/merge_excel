"""Created by jimw"""

import merge_excel_comm as merge_comm
import merge_excel_api as merge_api
import sys

SYS_VARB_INPUT = 4


"""
python merge_excel_run.py 电气部分.xlsx  电气部分.xlsx
python merge_excel_run.py /Users/jimeiwei/Desktop/book1.xlsx /Users/jimeiwei/Desktop/book2.xlsx
"""

if __name__ == "__main__":
    args = sys.argv         # 获取系统参数
    end_loop_flag = 1

    excel_file_a = ""
    excel_file_b = ""
    workbook_a = ""
    workbook_b = ""

    merge_comm.comm_init_prt()

    while end_loop_flag:
        num_varb_split = 0
        if args.__len__() != SYS_VARB_INPUT:
            print("Current system varb is less than 3, retype here:")
            varb_input = input()

            varb_splits = varb_input.split(" ")
            list_sys_varb = []
            list_sys_varb = [varb_split for varb_split in varb_splits if varb_split != ""]
            len_sys_varb = len(list_sys_varb)  # 输入的变量分割后有效个数
            if len_sys_varb == SYS_VARB_INPUT and "python" == list_sys_varb[0] and ".py" in list_sys_varb[1]:
                excel_file_a = list_sys_varb[2]
                excel_file_b = list_sys_varb[3]
                end_loop_flag = 0
        else:
            excel_file_a = sys.argv[1]
            excel_file_b = sys.argv[2]
            break

    workbook_a = merge_api.merge_excel_load_one_excel(excel_file_a)
    merge_comm.comm_check_rc(workbook_a)

    sheet_names_a = merge_api.merge_excel_load_sheetnames_by_wb(workbook_a)
    merge_comm.comm_check_rc(sheet_names_a)

    sheet_content_a = merge_api.merge_excel_load_one_sheet(workbook_a, sheet_names_a[0])
    merge_comm.comm_check_rc(sheet_content_a)

    cell_value_a = merge_api.merge_excel_value_get_by_row_column(sheet_content_a, 2, 1)
    merge_comm.comm_check_rc(cell_value_a)


