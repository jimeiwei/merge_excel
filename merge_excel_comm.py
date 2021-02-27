"""Created by jimw"""
import inspect

"""宏定义"""
TT_ERR = 1
TT_OK  = 0

COMM_EXIT_FLAG = 1
Exception = ''
EXCEPTION_HAPPEN_FLAG = 0

"""
初始化打印函数
@:parameter     
@:remake        
@:return        
@:exception     
"""
def comm_init_prt():
    print("-----------------------------------------------------------------------")   #70 个 _
    print("_________________________mergr excel file init ________________________")
    print("-----------------------------------------------------------------------")


"""
exit()使能接口
@:parameter     
@:remake        
@:return        
@:exception  
"""
def comm_exit():
    if COMM_EXIT_FLAG:
        exit()
    else:
        return TT_ERR

"""
失败原因显示
@:parameter     
@:remake        
@:return        
@:exception  
"""
def comm_check_err_prt():
    callerframerecords = inspect.stack()
    indexs_calls = list(range(0, callerframerecords.__len__()))
    indexs_calls.reverse()

    for index in indexs_calls:
        callerframerecord = inspect.stack()[index]
        filename = callerframerecord.filename
        function_name = callerframerecord.function
        lineno = callerframerecord.lineno
        print(filename + ", " + function_name + " fail," + "happens at line:" + str(lineno))

"""
exit()使能接口
@:parameter     
@:remake        
@:return        
@:exception  
"""
def comm_check_rc(rtn):
    if rtn == TT_ERR and EXCEPTION_HAPPEN_FLAG == 1:
        comm_check_err_prt()
        exit()
    elif ( type(rtn) == int and rtn == TT_OK and EXCEPTION_HAPPEN_FLAG == 0):
        return TT_OK

"""
当发生错误，打印错误的函数名和失败的原因
@:parameter     
@:remake        如果失败返回数字
@:return        返回excel词典
@:exception     IOError
"""
def comm_check_err(e):
    print("assert reason:", e.args[0])
    comm_check_err_prt()


"""
范围检查
@:parameter     
@:remake        如果失败返回数字 TT_ERR
@:return        返回excel词典
@:exception     IOError
"""
def comm_check_index(value, min, max):
    comm_check_type_int(value)
    comm_check_type_int(min)
    comm_check_type_int(max)

    if (value >=  min and value <= max):
        return TT_OK
    else:
        print("ERROR INFO:")
        callerframerecords = inspect.stack()
        indexs_calls = list(range(0, callerframerecords.__len__()))
        indexs_calls.reverse()

        for index in indexs_calls:
            callerframerecord = inspect.stack()[index]
            filename = callerframerecord.filename
            function_name = callerframerecord.function
            lineno = callerframerecord.lineno
            if index:
                print(filename + ", " + function_name + " fail," + "happens at line:" + str(lineno))
            else:
                print(filename + ", " + function_name + " fail," + "happens at line:" + str(lineno) + ", value: " + str(value) + ", max: " + str(max) + ", min: " + str(min))
        comm_exit()


"""
检查传参是否是字符串
@:parameter     varb_str
@:remake        如果失败返回数字 TT_ERR
@:return        
@:exception     IOError
"""
def comm_check_type_str(varb_str):
    if (type(varb_str) != str):
        comm_check_err_prt()
        comm_exit()


"""
检查传参是否是数字
@:parameter     varb_int
@:remake        如果失败返回数字 TT_ERR
@:return        TT_OK
@:exception     IOError
"""
def comm_check_type_int(varb_int):
    if (type(varb_int) != int):
        comm_check_err_prt()
        comm_exit()



"""
检查传参是否是列表
@:parameter     varb_list
@:remake        如果失败返回数字 TT_ERR
@:return        TT_OK
@:exception     IOError
"""
def comm_check_type_list(varb_list):
    if (type(varb_list) != list):
        comm_check_err_prt()
        comm_exit()