#用于从EXCEL中计算绩效
import xlrd
from datetime import datetime
import time
#import time.time()
###
import sys
f = open('jixiaonew.log', 'a')
sys.stdout = f
sys.stderr = f
###
path1 = './test.xlsx'
data = xlrd.open_workbook(path1)

table1 = data.sheet_by_index(0)
#print(table1)

nrows = table1.nrows
ncols = table1.ncols
print(nrows)

#print(table1.row_values(1))

def verb_name_check():
    ##define 出勤检查函数
    #verb_name = 
    verb_name = []
    i = 0
    for row in range(2,nrows):
        for col in range(1,ncols):
            
            cell_value_name = table1.cell(row,col).value
            #print(cell_value_name)
            if cell_value_name != '' :
                verb_name.append(cell_value_name)
                i = i + 1 

    print("出勤情况一共",i,"人/次")
    return verb_name

def verb_job(y):
    ##计算个人的每个月出勤职位
    for one in y :
        print(one," 出勤情况为")
        jixiao = 0 
        #for row in range(1,nrows):
        for col in range(0,ncols):
            for row in range(0,nrows):
                ceshi2 = table1.cell(row,col).value
                if one == table1.cell(row,col).value:

                    #verbtime = xlrd.xldate_as_tuple(table1.cell(0,col).value, 0)
                    verbtime = xlrd.xldate.xldate_as_datetime(table1.cell(1,col).value, 0)
                    #time1 = datetime.datetime(verbtime)
                    print(verbtime,table1.cell(row,0).value)
                    #gangwei =  table1.cell(row,0).value
                    #xldate_as_tuple(d,0)
                    if table1.cell(row,0).value == '旅检操机':
                        jixiao = jixiao + 30
                    elif table1.cell(row,0).value == '主人身' or table1.cell(row,0).value == '副人身' or table1.cell(row,0).value == '监护' or table1.cell(row,0).value == '开包' or table1.cell(row,0).value == '廊桥':
                        jixiao = jixiao + 18
                    elif table1.cell(row,0).value == '前传' or table1.cell(row,0).value == '旅检验证':
                        jixiao = jixiao + 12
                    elif table1.cell(row,0).value == '道口白班':
                        jixiao = jixiao + 6
                    elif table1.cell(row,0).value == '道口夜班':
                        jixiao = jixiao + 12
                    
        print("###############",one,"综合绩效为",jixiao)

 ## season is chenged 
def verb_job2(y):
    ##计算个人的每个月出勤职位
    for one in y :
        print(one," 出勤情况为")
        jixiao = 0 
        #for row in range(1,nrows):
        for col in range(0,ncols):
            for row in range(0,nrows):
                #ceshi = table1.cell(row,col).value
                #print(ceshi)

                if one == table1.cell(row,col).value:


                    #verbtime = xlrd.xldate_as_tuple(table1.cell(0,col).value, 0)
                    verbtime = xlrd.xldate.xldate_as_datetime(table1.cell(1,col).value, 0)
                    #time1 = datetime.datetime(verbtime)
                    print(verbtime,table1.cell(row,0).value)
                    #gangwei =  table1.cell(row,0).value
                    #xldate_as_tuple(d,0)
                    xishu = 0 
                    if table1.cell(row,0).value == '旅检操机':
                        xishu = 3.5
                    elif table1.cell(row,0).value == '主人身' or table1.cell(row,0).value == '副人身' or  table1.cell(row,0).value == '开包':
                        xishu =  2.5
                    elif table1.cell(row,0).value == '前传' or table1.cell(row,0).value == '旅检验证' or table1.cell(row,0).value == '廊桥' or table1.cell(row,0).value == '监护':
                        xishu = 2
                    elif table1.cell(row,0).value == '道口白班' or table1.cell(row,0).value == '道口晚班1'  or table1.cell(row,0).value == '道口晚班2':
                        xishu = 1
                    elif table1.cell(row,0).value == '行检':
                        xishu = 3
                    
                    if table1.cell(0,col).value == '星期一':
                        numfly  = 4 
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期二':
                        numfly = 3
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期三':
                        numfly = 4
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期四':
                        numfly = 3
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期五':
                        numfly = 4
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期六':
                        numfly = 3
                        gradefly = 13
                    elif  table1.cell(0,col).value == '星期日':
                        numfly = 5
                        gradefly = 19
                    
                    ##
                    jixiaoday = xishu * numfly + gradefly
                    jixiao = jixiao + jixiaoday

                    


        print("###############",one,"综合绩效为",jixiao)


                    
 ###################################################### main #######################                   
if __name__ == "__main__":
   # print(table1.cell(10,0).value)
    print("绩效计算使用换季后的方式，本次绩效计算操作时间是：")
    print(time.strftime('%Y.%m.%d',time.localtime(time.time())))
    verbname1 = verb_name_check()
    #print("verb1")
    #print(verbname1)
    verbname2 = list(set(verbname1))

    print("本月参加出勤人员",verbname2)

    print("本月参加出勤人数",len(verbname2))
    print(" ")
    print("#########具体出勤情况如下###########")
    print(" ")
    #换季前绩效计算
    #verb_job(verbname2)
    #换季后绩效计算
    verb_job2(verbname2)