#用于从EXCEL中计算绩效
import xlrd
from datetime import datetime
import time
###
import sys
f = open('jixiao.log', 'a')
sys.stdout = f
sys.stderr = f
###
path1 = '/Users/mac/Desktop/calculate/test.xlsx'
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
    for row in range(1,nrows):
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
        for col in range(1,ncols):
            for row in range(1,nrows):
                if one == table1.cell(row,col).value:

                    #verbtime = xlrd.xldate_as_tuple(table1.cell(0,col).value, 0)
                    verbtime = xlrd.xldate.xldate_as_datetime(table1.cell(0,col).value, 0)
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
    
                    
                    
                        
                    
                    


        




    



if __name__ == "__main__":
   # print(table1.cell(10,0).value)

    verbname1 = verb_name_check()
    #print("verb1")
    #print(verbname1)
    verbname2 = list(set(verbname1))

    print("本月参加出勤人员",verbname2)

    print("本月参加出勤人数",len(verbname2))
    print(" ")
    print("#########具体出勤情况如下###########")
    print(" ")

    verb_job(verbname2)
    








