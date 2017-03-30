#coding = utf8
import xlrd
import logging
import urllib
import json
import sys
from pylsy import pylsytable
import requests
import xlwt





#定义日志输出，其实这个日志可以增强，有时间要去研究要这个自带的logging库
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename = 'myapp.log',
                    filemode='w')
#定义一个streamHandler,将INFO（）级别或更高的日志信息打印到标准错误，并将其添加到当前的日志处理对象
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('(name)-12s:%(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)

#处理excel表格

data = xlrd.open_workbook("C:\\Users\\admin\\Desktop\\python_improve\\jiekou_test\\API.xlsx")#打开文件
logging.info("打开%s excel表格成功" %data)

table = data.sheet_by_name(u'Sheet1')#根据工作名称打开表
logging.info("打开%s表成功" %table)

nrows = table.nrows#对行数的统计
logging.info("表中有%s行"%nrows)

ncols = table.ncols#对列数的统计
logging.info("表中有%s列"%ncols)
logging.info("开始进行循环")
#定义列表，用来存储从excel中读取的对应内容
name_1 = []
url_1 = []
params_1 = []
type_1 = []
Expected_result_1 = []
Actual_result_1 = []
test_result_1 = []
Remarks_1 = []
Success=0
fail = 0
#for循环进行文件内容的读取
for i in range(1,nrows):
    cell_A3 = table.row_values(i)
    name = cell_A3[0]
    url = cell_A3[1]
    params = eval(cell_A3[2])
    type = cell_A3[3]
    error_code = cell_A3[4]
    Remarks = cell_A3[5]
    logging.info(url)
	#判断请求类型
    if type == "GET":
        response = requests.get(url,params)
    else:
        response = requests.post(url,params)
    apicontent = response.text
    print(apicontent)
    apicontent = json.loads(apicontent)
	#开始断言,其实这里可以if else 很多的条件进行判断！sql的话，在功能测试阶段就需要实现。
    if apicontent["code"] == int(error_code):
        name2 = "通过"
        print(name+"测试通过")
    else:
        name2 = "失败"
        print(name+"测试失败")
	#数据存储
    name_1.append(name)
    url_1.append(url)
    params_1.append(params)
    type_1.append(type)
    Expected_result_1.append(int(error_code))
    Actual_result_1.append(apicontent["code"])
    test_result_1.append(name2)
    Remarks_1.append(Remarks)
	#这段代码，就是用作统计
    if name2 == "通过":
        Success+=1
    elif name2 == "失败":
        fail+=1
    else:
        print("测试结果异常")
		
				
#将结果写入excel，构造数据结构
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True)
row0 = [u'接口名称',u'接口地址',u'接口参数',u'请求类型',u'期望结果',u'实际结果',u'测试结果',u'备注']
result = {"name":name_1,"url":url_1,"params":params_1,"type":type_1,"Expected_result":Expected_result_1,"Actual_result":Actual_result_1,"test_result":test_result_1,"Remarks":Remarks_1}
#写入第一行，也就是row0,标题
for i in range(0,len(row0)):
    sheet1.write(0,i,row0[i])
#写入测试过程以及结果
for m in range(0,nrows-1):
    sheet1.write(m+1,0,result['name'][m])
    sheet1.write(m+1,1,result['url'][m])
    sheet1.write(m+1,2,str(result['params'][m]))
    sheet1.write(m+1,3,result['type'][m])
    sheet1.write(m+1,4,result['Expected_result'][m])
    sheet1.write(m+1,5,result['Actual_result'][m])
    sheet1.write(m+1,6,result['test_result'][m])
    sheet1.write(m+1,7,result['Remarks'][m])
#写入统计结果
sheet1.write(nrows+1,0,"成功用例数："+str(Success)+"个")
sheet1.write(nrows+2,0,"失败用例数："+str(fail)+"个")
sheet1.write(nrows+3,0,"***************执行完毕****************")
#对结果进行保存
f.save('测试结果.xls')
		
	
		
#输出表格形式，这个是python库的使用。输出表格。
##attributes =["urlname","url","params","type","Expected_result","Actual_result","test_result","Remarks"]
##table =pylsytable(attributes)
##name =name_1
##url =url_1
##params=params_1
##type=type_1
##Expected_result=Expected_result_1
##Actual_result =Actual_result_1
##test_result=test_result_1
##Remarks=Remarks_1
##table.add_data("urlname",name)
##table.add_data("url",url)
##table.add_data("params",params)
##table.add_data("type",type)
##table.add_data("Expected_result",Expected_result)
##table.add_data("Actual_result",Actual_result)
##table.add_data("test_result",test_result)
##table.add_data("Remarks",Remarks)
##table._create_table()
##print(table)
##print ("成功的用例个数为：%s"%Success,"失败的用例个数为：%s"%fail)
##print( "***********执行测试成功************")























