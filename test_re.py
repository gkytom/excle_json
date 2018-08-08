#coding=utf-8
from handle_execl import *
import re
from xlrd import *
# tag=1
# target=['BUS','','BUS','BUS','dfsafsaf','dddd','']
# for each in target:
#     if tag==1:   ##开始匹配标志
#         if re.findall('BUS',each):
#             tag=tag+1
#             continue
#     else:
#         if re.findall('BUS',each) or re.findall('',each):
#             tag=tag+1
#
# print tag
xls=open_workbook('ACP0.xlsx')
sh=xls.sheet_by_index(0)

print sh.merged_cells
print sh.merged_cells[1][2:4]
print sh.merged_cells[1][0:2]
print sh.merged_cells[1][0]
print sh.merged_cells[1][1]

if sh.merged_cells[1][2:4]==(1,2):
    print '匹配成功'


data_list=[]
data=[(1,21,1,2),(22,49,1,2)]
for each in data:
    data_list.append(each[0:2])
print data_list
limited_list=[]
data1={"first":[{"t1":1},{"t2":2},{}]}
print data1["first"][1]["t2"]
data1["first"][2]["t3"]=3
print data1


# t1={'modules': [{'moduleName': u'zcfes_upel104', 'groups': [{}, {}, {}, {}, {}]}], 'systemName': u'ZCFES'}
# t1['modules'][0]['groups'][0]['pit']=1
# print t1
#
# t2=[(1, 7), (7, 13)]
#
# for i in range(0,len(t2)):
#     print t2[i][1]-t2[i][0]
#排序
data_order=[(1, 7), (7, 13), (37, 49), (13, 19), (19, 37)]
print sorted(data_order)
list_order=[]
for i in range(0,len(data_order)):
    if i<4:
        the_first = data_order[i][0]
        the_second=data_order[i+1][0]
        print the_first,the_second
        if the_first<the_second:
            list_order.append(data_order[i])



    # the_first=data_order[i-1][0]
    # the_second = data_order[i][0]
    # print the_first,the_second
    # if the_first>the_second:
    #     list_order.append(data_order[i])
# print list_order
#
# a=[1,7,37,13,19]
# b=[]
# print sorted(a)
# c=[u'qz_busas']
# print c[0].encode('utf-8')
#
# d= {'a':[{'a0':u'1234'}],'b':u'habus'}
# print type(d)
from handle_execl import coding
from xlrd import *
Excelfile=open_workbook('ACP0.xlsx')
sheet_include=Excelfile.sheet_by_name('ACP')


sys_name=sheet_include.col_values(0)
# print sys_name
mod_name=coding(sheet_include.col_values(2))
# print mod_name
def Fulling(list):
    """
    填充数组中的空白项目，补全分组名称
    :param list:
    :return:
    """
    sum=0
    repeat=0
    num=0
    repeat_num=[]
    group_num=0
    List=[]
    List_result=[]
    for i in range(1,len(list)):
        if list[i]!='':
            List.append(list[i])
        else:
            list[i]=list[i-1]
            # print '变化后的list为',list[i]
            List.append(list[i])
        if list[i]!=list[i-1]:
            group_num=group_num+1
            List_result.append(list[i])
        else:
            continue
    # print '获得的列表是：',List
    # print '列表长度是',len(List)
    for j in range(0,len(List)):
        if num<len(List)-1:
            num=j
            if List[num]==List[num+1]:
                repeat=repeat+1
            else:
                repeat_num.append(repeat)
                repeat=0
            num=num+1
        elif num==len(List)-1:
            for k in range(0,len(repeat_num)):
                repeat_num[k]=repeat_num[k]+1
                sum=sum+repeat_num[k]
            remain=len(List)-sum
            repeat_num.append(remain)
            # print 'num是',num


    return group_num,List,List_result,repeat_num
# print Fulling(sheet_include.col_values(2))[3]
# print Fulling(sheet_include.col_values(2))[0],Fulling(sheet_include.col_values(2))[1],Fulling(sheet_include.col_values(2))[2]
# print len(Fulling(sheet_include.col_values(2))[1])
# print Fulling(sheet_include.col_values(2))[3]

# test_list=[1,1,1,2,2]
# num1=0
# for i in range(len(test_list)):
#     if

def including():
    """
    解决一个模块下有几个组的问题
    :param list:
    :return:
    """
    mod_name = coding(sheet_include.col_values(1))
    return Fulling(mod_name)
#
# print including()[0]
# print including()[1]
# print including()[2]
# List1=including()[3]   #模块角标分布

def collecting(group_num,hosts):
    """

    :param group_num:
    :param hosts:
    :return:
    """
    result=[]
    result_list=[]
    flag=0
    t=0

    for i in range(0,len(group_num)):   #14
        for j in range(flag,flag+group_num[i]):   #10
            result.append(hosts[j])  #第一组数据插入完毕
            t=t+1
        result_list.append(result)
        flag=t
        print '第%d组插入完毕'%flag
        result=[]

    return result_list
print '参数1',Fulling(coding(sheet_include.col_values(2)))[3]
# print '参数2',Fulling(coding(sheet_include.col_values(3)))[1]
print collecting(Fulling(coding(sheet_include.col_values(2)))[3],Fulling(coding(sheet_include.col_values(3)))[1])



#
# print Fulling(coding(sheet_include.col_values(2)))[0]
# print Fulling(coding(sheet_include.col_values(2)))[1]
# print Fulling(coding(sheet_include.col_values(2)))[2]
# print Fulling(coding(sheet_include.col_values(2)))[3]
#
# print Fulling(coding(sheet_include.col_values(1)))[0]
# print Fulling(coding(sheet_include.col_values(1)))[1]
# print Fulling(coding(sheet_include.col_values(1)))[2]
# print Fulling(coding(sheet_include.col_values(1)))[3]
#
print Fulling(coding(sheet_include.col_values(2)))[0]
print Fulling(coding(sheet_include.col_values(2)))[1]
print Fulling(coding(sheet_include.col_values(2)))[2]
print Fulling(coding(sheet_include.col_values(2)))[3]
#
# print Fulling(coding(sheet_include.col_values(6)))[1]
# List2=Fulling(sheet_include.col_values(2))[3]   #分组角标分布
# def compare(mokuai,fenzu):
#     """
#     输出一个模块下有几个分组
#     :param mokuai:
#     :param fenzu:
#     :return:
#     """
#     for i in range(len(mokuai)):

print Fulling(coding(sheet_include.col_values(2)))[2]

import time


def deco(func):
    def wrapper(a, b):
        startTime = time.time()
        func(a, b)
        print a*b
        # print("time is %d ms" % msecs)

    return wrapper


@deco
def func(a, b):
    print("hello，here is a func for add :")
    time.sleep(1)
    print("result is %d" % (a + b))


if __name__ == '__main__':
    f = func
    f(3, 4)