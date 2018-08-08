#coding=utf-8
from xlrd import *
import json

def handle_repeat(col_name):
    """
    去重复用模块
    :param col_name:
    :return:
    """
    mod_name_list = []
    for each_mod in col_name:
        if each_mod == '':
            continue
        else:
            if len(mod_name_list) == 0:
                mod_name_list.append(each_mod)
            else:
                if each_mod not in mod_name_list:
                    mod_name_list.append(each_mod)
                else:
                    continue
    return mod_name_list

def gain_clonum(List):
    """
    获取合并单元格对应的列元组
    :param list:
    :return:
    """
    list=[]
    for each in List:
        list.append(each[0:2])
    return list


def gain_num(excel_name):
    """
    获取单元格索引信息，返回适配JSON字典，方便分组使用
    :param col_name:
    :return:
    """
    list=[]
    mode_list=[]
    group_list=[]

    list=excel_name.merged_cells   #返回合并单元格元祖索引字典
    # print list
    for each in list:
        if each[2:4]==(1,2):   #匹配模块成功
            mode_list.append(each)
        elif each[2:4]==(2,3):  #匹配组成功
            group_list.append(each)
        else:
            continue

    return gain_clonum(mode_list),gain_clonum(group_list)


def build_dict(mode):
    """
    按照所搜的各部分组件数量，动态在数组中插入字典项并返回
    :param mode:
    :return:
    """
    list_1 = []
    for i in range(0, len(mode)-1):
        list_1.append({})
    return list_1



def build_dict_int(mode):
    """
    按照所搜的各部分组件数量，动态在数组中插入字典项并返回
    :param mode:
    :return:
    """
    list_1 = []
    for i in range(0, mode):
        list_1.append({})
    return list_1

def collecting_Inf(info,sheet_num):
    """
    复用模块，用于查询对应组下主机，应用用户，主机IP等信息，返回二维数组
    :param info: sorted(gain_num(sheet_include)[1])
    :param sheet_num: sheet_include.col_values(3)
    :return:
    """
    global host_list_data
    host_list_data = []
    for i in range(0, len(info)):
        host_list_data.append([])
    # print host_list_data

    host = []
    for i in range(0, len(info)):
        length = info[i]
        print length
        # print gain_num(sheet_include)[1][i][0]
        # print gain_num(sheet_include)[1][i][1]-1
        for j in range(info[i][0], info[i][1]):
            host.append(sheet_num[j])
            # print host
            host_list_data[i].append(host)
            host = []
    print host_list_data
    print len(host_list_data)
    return host_list_data

def coding(list):
    """
    编码模块，返回去除u之后的数组
    :param list:
    :return:
    """
    List=[]
    for each in list:
        List.append(each.encode('utf-8'))
    return List

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
        # print repeat_num
        # print 'num是',num


    return group_num,List,List_result,repeat_num

def gain_list_data(list_info):
    """
    循环查询后面七项的详细信息
    :param list_info:编码后的list
    :return:
    """
    list=[]
    for i in range(1,len(list_info)):
        list.append(list_info[i])
    return list

def cor_grou(list1,list2):
    """
    输出一个模块下具有几个分组
    target=[1,3,1,1,1,3,2,1,2,2]
    :param list1:
    :param list2:
    :return:
    """
    remain=[0]
    group_num=0
    zancun=0
    zancun2=0
    target=[]
    biao=0
    zancunshuzu=[]

    while True:
        if len(list1)>0:
            if list1[0]==list2[0]:
                group_num=group_num+1
                target.append(group_num)
                # print target
                list1.pop(0)
                list2.pop(0)
                group_num=0
            elif list1[0]>list2[0]:
                zancun=list1[0]
                list1.pop(0)
                zancun2=list2[0]
                list2.pop(0)
                while True:
                    if zancun>zancun2:
                        zancun2=zancun2+list2[0]
                        group_num=group_num+1
                        list2.pop(0)
                        print list2
                    else:
                        group_num=group_num+1
                        remain=zancun2-zancun
                        if remain>0:
                            list2.insert(0,remain)
                        # print list2
                        break
                target.append(group_num)
                # print '模块长度大于分组长度'
                group_num=0
            else:
                group_num=group_num+1
                target.append(group_num)
                remain=list2[0]-list1[0]
                # print remain
                list1.pop(0)
                list2.pop(0)
                list2.insert(0,remain)
                group_num=0
                # print list2
        else:
            break

    return target

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

def group_len(list1,list2):
    """
    在模块的约束下，计算每个分组占用的空间

    :param list1:
    :param list2:  0-13
    :param list3:
    :return:
    list1=[10,62,4,6,7,66,4,10,22,25]
    list2=[10,2,1,76,2,1,63,1,3,10,1,21,1,25]
    list3=[1,3,1,1,1,3,2,1,2,2]
    target=[10,2,1,59,]
    """
    target=[]
    tmp=0
    remain=0
    while True:
        if len(list1)>0:
            if list1[0]==list2[0]:
                target.append(list2[0])

                list1.pop(0)
                list2.pop(0)
            elif list1[0] > list2[0]:
                while True:
                    if list1[0]>list2[0]:
                        target.append(list2[0])
                        list1[0]=list1[0]-list2[0]
                        list2.pop(0)
                    elif list1[0]==list2[0]:
                        target.append(list2[0])
                        list1.pop(0)
                        list2.pop(0)
                        break
                    else:
                        target.append(list1[0])
                        remain=list2[0]-list1[0]
                        list1.pop(0)
                        list2.pop(0)
                        list2.insert(0, remain)
                        remain=0
                        break

            else:
                while True:
                    if list1[0] < list2[0]:
                        target.append(list1[0])
                        remain=list2[0]-list1[0]
                        list1.pop(0)
                        list2.pop(0)
                        list2.insert(0, remain)
                        remain=0
                    elif list1[0]==list2[0]:
                        target.append(list1[0])
                        list1.pop(0)
                        list2.pop(0)
                        break
                    else:
                        target.append(list2[0])
                        remain=list1[0]-list2[0]
                        list2.pop(0)
                        list1.pop(0)
                        list1.insert(0, remain)
                        remain=0
                        break
        else:
            break
    return target

def gain_group_num(list1,list2):
    """
    根据模块下的分组个数，取出每个分组中的主机个数，输出二维数组
    :param list1:  [1, 3, 1, 1, 1, 3, 2, 1, 2, 2]     Fulling(coding(sheet_include.col_values(2)))[3]
    :param list2: [10, 2, 1, 59, 4, 6, 7, 2, 1, 63, 1, 3, 10, 1, 21, 1, 24] group_len(Fulling(coding(sheet_include.col_values(1)))[3],Fulling(coding(sheet_include.col_values(2)))[3])
    :return:
    """
    target=[]
    target_list = []
    num=0
    while True:
        if len(list1)>0:
            num=list1[0]
            # print 'num为',num
            for i in range(0,num):
                target.append(list2[0])
                list2.pop(0)
            # print '第%d次的list2为'%i,list2
            target_list.append(target)
            target=[]
            list1.pop(0)
            # print list1


        else:
            break
    return target_list


def entire_group(list1,list2):
    """
    target=['05', '01', '02', '05','05','05','05','01',
    '02','05','GWA1','GWA2','05','02','05','01','05']
    :param list1:
    :param list2:
    :return:
    """
    target=[]
    for i in range(0,len(list1)):   #0-10
        for j in range(0,len(list1[i])):  #0-3
            target.append(list2[0])
            for k in range(0,list1[i][j]):
                list2.pop(0)
    return target



def gain_data(file_location,system):
    """
    读取excel数据集
    :param file_local:
    :return:
    """
    Excelfile=open_workbook(file_location)
    Sheet_name=Excelfile.sheet_names()
    # for sheet in Sheet_name:
    #     for each_sys in system:
    #         if sheet==:
    #             sheet_include=Excelfile.sheet_by_name(system)
    #
    #         else:
    #             continue    ##预留处理多sheet情况
    if Sheet_name[0]==system:
        sheet_include=Excelfile.sheet_by_name(system)
        Nrows=sheet_include.nrows
        Ncols=sheet_include.ncols
        '''获取系统名称列表'''
        sys_name=sheet_include.col_values(0)[1]
        print sys_name
        '''获取模块名称列表'''

        mod_name=sheet_include.col_values(1)
        print mod_name
        mod_name_code=[]
        for each in mod_name:
            mod_name_code.append(each.encode('utf-8'))


        print handle_repeat(mod_name_code)

        '''获取组名称列表'''
        group_name=sheet_include.col_values(2)
        print len(group_name)
        unrepeat_group=handle_repeat(group_name)
        print unrepeat_group
        unrepeat_group_list=[]
        for each in unrepeat_group:
            unrepeat_group_list.append(each.encode('utf-8'))




        '''获取模块，组的的列索引字典'''
        print "模块索引：",gain_num(sheet_include)[0]

        print "分组索引：", sorted(gain_num(sheet_include)[1])
        print '取消单元格的分组索引',sheet_include.col_values(2)



        '''获取所有组下7个字典项内容'''
        print 'sheet_num是',sheet_include.col_values(3)
        print sorted(gain_num(sheet_include)[1])
        print coding(sheet_include.col_values(3))



        '''开始构建字典'''
        global dict
        print '模块数量为',Fulling(coding(sheet_include.col_values(1)))[0]
        dict={"systemName":"",
              "modules":
                  build_dict_int(Fulling(coding(sheet_include.col_values(1)))[0])    #在明确模块个数的情况下，构造数组里面嵌入字典
              }
        '''填充次级项'''
        dict["systemName"]=sys_name
        Mod_name=Fulling(coding(sheet_include.col_values(1)))[2]
        Group_name_all=Fulling(coding(sheet_include.col_values(2)))[1]
        # Group_name=['05', '01', '02', '05','05','05','05','01','02','05','GWA1','GWA2','05','02','05','01','05']


        Host_name=Fulling(coding(sheet_include.col_values(3)))[1]
        manageIp=Fulling(coding(sheet_include.col_values(8)))[1]
        userName=Fulling(coding(sheet_include.col_values(4)))[1]
        serviceIp=Fulling(coding(sheet_include.col_values(5)))[1]
        # port=Fulling(coding(sheet_include.col_values(6)))[1]
        port = Fulling(sheet_include.col_values(6))[1]
        subunitName=Fulling(coding(sheet_include.col_values(7)))[1]
        unitName=Fulling(coding(sheet_include.col_values(9)))[1]


        mokuaichangdu_list = Fulling(coding(sheet_include.col_values(1)))[3]
        fenzuchangdu_list = Fulling(coding(sheet_include.col_values(2)))[3]
        Mokuai_inclu_fenzu=cor_grou(Fulling(coding(sheet_include.col_values(1)))[3],Fulling(coding(sheet_include.col_values(2)))[3])
        Mokuai_num=Fulling(coding(sheet_include.col_values(1)))[0]
        fenzu_num_eachgroup=group_len(Fulling(coding(sheet_include.col_values(1)))[3],Fulling(coding(sheet_include.col_values(2)))[3])

        Tianchong_fenzu=gain_group_num(
            cor_grou(Fulling(coding(sheet_include.col_values(1)))[3],Fulling(coding(sheet_include.col_values(2)))[3]),
            group_len(Fulling(coding(sheet_include.col_values(1)))[3],Fulling(coding(sheet_include.col_values(2)))[3]))

        Group_name = entire_group(Tianchong_fenzu,Group_name_all)

        print '模块名为',Group_name
        for i in range(0, Mokuai_num):  # 准备遍历每个模块,10个
            dict["modules"][i]["moduleName"]=Mod_name[i]
            dict["modules"][i]["groups"]=build_dict_int(Mokuai_inclu_fenzu[i])    ##明确分组个数的情况下，构造数组里面嵌入字典
            for grou_i in range(0, Mokuai_inclu_fenzu[i]):  ###遍历每个模块中的分组
                dict['modules'][i]["groups"][grou_i]["hosts"]=build_dict_int(Tianchong_fenzu[i][grou_i])
                if len(Group_name)>0:
                    dict['modules'][i]["groups"][grou_i]["groupName"] = Group_name[0]
                    Group_name.pop(0)
                    print '剩余的数组为',Group_name
                for grou_k in range(0,Tianchong_fenzu[i][grou_i]):
                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["hostName"]=Host_name[0]
                    Host_name.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["manageIp"] = manageIp[0]
                    manageIp.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["userName"] = userName[0]
                    userName.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["serviceIp"] = serviceIp[0]
                    serviceIp.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["port"] = port[0]
                    port.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["subunitName"] = subunitName[0]
                    subunitName.pop(0)

                    dict['modules'][i]["groups"][grou_i]["hosts"][grou_k]["unitName"] = unitName[0]
                    unitName.pop(0)


        print dict
        return dict

if __name__=="__main__":
    result=gain_data('ZCFES-0806-pit.xls','ZCFES')
    print type(result)

    with open('test.json', 'a') as f:
        f.write(json.dumps(result))


