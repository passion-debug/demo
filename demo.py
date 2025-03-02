import os
import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
import copy
import numpy as np

def process_excel():
    #获得当前地址
    cwd = os.getcwd()
    #创建结果文件
    file_name_in_this_path = [file for file in os.listdir(cwd) if file == 'result' and os.path.isdir(file)]
    result_path = os.path.join(cwd,'result')
    if len(file_name_in_this_path) == 0:
        os.makedirs(result_path)

    #输入文件地址
    data_file_path = os.path.join(cwd,'报账表格.xlsx')

    #sheet_name 报账单 ，通讯录

    check_sheet = pd.read_excel(data_file_path,sheet_name='报账单')
    address_book = pd.read_excel(data_file_path,sheet_name='通讯录')
    student_data = defaultdict(dict)
    for index_X in range(address_book.shape[0]):
        student_data_cloums ={_key:None   for _key in ['学号','联系电话']}
        if isinstance(address_book.iloc[index_X,0],int):
            student_data_cloums['学号'] = address_book.iloc[index_X,2]
            student_data_cloums['联系电话'] = address_book.iloc[index_X,4]
            student_data[address_book.iloc[index_X,1]] = student_data_cloums
    check_sheet_dict = defaultdict(list)
    

    for index_X in range(check_sheet.shape[0]):
        student_check_data_dict = {_key:None   for _key in ['姓名',	'报销类别',	'明细',	'金额',	'有无发票',	'报销资金来源',	'发票张数','消费记录数','运单明细','分类']}
        if   pd.isna(check_sheet.loc[index_X,'姓名']) == False:
            last_name = check_sheet.loc[index_X,'姓名']
        if pd.isna(check_sheet.loc[index_X,'报销类别']) == False:
            last_LABEL = check_sheet.loc[index_X,'报销类别']
        if pd.isna(check_sheet.loc[index_X,'报销资金来源']) == False:
            last_SOURCE = check_sheet.loc[index_X,'报销资金来源']
        check_sheet.loc[index_X,'日期'] = check_sheet.loc[index_X,'日期'].date()
        if check_sheet.loc[index_X,'有无发票'] == '有':#heck_sheet.loc[index_X,'办理情况'] != '已办理' and 
            student_check_data_dict['姓名'] = last_name
            student_check_data_dict['报销类别'] = last_LABEL
            student_check_data_dict['报销资金来源'] = last_SOURCE
            student_check_data_dict['明细'] = check_sheet.loc[index_X,'明细']
            student_check_data_dict['金额'] = check_sheet.loc[index_X,'金额']
            student_check_data_dict['有无发票'] = check_sheet.loc[index_X,'有无发票']
            # student_check_data_dict['小计'] = check_sheet.loc[index_X,'小计']
            # student_check_data_dict['学号'] = check_sheet.loc[index_X,'学号']
            check_sheet_dict[last_name].append(student_check_data_dict)
            student_check_data_dict['发票张数'] = check_sheet.loc[index_X,'发票张数']
            student_check_data_dict['运单明细'] = check_sheet.loc[index_X,'运单明细']
            student_check_data_dict['消费记录数'] = check_sheet.loc[index_X,'消费记录数']
            student_check_data_dict['分类'] = check_sheet.loc[index_X,'分类']
        else:
            continue
            
        
    result_dict = defaultdict(dict)
    for _key in check_sheet_dict:
        project_dict = defaultdict(dict)
        for item in check_sheet_dict[_key]:
            if item['有无发票'] != '有':
                continue
            
            if item['报销资金来源'] not in project_dict[item['报销类别']]:
                project_dict[item['报销类别']][item['报销资金来源']] = {}
            if item['有无发票'] not in project_dict[item['报销类别']][item['报销资金来源']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']] = {}

            if'发票张数' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['发票张数'] = 0
            ori_number = project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['发票张数']
            if ori_number <= 15 and ori_number + item['发票张数'] > 15:
                result_dict.setdefault(_key,list()).append(project_dict)
                project_dict = defaultdict(dict)
                if item['报销资金来源'] not in project_dict[item['报销类别']]:
                    project_dict[item['报销类别']][item['报销资金来源']] = {}
                if item['有无发票'] not in project_dict[item['报销类别']][item['报销资金来源']]:
                    project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']] = {}
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['发票张数'] = 0

            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['发票张数'] += item['发票张数']

            if'运单明细' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['运单明细'] = 0
            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['运单明细'] += item['运单明细']

            if'明细' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['明细'] = []
            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['明细'].append({item['明细']:item['金额']})

            if'消费记录数' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['消费记录数'] = 0
            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['消费记录数'] += item['消费记录数']

            if'分类' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['分类'] = []
            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['分类'].append({item['明细']:item['分类']})    
            
            if'小计' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['小计'] = 0
            project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['小计'] += item['金额']
            if'学号' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['学号'] = student_data.get(_key,{}).get('学号','查无此人')
            if'联系电话' not in project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]:
                project_dict[item['报销类别']][item['报销资金来源']][item['有无发票']]['联系电话'] = student_data.get(_key,{}).get('联系电话','查无此人')
                
            
                
        result_dict.setdefault(_key,list()).append(project_dict)
    project_dict = dict()
    result_dict2 = copy.deepcopy(result_dict)
    for name in result_dict2:
        for indx,source_dict in enumerate(result_dict2[name]):
            for source in source_dict:
                project_dict.setdefault(source,dict()).setdefault(indx,dict())
                for project in source_dict[source]:
                    if project not in project_dict[source]:
                        project_dict[source][indx][project] = {'明细':source_dict[source][project]['有']['明细'],'小计':source_dict[source][project]['有']['小计']}
                    else:
                        project_dict[source][indx][project]['明细'] += source_dict[source][project]['有']['明细']
                        # project_dict[source][project]['小计'] += result_dict[name][source][project]['有']['小计']
    project_list = list()
    for source in project_dict:
        for indx in project_dict[source]:
            for project in project_dict[source][indx]:
                project_list.append([project,source, ' '.join(['{}:{}'.format(list(item.keys())[0],list(item.values())[0])  for item in  project_dict[source][indx][project]['明细']]),project_dict[source][indx][project]['小计']]) 
        
    res = pd.DataFrame()
    data_list = list()
    for name in result_dict:
        for  data_dict in result_dict[name]:
            for label in data_dict:
                for source in data_dict[label]:
                    for invoice in data_dict[label][source]:
                        new_data = {'姓名':name,'报销类别':label,'报销资金来源':source,'有无发票':invoice,
                                    '小计':data_dict[label][source][invoice]['小计'],
                                '明细':' '.join(['{}:{}'.format(list(item.keys())[0],list(item.values())[0])  for item in  data_dict[label][source][invoice]['明细']]),'学号':data_dict[label][source][invoice]['学号'],'联系电话':str(data_dict[label][source][invoice]['联系电话']),
                                '发票张数':data_dict[label][source][invoice]['发票张数'],'消费记录数':data_dict[label][source][invoice]['消费记录数'],'运单明细':data_dict[label][source][invoice]['运单明细'],'分类':data_dict[label][source][invoice]['分类'],
                                '票据数小计':data_dict[label][source][invoice]['发票张数'] + data_dict[label][source][invoice]['消费记录数'] + data_dict[label][source][invoice]['运单明细']}
                        data_list.append(new_data)
                        
    data_list2 = sorted(data_list,key= lambda x:(x['报销资金来源'],x['分类']))    
    add_data_dict = dict() 
    for item in data_list2:
        if  pd.isna(item['分类']) == False:           
            if item['报销资金来源'] not in add_data_dict:
                add_data_dict[item['报销资金来源']] = {}
            if item['分类'] not in add_data_dict[item['报销资金来源']]:
                add_data_dict[item['报销资金来源']][item['分类']] = {}
                add_data_dict[item['报销资金来源']][item['分类']]['票据数合计'] = 0
                add_data_dict[item['报销资金来源']][item['分类']]['金额合计'] = 0
            add_data_dict[item['报销资金来源']][item['分类']]['金额合计'] = add_data_dict[item['报销资金来源']][item['分类']]['金额合计'] + item['小计'] 
            add_data_dict[item['报销资金来源']][item['分类']]['票据数合计'] = add_data_dict[item['报销资金来源']][item['分类']]['票据数合计'] + item['发票张数'] + item['消费记录数'] + item['运单明细']
    write = set()
    data_list3 = copy.deepcopy(data_list2)
    for _item in data_list2:
        if  pd.isna(_item['分类']) == False: 
            if (_item['报销资金来源'],_item['分类']) not in write:
                _item['金额合计'] = add_data_dict[_item['报销资金来源']][_item['分类']]['金额合计']
                _item['票据数合计'] = add_data_dict[_item['报销资金来源']][_item['分类']]['票据数合计']
                write.add((_item['报销资金来源'],_item['分类']))
            else:
                _item['金额合计'] =  np.nan
                _item['票据数合计'] = np.nan
        else:
                _item['金额合计'] =  np.nan
                _item['票据数合计'] = np.nan
    
    for item in data_list3:
        c = 1
    
    
    res = pd.concat([res, pd.DataFrame(data_list2)], ignore_index=True)
    res.columns = ['姓名','报销类别','报销资金来源','有无发票','小计','明细','学号','联系电话','发票张数','消费记录数','运单明细','分类','票据数小计','金额合计','票据数合计']
    res2 = copy.deepcopy(res)
    res2 = res2.sort_values(by=['报销资金来源','报销类别','分类','金额合计',],ascending=[True,True,True,False])
    res2 = res2[['姓名','报销类别','报销资金来源','有无发票','小计','金额合计','明细','学号','联系电话','发票张数','消费记录数','运单明细','票据数小计','票据数合计','分类']]
    file_path = os.path.join(result_path,'process Table.xlsx')
    res3 = pd.DataFrame(project_list)
    res3.columns = ['报销资金来源','报销类别','明细','小计']
    res3 = res3.sort_values(by=['报销资金来源','报销类别','小计'],ascending=[True,True,False])
    # res.to_excel(file_path,sheet_name='总计')
    with pd.ExcelWriter(file_path) as writer:
        # check_sheet.to_excel(writer, sheet_name='报账单', index=False)
        res2.to_excel(writer, sheet_name='总计_by项目', index=False)
        res.to_excel(writer, sheet_name='总计_by人员', index=False)
        res3.to_excel(writer, sheet_name='项目总计', index=False)
        

        


if __name__ == '__main__':   
    process_excel()

  
