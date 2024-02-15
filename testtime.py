import os
import sys
import threading
import gitlab
import xlwt
import datetime
from dateutil.relativedelta import relativedelta
import const

const.WEEK1_COMMIT_INDEX = 0
const.WEEK1_ADD_INDEX = 1
const.WEEK1_DEL_INDEX = 2
const.WEEK1_CHG_INDEX = 3
const.WEEK2_COMMIT_INDEX = 4
const.WEEK2_ADD_INDEX = 5
const.WEEK2_DEL_INDEX = 6
const.WEEK2_CHG_INDEX = 7
const.WEEK3_COMMIT_INDEX = 8
const.WEEK3_ADD_INDEX = 9
const.WEEK3_DEL_INDEX = 10
const.WEEK3_CHG_INDEX = 11
const.WEEK4_COMMIT_INDEX = 12
const.WEEK4_ADD_INDEX = 13
const.WEEK4_DEL_INDEX = 14
const.WEEK4_CHG_INDEX = 15
const.TOTAL_COMMIT_INDEX = 16
const.TOTAL_ADD_INDEX = 17
const.TOTAL_DEL_INDEX = 18
const.TOTAL_CHG_INDEX = 19
const.TOTAL_RECORD_ITEM_NUM = 20

#写入execl
def WriteAllDataToExcel(excelPath, statistic_time, data):
    save_path_prefix = excelPath + "\\" + statistic_time.strftime('%Y-%m-')
    for one_groupname in data.keys():
        if one_groupname == 'timesplitarray':
            pass
        else:
            workbook = xlwt.Workbook()
            # 写第一个sheet页，主要是群组的用户信息
            sheet = workbook.add_sheet('groupmembers')
            row0 = ['用户名', '是否为群组成员', '第1周提交次数', '第1周新增代码', '第1周删除代码', '第1周总计代码', \
                '第2周提交次数', '第2周新增代码', '第2周删除代码', '第2周总计代码', \
                '第3周提交次数', '第3周新增代码', '第3周删除代码', '第3周总计代码', \
                '第4周提交次数', '第4周新增代码', '第4周删除代码', '第4周总计代码', \
                '本月提交次数', '本月新增代码', '本月删除代码', '本月总计代码']
            
            for i in range(0, len(row0)):
                sheet.write(0, i, row0[i])
            
            member_info = data[one_groupname]['all_group_members']
            i = 1
            for one_username in member_info.keys():
                j = 0
                sheet.write(i, j, one_username)
                j += 1
                sheet.write(i, j, str(member_info[one_username]['belongto']))
                for j in range(0,const.TOTAL_RECORD_ITEM_NUM):
                    sheet.write(i, j+2, str(member_info[one_username]['statistic'][j]))
                i += 1
            
            # 写后面每个项目的sheet页
            for one_project_name in data[one_groupname].keys():
                if one_project_name == 'all_group_members':
                    pass
                else:
                    sheet = workbook.add_sheet(one_project_name)
                    row0 = ['用户名', '第1周提交次数', '第1周新增代码', '第1周删除代码', '第1周总计代码', \
                        '第2周提交次数', '第2周新增代码', '第2周删除代码', '第2周总计代码', \
                        '第3周提交次数', '第3周新增代码', '第3周删除代码', '第3周总计代码', \
                        '第4周提交次数', '第4周新增代码', '第4周删除代码', '第4周总计代码', \
                        '本月提交次数', '本月新增代码', '本月删除代码', '本月总计代码']
                    for i in range(0, len(row0)):
                        sheet.write(0, i, row0[i])
                    
                    project_info = data[one_groupname][one_project_name]
                    i = 1
                    for one_username in project_info.keys():
                        j = 0
                        sheet.write(i, j, one_username)
                        for j in range(0,const.TOTAL_RECORD_ITEM_NUM):
                            sheet.write(i, j+1, str(project_info[one_username][j]))
                        i += 1
            # 生成excel文件名称
            save_filename = save_path_prefix + one_groupname + ".xls"
            workbook.save(save_filename)
 

if __name__ == '__main__':
    curtime = datetime.datetime.now()
    strtime = curtime.strftime('%Y-%m-%d 0:0:0')
    curtime = datetime.datetime.strptime(strtime,'%Y-%m-%d %H:%M:%S')
    sincetime = curtime - relativedelta(months=+1)
    untiltime = curtime - relativedelta(days=+1)
    print(sincetime)
    print(untiltime)
    #初始化data结构
    time_split_array = []
    time_split_array.append(sincetime+relativedelta(days=+6))
    time_split_array.append(sincetime+relativedelta(days=+13))
    time_split_array.append(sincetime+relativedelta(days=+20))
    time_split_array.append(untiltime)
    print(time_split_array)
    git_commit_timestr = "2021-09-20T11:50:22.001+00:00"
    git_commit_timestr = git_commit_timestr.split('.')[0]
    print(git_commit_timestr)
    testtime = datetime.datetime.strptime(git_commit_timestr,'%Y-%m-%dT%H:%M:%S')
    print(testtime)
    excelPath = 'E:\codes\python-gitlab'
    save_path_prefix = excelPath + "\\" + sincetime.strftime('%Y-%m-')
    print(save_path_prefix)
    
    data = {}
    dataLock = threading.Lock()
    thread_list = []
    # print(gitlab_projects_dict)
 
    #初始化data结构
    time_split_array = []
    time_split_array.append(sincetime+relativedelta(days=+6))
    time_split_array.append(sincetime+relativedelta(days=+13))
    time_split_array.append(sincetime+relativedelta(days=+20))
    time_split_array.append(untiltime)
    gitlab_projects_dict = dict(protocol = ['wenos-proto', 'ios-proto', 'sp5148-HAA', 'testproj' ], sdkdev = ['sdi-sdk', 'pcie-sdk', 'se5154-sdk', 'prb0400-linux', 'prb0400-acore'])
    data['timesplitarray'] = time_split_array
    for each_group in gitlab_projects_dict.keys():
        data[each_group] = dict()
        data[each_group]['all_group_members'] = dict()
        for each_project in gitlab_projects_dict[each_group]:
            data[each_group][each_project] = dict()
            data[each_group][each_project]['张三'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['张六六'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['王麻子'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['刘伟'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['王天琪'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['张明伟'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['罗秀春'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['刘翔'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['王伟'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['张晓雪'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['tommy'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['vector'] = [1, 3, 5, 8, 2, 6, 3, 9, 5, 100, 20, 120, 4, 2000, 150, 2150, 12, 2109, 178, 2287]
            data[each_group][each_project]['total_commit_data'] = [12, 36, 60, 96, 24, 72, 36, 108, 60, 1200, 240, 1440, 48, 24000, 1800, 25800, 144, 25308, 2136, 27444]
           
    
    #完成初始化data结构
    
    # 先初始化群组内的所有具有直接归属关系的用户名
    temp_user_dict = dict(protocol = ['张三', '张六六', '王麻子', '刘伟', '王天琪'], sdkdev = ['张明伟', '罗秀春', '刘翔', '王伟'])
    for one_groupname in temp_user_dict.keys():
        if temp_user_dict[one_groupname] == "":
            pass
        else:
            for one_member in temp_user_dict[one_groupname]:
                if not one_member in data[one_groupname]['all_group_members'].keys():
                    data[one_groupname]['all_group_members'][one_member] = dict( belongto = True, statistic = [0] * const.TOTAL_RECORD_ITEM_NUM )
        
        # 计算群组内各个项目的用户合计提交数据量，如果用户名不存在，则创建新的用户名
        for one_project in data[one_groupname].keys():
            if one_project == 'all_group_members':
                pass
            else:
                for one_username in data[one_groupname][one_project].keys():
                    if one_username == 'total_commit_data':
                        pass
                    else:
                        if not one_username in data[one_groupname]['all_group_members'].keys():
                            data[one_groupname]['all_group_members'][one_username] = dict( belongto = False, statistic = [0] * const.TOTAL_RECORD_ITEM_NUM )
                        
                        for i in range(0,const.TOTAL_RECORD_ITEM_NUM):
                            data[one_groupname]['all_group_members'][one_username]['statistic'][i] += data[one_groupname][one_project][one_username][i]
    #完成群组内的所有成员工作量计算
    
    WriteAllDataToExcel('E:\codes\python-gitlab', sincetime, data)