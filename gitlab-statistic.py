import os
import sys
import threading
import gitlab
import xlwt
import datetime
from dateutil.relativedelta import relativedelta
import const

#获取所有的群组
def getAllGroups():
    groups_objlist = []
    client = gitlab.Gitlab(private_host, private_token=private_token)
    groups = client.groups.list(all=True)
    for group in groups:
        groups_objlist.append(group)
    return groups_objlist
  
#获取所有群组下的所有的project
def getAllProjects(groups_objlist):
    projects_dict = { }
    #client = gitlab.Gitlab(private_host, private_token=private_token)
    #projects = client.projects.list(all=True)
    for group in groups_objlist:
        group_projects = group.projects.list(all=True)
        projects_dict[group.name] = group_projects
    return projects_dict


#获取所有群组下的所有的user
def getAllUsers(groups_objlist):
    users_dict = { }
    #client = gitlab.Gitlab(private_host, private_token=private_token)
    #users = client.users.list(all=True)
    for group in groups_objlist:
        try:
            group_members = group.members.list(all=True)
            users_dict[group.name] = group_members
        except:
            users_dict[group.name] = ""
    return users_dict


#获取project下所有的branche
def getAllBranchByProject(project):
    try:
        branches = project.branches.list()
        return branches
    except:
        return ""
 
#获取project和branch下的commit
def getCommitByBranch(project, branch, begin_time, end_time):
    author_commits = []
    commits = project.commits.list(all=True, ref_name=branch.name, since=begin_time.strftime('%Y-%m-%d %H:%M:%S'), until=end_time.strftime('%Y-%m-%d %H:%M:%S'))
    for commit in commits:
        #committer_email = commit.committer_email
        #title = commit.title
        #message = commit.message
        #if ('Merge' in message) or ('Merge' in title):
        #    print('Merge跳过')
        #    continue
        #else:
        author_commits.append(commit)
    return author_commits
 
#获取project项目下commit对应的code
def getCodeByCommit(commit, project):
    commit_info = project.commits.get(commit.id)
    code = commit_info.stats
    return code
 
#  data dictorary is defined as following
#  data['timesplitarray'] --> timedata[4]  ---> timedata[0] is end time of week1, timedata[1] is end time of week2,... timedata[3] is end time of week4
#  data[groupname] --> groupdata{}
#  
#  groupdata['all_group_members'] --> memberinfo{}         
#     memberinfo[username] --> memberdata{}
#        memberdata['belongto'] ---> True/False
#        memberdata['statistic'] ---> user_commit_list[4*5] 
#        user_commit_list[4*5]--->[commitcount at week1,total add lines at week1, total del lines at week1, total changed lines at week1,..... week4, and four total count for month]
#  groupdata[project.name] --> projectdata{}
#     projectdata['total_commit_data'] --> total_commit_data[4*5]---> defined the same as user_commit_list
#     projectdata[username] --> user_commit_data[4*5]---> defined the same as previous user_commit_list
#
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


def getOneProjectAuthorCode(groupname, project, begin_time, end_time):
    # print("project:%s" % project)
    one_project_data = dict()
    one_project_data['total_commit_data'] = [0] * const.TOTAL_RECORD_ITEM_NUM
    
    branches = getAllBranchByProject(project)
    if branches == "":
        pass
    else:
        for branch in branches:
            #print("branch#####",branch.name)
            #print("branch:%s" % branch)
            #print('获取工程', project.name, '分支', branch.name, "的提交记录")
            author_commits = getCommitByBranch(project, branch,begin_time,end_time)
            # print(author_commits)
            for commit in author_commits:
                #print('获取提交', commit.id, "的代码量")
                # committer_name 可以替换为 author_name
                temp_user_name = commit.committer_name
                if not temp_user_name in one_project_data.keys():
                    one_project_data[temp_user_name] = [0] * const.TOTAL_RECORD_ITEM_NUM
                
                temp_commit_date = commit.committed_date
                temp_commit_date = temp_commit_date.split('.')[0]
                commit_time = datetime.datetime.strptime(temp_commit_date,'%Y-%m-%dT%H:%M:%S')
                
                code = getCodeByCommit(commit, project)
                temp_addacount = int(code['additions'])
                temp_deletecount = int(code['deletions'])
                temp_totalcount = int(code['total'])
                one_project_data['total_commit_data'][const.TOTAL_COMMIT_INDEX] += 1
                one_project_data['total_commit_data'][const.TOTAL_ADD_INDEX] += temp_addacount
                one_project_data['total_commit_data'][const.TOTAL_DEL_INDEX] += temp_deletecount
                one_project_data['total_commit_data'][const.TOTAL_CHG_INDEX] += temp_totalcount
                
                one_project_data[temp_user_name][const.TOTAL_COMMIT_INDEX] += 1
                one_project_data[temp_user_name][const.TOTAL_ADD_INDEX] += temp_addacount
                one_project_data[temp_user_name][const.TOTAL_DEL_INDEX] += temp_deletecount
                one_project_data[temp_user_name][const.TOTAL_CHG_INDEX] += temp_totalcount
               
                if commit_time <= data['timesplitarray'][0]:
                    one_project_data['total_commit_data'][const.WEEK1_COMMIT_INDEX] += 1
                    one_project_data['total_commit_data'][const.WEEK1_ADD_INDEX] += temp_addacount
                    one_project_data['total_commit_data'][const.WEEK1_DEL_INDEX] += temp_deletecount
                    one_project_data['total_commit_data'][const.WEEK1_CHG_INDEX] += temp_totalcount
                    
                    one_project_data[temp_user_name][const.WEEK1_COMMIT_INDEX] += 1
                    one_project_data[temp_user_name][const.WEEK1_ADD_INDEX] += temp_addacount
                    one_project_data[temp_user_name][const.WEEK1_DEL_INDEX] += temp_deletecount
                    one_project_data[temp_user_name][const.WEEK1_CHG_INDEX] += temp_totalcount
                elif commit_time <= data['timesplitarray'][1]:
                    one_project_data['total_commit_data'][const.WEEK2_COMMIT_INDEX] += 1
                    one_project_data['total_commit_data'][const.WEEK2_ADD_INDEX] += temp_addacount
                    one_project_data['total_commit_data'][const.WEEK2_DEL_INDEX] += temp_deletecount
                    one_project_data['total_commit_data'][const.WEEK2_CHG_INDEX] += temp_totalcount
                    
                    one_project_data[temp_user_name][const.WEEK2_COMMIT_INDEX] += 1
                    one_project_data[temp_user_name][const.WEEK2_ADD_INDEX] += temp_addacount
                    one_project_data[temp_user_name][const.WEEK2_DEL_INDEX] += temp_deletecount
                    one_project_data[temp_user_name][const.WEEK2_CHG_INDEX] += temp_totalcount
                elif commit_time <= data['timesplitarray'][2]:
                    one_project_data['total_commit_data'][const.WEEK3_COMMIT_INDEX] += 1
                    one_project_data['total_commit_data'][const.WEEK3_ADD_INDEX] += temp_addacount
                    one_project_data['total_commit_data'][const.WEEK3_DEL_INDEX] += temp_deletecount
                    one_project_data['total_commit_data'][const.WEEK3_CHG_INDEX] += temp_totalcount
                    
                    one_project_data[temp_user_name][const.WEEK3_COMMIT_INDEX] += 1
                    one_project_data[temp_user_name][const.WEEK3_ADD_INDEX] += temp_addacount
                    one_project_data[temp_user_name][const.WEEK3_DEL_INDEX] += temp_deletecount
                    one_project_data[temp_user_name][const.WEEK3_CHG_INDEX] += temp_totalcount
                else:
                    one_project_data['total_commit_data'][const.WEEK4_COMMIT_INDEX] += 1
                    one_project_data['total_commit_data'][const.WEEK4_ADD_INDEX] += temp_addacount
                    one_project_data['total_commit_data'][const.WEEK4_DEL_INDEX] += temp_deletecount
                    one_project_data['total_commit_data'][const.WEEK4_CHG_INDEX] += temp_totalcount
                    
                    one_project_data[temp_user_name][const.WEEK4_COMMIT_INDEX] += 1
                    one_project_data[temp_user_name][const.WEEK4_ADD_INDEX] += temp_addacount
                    one_project_data[temp_user_name][const.WEEK4_DEL_INDEX] += temp_deletecount
                    one_project_data[temp_user_name][const.WEEK4_CHG_INDEX] += temp_totalcount            

    dataLock.acquire()
    data[groupname][project.name] = one_project_data
    dataLock.release()
    return data

#写入execl
def writeExcel(excelPath, data):
    workbook = xlwt.Workbook()
    # 获取第一个sheet页
    sheet = workbook.add_sheet('git')
    row0 = ['项目组', '工程名称', '分支名称', '提交次数', '新增代码', '删除代码', '总计代码']
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i])
    addcount = 0
    delcount = 0
    totalcount = 0
    commitcount = 0
    for i in range(0, len(data)):
        recode = data[i]
        j = 0
        sheet.write(i + 1, j, recode['group'])
        sheet.write(i + 1, j + 1, recode['projectname'])
        sheet.write(i + 1, j + 2, recode['branchename'])
        commitcount += (int)(recode['commitcount'])
        sheet.write(i + 1, j + 3, recode['commitcount'])
        addcount += (int)(recode['codecount']['additions'])
        sheet.write(i + 1, j + 4, recode['codecount']['additions'])
        delcount += (int)(recode['codecount']['deletions'])
        sheet.write(i + 1, j + 5, recode['codecount']['deletions'])
        totalcount += (int)(recode['codecount']['total'])
        sheet.write(i + 1, j + 6, recode['codecount']['total'])
 
    sheet.write(len(data) + 1, 3, commitcount)
    sheet.write(len(data) + 1, 4, addcount)
    sheet.write(len(data) + 1, 5, delcount)
    sheet.write(len(data) + 1, 6, totalcount)
    workbook.save(excelPath)
 
 
if __name__ == '__main__':
    # 用户git账户的token 6S7jy689FeCrP5w_UwgZ
    private_token = 'T3Nz2xCxq4FcVQ4wytr1'       #gitlab用户tonken
    # git地址
    private_host = 'http://10.0.0.1:8888/'       #gitlab地址
    curtime = datetime.datetime.now()
    strtime = curtime.strftime('%Y-%m-%d 0:0:0')
    curtime = datetime.datetime.strptime(strtime,'%Y-%m-%d %H:%M:%S')
    sincetime = curtime - relativedelta(months=+1)
    untiltime = curtime - relativedelta(days=+1)
    data = {}
    dataLock = threading.Lock()
    thread_list = []
    all_gitlab_groups = getAllGroups()
    gitlab_projects_dict = getAllProjects(all_gitlab_groups)
    # print(gitlab_projects_dict)
 
    #初始化data结构
    time_split_array = []
    time_split_array.append(sincetime+relativedelta(days=+6))
    time_split_array.append(sincetime+relativedelta(days=+13))
    time_split_array.append(sincetime+relativedelta(days=+20))
    time_split_array.append(untiltime)
    
    data['timesplitarray'] = time_split_array
    for each_group in gitlab_projects_dict.keys():
        data[each_group.name] = dict()
        data[each_group.name]['all_group_members'] = ""
        for each_project in gitlab_projects_dict[each_group]:
            data[each_group.name][each_project.name] = ""
    
    #完成初始化data结构
    
    for each_group in gitlab_projects_dict.keys():
         for each_project in gitlab_projects_dict[each_group]:
            t = threading.Thread(target=getOneProjectAuthorCode, args=(each_group.name,each_project,sincetime,untiltime))
            thread_list.append(t)
 
    for threadname in thread_list: threadname.start()
    for threadname in thread_list: threadname.join()
    # print(data)
    # 统计群组内的所有成员工作量
    temp_user_dict = getAllUsers(all_gitlab_groups)
    
    writeExcel('E:\codes\python-gitlab\code_count.xls', data)