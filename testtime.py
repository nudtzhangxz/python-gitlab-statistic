import os
import sys
import threading
import gitlab
import xlwt
import datetime
from dateutil.relativedelta import relativedelta

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
