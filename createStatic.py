import os
import re
import datetime
import time
import pandas as pd
import numpy as np

allIssue = []
openIssue = []
ongoingIssue = []
closeIssue = []
abnormalIssue = []
delayIssue = []
weekm1Issue = []
weekm2Issue = []
issueStatic = []

class SWIMExcel():
    def __init__(self,targetFile,ECUlist):
        self.targetFile = targetFile
        self.ECUlist = ECUlist
    def readSWIMs(self):
        SWIMs = pd.read_excel(self.targetFile)
        return SWIMs
    def readECUlist(self):
        INFO = pd.read_excel(self.ECUlist)
        return INFO
class SWIM():
    def __init__(self,**kwargs):

        self.Status_open = ("New", "Reopen", "Analysis", "Supplier inbox", "Supplier In Progress")
        self.Status_ongoing = ("Solution Identified", "Solved", "Ready for Test", "Under observation", "Testing On Hold")
        self.Status_close = ("Closed", "Cancelled")
        self.ValidTargetserie = ("VP2", "TT", "E4-2", "E4-2/VP2", "Pre TT", "PreTT", "vp2","VP")
        self.name = kwargs["name"]
        self.type = [allIssue, openIssue, ongoingIssue, closeIssue, abnormalIssue, weekm2Issue, weekm1Issue,
                     delayIssue]
        self.list = []


    def createIssue(self,num,SWIMs):
        self.Issue = {}
        self.Issue["Key"] = SWIMs["Key"][num]
        self.Issue["ECU"] = SWIMs["ECU"][num]
        self.Issue["Status"] = SWIMs["Status"][num]
        self.Issue["Created"] = SWIMs["Created"][num]
        self.Issue["Target Serie"] = SWIMs["Target Serie"][num]
        self.Issue["Found in Serie"] = SWIMs["Found in Serie"][num]
        return self.Issue
        print(self.Issue)




    def addIssue(self,num,SWIMs,INFO):
        if re.match("SWIM-", str(SWIMs["Key"][num])):
            allIssue.append(self.createIssue(num,SWIMs))
            createTime = SWIMs["Created"][num]
            # print (createTime)
            z = time.strptime(str(createTime), '%Y-%m-%d %H:%M:%S')
            createWeek = time.strftime("%W", z)
            createYear = time.strftime("%Y", z)
            today = time.localtime(time.time())
            todayWeek = time.strftime("%W", today)
            todayYear = time.strftime("%Y", today)
            if int(todayYear) - int(createYear) == 0:
                if int(todayWeek) - int(createWeek) == 1:
                    weekm1Issue.append(self.createIssue(num,SWIMs))
                elif int(todayWeek) - int(createWeek) == 2:
                    weekm2Issue.append(self.createIssue(num,SWIMs))
                else:
                    pass
            else:
                pass
            # create abnormalIssue list
            diff = datetime.datetime.now() - createTime
            if diff.days > 9 and SWIMs["Status"][num] in self.Status_open and SWIMs["ECU"][num] in list(INFO["ECU"]):
                self.createIssue(num,SWIMs)["abnormal reason"] = "delay"
                self.createIssue(num,SWIMs)["delay days"] = diff.days - 9
                delayIssue.append(self.Issue)
            elif SWIMs["Target Serie"][num] not in self.ValidTargetserie and SWIMs["Status"][num] in self.Status_ongoing \
                    and SWIMs["ECU"][num] in list(INFO["ECU"]) and SWIMs["Target Serie"][num] is not np.nan  :
                self.createIssue(num,SWIMs)["abnormal reason"] = "wrong target serie"
                abnormalIssue.append(self.Issue)
            else:
                pass
            #  create openIssue, ongoingIssue,closeIssue list
        if SWIMs["Status"][num] in self.Status_open:
            openIssue.append(self.createIssue(num,SWIMs))
        if SWIMs["Status"][num] in self.Status_ongoing:
            ongoingIssue.append(self.createIssue(num,SWIMs))
        if SWIMs["Status"][num] in self.Status_close:
            closeIssue.append(self.createIssue(num,SWIMs))

    def createIssuelist(self,SWIMs,INFO):
        i = 0
        while i < len(SWIMs["Key"]):
            self.addIssue(i,SWIMs,INFO)
            # print(openIssue,ongoingIssue,allIssue)
            i+=1
        else:
            i+=1


    def addSpecielIssue(self,num,IssueList):
        if IssueList in self.type:
            self.list.append(IssueList[num])
        else:
            pass

    def createSpecialIssuelist(self,LISTNAME, type, IssueList):
        issue_LISTNAME = []
        issue = SWIM(name=LISTNAME)
        num = 0
        if IssueList in issue.type:
            while num < len(IssueList):
                if IssueList[num][type] == LISTNAME:
                    issue.addSpecielIssue(num,IssueList)
                    num += 1
                else:
                    num += 1
        else:
            print("please input the right IssueList(allIssue,openIssue,ongoingIssue,"
                  "closeIssue,abnormalIssue,weekm2Issue,weekm1Issue)")
        issue_LISTNAME = issue.list
        return issue_LISTNAME

    def Static(self,num,INFO):
        ECU_num = {}
        # print(self.createSpecialIssuelist("SWSM","ECU",allIssue))
        # print(self.INFO["ECU"][num])
        ECU_num["ECU"] = INFO["ECU"][num]
        ECU_num["department"] = INFO["department"][num]
        ECU_num["leader"] = INFO["leader"][num]
        ECU_num["All"] = len(self.createSpecialIssuelist(INFO["ECU"][num],"ECU",allIssue))
        ECU_num["Open"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", openIssue))
        ECU_num["Onoing"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", ongoingIssue))
        ECU_num["Close"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", closeIssue))
        ECU_num["Issues in -1week"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", weekm1Issue))
        ECU_num["Issues in -2week"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", weekm2Issue))
        ECU_num["Abnormal Issues"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", abnormalIssue))
        ECU_num["Delay Issues"] = len(self.createSpecialIssuelist(INFO["ECU"][num], "ECU", delayIssue))
        if ECU_num["All"] != 0:
            issueStatic.append(ECU_num)
        else:
            pass

    def createStatic(self,INFO):
        i = 0
        while i < len(INFO["ECU"]):
            self.Static(i,INFO)
            i += 1



if __name__ == '__main__':
    targetfile ="D:\\xujunjie D盘\\项目（KX11）\\SWIM 问题推进\\KX11- 问题推进-W33D5.xlsx"
    ECUList = "D:\\xujunjie D盘\\项目（KX11）\\SWIM 问题推进\\KX11ECULIST.xlsx"
    swim = SWIM(name="swim")
    sexcel = SWIMExcel(targetfile, ECUList)
    SWIMs = sexcel.readSWIMs()
    INFO = sexcel.readECUlist()
    swim.createIssuelist(SWIMs,INFO)
    swim.createStatic(INFO)
    print(issueStatic)
    data = pd.DataFrame.from_dict(issueStatic,orient="columns")
    data1 = pd.DataFrame.from_dict(abnormalIssue,orient="columns")
    data2 = pd.DataFrame.from_dict(delayIssue,orient="columns")
    print(data)
    # print(data1)
    print(data2)
    # print(allIssue)
    # print(allIssue[10])

