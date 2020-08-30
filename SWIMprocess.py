import pandas as pd
import os
import re
import datetime
import time
from unittest import TestCase,main
from tkinter import filedialog
# ----------------------------------------------------------------------------------------------------------------------
#open the SWIM worksheet
# SWIMsourceFolder = "C:\\Users\\xujunjie2\\PycharmProjects\\EXCEL"
# ECUlistFolder= "C:\\Users\\xujunjie2\\PycharmProjects\\EXCEL"
# ECUexcelName = "KX11ECULIST.xlsx"
# SWIMexcelName = "KX11 W32D5.xlsx"
# targetFile = os.path.join(SWIMsourceFolder,SWIMexcelName)
# ECUlistFile = os.path.join(ECUlistFolder,ECUexcelName)
# targetFile = filedialog.askopenfilename()
# ECUlistFile = filedialog.askopenfilename()
# SWIM = pd.read_excel(targetFile)
# INFO= pd.read_excel(ECUlistFile)
# ----------------------------------------------------------------------------------------------------------------------
# define all the list that we need ,to make a list in future.

def createStaticFile():
    allIssue = []
    openIssue = []
    ongoingIssue = []
    closeIssue = []
    abnormalIssue = []
    delayIssue = []
    weekm1Issue = []
    weekm2Issue = []
    Status_open = ("New","Reopen","Analysis","Supplier inbox","Supplier In Progress")
    Status_ongoing = ("Solution Identified","Solved","Ready for Test","Under observation","Testing On Hold")
    Status_close = ("Closed","Cancelled")
    ValidTargetserie = ("VP2","TT","E4-2","E4-2/VP2","Pre TT","PreTT","vp2")
    # ----------------------------------------------------------------------------------------------------------------------
    # create a issuedict per row, and add all the issue to the allIssue, openIssue,ongoingIssue,closeIssue abnormalIssue,
    # weekm2Issue,weekm1Issue list
    # use issue_num as a dict, to insure that every dict is a new dict,will not change the other dict
    def createIssue(num):
        issue_num = {}
        issue_num["Key"] = SWIM["Key"][num]
        issue_num["ECU"] = SWIM["ECU"][num]
        issue_num["Status"] = SWIM["Status"][num]
        issue_num["Created"] = SWIM["Created"][num]
        issue_num["Target Serie"] = SWIM["Target Serie"][num]
        issue_num["Found in Serie"] = SWIM["Found in Serie"][num]

        if re.match("SWIM-",str(SWIM["Key"][num])):
         #    creat allIssue list
         allIssue.append(issue_num)
         # create weekm2Issue and weekm1Issue list
         createTime = issue_num["Created"]
         # print (createTime)
         z = time.strptime(str(createTime),'%Y-%m-%d %H:%M:%S')
         createWeek = time.strftime("%W", z)
         createYear = time.strftime("%Y", z)
         today = time.localtime(time.time())
         todayWeek = time.strftime("%W",today)
         todayYear = time.strftime("%Y",today)
         if int(todayYear) - int(createYear) ==0:
             if int(todayWeek)-int(createWeek) == 1:
                 weekm1Issue.append(issue_num)
             elif int(todayWeek) - int(createWeek) == 2:
                 weekm2Issue.append(issue_num)
             else:
                 pass
         else:
             pass
         # create abnormalIssue list
         diff = datetime.datetime.now()-createTime
         if diff.days >9 and SWIM["Status"][num] in Status_open and SWIM["ECU"][num] in list(INFO["ECU"]):
             issue_num["abnormal reason"] = "delay"
             issue_num["delay days"] = diff.days - 9
             delayIssue.append(issue_num)
         elif SWIM["Target Serie"][num] not in ValidTargetserie and SWIM["Status"][num] in Status_ongoing\
                 and SWIM["ECU"][num] in list(INFO["ECU"]):
             issue_num["abnormal reason"] = "wrong target serie"
             abnormalIssue.append(issue_num)
         else:
             pass
        #  create openIssue, ongoingIssue,closeIssue list
        if issue_num["Status"] in Status_open:
            openIssue.append(issue_num)
        if issue_num["Status"] in Status_ongoing:
            ongoingIssue.append(issue_num)
        if issue_num["Status"] in Status_close:
            closeIssue.append(issue_num)


    i = 0
    while i < len(SWIM["Key"]):
         createIssue(i)
         i+=1
    # ----------------------------------------------------------------------------------------------------------------------
    # class Issue: Issue.type =  [allIssue,openIssue,ongoingIssue,closeIssue,abnormalIssue,weekm2Issue,weekm1Issue]
    # addIssue: include Issue.type[num] to Issue.list[] ,should input num and one list in Issue.type
    class Issue:
        def __init__(self,*args,**kwargs):
            self.name = kwargs["name"]
            self.type = [allIssue,openIssue,ongoingIssue,closeIssue,abnormalIssue,weekm2Issue,weekm1Issue,delayIssue]
            self.list = []
        def addIssue(self,num,IssueList):
            if IssueList in self.type:
                self.list.append(IssueList[num])
            else :
                pass
    #-----------------------------------------------------------------------------------------------------------------------
    # def CreateIssueList: include LISTNAME,type,IssueList
    # LISTNAME: ECU name or Status name or the other specific value in one SWIM dict
    # type:  # Keys in SWIM
    # IssueList: allIssue,openIssue,ongoingIssue,closeIssue,abnormalIssue,weekm2Issue,weekm1Issue
    # from the IssueList return a list that the value in the key is equal to a specific value.
    # such as return a ECU issuelist from openIssue
    def createIssueList(LISTNAME,type,IssueList):
        issue_LISTNAME = []
        issue = Issue(name = LISTNAME)
        num = 0
        if IssueList in issue.type:
         while num < len(IssueList):
            if IssueList[num][type] == LISTNAME:
                issue.addIssue(num,IssueList)
                num+=1
            else:
                num+=1
        else:
            print("please input the right IssueList(allIssue,openIssue,ongoingIssue,"
                  "closeIssue,abnormalIssue,weekm2Issue,weekm1Issue)")
        issue_LISTNAME = issue.list
        return issue_LISTNAME
    # ----------------------------------------------------------------------------------------------------------------------
    # use CreateIssueList and Static functions to create a issueStatic dict List
    # each issue static include ECU,department,leader,ALL issue number, open issue number, ongoing issue number
    # close issue number , -1week issue number, -2week issue number,abnormal issue number
    issueStatic = []
    def Static(num):
        ECU_num = {}
        ECU_num ["ECU"]= INFO["ECU"][num]
        ECU_num["department"] = INFO["department"][num]
        ECU_num["leader"] = INFO["leader"][num]
        ECU_num["All"] = len(createIssueList(INFO["ECU"][num],"ECU",allIssue))
        ECU_num["Open"] = len(createIssueList(INFO["ECU"][num], "ECU", openIssue))
        ECU_num["Onoing"] = len(createIssueList(INFO["ECU"][num], "ECU", ongoingIssue))
        ECU_num["Close"] = len(createIssueList(INFO["ECU"][num], "ECU", closeIssue))
        ECU_num["Issues in -1week"] = len(createIssueList(INFO["ECU"][num], "ECU", weekm1Issue))
        ECU_num["Issues in -2week"] = len(createIssueList(INFO["ECU"][num], "ECU", weekm2Issue))
        ECU_num["Abnormal Issues"] = len(createIssueList(INFO["ECU"][num], "ECU", abnormalIssue))
        ECU_num["Delay Issues"] = len(createIssueList(INFO["ECU"][num],"ECU",delayIssue))
        if ECU_num["All"] != 0:
            issueStatic.append(ECU_num)
        else:
            pass

    i = 0
    while i <len(INFO["ECU"]):
            Static(i)
            i+=1

    # ---------------------------------------------------------------------------------------------------------------------
    data = pd.DataFrame.from_dict(issueStatic,orient = "columns")
    data2 = pd.DataFrame.from_dict(abnormalIssue,orient="columns")
    data3 = pd.DataFrame.from_dict(delayIssue,orient = "columns")
    data.to_excel("Issue Static.xlsx")
    data2.to_excel("AbnormalIssue.xlsx")
    data3.to_excel("DelayIssue.xlsx")

if __name__ == "__main__":
    main()
