from datetime import datetime
import os
import sys


gbl_blnResult = ""
g_iPass_Count = 0
g_iFail_Count = 0
g_iScenario_Pass_Count = 0
g_iScenario_Fail_Count = 0
g_tStart_Time = datetime.now()

ioMode_ForWriting = 2
gbl_intScreenCount = 0
g_iImage_Capture = 1
RunMode = "Server"
BlnFailStatue = False
strScriptName = "tc1"
#Name of Report-folders and Report-File-Name for this Run
arrStartTime = str(g_tStart_Time).split(" ")
strname1 = arrStartTime[0]
strname1 = strname1.replace("-","")
print(strname1)
strname2 = arrStartTime[1]
strname2 = strname2.replace(":","")
strname2 = strname2.split(".")
strname2 = strname2[0]
strname = strname1 + "_" + strname2
print(strname)
strEnvironment =""
if not os.path.exists("E:\\actitimeAutomation\\Results"):
    os.mkdir("E:\\actitimeAutomation\\Results")

TCName = "Dummy"
ReportFolder = "E:\\actitimeAutomation\\Results\\" + TCName + "_" +  strname

if not os.path.exists(ReportFolder):
    os.mkdir(ReportFolder)

Reportfile = "E:\\actitimeAutomation\\Results\\" + TCName + "_" +  strname + ".html"
screenshotfolder = "E:\\actitimeAutomation\\Results\\" + TCName + "_" +  strname

if not os.path.exists(screenshotfolder):
    os.mkdir(screenshotfolder)

resfile = open(Reportfile,"w")
#Write header
resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
Test_Automation_Test_Report_Logo = "E:\\actitimeAutomation\\Logo.png"
dttime = datetime.now()
dttime= str(dttime)
#Write Report - Header
resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
resfile.write("<TR COLS=2><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='" + Test_Automation_Test_Report_Logo + "'></TD><TD WIDTH=100% BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=NAVY SIZE=4><B>&nbspactiTime Test Automation Results - [" + dttime + "] </B></FONT></TD></TR></TABLE>")
resfile.write("<TABLE BORDER=1 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
resfile.write("</TABLE></BODY></HTML>")

#Write Report - Test-Set Name OR Test-Script Name
resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
resfile.write("<TR COLS=1>"\
                    "<TD ALIGN=LEFT BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=3><B>" + TCName + "</BR>" + "</B></FONT></TD>"\
                 "</TR>")
resfile.write("</TABLE></BODY></HTML>")

#Write Report - Column Headers
resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
resfile.write("<TR COLS=4>"\
                        "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Test Step</B></FONT></TD>"\
                        "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Expected Result</B></FONT></TD>"\
                        "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Actual Result</B></FONT></TD>"\
                        "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99   WIDTH=7%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Step-Result</B></FONT></TD>"\
                 "</TR>")
resfile.close()
# Set
# objFS = CreateObject("Scripting.FileSystemObject")
# ReportFolder = gblstrReportFolder & "_" & strEnvironment & "_" & replace(date, "/", "")
# TCName = Replace(DataTable("in_Test_Case_Name"), " ", "_")
# strReportFolder = ReportFolder & "\" & TCName & "\" & "
# Results
# "
# gbl_RepFolder = ReportFolder


def fn_HtmlReport_TestStep(strRepfilepath,strScreenshotfolder,gbl_intScreenCount,strDesc, strExpected, strActual, strResult):

    #***** Set Result parameters
    if str(strResult).upper() == "PASS":
        strResultColor = "GREEN"
        strResultSign = "P"
        blnCaptureImsge = True

    elif str(strResult).upper() == "FAIL":
        strResultColor = "RED"
        strResultSign = "O"
        blnCaptureImsge = True

    else:
        blnCaptureImsge = False
        strResultColor = "GREEN"
        strResultSign = "P"
        strActualHREF = strActual

    #***** Set Image Path and capture image
    if (blnCaptureImsge == True):
        gbl_intScreenCount = gbl_intScreenCount + 1
        #Capture Image
        strImagePath = strScreenshotfolder & "\\Screen_000" + str(gbl_intScreenCount, 3) & ".png"
        self.browser.get_screenshot_as_file(strImagePath)
        strActualHREF = "<A HREF='" & strImagePath & "'>" & strActual & "</A>"

    elif blnCaptureImsge == "False":
        strActualHREF = "<A>" & strActual & "</A>"

    #***** Open HTML Report file
        objReport = open(strRepfilepath,'a')

    #***** Update HTML Report

    if not strExpected is None:
        objReport.write(
        "<TR COLS=4>"\
        "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>" & strDesc & "</FONT></TD>"\
        "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>" & strExpected & "</FONT></TD>"\
        "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2>" & strActualHREF & "</FONT></TD>"\
        "<TD ALIGN=MIDDLE BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=" & strResultColor & ">" & strResultSign & "</FONT><FONT FACE=VERDANA SIZE=2 COLOR=" & strResultColor & "><B>" & strResult & "</B></FONT></TD>"\
        "</TR>")
    if strExpected is None:
        objReport.write(
        "<TR COLS=4>"\
        "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=5 COLOR=GREEN>" & strDesc & "</FONT></TD>"\
        "</TR>")

    objReport.Close


