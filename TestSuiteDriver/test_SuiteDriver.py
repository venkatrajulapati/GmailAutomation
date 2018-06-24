import pytest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from Lib.commonUtils import *
from Lib.App_CommonUtils import *
import time
import xlrd
import sys
import xlwt
from xlwt import *
from xlutils.copy import copy
import logging
from datetime import datetime


######################################################### Test Suite Driver #######################################################################################

class SuiteDriver(App_Common_utils):
    def __init__(self,Rootpah):
        super(SuiteDriver,self).__init__(Rootpah)

###################################################################################################################################################################
def mydecoratoe(func):
        print("###########################################################")
        print("started the execution " + func.__name__)
        res=func()
        print("###########################################################")
@mydecoratoe
def test_Runsuite():

    # Connect to the Test test case repository
    #print(os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))))
    logger = logging.getLogger(__name__)
    Rootpath = os.getcwd()
    #Rootpath="C:\\AutomationSuite\\"
    Rootpath = str(Rootpath).replace("TestSuiteDriver","")
    oRb = xlrd.open_workbook(Rootpath + "TestSuite\\TestSuite.xls")
    oWb = copy(oRb)
    #Read only copy of Test Suite
    oRTestsuite = oRb.sheet_by_name("TestSuite")
    #Editable copy of Test suite
    oWTestsuite = oWb.get_sheet(0)
    #Business Flow sheet
    oBusinessFlow = oRb.sheet_by_name("BusinessFlow")
    nooftcs = oRTestsuite.nrows
    #oui = UIdriver(Rootpath)
    obj=SuiteDriver(Rootpath)

    # Run the Test set
    for i in range(1,nooftcs):
        TCID = oRTestsuite.cell(i,0).value
        #TDID = oRTestsuite.cell(i,1).value
        TCName = oRTestsuite.cell(i,1).value

        tobeexecute = oRTestsuite.cell(i,2).value
        if str(tobeexecute).lower() == "y":
            dt1=datetime.now()
            print("Initiated the execution of the Test Case : " + TCID + " : " + TCName + ".................")
            #Create HTML Report
            f = obj.Create_HTML_Report(TCName)
            reppath = obj.Reportfile
            screenshotpath = obj.screenshotfolder
            reqrow = obj.GetxlRowNumber(oBusinessFlow,"TC_ID",TCID)
            noofsteps = oBusinessFlow.ncols
            iPasscount = 0
            stepcount = 0
            #Get Data sheetName
            DatasheetName = oRTestsuite.cell(reqrow,1).value
            screenshotcount=0

            noofDatarows = obj.GetNumberofrowsByXlCelltext(DatasheetName,"TC_ID",TCID)
            #Execute the Business flow Keywords and update the HTML Report
            for row in range(1,noofDatarows+1):
                if row <noofDatarows+1:
                    obj.fn_HtmlReport_TestStep(f,"","","Runnig the Iteration : " + str(row+1),"","","")
                for j in range(1,noofsteps):
                    Keyword = oBusinessFlow.cell(reqrow,j).value
                    temp = Keyword
                    print("=======================================================================================")
                    if not Keyword == "end":
                        print("running Keyword : " + Keyword)
                        Keyword = Keyword + "(" + chr(34)+DatasheetName+chr(34) + "," + str(row) + "," + chr(34) + TCID + chr(34) +" )"
                        #keyword = "%s%s%s%s" %(Keyword ,"(",oDataset,")")
                        if eval ("obj." + Keyword):
                            print(temp + " Keyword Passed")
                            screenshotcount = screenshotcount + 1
                            iPasscount = iPasscount + 1
                            stepcount = stepcount + 1
                            logger.info("pass", temp + " Keyword passed")
                            obj.fn_HtmlReport_TestStep(f,screenshotpath,screenshotcount,"running Step : " + str(temp),str(temp) + " Should be Passed",str(temp) + " is Passed","PASS")
                        else:
                            print(temp+" Keyword Failed hence Quitting the current test execution")
                            logger.info("fail",temp+" Keyword Failed hence Quitting the current test execution")
                            obj.fn_HtmlReport_TestStep(f, screenshotpath, screenshotcount, "running Step : " + str(temp),str(temp) + " Should be Passed", str(temp) + " is Failed", "FAIL")
                            stepcount = stepcount + 1
                            obj.browser.quit()
                            break
                    elif Keyword == "end":
                        print("end of the test")
                        obj.browser.quit()

                        #f.close()
                        break

            # Update The Testcase wise status and Report path in the test suite(Detailed summary Report
            dt2 = datetime.now()
            d1 = datetime(dt1.year, dt1.month, dt1.day, dt1.hour, dt1.minute, dt1.second, dt1.microsecond)
            d2 = datetime(dt2.year, dt2.month, dt2.day, dt2.hour, dt2.minute, dt2.second, dt2.microsecond)
            diff = d2 - d1
            if iPasscount==stepcount:
                oWTestsuite.write(i,4,"Pass")
                oWTestsuite.write(i, 5, str(dt1))
                oWTestsuite.write(i, 6, str(dt2))
                oWTestsuite.write(i, 7, str(diff))
                oWTestsuite.write(i, 8, xlwt.Formula('HYPERLINK("%s";"Clickto view report")' % reppath))
                f.close()
            else:
                oWTestsuite.write(i,4,"Fail")
                oWTestsuite.write(i, 5, str(dt1))
                oWTestsuite.write(i, 6, str(dt2))
                oWTestsuite.write(i, 7, str(diff))
                oWTestsuite.write(i, 8, xlwt.Formula('HYPERLINK("%s";"Clickto view report")' % reppath))
                f.close()

        #Create Detailed summary Report
        oWb.save(os.path.join(obj.Rootpath, "TestSuite\\DetailedSummaryReport.xls"))
    #Clean up
    obj=None
    oui=None
    oRTestsuite=None
    oWTestsuite=None
    oRb=None
    oWb=None
    #return True
test_Runsuite()
