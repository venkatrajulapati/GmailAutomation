from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

import time
import xlrd
import os
from datetime import datetime
import sys


class UIdriver(object):

    # Common utility functions for python selenium
    # Get the Browser driver
    #Initialize to Read the global configuration.
    def __init__(self,rootpath):

        self.Rootpath = rootpath
        #Open Master Data Excel workbook
        oWB = xlrd.open_workbook( self.Rootpath+"TestData\\MasterData.xls")
        #Connect to Environment sheet
        oEnvSheet = oWB.sheet_by_name("EnvironmentSetpup")
        #Fetch the Execution Environment Name
        env = self.GetxlColumnNumber(oEnvSheet, "Environment")
        self.envname=str(oEnvSheet.cell(1,env).value)
        #Connect to Master data Sheet
        oSheet = oWB.sheet_by_name("MasterData")
        #Fetch the Details and store in Global variables
        reqrow = self.GetxlRowNumber(oSheet,"Environment",self.envname)
        URL = self.GetxlColumnNumber(oSheet, "URL")
        BR = self.GetxlColumnNumber(oSheet, "Browser")
        UN = self.GetxlColumnNumber(oSheet, "UserName")
        PWD = self.GetxlColumnNumber(oSheet, "Password")
        self.BrName = str(oSheet.cell(reqrow,BR).value)
        self.username = str(oSheet.cell(reqrow,UN).value)
        self.password = str(oSheet.cell(reqrow,PWD).value)
        self.url = str(oSheet.cell(reqrow,URL).value)

        #global Variables Reports
        self.Reportfile = ""
        self.screenshotfolder = ""



    def Get_Browser(self):

        if str(self.BrName).lower() == 'chrome':
            self.browser = webdriver.Chrome()
        elif self.BrName.lower() == 'ie':
            capabilities = DesiredCapabilities.INTERNETEXPLORER
            #capabilities.pop('platform',None)
            #capabilities.pop('version',None)

            capabilities['ignoreProtectedModeSettings'] = True
            capabilities['ignoreZoomSetting'] = True
            #capabilities["requireWindowFocus"] = True
            #urllib.request.getproxies = lambda: {}
            self.browser = webdriver.Ie()#executable_path="C:\\Python27\\Scripts\\IEDriverServer.exe",capabilities=capabilities)


        else:
            self.browser = webdriver.Firefox()
        self.browser.maximize_window()
        return self.browser

    # To launch Application

    def Launch_Application(self,driver):

        driver.get(self.url)
    # Verify Page load status
    def wait_pageload(self,driver):
        page_status = driver.execute_script('return document.readyState')
        while str(page_status) != 'complete':
            page_status = driver.execute_script('return document.readyState')

        print(page_status)

        # try:
        #     myElem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, Locatorval)))
        #     print("Page is ready!")
        # except TimeoutException:
        #     print("Loading took too much time!")

    def App_Sync(self,driver,strPageName,strObjectName,args=None):

        objArr = []
        objArr = self.Get_Object_ObjectRepository(strPageName, strObjectName)
        objType = objArr.__getitem__(0)
        Locator = objArr.__getitem__(1)
        Locatorval = objArr.__getitem__(2)
        if not args is None:
            for idx in range(0,len(args)):
                Locatorval = Locatorval.replace('$$'+str(idx+1),args[idx])

        for Lpc in range(1, 60, 1):
            print("Waiting for the Object " + strPageName + "." + strObjectName)
            try:
                elem = self.Get_UIObject(driver,Locator,Locatorval)
                if elem.is_displayed() and elem.is_enabled():
                    print(strPageName + "." + strObjectName + " object Found")
                    break
            except:
                print("Please wait element Not Found")
        if not elem:
            return False
        else:
            return True


    def Get_UIObject(self,driver,Locator,Locatorvalue):
        Locator = str(Locator)
        Element = ""
        try:
            if Locator.lower() == 'id':
                Element = driver.find_element_by_id(Locatorvalue)
            elif Locator.lower() == 'name':
                Element = driver.find_element_by_name(Locatorvalue)
            elif Locator.lower()== 'css':
                Element = driver.find_element_by_css_selector(Locatorvalue)
            elif Locator.lower()=='Linktext':
                Element=driver.find_element_by_link_text(Locatorvalue)
            elif Locator.lower()=='xpath':
                Element = driver.find_element_by_xpath(Locatorvalue)
            return Element
        except:
            print("unable to find the Locator : "+ Locatorvalue)
            return False

    def ClickObject(self,driver,pagename,objName,args=None):
        try:
            objArr = []
            objArr = self.Get_Object_ObjectRepository(pagename,objName)
            objType = objArr.__getitem__(0)
            Locator = objArr.__getitem__(1)
            Locatorval = objArr.__getitem__(2)
            if not args is None:
                for idx in range(0,len(args)):
                    Locatorval = Locatorval.replace('$$'+str(idx+1),args[idx])
            result = False
            elem = self.Get_UIObject(driver,Locator,Locatorval)
            if not elem is None:
                # actions = webdriver.ActionChains(driver)
                # time.sleep(1)
                # actions.move_to_element(elem)
                # time.sleep(1)
                # actions.click()
                # actions.perform()
                elem.click()
                print("Clicked on the Object : " + pagename + "." + objName)
                result = True
            else:
                print("element not found please check the object description")
                result = False
            return result
        except:
            print("some error occured while clicking object : " + pagename + "." + objName)
            return False

    def Switch_frame(self,driver,strPageName,strObjectName):
        objArr = self.Get_Object_ObjectRepository(strPageName, strObjectName)
        objType = objArr.__getitem__(0)
        Locator = objArr.__getitem__(1)
        Locatorval = objArr.__getitem__(2)
        try:
            elem = self.Get_UIObject(driver,Locator,Locatorval)
            driver.switch_to.frame(elem)
            print('Successfully switched to the Frame')
        except:
            print('Failed to switched to the Frame')

    def Switch_defaultframe(self,driver):

        try:
            driver.switch_to.default_content()
            print('Successfully switched to the default Frame')
        except:
            print('Failed to switched to default Frame')


    def Switch_window(self,driver):
        main_window_handle = driver.current_window_handle
        child_window_handle = None
        while child_window_handle is None:
             for handle in driver.window_handles:
                 if handle !=main_window_handle:
                     child_window_handle=handle
                     break
             driver.switch_to.window(child_window_handle)
        return driver

    def GetxlColumnNumber(self,oSheet,strColName):
        noofcols = oSheet.ncols
        for c in range(0,noofcols):
            if(oSheet.cell(0,c).value == strColName):
                return c
                break

    def GetxlRowNumber(self,oSheet,strColName,strColumnvalue):
        reqcol = self.GetxlColumnNumber(oSheet, strColName)
        noofrows = oSheet.nrows
        for r in range(0,noofrows):
            if(oSheet.cell(r,reqcol).value == strColumnvalue):
                return r
                break

    def GetxlRowNumberbytwocolvals(self,oSheet,strColName1,strColumnvalue1,strColName2,strColumnvalue12):
        reqcol1 = self.GetxlColumnNumber(oSheet, strColName1)
        reqcol2 = self.GetxlColumnNumber(oSheet, strColName2)
        noofrows = oSheet.nrows
        for r in range(0,noofrows):
            if(oSheet.cell(r,reqcol1).value == strColumnvalue1 and oSheet.cell(r,reqcol2).value == strColumnvalue12):
                return r
                break
    def GetNumberofrowsByXlCelltext(self,owb,strColName1,strColumnvalue1):
        owb = xlrd.open_workbook(self.Rootpath+"TestData\\" + owb + ".xls")
        oDataset = owb.sheet_by_index(0)
        reqcol1 = self.GetxlColumnNumber(oDataset, strColName1)
        noofrows = oDataset.nrows
        cnt = 0
        for r in range(0,noofrows):
            if(oDataset.cell(r,reqcol1).value == strColumnvalue1):
                cnt = cnt+1
        return cnt
    def Get_Object_ObjectRepository(self,strPageName,strObjectName):

        try:
            oWB = xlrd.open_workbook(self.Rootpath + "\\ObjectRepository\\ObjectRepository.xls")
            oSheet = oWB.sheet_by_name("ObjectRepository")
            r = self.GetxlRowNumberbytwocolvals(oSheet, "PageName", strPageName, "ObjectName", strObjectName)
            st = []
            #if r1==r2:
            Objtypecolnum = self.GetxlColumnNumber(oSheet, "ObjectType")
            ObjLocatorcolnum = self.GetxlColumnNumber(oSheet, "Locator")
            ObjLocatorvalcolnum= self.GetxlColumnNumber(oSheet, "LocatorValue")
            strObjectType = oSheet.cell(r,Objtypecolnum).value
            strLocator = oSheet.cell(r, ObjLocatorcolnum).value
            strLocatorval = oSheet.cell(r, ObjLocatorvalcolnum).value
            st.append(strObjectType)
            st.append(strLocator)
            st.append(strLocatorval)
            #print(st)
            return st
        except:
            print("Failed to load Object repository please check the path")

        #else:
         #   print("object not found please check the PageName and Object Name you are looking for")

    def SetFieldValue(self,driver,strPageName,strObjectName,fval,args=None):

        try:

            objArr =[]
            objArr= self.Get_Object_ObjectRepository(strPageName, strObjectName)
            objType = objArr.__getitem__(0)
            Locator = objArr.__getitem__(1)
            Locatorval = objArr.__getitem__(2)
            if not args is None:
                for idx in range(0,len(args)):
                    Locatorval = Locatorval.replace('$$'+str(idx+1),args[idx])
            result = False
            objType = str(objType)
            if objType.lower() == "editfield":
                elem = self.Get_UIObject(driver,Locator,Locatorval)
                if elem.is_displayed():
                    elem.clear()
                    elem.send_keys(fval)
                    print(strObjectName + "value is entered as " + fval)
                    result=True
                else:
                    print(strObjectName+ " element not displayed")
            elif objType.lower() == "dropdown":
                elem = self.Get_UIObject(driver, Locator, Locatorval)
                if elem.is_displayed:
                    selectoption = Select(elem)
                    selectoption.select_by_visible_text(fval)
                    print(strObjectName + "value is selected as " + fval)
                    result=True
                else:
                    print(strObjectName + "element not displayed")
            elif objType.lower() == "chkbox":
                elem = self.Get_UIObject(driver, Locator, Locatorval)
                if elem.is_displayed:
                    elem.click()
                    print(strObjectName + "check box is selected")
                    result= True
                else:
                    print(strObjectName + "element not displayed")
            return result
        except:
            print("failed to set the field value of the filed : " + strPageName + "." + strObjectName + objType)
            return False

    def Create_HTML_Report(self,TCName):

        g_tStart_Time = datetime.now()

        # Name of Report-folders and Report-File-Name for this Run
        arrStartTime = str(g_tStart_Time).split(" ")
        strname1 = arrStartTime[0]
        strname1 = strname1.replace("-", "")
        print(strname1)
        strname2 = arrStartTime[1]
        strname2 = strname2.replace(":", "")
        strname2 = strname2.split(".")
        strname2 = strname2[0]
        strname = strname1 + "_" + strname2
        print(strname)
        strEnvironment = ""
        Rp = self.Rootpath
        if not os.path.exists(Rp + "Results"):
            os.mkdir(Rp + "Results")

        #TCName = "Dummy"
        ReportFolder = Rp + "Results\\" + TCName + "_" + strname

        if not os.path.exists(ReportFolder):
            os.mkdir(ReportFolder)

        Reportfile = Rp + "Results\\" + TCName + "_" + strname + ".html"
        screenshotfolder = Rp + "Results\\" + TCName + "_" + strname

        self.Reportfile = Reportfile
        self.screenshotfolder = screenshotfolder
        if not os.path.exists(screenshotfolder):
            os.mkdir(screenshotfolder)

        resfile = open(Reportfile, "a")
        # Write header
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        Test_Automation_Test_Report_Logo = Rp + "Logo.png"
        dttime = datetime.now()
        dttime = str(dttime)
        # Write Report - Header
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write(
            "<TR COLS=2><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='" + Test_Automation_Test_Report_Logo + "'></TD><TD WIDTH=100% BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=NAVY SIZE=4><B>&nbspactiTime Test Automation Results - [" + dttime + "] </B></FONT></TD></TR></TABLE>")
        resfile.write("<TABLE BORDER=1 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("</TABLE></BODY></HTML>")

        # Write Report - Test-Set Name OR Test-Script Name
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("<TR COLS=1>" \
                      "<TD ALIGN=LEFT BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=3><B>" + TCName + "</BR>" + "</B></FONT></TD>" \
                                                                                                                     "</TR>")
        resfile.write("</TABLE></BODY></HTML>")

        # Write Report - Column Headers
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("<TR COLS=4>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Test Step</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Expected Result</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Actual Result</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99   WIDTH=7%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Step-Result</B></FONT></TD>" \
                      "</TR>")
        return resfile
        #resfile.close()



    def fn_HtmlReport_TestStep(self,strRepfilepath,strScreenshotfolder,gbl_intScreenCount,strDesc, strExpected, strActual, strResult):

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
        #Set Image Path and capture image
        if (blnCaptureImsge == True):
            #gbl_intScreenCount = gbl_intScreenCount + 1
            #Capture Image
            strImagePath = strScreenshotfolder + "\\Screen_000" + str(gbl_intScreenCount) + ".png"
            self.browser.get_screenshot_as_file(strImagePath)
            strActualHREF = "<A HREF='" + strImagePath + "'>" + strActual + "</A>"

        elif blnCaptureImsge == "False":
            strActualHREF = "<A>" + strActual + "</A>"
        #Update HTML Report
        if not strExpected is None:
            strRepfilepath.write("<TR COLS=4>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>" + strDesc + "</FONT></TD>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strExpected + "</FONT></TD>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2>" + strActualHREF + "</FONT></TD>"\
            "<TD ALIGN=MIDDLE BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=" + strResultColor + ">" + strResultSign + "</FONT><FONT FACE=VERDANA SIZE=2 COLOR=" + strResultColor + "><B>" + strResult + "</B></FONT></TD>"\
            "</TR>")
        if strExpected is None:
            strRepfilepath.write("<TR COLS=4>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=5 COLOR=GREEN>" + strDesc + "</FONT></TD>"\
            "</TR>")




