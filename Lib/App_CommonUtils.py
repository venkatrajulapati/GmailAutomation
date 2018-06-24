from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from Lib.commonUtils import *
import time
import xlrd
import sys

class App_Common_utils(UIdriver):

    def __init__(self,rootpath):
        super(App_Common_utils,self).__init__(rootpath)


    def Login_App(self,odata,reqrow,TCID):
        #try:
        driver = self.Get_Browser()
        self.Launch_Application(driver)
        if self.App_Sync(driver,"Login","GenericEditField",args=['username']):
            self.SetFieldValue(driver,"Login","GenericEditField",self.username,args=['username'])
            self.SetFieldValue(driver,"Login","GenericEditField",self.password,args=['password'])
            self.ClickObject(driver,"Login","btn_LoginBtn",args=['sign in'])
            print("Application Login Successful")
            return True
        else:
            print("Failed to load Login page")
            return False

        #except:
            #print("Login failed")
            #return False


    def SendEmail(self,odata,reqrow,TCID):
        #odata=str(odata).replace("'","")
        try:
            owb = xlrd.open_workbook(self.Rootpath+"TestData\\" + odata + ".xls")
            oDataset = owb.sheet_by_index(0)
            reqrow = int(reqrow)
            #reqrow = self.GetxlRowNumberbytwocolvals(oDataset, "TC_ID", TCID, "TD_ID", TDID)
            print(oDataset.nrows)
            self.wait_pageload(self.browser)
            if self.App_Sync(self.browser, "MailBox", "lnk_NewMail"):
                #time.sleep(3)
                self.ClickObject(self.browser,"MailBox","lnk_NewMail")
                Tocol = self.GetxlColumnNumber(oDataset,"To")
                Cccol = self.GetxlColumnNumber(oDataset,"Cc")
                Subjectcol = self.GetxlColumnNumber(oDataset,"Subject")
                #Parameters
                toAddress = oDataset.cell(reqrow,Tocol).value
                ccAddress = oDataset.cell(reqrow,Cccol).value
                subJect = oDataset.cell(reqrow,Subjectcol).value
                self.App_Sync(self.browser, "MailBox", "GenericEditField",args=['To'])
                #Edit To Address
                if toAddress!="":
                    result = self.SetFieldValue(self.browser, "MailBox", "GenericEditField",toAddress,args=['To'])
                #Edit cc Address
                if ccAddress!="":
                    result = self.SetFieldValue(self.browser, "MailBox", "GenericEditField",ccAddress,args=['Cc'])
                #Edit Subject
                if subJect!="":
                    result = self.SetFieldValue(self.browser, "MailBox", "edt_Subject",subJect )
                #Click on Send
                result = self.ClickObject(self.browser,"MailBox","lnk_Send")
                return result

            else:
                print("Failed to load MailBox")
                return False
        except:
            print("Failed to Send Email")
            return False

    def Logout(self,odata,reqrow,TCID):
        try:
            if self.App_Sync(self.browser,"MailBox","lnk_UserName"):
                result = self.ClickObject(self.browser,"MailBox","lnk_UserName")
                time.sleep(1)
                if result:
                    result = self.ClickObject(self.browser,"MailBox","lnk_Signout")
                    if result:
                        result = self.App_Sync(self.browser,"Login","edt_User Name")

            if result:
                print("Application Successfully Logged out")
                return True
            else:
                print("Logout failed")
                return False

        except:
            print('Failed to click logout')
            return False
