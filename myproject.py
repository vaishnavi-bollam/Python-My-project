import subprocess, sys, win32com.client
import time, datetime
import base64
from ms_active_directory import ADDomain
import random, string
import ix_m
from bson.objectid import ObjectId
import time

from selenium import webdriver
from datetime import datetime
from selenium.webdriver import ActionChains
from selenium.common import exceptions
from selenium.webdriver.common.by import By
import urllib.parse

def Encrypt_key(source_string=""):
   return base64.b64encode(source_string.encode("ascii")).decode("ascii")
def Decrypt_key(self,enc_string="RDhOUFA2aHRHODE1dGxN"):
    return base64.b64decode(enc_string.encode("ascii")).decode("ascii")

try:
    flag=0
    ix = ix_m.IxAgent()
    ticketprofile_id = sys.argv[1]
    ix.Init_SNOW_Session()
    db = ix.db
    tck = db['ticket_profiles']
    ticket_data = tck.find_one({"_id": ObjectId(ticketprofile_id)})
    sys_id         = ticket_data['validated_inputs']['sys_id']
    ticket_number  = ticket_data['number']
    ritm_number  = ticket_data['validated_inputs']['parent.number']
    mail=ticket_data['validated_inputs']['parent.variables.email']
    mail_id="veluxprintservice@hcl.com",mail
    printer_name        = (ticket_data['validated_inputs']['parent.variables.printer_name']).replace(" ","_")
    printer_long_name = printer_name
    printer_model    = ticket_data['validated_inputs']['parent.variables.printer_model_name']
    host_ip    = ticket_data['validated_inputs']['parent.variables.host_or_printer_ip']
    printer_workflow    = ticket_data['validated_inputs']['parent.variables.printer_workflow_deployment']
    location  = ticket_data['validated_inputs']['parent.variables.printer_location']
    drivers=ticket_data['validated_inputs']['parent.variables.drivers']
    vpsx_name="VPS2"
    com_type="TCPIP/PJL"
    ticket_reassignmentgroup      = ticket_data['fulfillment_group_sysid']
    snow_table                    = "sc_task"
    status = True
except:
    status = False

usr = "ad-svc-automation"
pwd = Decrypt_key("RDhOUFA2aHRHODE1dGxN")
password=urllib.parse.quote(pwd)

def printer_mailbody(printername,printerip,ticketno):
   subject = "New Printer Onboarded | "+ printername + " | " + ticketno
   body = """
          <html>  
                    <body style=`"color:#455A64;font-family: 'Calibri Light', sans-serif;`">
                    <p>Hi <b>Requester</b>,</p>
                    <div>
                        <p>As per the request (<b>"""+ ticketno +"""</b>) we have successfully onboarded the New printer : <b>"""+ printername +"""</b>. Please check the below details</p>
                    </div>
                    <ul style=`"list-style-type: none;font-size:14px;`">
                        <li>Printer Name&nbsp;&nbsp;: <b>"""+ printername +""" </b></li>
                        <li>Printer IP&emsp;&nbsp;&nbsp;&nbsp;: <b>"""+ printerip +"""</b></li>
                    </ul>
                    <br><br>
                    <em style='color:#CE1029;font-size:18px;'>CONTACT US</em>
                    <p>In case of any enquiry <a href='https://velux.service-now.com/sp?id=sc_cat_item&sys_id=386ec9cb37bb7500d60498a543990ec0&referrer=popular_items'>click here</a> to raise the ticket</p>
                    <br><br>
                    <em style='color:#1565c0;font-size:11px;'>***Please note this is auto-generated mail. Do not Reply to this mail***</em>
                    <br><br><br>
                    <em style='font-size:12px;'>Thanks & Regards,</em>
                    <p style='color:#1565c0;'><b>VELUX AUTOMATION TEAM</b></p>
                    </body>
                    </html>
         """
   return subject,body
def printercheck(check):
    driver.find_element("xpath", '//*[@id="VXHOME"]/body/div/div[2]/div/ul/li[1]/a').click()
    time.sleep(2)
    driver.find_element("xpath",'/html/body/div[13]/div[2]/div/div[1]/form/fieldset/div[1]/a/span').click()
    time.sleep(2)
    driver.find_element("xpath",'/html/body/div[13]/div[2]/div/div[1]/form/fieldset/div[1]/ul/li[5]/a').click()
    time.sleep(2)
    driver.find_element("xpath",'//*[@id="txtSearch"]').send_keys(host_ip)
    time.sleep(2)
    table_data=driver.find_element("xpath",'/html/body/div[13]/div[2]/div/div[4]/table/tbody')
    row_data=table_data.find_elements(By.TAG_NAME, "tr")
    if row_data:
        printer_data=[]
        actionChains = ActionChains(driver)
        time.sleep(2)
        row_data.reverse()
        for i in row_data:
            try:
                data=i.text
                if data==printer_name.upper():
                    printer_data.append(data)
                    if check=="postcheck":
                        actionChains.context_click(i).perform()
                    break
            except exceptions.StaleElementReferenceException:
                pass
        if printer_data:
            return True
        else:
            return False
    else:
        return False

if status:
    try:
        worknote = ix.set_IxWorkNotesText("Automation Work in progress","Info")
        ix.TicketLog(ObjectId(ticketprofile_id),'workinprogress')
        ix.SNOWupdate(sys_id,snow_table,worknote,"wip",ticket_reassignmentgroup,"")
        worknotes = []
        worknotes += ix.set_IxWorkNotesText("Inputs Extracted Successfully","Info") 
        url=('https://{0}:{1}@dnkgevelmcn-002.velux.org/lrs/nlrswc2.exe/vpsx/basic_auth'.format(usr,password))
        options=webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.binary_location="C:/Program Files/Google/Chrome/Application/chrome.exe"
        chrome_driver_binary="C:/Program Files (x86)/chromedriver_win32/chromedriver.exe"
        driver=webdriver.Chrome(chrome_driver_binary,options=options)
        driver.get(url)
        main_page = driver.current_window_handle
        time.sleep(5)
        driver.refresh()
        time.sleep(2)
        #precheck
        precheck=printercheck("precheck")
        #precheck=False
        if not precheck:
            print("[PRECHECK] : printer details not available in LRS portal")
            worknotes += ix.set_IxWorkNotesText("[PRECHECK] : printer {} details not available in LRS portal".format(printer_name),"SUCCESS") 
            time.sleep(3)
            print("[INFO] : Process begin to add the printer details in LRS portal")
            worknotes += ix.set_IxWorkNotesText("Process begin to add the printer {} details in LRS portal".format(printer_name),"Info")
            driver.execute_script("window.open('https://dnkgevelmcn-002.velux.org/lrs/nlrswc2.exe/vpsx');")
            time.sleep(2)
            driver.switch_to.window(driver.window_handles[1])
            driver.find_element("xpath", '//*[@id="VXHOME"]/body/div/div[2]/div/ul/li[1]/a/span[2]').click()
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="option-admin"]').click()
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="btnAddConfig"]/span').click()
            time.sleep(3)
            driver.find_element("xpath",'//*[@id="dialog-add-printer-wizard"]')
            time.sleep(2)
            driver.find_element("xpath",'/html/body/div[8]/div[6]/div[2]/div/ul/li[1]/a').click()
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="prtname"]').send_keys(printer_name)
            time.sleep(2)

            button = driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/table/tbody/tr[2]/td/div/div/a')
            driver.execute_script("arguments[0].click();", button)
            time.sleep(3)
            driver.find_element("xpath",'//*[@id="dialog-vpsx-selection"]')
            time.sleep(2)
            driver.find_element("xpath",'/html/body/div[9]/div/div[2]/div/div[1]/div/div/div[1]/input').send_keys(vpsx_name)
            time.sleep(2)
            driver.find_element("xpath",'/html/body/div[9]/div/div[2]/div/div[1]/div/div/div[1]/a').click()
            time.sleep(2)
            driver.find_element("xpath",'/html/body/div[9]/div/div[2]/div/div[2]/table/tbody/tr/td[1]').click()

            time.sleep(2)
            driver.find_element("xpath",'/html/body/div[9]/div/ul/li[2]/a').click()



            time.sleep(2)
            driver.find_element("xpath",'//*[@id="commtype"]').send_keys(com_type)
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="tcphost"]').send_keys(host_ip)
            time.sleep(2)
            if printer_workflow=="Direct":
                button = driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/div[3]/div[1]/div[2]/div[1]/a')
                driver.execute_script("arguments[0].click();", button)
                time.sleep(3)
                driver.find_element("xpath",'//*[@id="dialog-driver-selection"]')
                time.sleep(2)

                driver.find_element("xpath",'/html/body/div[5]/div/div[2]/div/div[1]/div/div/div[1]/input').send_keys(printer_model)
                time.sleep(2)
                driver.find_element("xpath",'/html/body/div[5]/div/div[2]/div/div[1]/div/div/div[1]/a').click()
                time.sleep(2)
                table_data1=driver.find_element("xpath",'/html/body/div[5]/div/div[2]/div/div[2]/table/tbody')
                row_data1=table_data1.find_elements(By.TAG_NAME, "td")
                actionChains1 = ActionChains(driver)
                time.sleep(2)
                printerdata1=[]
                if row_data1:
                    for i in row_data1:
                        try:
                            data=i.text
                            if data==drivers:
                                printerdata1.append(data)
                                actionChains1.click(i)
                                actionChains1.perform()
                                break
                        except exceptions.StaleElementReferenceException:
                            pass
                    if printerdata1:
                        print("Driver Found")
                        worknotes += ix.set_IxWorkNotesText("Driver {} details Found".format(drivers),"INFO")
                    else:
                        print("Unable to found the driver name")
                        worknotes += ix.set_IxWorkNotesText("Unable to found the driver {} details under {}".format(drivers,printer_model),"Error")
                else:
                    print("Unable to found the driver name")
                    worknotes += ix.set_IxWorkNotesText("Unable to found the driver details".format(drivers),"Error")
                time.sleep(2)
                driver.find_element("xpath",'/html/body/div[5]/div/ul/li[2]/a').click()

            time.sleep(2)
            driver.find_element("xpath",'//*[@id="location"]').send_keys(location)
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="prtlname"]').send_keys(printer_long_name)
            time.sleep(2)
            if printer_workflow=="Direct":
                driver.find_element("xpath",'//*[@id="grpname"]').send_keys(printer_workflow)
                time.sleep(1)
            driver.find_element("xpath",'//*[@id="color"]').click()
            time.sleep(1)
            driver.find_element("xpath",'//*[@id="duplex"]').click()
            time.sleep(1)
            driver.find_element("xpath",'//*[@id="staple"]').click()
            if printer_workflow=="Secure":
                time.sleep(1)
                driver.find_element("xpath",'//*[@id="pullprint"]').click()
            time.sleep(3)
            button = driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/ul/li[2]/a')
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            if printer_workflow=="Direct":
                time.sleep(1)
                driver.find_element("xpath",'//*[@id="lics-list1"]/li[2]/label').click()
            time.sleep(1)
            driver.find_element("xpath",'//*[@id="lics-list2"]/li[1]/label').click()
            if printer_workflow=="Secure":
                time.sleep(1)
                driver.find_element("xpath",'//*[@id="lics-list2"]/li[2]/label').click()
            time.sleep(3)
            button = driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/ul/li[3]/a')
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="snmp"]').click()

            time.sleep(3)
            time.sleep(3)
            button = driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/ul/li[5]/a')
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/div[3]/div[5]/a[1]').click()
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="acct"]').click()
            time.sleep(2)
            driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/div[3]/div[5]/a[1]').click()
            time.sleep(2)
            driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/div[3]/div[5]/a[5]').click()
            time.sleep(2)
            if printer_workflow=="Direct":
                driver.find_element("xpath",'//*[@id="winopt1B"]').click()
            else:
                driver.find_element("xpath",'//*[@id="winopt1A"]').click()
            time.sleep(2)
            driver.find_element("xpath",'/html/body/table/tbody/tr/td/div[2]/div[2]/form/div[3]/div[5]/a[5]').click()
            time.sleep(2)
            driver.find_element("xpath",'//*[@id="btnUpdate"]').click()
            if printer_workflow=="Secure":
                time.sleep(2)
                poptext=driver.find_element("xpath",'/html/body/div[8]/div[7]/div[2]/form/div[1]/span[3]').text
                worknotes += ix.set_IxWorkNotesText(poptext,"INFO")
                driver.find_element("xpath",'/html/body/div[8]/div[7]/ul/li[3]/a').click()
                time.sleep(2)
                driver.find_element("xpath",'//*[@id="btnUpdate"]').click()
            
            print("Printer Added")
            worknotes += ix.set_IxWorkNotesText("Printer {} Onboarded successfully".format(printer_name),"SUCCESS")
            time.sleep(90)
            #post_check
            driver.execute_script("window.open('https://dnkgevelmcn-002.velux.org/lrs/nlrswc2.exe/vpsx');")
            time.sleep(2)
            driver.switch_to.window(driver.window_handles[2])
            postcheck=printercheck("postcheck")
            if postcheck:
                print("[POSTCHECK] : Printer Added Successfuly in LRS portal")
                worknotes += ix.set_IxWorkNotesText("[POSTCHECK] : Printer {} Added Successfuly in LRS portal".format(printer_name),"SUCCESS")
                if printer_workflow=="Secure":
                    time.sleep(2)
                    driver.find_element("xpath",'/html/body/div[1]/table/tbody/tr[6]/td/div/a').click()
                    time.sleep(2)
                    driver.find_element("xpath",'/html/body/div[1]/table/tbody/tr[6]/td/div/div/table/tbody/tr[1]/td/div/a').click()
                    time.sleep(10)
                    driver.find_element("xpath",'/html/body/div[5]/div[9]/ul/li[2]/a').click()
                    print("Printer workflow configuration done")
                    worknotes += ix.set_IxWorkNotesText("Printer {} workflow configuration done".format(printer_name),"SUCCESS")
                    time.sleep(400)
            else:
                print("[POSTCHECK] : Unable to find the printer details, printer not added or not sync")
                worknotes += ix.set_IxWorkNotesText("[POSTCHECK] : Unable to find the printer {} details, maybe printer not added due to invaild input or not sync".format(printer_name),"Error")
                flag=2
        else:
            print("[PRECHECK] : printer details already available in LRS portal")
            worknotes += ix.set_IxWorkNotesText("[PRECHECK] : printer {} details already available in LRS portal".format(printer_name),"Error")
            flag=2
    except Exception as e:
        print(e)
        flag=2
        worknotes += ix.set_IxWorkNotesText("Script fail to complete","Error")
        
    finally:
        driver.quit()
        if flag==0:
            subject,body = printer_mailbody(printer_name,host_ip,ritm_number)
            ix.SendMail(mail_id,subject,body,"html")
            ix.SNOWupdate(sys_id,snow_table,worknotes,"success",ticket_reassignmentgroup,"")
            ix.TicketLog(ObjectId(ticketprofile_id),'completed')
        else:
            ix.SNOWupdate(sys_id,snow_table,worknotes,"failed",ticket_reassignmentgroup,"")
            ix.TicketLog(ObjectId(ticketprofile_id),'failed')
else:
    worknotes = ix.set_IxWorkNotesText("Unable to Extract the Inputs","Error") 
    ix.SNOWupdate(sys_id,snow_table,worknotes,"failed",ticket_reassignmentgroup,"")
    ix.TicketLog(ObjectId(ticketprofile_id),'failed')