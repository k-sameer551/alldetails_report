import os, time, ctypes, re
from datetime import datetime, timedelta
from pathlib import Path
import psutil
from selenium.webdriver.common.by import By
from selenium import webdriver
import win32com.client as win32
import alldetails.constants as const
# import edgedriver_autoinstaller


class Webtrax(webdriver.Edge):
    """download alldetails"""
    def __init__(self, username, password, teardown=False,  driver_path = const.DRIVER_PATH): #, driver_path = edgedriver_autoinstaller.install()
        """init"""
        self.username = username
        self.password = password
        self.teardown = teardown
        self.driver_path = driver_path
        oShell = win32.Dispatch('WScript.Shell')
        # path = os.path.join(os.environ['userprofile'], "Downloads")
        path = oShell.SpecialFolders("MyDocuments")
        prefs = {'download.default_directory' : path}
        options = webdriver.EdgeOptions()
        options.add_experimental_option('detach', True)
        # options.add_experimental_option('--headless', True)
        options.add_experimental_option('prefs', prefs)
        super(Webtrax, self).__init__(options=options)
        self.maximize_window()
        self.implicitly_wait(30)

    def __exit__(self, exc_type, exc, traceback):
        """exit"""
        if self.teardown:
            self.quit()

    def load_webpage(self):
        """download alldetails"""
        self.get(const.BASE_URL)
        self.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Login1_UserName').send_keys(self.username)
        self.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Login1_Password').send_keys(self.password)
        self.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Login1_LoginButton').click()
        self.get(const.QUEUE_GROUP)
        self.navigate_to_link('a', "UBH UNET Claims Workflow")
        self.navigate_to_link('a', "Work")
        report_date = self.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_GridView2 > tbody > tr:nth-child(1) > td:nth-child(2)').text
        report_date = datetime.strptime(report_date, r"%m/%d/%Y %H:%M:%S %p")
        report_date = report_date.strftime('%Y-%m-%d')
        todays_date =  (datetime.now() - timedelta(days=0)).strftime('%Y-%m-%d')
        if report_date == todays_date:
            self.get(const.DETAILS_FILE)
        else: pass

    def navigate_to_link(self, tag_name, tag_text):
        """navigate to the link"""
        for link in self.find_elements(By.TAG_NAME, tag_name):
            if link.text == tag_text:
                link.click()
                return True
        return True
    
    def download_file(self, element, directory, timeout):
        """wait for download"""
        self.find_element(By.CSS_SELECTOR, element).click()
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < timeout:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(directory)
            for fname in files:
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds +=1
        return True

    @classmethod
    def clear_files(cls, download_folder):
        """ clear previously downloaded file"""
        for file in Path(download_folder).iterdir():
            if file.name.startswith('Unet All Details') or file.name.startswith('Unconfirmed'):
                file.unlink()
        return True
    
    @classmethod
    def share_alldetails(cls, files_list: list):
        """share alldetails through outlook"""
        outlook = win32.Dispatch('Outlook.Application')
        wsh = win32.Dispatch('WScript.Shell')
        olmailitem=0x0
        mail_item = outlook.CreateItem(olmailitem)
        mail_item.Display()
        # signature = mail_item.HTMLBody
        mail_item.To = ''
        mail_item.CC = ''
        mail_item.Subject = 'Unet All Details_' + datetime.now().strftime(r'%m%d%Y')
        mail_item.HTMLBody = r"""
                Dear All,<br><br>
                Please find attached today’s All Detail files.<br><br>
                Best regards,<br><br>Regards<br>Unet Auto Reporting
                """
        for file in files_list:
            mail_item.Attachments.Add(file)
        mail_item.Display()
        wsh.AppActivate("Outlook")
        time.sleep(2)
        wsh.SendKeys("%s", 0)
        
    @classmethod
    def get_alldetail_files_path(cls, download_folder):
        """ return list of alldetails files"""
        files = []
        for file in Path(download_folder).iterdir():
            if file.name.startswith('Unet All Details'):
                files.append(file)    
        return files

    @classmethod
    def close_app(cls, app_name):
        """close app"""
        running_apps=psutil.process_iter(['pid','name']) #returns names of running processes
        found=False
        for app in running_apps:
            sys_app=app.info.get('name').split('.')[0].lower()
            if sys_app in app_name.split() or app_name in sys_app:
                pid=app.info.get('pid') #returns PID of the given app if found running
                try: #deleting the app if asked app is running.(It raises error for some windows apps)
                    app_pid = psutil.Process(pid)
                    app_pid.terminate()
                    found=True
                except: pass
            else: pass
        if not found:
            return False
        else:
            return True

    @classmethod
    def Mbox(cls, title, text, style):
        """message box"""
        return ctypes.windll.user32.MessageBoxW(0, text, title, style)

    @classmethod
    def Convert_xls_xlsb(cls, files_list: list, location):
        """convert file"""
        xlExcel12 = 50
        all_details_files = []
        xl = win32.Dispatch('Excel.Application')
        # xl.Visible = False
        for file in files_list:
            filename = re.search(r"(.*)\.xls\.xls", file.name).group(1)
            newfilename = str(Path.joinpath(location, str(filename) + "_" + datetime.now().strftime(r"%m%d%Y") + ".xlsb"))
            wb = xl.Workbooks.Open(file)
            try:
                Webtrax.set_sensitiviy_label(wb)
            except: pass
            wb.SaveAs(newfilename, FileFormat=xlExcel12)
            wb.Close()
            all_details_files.append(newfilename)
        return all_details_files

    @classmethod
    def set_sensitiviy_label(cls, wbook):
        """set sensitivity label"""
        label = wbook.SensitivityLabel.CreateLabelInfo()
        label.AssignmentMethod = 1  #MsoAssignmentMethod.PRIVILEGED
        label.LabelId = "a8a73c85-e524-44a6-bd58-7df7ef87be8f"
        label.SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
        wbook.SensitivityLabel.SetLabel(label, label)