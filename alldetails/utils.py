import ctypes, os
from datetime import datetime
from pathlib import Path
import psutil
import win32com.client as win32


class Utils():
    """utililty support functions"""
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
        olmailitem=0x0
        mail_item = outlook.CreateItem(olmailitem)
        mail_item.Display()
        mail_item.To = ''
        mail_item.CC = ''
        mail_item.Subject = 'Unet All Details_' + datetime.now().strftime('%m%d%Y')
        mail_item.HTMLBody = r"""
                Dear All,<br><br>
                Please find attached today’s All Detail files.<br><br>
                Best regards,<br>
                """
        for file in files_list:
            mail_item.Attachments.Add(file)
        mail_item.Display()
        mail_item.Send()
        
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
    def set_dest_folder(cls, path):
        """path to store downloaded file"""
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    @classmethod
    def get_directory(cls, folder_name):
        """my document path"""
        wsh = win32.Dispatch("WScript.Shell")
        return wsh.SpecialFolders(folder_name)
    
