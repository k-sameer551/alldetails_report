from configparser import ConfigParser
import os.path
# from win32com.shell import shell, shellcon
from win32comext.shell import shell, shellcon

class Credentails(ConfigParser):
    """class to get credentails"""
    def __init__(self):
        super(Credentails, self).__init__()

    def get_credentails(self):
        """get credentails"""
        # file_path = os.path.abspath(os.path.realpath(r'resources\credentials.ini'))
        try:
            # file_path = os.path.expanduser(r"~\Documents\Auto_Settings\CFCofig.ini")
            documents_path = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0)
            file_path = os.path.join(documents_path, r'Auto_Settings\CFCofig.ini')
            # print(file_path)
            self.read(filenames= file_path, encoding='utf-8')
            username = self.get('Details', 'username')
            password = self.get('Details', 'password')
            return username, password
        except: ValueError("CFConfig file not found")

        # self['Details']['username']
        # self['Details']['password']
        # return self.items('detials')
        # uname, pword = Credentails().get_credentails()
        # print(uname, pword)