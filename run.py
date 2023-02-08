import os.path
from pathlib import Path
import win32com.client as win32
from alldetails.alldetails import Webtrax
from alldetails.credentails import Credentails
from alldetails.utils import Utils


if __name__ == "__main__":
    username, password = Credentails().get_credentails()
    with Webtrax(username=username, password=password, teardown=False) as bot:
        oShell = win32.Dispatch('WScript.Shell')
        user_profile = Path(os.path.dirname(oShell.SpecialFolders("MyDocuments")))
        download_directory = user_profile.joinpath('Documents')
        all_details = "#ctl00_ContentPlaceHolder1_GridView1 > tbody > tr:nth-child(2) > td:nth-child(3) > a"
        all_details_medicaid = "#ctl00_ContentPlaceHolder1_GridView1 > tbody > tr:nth-child(3) > td:nth-child(3) > a"
        bot.clear_files(download_directory)
        bot.load_webpage()
        bot.download_file(all_details, download_directory, 420)
        bot.download_file(all_details_medicaid, download_directory, 180)
        files_list = bot.get_alldetail_files_path(download_directory)
        destination_directory = user_profile.joinpath("Documents") #"WorkInventory"
        destination_directory = Utils.set_dest_folder(destination_directory)
        alldetails_files = bot.Convert_xls_xlsb(files_list, destination_directory)
        if len(alldetails_files) > 0:
            bot.close_app('outlook')
            bot.share_alldetails(alldetails_files)
        else:
            bot.Mbox('All details', 'All Details files download terminated unexpectedly.', 1)
        bot.quit()