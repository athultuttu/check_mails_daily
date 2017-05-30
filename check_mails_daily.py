__author__ = 'athultuttu'

# PURPOSE:
# TO MONITOR MAILS IN OUTLOOK APPLICATION ON WINDOWS AND CHECK IF MAIL HAS BEEN 
# BEEN RECEIVED YESTERDAY IN SOME PARICULAR FOLDER
#
# https://github.com/athultuttu

#from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch
import datetime
import ctypes  # An included library with Python install.

my_mailbox = 'mail_id@mailbox.com'          #provide the mail id of mailbox to monitor
#This is assuming that some rule has been defined to move particular mailsto some folder
my_folder = 'the folder to search for'      #give the folder name to check if new mails has 
                                            #been received yesterday

today = datetime.date.today()
yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
yesterday = yesterday.strftime('%Y-%m-%d')  #date format - 2017-05-30

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class Oli():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1, array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted(self._obj._prop_map_get_.keys())

folder_found = False

for inx, folder in Oli(mapi.Folders).items():
    # iterate all Outlook folders (top level)
    print("-"*70)
    print(folder.Name)
    if folder.Name == my_mailbox:                           #find mailbox
        for inx, subfolder in Oli(folder.Folders).items():
            print("(%i)" % inx, subfolder.Name)             #print folder name
            if subfolder.Name == 'Inbox':                   #find Inbox
                for inx, subsubfolder in Oli(subfolder.Folders).items():
                    print("(%i)" % inx, subsubfolder.Name)  #print folder name
                    if subsubfolder.Name == my_folder:      #find the folder
                        folder_found = True                 #set indicator that folder is found
                        break                               #else error out
            if folder_found is True:                        #exit if folder found
                break
    if folder_found is True:                                #exit if folder found
        break
if folder_found is True:                                    #if the folder was found
    print(subsubfolder.Name)                                #print folder name
    msgs = subsubfolder.Items                               #all mails inside folder
    msg_cnt = msgs.Count                                    #number of mails inside folder
    server_up = False                                       #initialise variable
    for msg in msgs:                                        #loop through each message
        date = str(msg.ReceivedTime).split()[0]             #get mail date
        #print date                                         #print date
        if date == yesterday:                               #if mail was received yesterday
            server_up = True                                #set indicator to say server has sent mails
    if server_up is False:                                  #if no mails received, display error in popup
        print("Server Down. No mails received yesterday. Check Server.")
        ctypes.windll.user32.MessageBoxA(0, "No mails received yesterday. Check Server.", "Server Down", 0)
