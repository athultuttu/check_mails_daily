# check_mails_daily
Python script to check Microsoft Outlook folders to see if new mails has been received.

I had been maintaining a server which sents out report mail based on collected data. As it is a remote server, I had no idea how to monitor if server was up always. As the mails were moved to a folder in mailbox using rules, i wont notice incase mails were not received for few days. So, i had to figure out a way to notify myself in case mails were not received yesterday.

This script was created with this purpose in mind to schedule a task and monitor my mailbox everyday to check if mails has been received in the respective folder on the previous day. If mails were received, script would end normally and if not, script will show a popup with warning message and stay open until I acknowlwdge. This will help me check the server immideately and generate pending reports if any.

Dependencies: None
