
********************************************
** Flexera-deployment-Windows SCRIPT **
********************************************

This script allows CloudOPS to deploy Flexera Agent on Windows server by subscription or from an import csv file.
Azure Run command is used to do it from a central point and as a root user.
Update the software share location on the script line no 194 ie: $UNCPath="\\10.2.0.7\Share\CPBWinAgent\"

STEPS:

1 - Open PowershellISE on your local computer
2 - Load the script
3 - Run the script
4 - Login to azure environment: login-AzAccount
5 - Choose the scope: 1: Full subscription / 2: Input CSV File
6 - If you choose Full subscription, a graphical unit interface shows the list of subscription available. Select one
     If you choose Input CSV File, an explorer opens and navigate to select your csv. CSV Headers required: 'Name', 'ResourceGroupName', 'SubscriptionName' (Name = Azure VM Name)
7 - Script uses Azure Run Command to process
8 - Log file is created on your local PC in C:\Campbell\Logs

CHANGELOG:
- v1: init

