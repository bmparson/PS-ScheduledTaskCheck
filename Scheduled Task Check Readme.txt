This script will check all scheduled tasks running on a Windows server and alert on any failures.
The script is intended to be run as a scheduled task by an account with domain admin privileges.
There are some files it creates and it requires some files to be present in order to run properly.
To install you need to copy the script to a folder, by default this is C:\Scripts\ScheduledTaskCheck\.
You can edit the default folder if you want but make sure you find and replace each instance in the script.
You'll need to edit the smtp server and email address at the bottom to match your environment before use.
You also need to create a servers.txt file in the install directory which is just the DNS Host Name of the servers, one per line.
If you want to set a threshold of failures before sending an alert create a csv file named taskfailurethreshold.csv in the install directory.
The threshold csv need to have two columns in this order Taskname Threshold.
Taskname contains the name of the task you want to set a threshold on and threshold is the number of sequential failures before triggering an alert.
This script will also generate a report in the install directory of all tasks which includes the taskname, runas user, command run, arguments to the command run, triggers,
and server the task runs from.
