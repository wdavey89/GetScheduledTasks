import win32com.client, datetime

# If using in a domain environment, fill out the below and you can remote connect to machines. Then uncomment the 'schedule.Connect(computername, username, domain, password) line
# and comment out 'schedule.Connect() line
# For local machine use, leave as is
computername = 'computername'
username = 'username'
password = 'password'
domain = 'domain'

schedule = win32com.client.Dispatch('Schedule.Service')
#schedule.Connect(computername, username, domain, password)
schedule.Connect()

folders = [schedule.GetFolder('\\')] # Get the root folder.
taskname = []

# Populates the folders with from the root of the task scheduler and adds to a list. Then, ignores the folders with the names Microsoft, Google & Optimize
# as well as ignores all tasks starting with 'OneDrive'
# Information of each task that doesn't match that criteria is returned to the console, with its 'task name', 'enabled or not', 'last run time', and 'current state' 
while folders:
    folder = folders.pop(0)
    folders += list(folder.GetFolders(0))
    for task in folder.GetTasks(0):
        tpath = str(task.Path)
        tname = str(task.Name)
        if not tpath.startswith(('\Microsoft', '\Google', '\Optimize')) and not tname.startswith(('OneDrive')):
            name = task.Name
            enabled = task.Enabled
            path = task.Path                
            lastrun = str(task.LastRunTime.strftime('%d-%m-%y %H:%m'))
            missedruns = task.NumberOfMissedRuns
            state = task.State
            print(name, enabled, lastrun, missedruns, state)
