# O365-Groups-and-Members

Attached script to export Microsoft 365 Groups, Distribution Lists, and Security Groups. They get saved into 3 different xlsx files at C:\scripts.

Within each Excel file, each group is separated into a new sheet with all usernames and email addresses in the group.

You'll need to login twice. Once for Microsoft 365 groups and Distribution lists and once for the Security groups. Because security groups need a different cmdlet to pull the info for some reason. 

If the group name is over 31 characters long or contains any special characters (\  /  ?  *  [  ]), it cannot name the sheet to that group name due Excel's limitations. There shouldn't be too many so just manually edit those few if needed. I'll find a workaround if it becomes a problem.
