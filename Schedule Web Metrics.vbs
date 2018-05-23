'=============================
'  Declarations
'=============================
Dim WshShell  
Dim fso
Dim myfile
Dim myPath
Dim myReportPath
Dim myReport
Dim myViewPath, myViewPath2

'=============================
'  Set up paths
'=============================

myPath = "\\view\general\TEST_DISCIPLINE\4.Facilities\Web Based Metrics\Web based TM Metrics v3.0 Overnight.xlsm"
myRunScript = "M:\Run_Web_Metrics.vbs"

'=============================
'  Create the run file on M:\
'=============================

	'  Create the file system object to create a new schedule file to run
	set fso = CreateObject("Scripting.FileSystemObject")

	'  Create the run script on the M drive
	Set myfile = fso.OpenTextFile(myRunScript,2,True)

	myfile.writeline ("'=============================")
	myfile.writeline ("'  Declarations")
	myfile.writeline ("'=============================")
	myfile.writeline ("dim objXL, objWB")

	myfile.writeline ("'=============================")
	myfile.writeline ("'  Run the Excel section")
	myfile.writeline ("'=============================")
	myfile.writeline ("'  Set up Excel")
	myfile.writeline ("Set objXL = CreateObject(""Excel.Application"")")
	myfile.writeline ("'  Open the file which runs the metrics code")
	myfile.writeline ("Set objWB = objXL.WorkBooks.Open(""" & myPath & """)")
	myfile.writeline ("'  Quit Excel")
	myfile.writeline ("objXL.Application.Quit")
	myfile.writeline ("Set objXL = Nothing")
	myfile.writeline ("Set objWB = Nothing")

	myfile.close

	set fso = Nothing

'=============================
'  Set up the schedule
'=============================
	
	'  Now create the schedule
	Set WshShell = CreateObject("WScript.Shell")

	'  Run the schedule
	wshell_cmd = "/create /tn " & """Run_Web_Metrics""" & " /tr " & myRunScript & " /sc weekly /d mon,tue,wed,thu,fri /st 23:30:00"
	WshShell.Run  "cmd /c c:\windows\system32\schtasks.exe " & wshell_cmd