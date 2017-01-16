OPTION EXPLICIT
DIM strComputer,strProcess

strComputer = "." ' local computer
strProcess = "smc.exe"

' Check if Calculator is running on specified computer (. = local computer)
IF Not isProcessRunning(strComputer,strProcess) THEN
	wscript.echo strProcess & " is NOT running on computer "
ELSE
	wscript.echo strProcess & " is running on computer "
END IF

' Function to check if a process is running
FUNCTION isProcessRunning(BYVAL strComputer,BYVAL strProcessName)

	DIM objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	
	SET objWMIService = GETOBJECT("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 


	IF objWMIService.ExecQuery(strWMIQuery).Count > 0 THEN
		isProcessRunning = TRUE
	ELSE
		isProcessRunning = FALSE
	END IF

END FUNCTION