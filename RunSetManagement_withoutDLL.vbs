On Error Resume Next


ReadParameters "RunSetManagement",  sCommons, scriptfolder,  sParameters
vbsFolderPath=CreateObject("Wscript.Shell").CurrentDirectory & "\"

arrPara=split(sParameters,"|")
For i=0 to (ubound(arrPara)-1)/2
	scriptRowNo=arrPara(i*2)
	scriptName=arrPara(i*2+1)
	scriptFullName=vbsFolderPath & scriptfolder & "\" & scriptName & ".vbs"
	CreateObject("WSCript.shell").Run chr(34) & scriptFullName & chr(34)
	WriteInfo2Excel "RunManager",scriptRowNo,scriptFullName
Next




























