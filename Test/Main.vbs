'On Error Resume Next



dataVersion="version1.0"

scriptName=wscript.scriptname


set obj=CreateObject("Excite.l022")


if err.number<>0 Then


	dllFile=Replace(CreateObject("Wscript.Shell").CurrentDirectory, split(CreateObject("Wscript.Shell").CurrentDirectory,"\")(ubound(split(CreateObject("Wscript.Shell").CurrentDirectory,"\"))),"") & "Excite.dll"


	CreateObject("WSCript.shell").Run "cmd /c regsvr32 /s " & Chr(34) & dllFile & Chr(34)

	
obj.wait 0.4


end
 If


reportFileName=obj.CurrentReportFile(scriptName)


call obj.InitialTestData (scriptName,dataVersion,strConfig,arrParameters)



obj.RunComponent "Login","reportFileName",reportFileName,"StrFromConfig",strConfig,"","","ActionResult",runResult

obj.wait 1



For i=0 To ubound(arrParameters)
	
	obj.RunComponent "SelectAccess","","","","","StrFromPara",arrParameters(i) ,"",""
	
obj.wait 0.6

	obj.RunComponent "SelectCustomer","","","","","StrFromPara",arrParameters(i) ,"",""

	obj.wait 0.6

	obj.RunComponent "CreateMasterCatalog","","","","","StrFromPara",arrParameters(i) ,"O_MCName",masterCatalogName

	obj.wait 0.6

	obj.RunComponent "DeleteMasterCatalog","","","masterCatalogForDel",masterCatalogName,"","" ,"",""

	obj.wait 0.6

	obj.RunComponent "ChangeAccess","","","htmlTitle","Price Master Summary","","" ,"",""
Next

obj.RunComponent "Logout","","","","","","" ,"",""


obj.wait 1


obj.Reporter  reportFileName, "", "", "Script Done","The script "&scriptName&" has finished!"


obj.OutputTheResult reportFileName



CreateObject("Wscript.Shell").popup "The script "&scriptName&" has finished.",3,"The End",64

