Sub ClickPopupWinBtn(byval caption)
	Set obj=CreateObject("Excite.l022")
	If  not obj.IsRegistered("DynamicWrapper") Then
		sourceFolder=obj.GetSourceFolder
		If right(sourceFolder,1)<>"\" Then
			sourceFolder=sourceFolder & "\"
		End If
		CreateObject("WSCript.shell").Run "cmd /c regsvr32 /s "& Chr(34) & sourceFolder & "dynwrap.dll" & Chr(34)
	End If
	Set dWrap = CreateObject("DynamicWrapper")
	dWrap.Register "user32.dll", "SendMessage", "i=hlll", "f=s", "r=l"
	dWrap.Register "user32.dll", "FindWindow", "i=ss", "f=s", "r=l"
	dWrap.Register "user32.dll", "FindWindowEx", "i=hhsl", "f=s", "r=h"
	if obj.IsWindowExist(dWrap,caption,winHnd) then
		obj.ClickPopupWinBtn dWrap,winHnd,"Button"
	else
		CreateObject("Wscript.Shell").popup "The Dialog box not found.",6,"The End",64
	end If
End Sub

Function ReturnRowNOByColumnValue(byval htmlTitle,byval objNameInOR,byval colNo,byval expectedValue,byref rowNo)
	Set obj=CreateObject("Excite.l022")
	set browser=obj.IEWindow(htmlTitle)
	set doc=obj.HtmlDoc(browser)
	set tblObj=obj.TestObj(doc,objNameInOR)
	rowsum=obj.GetTableRows(tblObj)
	For i=2 To rowsum
		strval=obj.GetChildItemFromTable(tblObj,i,colNo,nodetype)
		If StrComp(strval,expectedValue,1)=0 Then
			ReturnRowNOByColumnValue=true
			rowNo=i
			Exit for
		End if
	Next
	If i>rowsum Then
		ReturnRowNOByColumnValue=false
	end if
End Function



Function ReturnColNOByName(ByVal htmlTitle,byval objNameInOR,byval expectedValue,byref colNo)
	Set obj=CreateObject("Excite.l022")
	Set browser=obj.IEWindow(htmlTitle)
	Set doc=obj.HtmlDoc(browser)
	Set tblObj=obj.TestObj(doc,objNameInOR)
	colsum=obj.GetTableColumns(tblObj)
	For i=8 To colsum
		Call obj.GetChildItemFromTable(tblObj,1,i,nodetype)
		If instr(1,nodetype,"text",1)<=0 and nodetype<>empty Then
			set objColumn=obj.GetChildItemFromTable(tblObj,1,i,nodetype)
			strval=obj.GetTextOfObj(objColumn)
		ElseIf instr(1,nodetype,"text",1)>0 Then
			strval=obj.GetChildItemFromTable(tblObj,1,i,nodetype)
		Else
			strval=" "
		end If
		If instr(strval,chr(63))>1 Or instr(strval,chr(160))>1 Then
			strval=Replace(strval,chr(160)," ")
			strval=Replace(strval,chr(63)," ")
		End If
		If StrComp(strval,expectedValue,1)=0 Then
			ReturnColNOByName=True
			colNo=i
			Exit For
		End If
	Next
	If i>colsum Then
		ReturnColNOByName=False
	End If
End Function