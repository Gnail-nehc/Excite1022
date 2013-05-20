Set obj=CreateObject("Excite.l022")
htmlTitle="Welcome to WW Comcat"
set browser=obj.IEWindow(htmlTitle)

set doc=obj.HtmlDoc(browser)

set linkSelCust=obj.TestObj(doc,"linkSelCust")

custName=obj.GetOperationData(StrFromPara,"CustomerName")
custID=obj.GetOperationData(StrFromPara,"CustomerID")

linkSelCust.Click

obj.WaitForPageLoad browser,timelast

set txtCustName=obj.TestObj(doc,"txtCustName")

set btnSearch=obj.TestObj(doc,"btnSearch")

obj.SetValue txtCustName,custName

obj.wait 0.6

btnSearch.click

obj.WaitForPageLoad browser,timelast

htmlTitle="Select Customer"
calledStatement="ReturnRowNOByColumnValue " & chr(34) & htmlTitle & chr(34) & "," & chr(34) & "tblCat" & chr(34) & ",3," & chr(34) & custID & chr(34) & ",rowNo"
result=obj.InvokeFunctionOrSub( "customize", calledStatement, strRefPara)

if strcomp(result,"true",1)=0 then
	rowNo=CLng(strRefPara)
	set tblCat=obj.TestObj(doc,"tblCat")
	set linkSelect=obj.GetChildItemFromTable(tblCat, rowNo, 1, nodetype)

	linkSelect.click
	obj.WaitForPageLoad browser,timelast
end if




































