Set obj=CreateObject("Excite.l022")

htmlTitle="Product list for catalog master"
set browser=obj.IEWindow(htmlTitle)

set doc=obj.HtmlDoc(browser)

set menuMaintainCatalogs=obj.TestObj(doc,"menuMaintainCatalogs")

menuMaintainCatalogs.Click

obj.WaitForPageLoad browser,timelast

set listSort=obj.TestObj(doc,"listSort")

Set tblCat=obj.TestObj(doc,"tblCat")

obj.SelectValue listSort,"Descending"

Set linkLastDate=obj.GetChildItemFromTable(tblCat, 1, 8, nodetype)

linkLastDate.Click

obj.WaitForPageLoad browser,timelast

htmlTitle="Price Master Summary"
calledStatement="ReturnRowNOByColumnValue " & chr(34) & htmlTitle & chr(34) & "," & chr(34) & "tblCat" & chr(34) & ",3," & chr(34) & masterCatalogForDel & chr(34) & ",rowNo"
result=obj.InvokeFunctionOrSub( "customize", calledStatement, strRefPara)

if strcomp(result,"true",1)=0 Then

    rowNo=CLng(strRefPara)

    Set tblCat=obj.TestObj(doc,"tblCat")

    Set chkboxMC=obj.GetChildItemFromTable(tblCat, rowNo, 1, nodetype)
    obj.ChangeCheckbox chkboxMC,"on"

    Set btnDel=obj.TestObj(doc,"btnDel")
    btnDel.click

    obj.WaitForPageLoad browser,timelast

end If


































