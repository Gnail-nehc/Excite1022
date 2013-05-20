Set obj=CreateObject("Excite.l022")

saleCountry=obj.GetOperationData(StrFromPara,"SaleCountry")

priceDescriptor=obj.GetOperationData(StrFromPara,"PriceDescriptor")

deal=obj.GetOperationData(StrFromPara,"Deal")

includeStdCat=obj.GetOperationData(StrFromPara,"IncludeStdCat")

stdCatName=obj.GetOperationData(StrFromPara,"StdCatName")

priceTier=obj.GetOperationData(StrFromPara,"PriceTier")

plcCheck=obj.GetOperationData(StrFromPara,"PLCCheck")

htmlTitle="View Customer"
set browser=obj.IEWindow(htmlTitle)

set doc=obj.HtmlDoc(browser)

set menuCreateCatalogs=obj.TestObj(doc,"menuCreateCatalogs")

menuCreateCatalogs.click

obj.WaitForPageLoad browser,timelast

set btnCreate=obj.TestObj(doc,"btnCreate")

btnCreate.click

obj.WaitForPageLoad browser,timelast

if saleCountry<>"" then

    set listSaleCountry=obj.TestObj(doc,"listSaleCountry")

    obj.SelectValue listSaleCountry,saleCountry

    obj.WaitForPageLoad browser,timelast

end if

if priceDescriptor<>"" Then

    set listPriceDes=obj.TestObj(doc,"listPriceDes")

    obj.SelectValue listPriceDes,priceDescriptor

    obj.WaitForPageLoad browser,timelast

end If

masterCatalogName="qatest_"& obj.generateRandomString

set txtMCName=obj.TestObj(doc,"txtMCName")

set flagIncludePT=obj.TestObj(doc,"flagIncludePT")

set listAvailablePT=obj.TestObj(doc,"listAvailablePT")

set btnSelectPT=obj.TestObj(doc,"btnSelectPT")

Set btnSelectSC =obj.TestObj(doc,"btnSelectSC")
	
set tblDealList=obj.TestObj(doc,"tblDealList")

set listIncludeSC=obj.TestObj(doc,"listIncludeSC")

set listAvailableSC =obj.TestObj(doc,"listAvailableSC")

set flagPLCCheck =obj.TestObj(doc,"flagPLCCheck")

set btnSave =obj.TestObj(doc,"btnSave")

obj.SetValue txtMCName,masterCatalogName

if stdCatName<>"" then

    obj.ChangeCheckbox flagIncludePT,"off"

    obj.SelectValue listIncludeSC,includeStdCat
    listIncludeSC.click
    obj.SelectValue listIncludeSC,includeStdCat
    obj.Wait 1
	
    obj.SelectValue listAvailableSC,stdCatName

    obj.Wait 1
    btnSelectSC.click
    
    obj.ChangeCheckbox flagPLCCheck,plcCheck

elseif priceTier<>"" Then

    obj.ChangeCheckbox flagIncludePT,"on"

    obj.Wait 0.6

    obj.SelectValue listAvailablePT,priceTier

    obj.Wait 1
	
    btnSelectPT.click
    obj.Wait 1
	btnSelectPT.click
	
end If

htmlTitle="Create Catalog Master"
calledStatement="ReturnRowNOByColumnValue " & chr(34) & htmlTitle & chr(34) & "," & chr(34) & "tblDealList" & chr(34) & ",3," & chr(34) & deal & chr(34) & ",rowNo"
result=obj.InvokeFunctionOrSub( "customize", calledStatement, strRefPara)

If strcomp(result,"true",1)=0 Then
    rowNo=CLng(strRefPara)
    Set flagDeal=obj.GetChildItemFromTable(tblDealList, rowNo, 1, nodetype)
    obj.ChangeCheckbox flagDeal,"on"
End If

btnSave.click

obj.WaitForPageLoad browser,timelast

O_MCName=masterCatalogName