set obj=CreateObject("Excite.l022")

htmlTitle="Select Access"
set browser=obj.IEWindow(htmlTitle)
set doc=obj.HtmlDoc(browser)

set listAccessRole=obj.TestObj(doc,"listAccessRole")
set btnSubmit=obj.TestObj(doc,"btnSubmit")


access=obj.GetOperationData(StrFromPara,"AccessRole")

obj.SelectValue listAccessRole,access

obj.wait 1

btnSubmit.click

obj.WaitForPageLoad browser,timelast