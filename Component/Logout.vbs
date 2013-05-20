set obj=CreateObject("Excite.l022")

htmlTitle="Select Access"
set browser=obj.IEWindow(htmlTitle)
set doc=obj.HtmlDoc(browser)

set linkLogout=obj.TestObj(doc,"linkLogout")

linkLogout.click

obj.wait 1

calledStatement="ClickPopupWinBtn " & chr(34) & "Microsoft Internet Explorer" & chr(34)
obj.InvokeFunctionOrSub "customize", calledStatement, strRefPara
