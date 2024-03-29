set obj=CreateObject("Excite.l022")

obj.CloseProcessByName "IEXPLORE.EXE"
obj.wait 0.4

url=obj.GetOperationData(StrFromConfig,"url")

obj.Run url

htmlTitle="HP Employee Portal"
set browser=obj.IEWindow(htmlTitle)
set doc=obj.HtmlDoc(browser)

set userAcctObj=obj.TestObj(doc,"user")
set userPwdObj=obj.TestObj(doc,"pwd")
set btnLogin=obj.TestObj(doc,"btnLogon")

userAcct=obj.GetOperationData(StrFromConfig,"loginUser")
userPwd=obj.GetOperationData(StrFromConfig,"loginPwd")

obj.SetValue userAcctObj,userAcct

obj.wait 1

obj.SetSecure userPwdObj,userPwd

btnLogin.click

obj.wait 1

calledStatement="ClickPopupWinBtn " & chr(34) & "Security Alert" & chr(34)
obj.InvokeFunctionOrSub "customize", calledStatement, strRefPara

obj.WaitForPageLoad browser,timelast

if browser.locationURL<>"http://dcc.itg.gui.ecomcat.hp.com/eComCat/do/selectAccess" then
	obj.Reporter  reportFileName, false, "", "Check Login","Login failed!"
	ActionResult="fail"
	obj.ExitTest
else
	obj.Reporter  reportFileName, true, "", "Check Login","Login success!"
	ActionResult="pass"
end If