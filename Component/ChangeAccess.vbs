set obj=CreateObject("Excite.l022")

set browser=obj.IEWindow(htmlTitle)

set doc=obj.HtmlDoc(browser)

set linkChngAcss=obj.GetObjByTag(doc,"a", "Change Access")

linkChngAcss.click

obj.WaitForPageLoad browser,timelast

