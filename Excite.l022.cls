VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "l022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.1
'
' NAME:
'
' AUTHOR: Liang Chen , Hewlett-Packard Company
' DATE  : 2010-10-22
'
' COMMENT:
'
'==========================================================================


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Function Wait(ByVal sec)
    Sleep sec * 1000
End Function


Function Run(ByVal url)
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .navigate "about:blank"
        .Left = 0
        .Top = 0
        .Width = .document.parentWindow.Screen.availWidth
        .Height = .document.parentWindow.Screen.availHeight
        .Visible = True
    End With
    ie.navigate url
    Wait 1
    While ie.busy:     Wend
End Function


Function IEWindow(ByVal title)
    Set objShell = CreateObject("Shell.Application")
    Set objShellWindows = objShell.Windows
    'objShellWindows.count   means   amount of the IE windows and  Folders  in  Current Dssktop
    For i = 0 To objShellWindows.Count - 1
        On Error Resume Next
        If Not IsNothing(objShellWindows(i)) Then
            If InStr(1, objShellWindows(i).locationName, title, 1) > 0 Then
                Exit For
            End If
        End If
    Next
    If i < objShellWindows.Count Then
        Set IEWindow = objShellWindows(i)
    Else
        Set IEWindow = Nothing
    End If
End Function

Function HtmlDoc(ByVal browser)
    If Not IsNothing(browser) Then
        Set HtmlDoc = browser.document
    Else
        Set HtmlDoc = Nothing
    End If
End Function

Function IsNothing(ByVal obj)
    If TypeName(obj) = "Nothing" Then
        IsNothing = True
    Else
        IsNothing = False
    End If
End Function


Function WaitForPageLoad(browser, ByRef timeLast)
    waitTime = 15
    While browser.readystate = 4 And waitTime > 0
        Wait 1
        waitTime = waitTime - 1
    Wend
    begin = FormatCurrentTime(Time)
    While browser.readystate <> 4:   Wend
    ends = FormatCurrentTime(Time)
    sec = ends - begin
    If sec >= 0 Then
        timeLast = sec
    Else
        timeLast = sec + 24 * 3600
    End If
End Function

Function FormatCurrentTime(ByVal timing)
    timing = CStr(timing)
    h = Split(timing, ":")(0)
    m = Split(timing, ":")(1)
    s = Split(timing, ":")(2)
    h = CLng(h)
    m = CLng(m)
    s = CLng(s)
    FormatCurrentTime = h * 3600 + m * 60 + s
End Function

Function GetSourceFolder()
    Set objShell = CreateObject("Wscript.Shell")
    vbsFolder = objShell.CurrentDirectory
    arrpath = Split(vbsFolder, "\")
    find = arrpath(UBound(arrpath))
    GetSourceFolder = Replace(vbsFolder, find, "")
End Function

Function CurrentReportFile(ByVal scriptName)
   If StrComp(Right(scriptName, 4), ".vbs", 1) = 0 Then
        scriptName = Left(scriptName, Len(scriptName) - 4)
   End If
    CurrentReportFile = "Report for_" & scriptName & "_" & Replace(Date, "/", ".") & "_" & Replace(Time, ":", "-")
End Function



'@expectedAttributes :  name    class     id    value  ..etc
'@ Return the first object if the more than one object match the expected attribute .
Function WebObj(ByVal doc, ByVal htmlTag, ByVal attributeName, ByVal expectedAttrVal)
    On Error Resume Next
    Set objContainer = doc.getElementsByTagName(htmlTag)
    For cnt = 0 To objContainer.length - 1
        actualAttrVal = objContainer(cnt).Attributes.getNamedItem(attributeName).Value
        If StrComp(expectedAttrVal, actualAttrVal, 1) = 0 Then
            Set WebObj = objContainer(cnt)
            Exit Function
        End If
    Next
End Function




' a a a a a a a a a a a a a a a a a a a a a a a a a aRead  TestObj   Start

Function testObj(ByVal doc, ByVal names)
    xmlFile = "ObjectRepository.xml"
    sourceFolder = GetSourceFolder
    xmlPath = sourceFolder & xmlFile
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.async = False
    xmlDoc.Load xmlPath
    If xmlDoc.parseError.errorCode <> 0 Then
        MsgBox "XML file's format is wrong,the reason is:" & Chr(13) & xmlDoc.parseError.reason
        ExecuteByXML = False
        Exit Function
    End If
    Set xmlRoot = xmlDoc.documentElement
    nodes = xmlRoot.ChildNodes.length
    For i = 0 To nodes - 1
        Set testObj = xmlRoot.ChildNodes.Item(i)
        Set objName = testObj.ChildNodes.Item(0)
        If StrComp(objName.Text, names, 1) = 0 Then
            Exit For
        End If
    Next
    txtHtmlTag = testObj.ChildNodes.Item(1).Text
    txtAttribute = testObj.ChildNodes.Item(2).Text
    txtAttrValue = testObj.ChildNodes.Item(3).Text
    Set testObj = WebObj(doc, txtHtmlTag, txtAttribute, txtAttrValue)
End Function

' a a a a a a a a a a a a a a a a a a a a a a a a a a a aRead  TestObj   End


Function ReturnAttributeValue(ByVal testObj, ByVal attributeName, ByRef attrVal)
    strAttrs = ""
    For i = 0 To testObj.Attributes.length - 1
        strAttrs = strAttrs & testObj.Attributes(i).Name & "|"
    Next
    If InStr(strAttrs, attributeName) > 0 And testObj.Attributes.getNamedItem(attributeName).Value <> "" Then
        attrVal = testObj.Attributes.getNamedItem(attributeName).Value
        ReturnAttributeValue = True
    Else
        attrVal = ""
        ReturnAttributeValue = False
    End If
End Function



Function GetObjByTag(ByVal doc, ByVal htmlTag, ByVal descriptionInTag)
    Set objContainer = doc.getElementsByTagName(htmlTag)
    For cnt = 0 To objContainer.length - 1
        varOuterhtml = objContainer(cnt).outerhtml
        If InStr(varOuterhtml, descriptionInTag) > 0 Then
            Set GetObjByTag = objContainer(cnt)
            Exit For
        End If
    Next
    If cnt >= objContainer.length Then
        Set GetObjByTag = Nothing
    End If
End Function




Function SetValue(ByVal testObj, ByVal strValue)
    testObj.focus
    If testObj.Type = "text" Or testObj.Type = "password" Then
        testObj.Value = strValue
    End If
End Function

Function SetSecure(ByVal testObj, ByVal strValue)
   testObj.focus
    If testObj.Type = "password" Then
        strValue = Decoder(strValue)
        testObj.Value = strValue
    End If
End Function



Function SelectValue(ByVal testObj, ByVal strValue)
   testObj.focus
   testObj.Click
   If InStr(testObj.Type, "select") > 0 Then
        For i = 0 To testObj.ChildNodes.length - 1
            If InStr(1, testObj.ChildNodes.Item(i).nodename, "text", 1) <= 0 Then
                itemValue = testObj.ChildNodes.Item(i).Text
                If StrComp(itemValue, strValue, 1) = 0 Then
                    testObj.Value = testObj.ChildNodes.Item(i).Value
                    If i = testObj.ChildNodes.length - 1 Then
                        CreateObject("wscript.shell").SendKeys "{UP}"
                        CreateObject("wscript.shell").SendKeys "{DOWN}"
                    Else
                        CreateObject("wscript.shell").SendKeys "{DOWN}"
                        CreateObject("wscript.shell").SendKeys "{UP}"
                    End If
                    Exit For
                End If
            End If
        Next
   End If
End Function



Function ChangeCheckbox(ByVal testObj, ByVal boolOrStr)
    testObj.focus
    If StrComp(boolOrStr, "on", 1) = 0 Then
        boolOrStr = True
    ElseIf StrComp(boolOrStr, "off", 1) = 0 Then
        boolOrStr = False
    End If
    If VarType(boolOrStr) = 11 Then
        If InStr(testObj.Type, "checkbox") > 0 Then
            testObj.Checked = boolOrStr
            CreateObject("wscript.shell").SendKeys "  "
        End If
    End If
    Wait 0.5
End Function


Function ChangeRadio(ByVal testObj, ByVal boolOrStr)
   testObj.focus
    If StrComp(boolOrStr, "on", 1) = 0 Then
        boolOrStr = True
    ElseIf StrComp(boolOrStr, "off", 1) = 0 Then
        boolOrStr = False
    End If
    If VarType(boolOrStr) = 11 Then
        If InStr(testObj.Type, "radio") > 0 Then
            testObj.Checked = boolOrStr
            CreateObject("wscript.shell").SendKeys "  "
        End If
    End If
    Wait 0.5
End Function



Function GetChildItemFromTable(ByVal webTbl, ByVal rowNo, ByVal colNo, ByRef nodetype)
    expectedNodename = "TR"
    If rowNo > 1 Then
        expectedRowTag = "TD"
    Else
        expectedRowTag = "TH"
    End If
    Do While webTbl.ChildNodes.length > 0
        i = 0
        Do While i < webTbl.ChildNodes.length
            If StrComp(webTbl.ChildNodes.Item(i).nodename, expectedNodename, 1) = 0 Then
                rowNo = rowNo - 1
                If rowNo = 0 Then
                    If colNo = 0 Then
                        If webTbl.ChildNodes.Item(i).ChildNodes.length > 0 Then
                            Set node = webTbl.ChildNodes.Item(i).ChildNodes.Item(0)
                            Set GetChildItemFromTable = node
                            nodetype = node.nodename
                            If InStr(nodetype, "#text") > 0 Then
                                nodevalue = node.nodevalue
                                GetChildItemFromTable = Replace(nodevalue, Chr(160), Chr(32))
                            End If
                        End If
                        Exit Function
                    End If
                    Set webTbl = webTbl.ChildNodes.Item(i)
                    expectedNodename = expectedRowTag
                    rowNo = colNo
                    colNo = 0
                    i = -1
                End If
            End If
            i = i + 1
        Loop
        Set webTbl = webTbl.ChildNodes.Item(0)
    Loop
End Function


Function GetTableRows(ByVal webTbl)
    Dim counter
    counter = 0
    Do While webTbl.ChildNodes.length > 0
        For i = 0 To webTbl.ChildNodes.length - 1
            If StrComp(webTbl.ChildNodes.Item(i).nodename, "TR", 1) = 0 Then
                counter = counter + 1
            End If
        Next
        Set webTbl = webTbl.ChildNodes.Item(0)
    Loop
    GetTableRows = counter
End Function


Function GetTableColumns(ByVal webTbl)
    Dim counter
    counter = 0
    Do While webTbl.ChildNodes.length > 0
        For i = 0 To webTbl.ChildNodes.length - 1
            If StrComp(webTbl.ChildNodes.Item(i).nodename, "TR", 1) = 0 Then
                Set webTbl = webTbl.ChildNodes.Item(i)
                Exit Do
            End If
        Next
        Set webTbl = webTbl.ChildNodes.Item(0)
    Loop
    Do While webTbl.ChildNodes.length > 0
        For j = 0 To webTbl.ChildNodes.length - 1
            If StrComp(webTbl.ChildNodes.Item(j).nodename, "TH", 1) = 0 Or StrComp(webTbl.ChildNodes.Item(j).nodename, "TD", 1) = 0 Then
                counter = counter + 1
            End If
        Next
        Set webTbl = webTbl.ChildNodes.Item(0)
    Loop
    GetTableColumns = counter
End Function


Function GetTextOfObj(ByVal testObj)
    If testObj.canHaveChildren Then
        For i = 0 To testObj.ChildNodes.length - 1
            If InStr(testObj.ChildNodes.Item(i).nodename, "text") > 0 Then
                GetTextOfObj = testObj.ChildNodes.Item(i).nodevalue
                GetTextOfObj = Replace(GetTextOfObj, Chr(160), Chr(32))
                Exit For
            End If
        Next
    Else
        If testObj.parentnode.ChildNodes.length > 0 Then
            Do
                Set testObj = testObj.nextsibling
                If InStr(testObj.nodename, "text") > 0 Then
                    GetTextOfObj = testObj.nodevalue
                    GetTextOfObj = Replace(GetTextOfObj, Chr(160), Chr(32))
                    Exit Do
                End If
            Loop
        End If
    End If
End Function




'@ sender : destination email address , if  receiver  is more than one, using  ";"  to separated  each receiver
'@ attached :  if  the value is not empty, it  is attachment  filename include absoluted filepath
Sub SendEmail(ByVal sender, ByVal attached)
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMailItem = OutlookApp.CreateItem(olMailItem)
    OutlookMailItem.To = sender
    OutlookMailItem.subject = "Test Report" & Date & " " & Time
    OutlookMailItem.body = "Please refer to attachment."
    If attached <> "" Then
        OutlookMailItem.Attachments.Add attached
    End If
    OutlookMailItem.save
    OutlookMailItem.send
    Set OutlookMailItem = Nothing
    Set OutlookApp = Nothing
End Sub




'Customize the htm's report by recording stepname,testresult,checkcontent and description of the each step
'@  filepath="C:\Documents and Settings\lchen3\Desktop\results"
'@  filename="TestResult_" & Environment("TestName") & date & "_" & split(time,":")(0) & ".htm"
'@  results:  boolean
Function Reporter(ByVal filename, ByVal results, ByVal steps, ByVal title, ByVal contents)
   sourceFolder = GetSourceFolder
    If Right(sourceFolder, 1) <> "\" Then
        sourceFolder = sourceFolder & "\"
    End If
    fileFolder = sourceFolder & "Report" & "\"
    Dim column1, column2, column3, column4, subject
    subject = "Check Test Result : " & Date & " " & Time
    column1 = "Result"
    column2 = "Step"
    column3 = "Title"
    column4 = "Description"
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fileFolder) Then
         fso.CreateFolder (fileFolder)
    End If
    Dim abs_filepath
    If InStr(2, filename, ".htm", 1) <= 0 Then
        filename = filename & ".htm"
    End If
    abs_filepath = fileFolder & filename
    If Not (fso.FileExists(abs_filepath)) Then
        Dim f_htm, str
        Set f_htm = fso.CreateTextFile(abs_filepath, True)
        str = "<html xmlns:o=" & Chr(34) & "urn:schemas-microsoft-com:office:office" & Chr(34) & Chr(13)
        str = str & "xmlns:w=" & Chr(34) & "urn:schemas-microsoft-com:office:word" & Chr(34) & Chr(13)
        str = str & "xmlns=" & Chr(34) & "http://www.w3.org/TR/REC-html40" & Chr(34) & ">" & Chr(13)
        str = str & "<head>" & Chr(13)
        str = str & "<meta http-equiv=Content-Type content=" & Chr(34) & "text/html; charset=us-ascii" & Chr(34) & ">" & Chr(13)
        str = str & "<meta name=ProgId content=Word.Document>" & Chr(13)
        str = str & "<meta name=Generator content=" & Chr(34) & "Microsoft Word 9" & Chr(34) & ">" & Chr(13)
        str = str & "<meta name=Originator content=" & Chr(34) & "Microsoft Word 9" & Chr(34) & ">" & Chr(13)
        str = str & "<link rel=File-List href=" & Chr(34) & "./result.files/filelist.xml" & Chr(34) & ">" & Chr(13)
        str = str & "<!--[if gte mso 9]><xml>" & Chr(13)
        str = str & " <o:DocumentProperties>" & Chr(13)
        str = str & "  <o:Author>lio</o:Author>" & Chr(13)
        str = str & "  <o:LastAuthor>lio</o:LastAuthor>" & Chr(13)
        str = str & "  <o:Revision>4</o:Revision>" & Chr(13)
        str = str & "  <o:TotalTime>8</o:TotalTime>" & Chr(13)
        str = str & "  <o:Created>" & Date & Time & "</o:Created>" & Chr(13)
        str = str & "  <o:LastSaved>" & Date & Time & "</o:LastSaved>" & Chr(13)
        str = str & "  <o:Pages>1</o:Pages>" & Chr(13)
        str = str & "  <o:Words>17</o:Words>" & Chr(13)
        str = str & "  <o:Characters>99</o:Characters>" & Chr(13)
        str = str & "  <o:Company>&#38632;&#34183;&#22312;&#32447;</o:Company>" & Chr(13)
        str = str & "  <o:Lines>1</o:Lines>" & Chr(13)
        str = str & "  <o:Paragraphs>1</o:Paragraphs>" & Chr(13)
        str = str & "  <o:CharactersWithSpaces>121</o:CharactersWithSpaces>" & Chr(13)
        str = str & "  <o:Version>9.6926</o:Version>" & Chr(13)
        str = str & " </o:DocumentProperties>" & Chr(13)
        str = str & "</xml><![endif]--><!--[if gte mso 9]><xml>" & Chr(13)
        str = str & " <w:WordDocument>" & Chr(13)
        str = str & "  <w:DrawingGridVerticalSpacing>7.8 &#30917;</w:DrawingGridVerticalSpacing>" & Chr(13)
        str = str & "  <w:Compatibility>" & Chr(13)
        str = str & "   <w:UseFELayout/>" & Chr(13)
        str = str & "  </w:Compatibility>" & Chr(13)
        str = str & " </w:WordDocument>" & Chr(13)
        str = str & "</xml><![endif]-->" & Chr(13)
        str = str & "<style>" & Chr(13)
        str = str & "<!--" & Chr(13)
        str = str & " /* Font Definitions */" & Chr(13)
        str = str & "@font-face" & Chr(13)
        str = str & "   {font-family:SimSun;" & Chr(13)
        str = str & "   panose-1:2 1 6 0 3 1 1 1 1 1;" & Chr(13)
        str = str & "   mso-font-alt:SimSun;" & Chr(13)
        str = str & "   mso-font-charset:134;" & Chr(13)
        str = str & "   mso-generic-font-family:auto;" & Chr(13)
        str = str & "   mso-font-pitch:variable;" & Chr(13)
        str = str & "   mso-font-signature:3 135135232 16 0 262145 0;}" & Chr(13)
        str = str & "@font-face" & Chr(13)
        str = str & "   {font-family:SimSun;" & Chr(13)
        str = str & "   panose-1:2 1 6 0 3 1 1 1 1 1;" & Chr(13)
        str = str & "   mso-font-charset:134;" & Chr(13)
        str = str & "   mso-generic-font-family:auto;" & Chr(13)
        str = str & "   mso-font-pitch:variable;" & Chr(13)
        str = str & "   mso-font-signature:3 135135232 16 0 262145 0;}" & Chr(13)
        str = str & " /* Style Definitions */" & Chr(13)
        str = str & "p.MsoNormal, li.MsoNormal, div.MsoNormal" & Chr(13)
        str = str & "   {mso-style-parent:" & Chr(34) & Chr(34) & ";" & Chr(13)
        str = str & "   margin:0cm;" & Chr(13)
        str = str & "   margin-bottom:.0001pt;" & Chr(13)
        str = str & "   mso-pagination:widow-orphan;" & Chr(13)
        str = str & "   font-size:12.0pt;" & Chr(13)
        str = str & "   font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & Chr(13)
        str = str & "   mso-fareast-font-family:SimSun;}" & Chr(13)
        str = str & "p.MsoTitle, li.MsoTitle, div.MsoTitle" & Chr(13)
        str = str & "   {margin:0cm;" & Chr(13)
        str = str & "   margin-bottom:.0001pt;" & Chr(13)
        str = str & "   text-align:center;" & Chr(13)
        str = str & "   mso-pagination:widow-orphan;" & Chr(13)
        str = str & "   font-size:14.0pt;" & Chr(13)
        str = str & "   mso-bidi-font-size:12.0pt;" & Chr(13)
        str = str & "   font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & Chr(13)
        str = str & "   mso-fareast-font-family:SimSun;" & Chr(13)
        str = str & "   font-weight:bold;}" & Chr(13)
        str = str & " /* Page Definitions */" & Chr(13)
        str = str & "@page" & Chr(13)
        str = str & "   {mso-page-border-surround-header:no;" & Chr(13)
        str = str & "   mso-page-border-surround-footer:no;}" & Chr(13)
        str = str & "@page Section1" & Chr(13)
        str = str & "   {size:595.3pt 841.9pt;" & Chr(13)
        str = str & "   margin:72.0pt 90.0pt 72.0pt 90.0pt;" & Chr(13)
        str = str & "   mso-header-margin:42.55pt;" & Chr(13)
        str = str & "   mso-footer-margin:49.6pt;" & Chr(13)
        str = str & "   mso-paper-source:0;}" & Chr(13)
        str = str & "div.Section1" & Chr(13)
        str = str & "   {page:Section1;}" & Chr(13)
        str = str & "-->" & Chr(13)
        str = str & "</style>" & Chr(13)
        str = str & "</head>" & Chr(13)
        str = str & "<body lang=ZH-CN style='tab-interval:21.0pt'>" & Chr(13)
        str = str & "<div class=Section1>" & Chr(13)
        str = str & "<p class=MsoTitle><span lang=EN-US>" & subject & "</span></p>" & Chr(13)
        str = str & "<div align=center>" & Chr(13)
        str = str & "<table border=1 cellpadding=0 width=990 style='width:742.6pt;mso-cellspacing:" & Chr(13)
        str = str & " 1.5pt;border:none;border-bottom:solid #99CCFF .5pt;mso-padding-alt:0cm 0cm 0cm 0cm'>" & Chr(13)
        str = str & " <tr style='height:23.55pt'>" & Chr(13)
        str = str & "  <td width=70 style='width:52.2pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
        str = str & "  height:23.55pt'>" & Chr(13)
        str = str & "  <p class=MsoNormal align=center style='text-align:center'><b><span" & Chr(13)
        str = str & "  lang=EN-US style='font-size:14.0pt;mso-bidi-font-size:12.0pt;color:#666699'>" & column1 & "<o:p></o:p></span></b></p>" & Chr(13)
        str = str & "  </td>" & Chr(13)
        str = str & "  <td width=82 style='width:61.5pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
        str = str & "  height:23.55pt'>" & Chr(13)
        str = str & "  <p class=MsoNormal align=center style='text-align:center'><b><span" & Chr(13)
        str = str & "  lang=EN-US style='font-size:14.0pt;mso-bidi-font-size:12.0pt;color:#666699'>" & column2 & "<o:p></o:p></span></b></p>" & Chr(13)
        str = str & "  </td>" & Chr(13)
        str = str & "  <td width=298 style='width:223.5pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
        str = str & "  height:23.55pt'>" & Chr(13)
        str = str & "  <p class=MsoNormal align=center style='text-align:center'><b><span" & Chr(13)
        str = str & "  lang=EN-US style='font-size:14.0pt;mso-bidi-font-size:12.0pt;color:#666699'>" & column3 & "<o:p></o:p></span></b></p>" & Chr(13)
        str = str & "  </td>" & Chr(13)
        str = str & "  <td width=531 style='width:397.9pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
        str = str & "  height:23.55pt'>" & Chr(13)
        str = str & "  <p class=MsoNormal align=center style='text-align:center'><b><span" & Chr(13)
        str = str & "  lang=EN-US style='font-size:14.0pt;mso-bidi-font-size:12.0pt;color:#666699'>" & column4 & "<o:p></o:p></span></b></p>" & Chr(13)
        str = str & "  </td>" & Chr(13)
        str = str & " </tr>"
        f_htm.Write (str)
        f_htm.Close
    End If
    Dim fo, str1
    Set fo = fso.OpenTextFile(abs_filepath, 8, True)
    str1 = " <tr style='height:35.25pt'>" & Chr(13)
    str1 = str1 & "  <td width=82 style='width:61.5pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
    str1 = str1 & "  height:35.25pt'>" & Chr(13)
    If results <> "" Then
        If results Then
            str1 = str1 & "  <p class=MsoNormal><span lang=EN-US style='color:green'>Pass<o:p></o:p></span></p>" & Chr(13)
        Else
            str1 = str1 & "  <p class=MsoNormal><span lang=EN-US style='color:red'>Fail<o:p></o:p></span></p>" & Chr(13)
        End If
    Else
        str1 = str1 & "  <p class=MsoNormal><span lang=EN-US style='color:gray'>done<o:p></o:p></span></p>" & Chr(13)
    End If
    
    str1 = str1 & "  </td>" & Chr(13)
    str1 = str1 & "  <td width=70 style='width:52.2pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
    str1 = str1 & "  height:35.25pt'>" & Chr(13)
    str1 = str1 & "  <p class=MsoNormal><span lang=EN-US>" & steps & "</span></p>" & Chr(13)
    str1 = str1 & "  </td>" & Chr(13)
    str1 = str1 & "  <td width=298 style='width:223.5pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
    str1 = str1 & "  height:35.25pt'>" & Chr(13)
    str1 = str1 & "  <p class=MsoNormal><span lang=EN-US>" & title & "</span></p>" & Chr(13)
    str1 = str1 & "  </td>" & Chr(13)
    str1 = str1 & "  <td width=531 style='width:397.9pt;border:solid #99CCFF .5pt;padding:.75pt .75pt .75pt .75pt;" & Chr(13)
    str1 = str1 & "  height:35.25pt'>" & Chr(13)
    str1 = str1 & "  <p class=MsoNormal><span lang=EN-US>" & contents & "</span></p>" & Chr(13)
    str1 = str1 & "  </td>" & Chr(13)
    str1 = str1 & " </tr>" & Chr(13)
'   str1=str1&"</table>"
'   str1=str1&"</div>"
'   str1=str1&"<p class=MsoNormal><span lang=EN-US><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>"
'   str1=str1&"</div>"
'   str1=str1&"</body>"
'   str1=str1&"</html>"
    fo.Write (str1)
    fo.Close
End Function




Function generateRandomString()
    ' The single character 's  Asc Code  Must be  less then 110 ,that  n
   Randomize
   figure = Round(Rnd * 9, 0) + 5
   Randomize
   If figure = 1 Then
        generateRandomString = Round(Rnd * 9, 0)
    ElseIf figure > 1 Then
        letter = "0|1|2|3|4|5|6|7|8|9|A|a|B|b|C|c|D|d|E|e|F|f|G|g|H|h|I|i|J|j|K|k|L|l|M|m|N|n|O|P|Q|R|S|T|U|V|W|X|Y|Z"
        letterArray = Split(letter, "|")
        For i = 1 To figure
            indexLetter = Round(Rnd * 49, 0)
            generateRandomString = generateRandomString & letterArray(indexLetter)
        Next
    Else
        generateRandomString = ""
   End If
End Function




Function Encoder(ByVal str)
    encodeList = generateRandomString
    str = StrReverse(str)
    Encoder = ""
    j = 1
    For i = 1 To Len(str)
        Char = Mid(str, i, 1)
        If Asc(Char) > 99 Then
            c1 = Mid(CStr(Asc(Char)), 1, 2)
            If j > Len(encodeList) Then
                j = 1
            End If
            c1Transfer = Chr(Asc(Mid(encodeList, j, 1)) + CInt(c1))
            j = j + 1
            c2 = Mid(CStr(Asc(Char)), 3, 1)
            Encoder = Encoder & c1Transfer & c2
        Else
            Encoder = Encoder & CStr(Asc(Char))
        End If
    Next
End Function



Function Decoder(ByVal str)
    encodeList = Split(str, " ")(1)
    str = Split(str, " ")(0)
    Decoder = ""
    j = 1
    For i = 1 To Len(str) / 2
        decodeChar = Mid(str, i * 2 - 1, 2)
        ascDecodeChar = Asc(Left(decodeChar, 1))
        If ascDecodeChar > 47 And ascDecodeChar < 58 Then
            Decoder = Decoder & Chr(CInt(decodeChar))
        Else
            If j > Len(encodeList) Then
                j = 1
            End If
            encodeChar = Mid(encodeList, j, 1)
            decodeChar2 = Right(decodeChar, 1)
            numDecodeChar1 = ascDecodeChar - Asc(encodeChar)
            ascDecodeChar = numDecodeChar1 * 10 + CInt(decodeChar2)
            Decoder = Decoder & Chr(ascDecodeChar)
            j = j + 1
        End If
    Next
    Decoder = StrReverse(Decoder)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''Read  Data   Start
Function InitialTestData(ByVal scriptName, ByVal version, ByRef strCommons, ByRef arrParameters)
    InitialTestData = True
    filename = "DataTable.xls"
    If StrComp(Right(scriptName, 4), ".vbs", 1) = 0 Then
        scriptName = Left(scriptName, Len(scriptName) - 4)
    End If
    fileFolder = GetSourceFolder
    Dim xlsApp, xlsBook, xlsSheet, rows1, indxrow
   If Right(fileFolder, 1) <> "\" Then
        fileFolder = fileFolder & "\"
    End If
    ExeclFilePath = fileFolder & filename
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.application.Visible = False
    Set xlsBook = xlsApp.WorkBooks.Open(ExeclFilePath)
    
    'Get Common Parameters from Excel to {CommonData}
    Set xlsSheet = xlsBook.Worksheets(1)
    strCommons = ""
    rows1 = xlsSheet.UsedRange.rows.Count
    For indxrow = 2 To rows1
        strCommons = strCommons & Trim(xlsSheet.Cells(indxrow, 1).Value) & "|" & Trim(xlsSheet.Cells(indxrow, 2).Value) & "|"
    Next
    strCommons = Left(strCommons, Len(strCommons) - 1)
    
    'Get  Parameters  with relevant  sheet  from Excel to {Parameters}
    Dim scriptIndex, iterations, i, countParaPerIter, ChildParameters
    Set xlsSheet = xlsBook.Worksheets(version)
    scriptIndex = GetXlsIndexByScriptName(xlsSheet, scriptName)
    If scriptIndex <= xlsSheet.UsedRange.rows.Count Then
        iterations = GetScriptIteration(xlsSheet, scriptIndex, scriptName)
        ReDim arrParameters(iterations - 1)
        For i = 0 To iterations - 1
            Dim str
            str = ""
            j = 0
            Do While Trim(xlsSheet.Cells(scriptIndex + i, 3 + j * 2).Value) <> ""
                str = str & Trim(xlsSheet.Cells(scriptIndex + i, 3 + j * 2).Value) & "|" & Trim(xlsSheet.Cells(scriptIndex + i, 4 + j * 2).Value) & "|"
                j = j + 1
            Loop
            str = Left(str, Len(str) - 1)
            arrParameters(i) = str
        Next
    Else
        InitialTestData = False
        xlsBook.Close (True)
        xlsApp.quit
    End If

    xlsBook.Close (True)
    xlsApp.quit
End Function



Function GetXlsIndexByScriptName(ByVal xlsSheet, ByVal scriptName)
   Dim rowCount, irow
    rowCount = xlsSheet.UsedRange.rows.Count
    For irow = 2 To rowCount
        If StrComp(Trim(xlsSheet.Cells(irow, 1).Value), Trim(scriptName), 1) = 0 Then
            GetXlsIndexByScriptName = irow
            Exit Function
        End If
    Next
End Function


Function GetScriptIteration(ByVal xlsSheet, ByVal scriptRowIndx, ByVal scriptName)
   Dim Iteration, temp
    Iteration = 0
    temp = scriptRowIndx
    Do While xlsSheet.Cells(temp, 2).Value <> ""
        If StrComp(Trim(xlsSheet.Cells(temp, 1).Value), Trim(scriptName)) = 0 Or xlsSheet.Cells(temp, 1).Value = "" Then
            temp = temp + 1
            Iteration = Iteration + 1
        Else
            Exit Do
        End If
    Loop
    GetScriptIteration = Iteration
End Function



Function GetOperationData(ByVal Parameters, ByVal paraName)
    arrParameters = Split(Parameters, "|")
    For i = 0 To (UBound(arrParameters) - 1) / 2
        If StrComp(arrParameters(i * 2), paraName, 1) = 0 Then
            GetOperationData = arrParameters(i * 2 + 1)
            Exit Function
        End If
    Next
    GetOperationData = ""
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''Read  Data   End

'   comment  rule :  The comment  not allow to  write behind the statement at the same line,The comment can use both ' and rem, when use rem,it must add a blank  character behind rem
Sub RunComponent(ByVal componentName, ByVal varReportFile, ByVal reportFileName, ByVal variableStrConfig, ByVal strConfig, ByVal variableStrParameter, ByVal strParameter, ByVal strOutputVariable, ByRef strOutputValues)
    folderName = "Component"
    sourceFolder = GetSourceFolder
    If Right(sourceFolder, 1) <> "\" Then
        sourceFolder = sourceFolder & "\"
    End If
    scriptfolder = sourceFolder & folderName & "\"
    If StrComp(Right(componentName, 4), ".vbs", 1) <> 0 Then
        componentName = componentName & ".vbs"
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(scriptfolder & componentName, 1, True, -2)
    Set fso = Nothing
    code = ""
    Do While Not f.atEndOfStream
        lineContent = f.readline
        lineContent1 = Trim(lineContent)
        For i = 0 To Len(lineContent1) - 1
            If Left(lineContent1, i + 1) <> Chr(9) Then
                lineContent1 = Mid(lineContent1, i + 1, Len(lineContent1) - i)
                Exit For
            End If
        Next
        pos1 = InStr(lineContent1, Chr(39))
        pos2 = InStr(lineContent1, "rem ")
        If (Not (pos1 = 1)) And (Not (pos2 = 1)) Then
            code = code & lineContent & " :"
        End If
    Loop
    f.Close
    Set f = Nothing
    code = Replace(code, "ExitComponent", "Exit For", 1, -1, 1)
    If strConfig <> "" Then
        assignment = ":" & variableStrConfig & "=" & Chr(34) & strConfig & Chr(34) & ":"
    End If
    If strParameter <> "" Then
        assignment = assignment & ":" & variableStrParameter & "=" & Chr(34) & strParameter & Chr(34) & ":"
    End If
    If reportFileName <> "" Then
        assignment = assignment & ":" & varReportFile & "=" & Chr(34) & reportFileName & Chr(34) & ":"
    End If
    code = assignment & code
    code = "For r_u_n_t_i_m_e_s=0 to 0 :" & code
    Set oArrayList = CreateObject("System.Collections.ArrayList")
    oArrayList.Add InStr(1, code, ":function ", 1)
    oArrayList.Add InStr(1, code, Chr(9) & "function ", 1)
    oArrayList.Add InStr(1, code, " function ", 1)
    oArrayList.Sort
    If oArrayList(0) > 0 Then
        pos1 = oArrayList(0)
    ElseIf oArrayList(1) > 0 Then
        pos1 = oArrayList(1)
    Else
        pos1 = oArrayList(2)
    End If
    Set oArrayList = Nothing
    Set oAL = CreateObject("System.Collections.ArrayList")
    oAL.Add InStr(1, code, ":sub ", 1)
    oAL.Add InStr(1, code, Chr(9) & "sub ", 1)
    oAL.Add InStr(1, code, " sub ", 1)
    oAL.Sort
    If oAL(0) > 0 Then
        pos2 = oAL(0)
    ElseIf oAL(1) > 0 Then
        pos2 = oAL(1)
    Else
        pos2 = oAL(2)
    End If
    Set oAL = Nothing
    If pos1 = 0 And pos2 = 0 Then
        code = code & ": Next"
    ElseIf pos1 <> 0 And pos2 = 0 Then
        codeWithoutFunction = Left(code, pos1 - 1)
        codeFunction = Right(code, Len(code) - pos1 + 1)
        code = codeWithoutFunction & ": Next :" & codeFunction
    ElseIf pos1 <> 0 And pos2 <> 0 Then
        If pos1 > pos2 Then
            pos = pos2
        Else
            pos = pos1
        End If
        codeWithoutFS = Left(code, pos - 1)
        codeFS = Right(code, Len(code) - pos + 1)
        code = codeWithoutFS & ": Next :" & codeFS
    ElseIf pos1 = 0 And pos2 <> 0 Then
        codeWithoutSub = Left(code, pos2 - 1)
        codeSub = Right(code, Len(code) - pos2 + 1)
        code = codeWithoutSub & ": Next :" & codeSub
    End If
    code = Replace(code, "elseif ", "else if ", 1, -1, 1)
    'When the code contain  "else if",it must add "end if" before original "end if"
    elseifPos = InStr(code, "else if")
    Do While elseifPos > 0
        posEndifAfterElseif = InStr(elseifPos + 7, code, "end if", 1)
        endifBeforeCode = Left(code, posEndifAfterElseif - 1)
        endifAfterCode = Right(code, Len(code) - posEndifAfterElseif + 1)
        code = endifBeforeCode & ":end if:" & endifAfterCode
        elseifPos = InStr(elseifPos + 7, code, "else if", 1)
    Loop
    Set sco = CreateObject("ScriptControl")
    sco.Language = "vbscript"
    sco.ExecuteStatement code
    strOutputValues = ""
    If strOutputVariable <> "" Then
        For j = 0 To UBound(Split(strOutputVariable, "|"))
            strOutputValues = strOutputValues & sco.Eval(Split(strOutputVariable, "|")(j)) & "|"
        Next
        strOutputValues = Left(strOutputValues, Len(strOutputValues) - 1)
    End If
    Set sco = Nothing
End Sub




Function CloseProcessByName(ByVal processName)
    If StrComp(Right(processName, 4), ".exe", 1) <> 0 Then
        CloseProcessByName = CloseProcessByName & ".exe"
    End If
    On Error Resume Next
    For Each ps In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
        If StrComp(ps.Name, processName, 1) = 0 Then
            ps.Terminate
        End If
    Next
End Function


Function ExitTest()
    CloseProcessByName "wscript.exe"
End Function


Function IsRegistered(ByVal dllProgid)
    On Error Resume Next
    Set obj = CreateObject(dllProgid)
    If Err.Number = 0 Then
       IsRegistered = True
    Else
        IsRegistered = False
    End If
End Function




Function ClosePopupWindowByHandle(ByVal dWrap, ByVal wHnd)
    dWrap.SendMessage CLng(wHnd), &H10, 0, 0
End Function




Function IsWindowExist(ByVal dWrap, ByVal caption, ByRef winHnd)
    winHnd = dWrap.FindWindow(vbNullString, CStr(caption))
    If winHnd <> 0 Then
        IsWindowExist = True
    Else
        IsWindowExist = False
    End If
End Function




Function ClickPopupWinBtn(ByVal dWrap, ByVal parentWinHnd, ByVal btnClass)
    Const WM_LBUTTONDOWN = &H201
    Const WM_LBUTTONUP = &H202
    Const WM_ACTIVATE = &H6
    hBtn = dWrap.FindWindowEx(CLng(parentWinHnd), 0, CStr(btnClass), 0)
    dWrap.SendMessage hBtn, WM_ACTIVATE, 1, 0
    dWrap.SendMessage hBtn, WM_LBUTTONDOWN, 1, 0
    dWrap.SendMessage hBtn, WM_LBUTTONUP, 1, 0
End Function


'   comment  rule :  The comment  not allow to  write behind the statement at the same line,The comment can use both ' and rem, when use rem,it must add a blank  character behind rem
'   called  function or  output parameter  rule:  The type of  function or  output parameter  can only be a string or numeric,should not  be an object  type

Function ReturnFunctionOrSub(ByVal functionFile, ByVal calledStatement, ByRef tempOutputFile)
    folderName = "FunctionLibrary"
    sourceFolder = GetSourceFolder
    If Right(sourceFolder, 1) <> "\" Then
        sourceFolder = sourceFolder & "\"
    End If
    If StrComp(Right(functionFile, 4), ".vbs", 1) <> 0 Then
        functionFile = functionFile & ".vbs"
    End If
    libFilePath = sourceFolder & folderName & "\" & functionFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(libFilePath, 1, True, -2)
    If StrComp(Left(calledStatement, 5), "call ", 1) = 0 Then
        calledStatement = Trim(Right(calledStatement, Len(calledStatement) - 5))
    End If
    If InStr(calledStatement, "(") > 0 Then
        functionName = Trim(Left(calledStatement, InStr(calledStatement, "(")))
    ElseIf InStr(calledStatement, " ") > 0 Then
        functionName = Trim(Left(calledStatement, InStr(calledStatement, " ")))
    Else
        functionName = calledStatement
    End If
    tempFileName = CreateObject("Wscript.Shell").CurrentDirectory & "\" & functionName
    randomstr = CStr(Rnd * 9)
    tempOutputFile = tempFileName & randomstr & ".txt"
    code = ""
    lineno = 1
    Do While Not f.atEndOfStream
        content = f.readline
        contentTemp = Trim(content)
        contentTemp = Replace(contentTemp, Chr(9), "")
        pos1 = InStr(contentTemp, Chr(39))
        pos2 = InStr(contentTemp, "rem ")
        If (Not (pos1 = 1)) And (Not (pos2 = 1)) Then
            code = code & content & ":"
            If InStr(5, content, functionName, 1) > 0 And lineno = 1 Then
                functionDefinition = content
                lineno = lineno + 1
            End If
        End If
    Loop
    calledStatement = Trim(calledStatement)
    If InStr(1, calledStatement, "call ", 1) > 0 Then
        evalStatement = Replace(calledStatement, "call ", "", 1, -1, 1)
    ElseIf InStr(1, calledStatement, " ", 1) > 1 And InStr(1, calledStatement, "(", 1) <= 0 Then
        posBlank = InStr(1, calledStatement, " ", 1)
        evalStatement = Left(calledStatement, posBlank - 1) & "(" & Right(calledStatement, Len(calledStatement) - posBlank) & ")"
    Else
        evalStatement = calledStatement
    End If
    code = code & ":  full_na_me=" & Chr(34) & tempOutputFile & Chr(34) & ":"
    code = code & ":  lin_e_1=" & Chr(34) & functionDefinition & Chr(34) & ":"
    code = code & "Set fso =createobject(""Scripting.FileSystemObject""):Set fw=fso.OpenTextFile(full_na_me,2,True):posstart=instr(1,lin_e_1,""function "",1):If posstart>0 Then:   fw.writeline eval(b_e_E_v_a_l):Else:    fw.writeline """":End If:For k=0 To UBound(split(lin_e_1,chr(44))): temp=split(lin_e_1,chr(44))(k):             If instr(1,temp,""byref "",1)>0 Then:                   sPara=Replace(temp,left(temp,instr(1,temp,""byref"",1)+5),""""):                    If InStr(sPara,chr(41))>0 Then:                         sPara=Trim(replace(sPara,chr(41),"""")):                End If:                 fw.writeline eval(sPara):  End if:Next:fw.close:"
    code = Replace(code, "b_e_E_v_a_l", evalStatement)
    f.Close
    Set f = Nothing
    Set fso = Nothing
    code = Replace(code, "elseif ", "else if ", 1, -1, 1)
    'When the code contain  "else if",it must add "end if" before original "end if"
    elseifPos = InStr(code, "else if")
    Do While elseifPos > 0
        posEndifAfterElseif = InStr(elseifPos + 7, code, "end if", 1)
        endifBeforeCode = Left(code, posEndifAfterElseif - 1)
        endifAfterCode = Right(code, Len(code) - posEndifAfterElseif + 1)
        code = endifBeforeCode & ":end if:" & endifAfterCode
        elseifPos = InStr(elseifPos + 7, code, "else if", 1)
    Loop
    code = calledStatement & ":" & code & ":"
    Set sco = CreateObject("ScriptControl")
    sco.Language = "vbscript"
    sco.ExecuteStatement code
End Function


Function InvokeFunctionOrSub(ByVal functionFile, ByVal calledStatement, ByRef strRefPara)
    Call ReturnFunctionOrSub(functionFile, calledStatement, tempOutput)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(tempOutput, 1, True, -2)
    rowCount = 0
    strRefPara = ""
    Do While Not f.atEndOfStream
        If rowCount = 0 Then
            InvokeFunctionOrSub = f.readline
        Else
            strRefPara = strRefPara & f.readline & "|"
        End If
        rowCount = rowCount + 1
    Loop
    If strRefPara <> "" Then
        strRefPara = Left(strRefPara, Len(strRefPara) - 1)
    End If
    f.Close
    fso.DeleteFile tempOutput
End Function



Function GetRefPara(ByVal functionDefinition)
    paraArr = Split(functionDefinition, ",")
    paracount = UBound(paraArr) + 1
    GetRefPara = ""
    For k = 0 To paracount - 1
        position = InStr(1, paraArr(k), "byref ", 1)
        If position > 0 Then
            refPara = Trim(Right(paraArr(k), Len(paraArr(k)) - position - 5))
            If InStr(refPara, ")") > 0 Then
                refPara = Trim(Replace(refPara, ")", ""))
            End If
            GetRefPara = GetRefPara & refPara & Chr(9)
        End If
    Next
    If GetRefPara <> "" Then
        GetRefPara = Left(GetRefPara, Len(GetRefPara) - 1)
    End If
End Function


'sql="select count(*) from eccreg2.cat_info t where ct_type_id =2 and cust_key ='Y10090162579-Y10090162579' and cat_catalog_name like 'TAT 2%'"
Function ReturnArrayViaDBQuery(ByVal serviceName, ByVal uid, ByVal pwd, ByVal sql, ByRef resultArr)
    Dim conn, rs, ConnectionString, sqlSection, recordSum, columns
    ReturnArrayViaDBQuery = False
    Set conn = CreateObject("adodb.connection")
    Set rs = CreateObject("adodb.recordset")
    ConnectionString = "Driver={Microsoft ODBC for Oracle};Server=" & serviceName & ";Uid=" & uid & ";Pwd=" & pwd
    conn.Open ConnectionString
    Set rs = conn.execute(sql)
    If Not rs.BOF Then
        resultArr = rs.GetRows
        ReturnArrayViaDBQuery = True
    Else
        resultArr = ""
        ReturnArrayViaDBQuery = False
    End If
    rs.Close
    conn.Close
End Function



Function TestResult(ByVal filename)
    folderName = "Report"
   TestResult = True
   sourceFolder = GetSourceFolder
    If Right(sourceFolder, 1) <> "\" Then
        sourceFolder = sourceFolder & "\"
    End If
    fileFolder = sourceFolder & folderName & "\"
    If StrComp(Right(filename, 4), ".htm", 1) <> 0 Then
        filename = filename & ".htm"
    End If
    filepath = fileFolder & filename
    Set ieo = CreateObject("InternetExplorer.Application")
    ieo.navigate filepath
    Set docReport = ieo.document
    Set objTr = docReport.getElementsByTagName("TR")
    For cnt = 1 To objTr.length - 1
        Set objTd = objTr(cnt).ChildNodes.Item(0)
        Do While objTd.ChildNodes.Item(0).nodename <> "#text"
            Set objTd = objTd.ChildNodes.Item(0)
        Loop
        If StrComp(objTd.ChildNodes.Item(0).nodevalue, "Fail", 1) = 0 Then
            TestResult = False
            Exit For
        End If
    Next
End Function


Sub OutputTheResult(ByVal resultFileName)
   scriptName = Split(resultFileName, "_")(1) & ".vbs"
    tempOutputFile = CreateObject("Wscript.Shell").CurrentDirectory & "\" & scriptName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fw = fso.OpenTextFile(tempOutputFile & ".txt", 2, True)
    Status = TestResult(resultFileName)
    recentErrorMsg = " "
    If Not Status Then
        result = "Fail"
        If Err.Number <> 0 Then
           recentErrorMsg = Err.Description
        End If
        Err.Clear
    Else
        result = "Pass"
    End If
    fw.writeline result
    timeStr = Date & " " & Time
    fw.writeline timeStr
    fw.writeline recentErrorMsg
    fw.Close
End Sub


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' Input     ->excelFilePath: the Excel which to be readed.
' Output    ->sCommons :    string list with all generic parameters saved in sheet1 of excel
'           ->sParameters : string list with each running test script's row no. in excel and test script name which split as '|' character
Function ReadParameters(ByRef sCommons, ByRef scriptfolder, ByRef sParameters)
    excelName = "RunManager.xls"
    Set objShell = CreateObject("Wscript.Shell")
    vbsFolder = objShell.CurrentDirectory
    If Right(vbsFolder, 1) <> "\" Then
        vbsFolder = vbsFolder & "\"
    End If
    excelFilePath = vbsFolder & excelName
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.application.Visible = False
    Set xlsBook = xlsApp.WorkBooks.Open(excelFilePath)
    'Get Configuration Parameters from Excel's Sheet1  to sCommons
    Set xlsSheet = xlsBook.Worksheets(1)
    rows1 = xlsSheet.UsedRange.rows.Count
    sCommons = ""
    For indxrow = 2 To rows1
        sCommons = sCommons & Trim(xlsSheet.Cells(indxrow, 1).Value) & "|" & Trim(xlsSheet.Cells(indxrow, 2).Value) & "|"
    Next
    sCommons = Left(sCommons, Len(sCommons) - 1)
    
    Set oNetWork = CreateObject("Wscript.Network")
    localMachine = oNetWork.ComputerName
    scriptfolder = GetOperationData(sCommons, "RunScriptFolder")
    runset = GetOperationData(sCommons, "RunningSet")
    
    'Get running test info from Excel's Sheet2 to sParameters
    Set xlsSheet = xlsBook.Worksheets(2)
    For indxrow2 = 2 To xlsSheet.UsedRange.rows.Count
        If StrComp(Trim(xlsSheet.Cells(indxrow2, 1).Value), runset, 1) = 0 Then
            Exit For
        End If
    Next
    idx = indxrow2
    sParameters = ""
    Do While xlsSheet.Cells(idx, 1).Value = "" Or xlsSheet.Cells(idx, 1).Value = xlsSheet.Cells(indxrow2, 1).Value
        If xlsSheet.Cells(idx, 2).Value <> "" Then
            If StrComp(Trim(xlsSheet.Cells(idx, 3).Value), "Y", 1) = 0 Or StrComp(Trim(xlsSheet.Cells(idx, 3).Value), "yes", 1) = 0 Then
                runmachinePara = Trim(xlsSheet.Cells(idx, 4).Value)
                runmachine = GetOperationData(sCommons, runmachinePara)
                If StrComp(localMachine, runmachine, 1) = 0 Or InStr(1, runmachinePara, "local", 1) > 0 Then
                    sParameters = sParameters & CStr(idx) & "|" & Trim(xlsSheet.Cells(idx, 2).Value) & "|"
                End If
            End If
        End If
        idx = idx + 1
        If idx = 65536 Then
            Exit Do
        End If
    Loop
    sParameters = Left(sParameters, Len(sParameters) - 1)
    xlsBook.Close (True)
    xlsApp.quit
End Function



Function WriteInfo2Excel(ByVal scriptRowNo, ByVal scriptPath)
   excelName = "RunManager.xls"
    If StrComp(Right(scriptPath, 4), ".vbs") <> 0 Then
        scriptPath = scriptPath & ".vbs"
    End If
    scriptPath = scriptPath & ".txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fr = fso.OpenTextFile(scriptPath, 1, True, -2)
    sResult = fr.readline
    sTime = fr.readline
    sErrorMsg = fr.readline
    fr.Close
    fso.DeleteFile scriptPath
    Set fr = Nothing
    Set fso = Nothing
    
    Set objShell = CreateObject("Wscript.Shell")
    vbsFolder = objShell.CurrentDirectory
    If Right(vbsFolder, 1) <> "\" Then
        vbsFolder = vbsFolder & "\"
    End If
    excelFilePath = vbsFolder & excelName & ".xls"
    Set xlApp = CreateObject("Excel.Application")
    xlApp.application.Visible = False
    Set xlBook = xlApp.WorkBooks.Open(excelFilePath)
    Set xlSheet = xlBook.Worksheets(2)
    colorFlag = xlSheet.Cells(1, 8).Interior.Color
    Do While colorFlag = RGB(255, 0, 0)
        colorFlag = xlSheet.Cells(1, 8).Interior.Color
    Loop
    ''set to red
    xlSheet.Cells(1, 8).Interior.Color = RGB(255, 0, 0)
    xlSheet.Cells(scriptRowNo, 5).Value = sResult
    xlSheet.Cells(scriptRowNo, 6).Value = sTime
    xlSheet.Cells(scriptRowNo, 7).Value = sErrorMsg
    ''set to green
    xlSheet.Cells(1, 8).Interior.Color = RGB(0, 255, 0)
    xlBook.Close (True)
    xlApp.quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Function
