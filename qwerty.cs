'---  Force declaration of variables
'Option Explicit

Dim myApp, myDocument, myPages, myKeys

Set myApp = CreateObject("OnBase.Application")
Set myDocument = myApp.CurrentDocument
'Set docKeywords = myDocument.Keywords
Set myKeys = myDocument.Keywords
Set myPages = myDocument.Pages

Dim FirstPage, FilePath
Dim FSO, ResponseFile
Dim TheLine, KeyLoopCounter, linect

Dim strConnectTm,  strExplanation, strFaxLine, strFaxServer, strFailSuccess
Dim strJID, strPages, strRecipient, strResult, strRetryCt
Dim strStatusCd, strSubject, strTransmitTm
Dim strSubjectRGHID, strSubjectRGHType, strUID

Dim idxConnectTm, idxExplanation, idxFaxLine, idxFaxServer, idxFailSuccess
Dim idxJID, idxPages, idxRecipient, idxXmitResult, idxRetryCt
Dim idxStatusCd, idxSubject, idxTransmitTm
Dim idxSubjectRGHID, idxSubjectRGHType, idxUID

Dim lenKW, strCell, strSplitCell  
      
For KeyLoopCounter = 1 to myKeys.Count - 1
    'msgbox myKeys.item(KeyLoopCounter).name
    If myKeys.item(KeyLoopCounter).name = "XMIT - Connect Time" Then
       idxConnectTm = KeyLoopCounter
    End If
    If myKeys.item(KeyLoopCounter).name = "XMIT - Explanation" Then
       idxExplanation = KeyLoopCounter
    End If
    If myKeys.item(KeyLoopCounter).name = "XMIT - Fax Line" Then
       idxFaxLine = KeyLoopCounter
    End If
    If myKeys.item(KeyLoopCounter).name = "XMIT - Fax Server" Then
       idxFaxServer = KeyLoopCounter
    End If
    If myKeys.item(KeyLoopCounter).name = "XMIT - Job ID" Then
       idxJID = KeyLoopCounter
   End If 
   If myKeys.item(KeyLoopCounter).name = "XMIT - Pages" Then
       idxPages = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Recipient" Then
       idxRecipient = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Result" Then
       idxXmitResult = KeyLoopCounter
'msgbox "idxXmitResult" & idxXmitResult
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Retry Count" Then
       idxRetryCt = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Status Code" Then
       idxStatusCd = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Subject" Then
       idxSubject = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Transmit Time" Then
       idxTransmitTm = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - Unique ID" Then
       idxUID = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - RGH ID" Then
       idxSubjectRGHID = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - RGH Type" Then
       idxSubjectRGHType = KeyLoopCounter
   End if
   If myKeys.item(KeyLoopCounter).name = "XMIT - FailSuccess" Then
       idxFailSuccess = KeyLoopCounter
   End if
Next

'Look through the email
Set FirstPage = myPages.Item(0)
FilePath = FirstPage.SubPath()
FilePath = "\\Cardinalhealth.net\applications\EC500\ONB\PRD_SYS\Diskgroups\EP" & FilePath 
'msgbox "FilePath=" & FilePath & " strLN=" & strLN
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ResponseFile= FSO.OpenTextFile(FilePath, 1) 

linect = 0
Do Until ResponseFile.AtEndOfStream
   TheLine = ResponseFile.ReadLine
'msgbox "TheLine=" & TheLine 
   linect = linect + 1   
   If InStr(Ucase(TheLine), "RECIPIENT AT") Then
      strSplitCell = Split(TheLine, "at ")  
      strCell = trim(strSplitCell(1))
      strRecipient = strCell
     ' msgbox "strRecipient=" & strRecipient
   End if 
   If InStr(Ucase(TheLine), "SUBJECT:") Then
      If Instr(Ucase(TheLine),"FAX:") Then
      Else
         strSplitCell = Split(TheLine, ":")  
         strCell = trim(strSplitCell(1))
         strCell = Replace(strCell, " ", "*")
'msgbox strCell
         strSplitCell = Split(strCell, "*") 
         strSubject = strCell
'msgbox strSubject
         If Instr(Ucase(strSubject),"**-M") Then  'this is an M
            strSubjectRGHType = trim(strSplitCell(2))
            strSubjectRGHType = Replace(strSubjectRGHType, "-", "")
            strSubjectRGHID = trim(strSplitCell(3))
         else 'R or P

            strSubjectRGHType = trim(strSplitCell(1))
'msgbox "strSubjectRGHType B4" &  strSubjectRGHType         
            strSubjectRGHType = Replace(strSubjectRGHType, "-", "")

            strSubjectRGHID = trim(strSplitCell(2))
         end if 
'msgbox "strSubjectRGHType After" &  strSubjectRGHType 
'msgbox "strSubjectRGHType" & strSubjectRGHID
         strSubject = Replace(strSubject, "*", " ")
         'strSubject = TheLine

      End if
   End if
   
   If InStr(Ucase(TheLine), "RESULT:") Then
      strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strResult = strCell
      If InStr(Ucase(TheLine), "SUCCESSFUL") Then
         strFailSuccess = "SUCCESS"
      else
         strFailSuccess = "FAIL"
      end if 
      'strResult = TheLine
   End if
   If InStr(Ucase(TheLine), "EXPLANATION:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strExplanation = strCell
      'strExplanation = TheLine
   End if
   If InStr(Ucase(TheLine), "PAGES SENT:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strPages = strCell
      'strPages = TheLine
   End if
   If InStr(Ucase(TheLine), "CONNECT TIME:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strConnectTm = strCell
      'strConnectTm = TheLine
   End if 
   If InStr(Ucase(TheLine), "TRANSMIT TIME:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1)) & ":" & trim(strSplitCell(2))
      strTransmitTm = strCell
'msgbox strTransmitTm
      'strTransmitTm = TheLine
   End if
   
   If InStr(Ucase(TheLine), "STATUS CODE:") Then
      strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strStatusCD = strCell
      'strStatusCd = TheLine
   End if
   
   If InStr(Ucase(TheLine), "RETRY COUNT:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strRetryCt = strCell
      'strRetryCt = TheLine
   End if
   If InStr(Ucase(TheLine), "JOB ID:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strJID = strCell
      'strJID = TheLine
   End if
   If InStr(Ucase(TheLine), "UNIQUE ID:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strUID = strCell
      'strUID = TheLine
   End if
   If InStr(Ucase(TheLine), "FAX LINE:") Then
       strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strFaxLine = strCell
      'strFaxLine = TheLine
   End if 
   If InStr(Ucase(TheLine), "FAX SERVER:") Then
      strSplitCell = Split(TheLine, ":")  
      strCell = trim(strSplitCell(1))
      strFaxServer = strCell
      'strFaxServer = TheLine
   End if
loop
'msgbox "strRecipient=" & strRecipient & " strExplanation=" & strExplanation
'msgbox "strSubject=" & strSubject

lenKW = Len(trim(strConnectTm))
if lenKW > 0 then
   If idxConnectTm > 0  then
      myKeys.item(idxConnectTm).Value = strConnectTm
   Else
      Call myKeys.AddKeyword("XMIT - CONNECT TIME", strConnectTm)
   End if
End If

lenKW = Len(trim(strExplanation))
if lenKW > 0 then
   If idxExplanation > 0  then
      myKeys.item(idxExplanation).Value = strExplanation
   Else
      Call myKeys.AddKeyword("XMIT - EXPLANATION", strExplanation)
   End if
End If

lenKW = Len(trim(strFaxLine))
if lenKW > 0 then
   If idxFaxLine > 0  then
      myKeys.item(idxFaxLine).Value = strFaxLine
   Else
      Call myKeys.AddKeyword("XMIT - FAX LINE", strFaxLine)
   End if
End If

lenKW = Len(trim(strFaxServer))
if lenKW > 0 then
   If idxFaxServer > 0  then
      myKeys.item(idxFaxServer).Value = strFaxServer
   Else
      Call myKeys.AddKeyword("XMIT - FAX SERVER", strFaxServer)
   End if
End If

lenKW = Len(trim(strJID))
if lenKW > 0 then
   If idxJID > 0  then
      myKeys.item(idxJID).Value = strJID
   Else
      Call myKeys.AddKeyword("XMIT - JOB ID", strJID)
   End if
End If

lenKW = Len(trim(strPages))
if lenKW > 0 then
   If idxPages > 0  then
      myKeys.item(idxPages).Value = strPages
   Else
      Call myKeys.AddKeyword("XMIT - PAGES", strPages)
   End if
End If

lenKW = Len(trim(strRecipient))
if lenKW > 0 then
   If idxRecipient > 0  then
      myKeys.item(idxRecipient).Value = strRecipient
   Else
      Call myKeys.AddKeyword("XMIT - RECIPIENT", strRecipient)
   End if
End If

lenKW = Len(trim(strResult))
if lenKW > 0 then
   If idxXmitResult > 0  then
      myKeys.item(idxXmitResult).Value = strResult
   Else
      Call myKeys.AddKeyword("XMIT - RESULT", strResult)
   End if
end if

lenKW = Len(trim(strRetryCt))
if lenKW > 0 then
   If idxRetryCt > 0  then
      myKeys.item(idxRetryCt).Value = strRetryCt
   Else
      Call myKeys.AddKeyword("XMIT - RETRY COUNT", strRetryCt)
   End if
End If

lenKW = Len(trim(strStatusCd))
if lenKW > 0 then
   If idxStatusCd > 0  then
      myKeys.item(idxStatusCd).Value = strStatusCd
   Else
      Call myKeys.AddKeyword("XMIT - STATUS CODE", strStatusCd)
   End if
End If

lenKW = Len(trim(strSubject))
if lenKW > 0 then
   If idxSubject > 0  then
      myKeys.item(idxSubject).Value = strSubject
   Else
      Call myKeys.AddKeyword("XMIT - SUBJECT", strSubject)
   End if
end if

lenKW = Len(trim(strTransmitTm))
if lenKW > 0 then
   If idxTransmitTm > 0  then
      myKeys.item(idxTransmitTm).Value = strTransmitTm
   Else
      Call myKeys.AddKeyword("XMIT - TRANSMIT TIME", strTransmitTm)
   End if
End If

lenKW = Len(trim(strUID))
if lenKW > 0 then
   If idxUID > 0  then
      myKeys.item(idxUID).Value = strUID
   Else
      Call myKeys.AddKeyword("XMIT - UNIQUE ID", strUID)
   End if
End If

lenKW = Len(trim(strSubjectRGHID))
if lenKW > 0 then
   If idxSubjectRGHID > 0  then
      myKeys.item(idxSubjectRGHID).Value = strSubjectRGHID
   Else
      Call myKeys.AddKeyword("XMIT - RGH ID", strSubjectRGHID)
   End if
End If

lenKW = Len(trim(strSubjectRGHType))
if lenKW > 0 then
   If idxSubjectRGHType > 0  then
      myKeys.item(idxSubjectRGHType).Value = strSubjectRGHType
   Else
      Call myKeys.AddKeyword("XMIT - RGH TYPE", strSubjectRGHType)
   End if
End If

lenKW = Len(trim(strFailSuccess))
if lenKW > 0 then
   If idxFailSuccess > 0  then
      myKeys.item(idxFailSuccess).Value = strFailSuccess
   Else
      Call myKeys.AddKeyword("XMIT - FAILSUCCESS", strFailSuccess)
   End if
end if

myDocument.StoreKeywords

'Added Destroy Objects
Set myKeys = Nothing
Set myPages = Nothing
'Set docKeywords = Nothing
Set myDocument = Nothing
Set myApp = Nothing
