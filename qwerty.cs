
'---  This VBScript will retrieve the necessary document data and  
'---  Create a file for RGH with Wound Care Cancellation information
'---  Update the document with the filename
'
'---  Force declaration of variables
Option Explicit

'Constants for File System Object
Const ForReading = 1 
Const ForWriting = 2 
Const ForAppending = 8 

Dim idxCounterKW, KeyLoopEndValue, ReplaceKWLoopCount, ReplaceKWEndValue
Dim sAccount, sWoundCancelledBy, sWoundCareToRGH      
Dim foundAccount,  foundWoundCancelledBy,foundWoundCareToRGH
Dim idxWoundCareToRGH

Dim objFSO, objFldr, objFil, objFolder
'Dim strPatientInfo, strPatientFile, 
Dim sSplit, sTimeStamp, oFileName, oFilePath, oFilePathName, oPatientInfo, oWoundCareFile
Dim sWoundCareMMDDYYYY, sMMDDYYYY

Dim sTodaysDate, sTodaysDateM, sTodaysDateYY, sTodaysDateD
Dim sTodaysDateHH, sTodaysDateMM, sTodaysDateSS

foundAccount = "N"
foundWoundCancelledBy = "N" 
foundWoundCareToRGH = "N"

'Thick Client Entry Point
Sub Main35()

Dim objApp, currDoc, docKeys
Set objApp = CreateObject("Onbase.Application")
Set currDoc = objApp.CurrentDocument
Set docKeys = currDoc.Keywords

'Find the keywords on the doc
For idxCounterKW = 0 to docKeys.Count - 1
  
   'check Account
   If docKeys.item(idxCounterKW).name = "Account # A" Then
      sAccount = RTrim(docKeys.item(idxCounterKW).value)
     foundAccount = "Y"
   End if
   
   If docKeys.item(idxCounterKW).name = "MD - RGH File" Then
      sWoundCareToRGH = RTrim(docKeys.item(idxCounterKW).value)
      idxWoundCareToRGH =idxCounterKW
      foundWoundCare = "Y"
   End if
    
   If docKeys.item(idxCounterKW).name = "Cancelled By" Then
      sWoundCancelledBy = docKeys.item(idxCounterKW).value
      foundWoundCancelledBy = "Y"
   End If
'msgbox "Loop # " & idxCounterKW & " " & docKeys.item(idxCounterKW).name & sWoundCancelledBy
   
Next

If (foundAccount = "Y" and foundWoundCancelledBy = "Y") then
' build the filename and the record information
  
   sTodaysDate =now()
   sTodaysDateM = (DatePart("m",sTodaysDate))
   sTodaysDateYY = (DatePart("yyyy",sTodaysDate))
   sTodaysDateD = (DatePart("d",sTodaysDate))

   if (len(sTodaysDateM)) = 1 then sTodaysDateM = "0" & sTodaysDateM end if
   if (len(sTodaysDateD)) = 1 then sTodaysDateD = "0" & sTodaysDateD end if

'Wscript.Echo "todaysDate " & sTodaysDateM & sTodaysDateD & sTodaysDateYY
   sTodaysDateHH = Hour(Now())
   sTodaysDateMM= Minute(Now())
   sTodaysDateSS= Second(Now())

   if (len(sTodaysDateHH)) = 1 then sTodaysDateHH = "0" & sTodaysDateHH end if
   if (len(sTodaysDateMM)) = 1 then sTodaysDateMM = "0" & sTodaysDateMM end if
   if (len(sTodaysDateSS)) = 1 then sTodaysDateSS = "0" & sTodaysDateSS end if

'Wscript.Echo "todaysHHMMSS " & sTodaysDateHH & sTodaysDateMM & sTodaysDateSS
   sTimeStamp = sTodaysDateM & sTodaysDateD & sTodaysDateYY & sTodaysDateHH & sTodaysDateMM & sTodaysDateSS

   oFilePath = "\\CORP.RGHENT.COM\ONB\PRD_SRC\BKOEXP\WoundCare\"
   oFileName = "WoundCare" & "_" & sAccount & "_" & sTimeStamp & ".txt"
   oFilePathName = oFilePath & oFileName
   oPatientInfo = sAccount & "|N|" & sWoundCancelledBy
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   Set oWoundCareFile = objFSO.OpenTextFile(oFilePathName,ForAppending,True)	
   Call oWoundCareFile.WriteLine(oPatientInfo)
   oWoundCareFile.Close
  
   'store the RGH filename
   If foundWoundCareToRGH = "Y" then
     docKeys.item(idxWoundCareToRGH).value = oFilePathName
 '    msgbox foundWoundCareToRGH & "=" & oFilePathName
   else    
     Call docKeys.AddKeyword("MD - MWC Cancelled RGH File", oFilePathName)
 '    msgbox foundWoundCareToRGH & "=" & oFilePathName
   End if
   Call currDoc.StoreKeywords()
   
   'destroy objects
   Set docKeys = Nothing 
   Set currDoc = Nothing
   Set objApp = Nothing
   Set objFSO = Nothing
End if
End Sub               'Main35()

'
'Set oWoundCareFile = Nothing
