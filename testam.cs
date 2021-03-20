'*************************************'
'* Written by Jeremy Giaco 4/25/2005 *'
'*************************************'

'This script generates a review letter from an OnBase 835 (EOB)

'Constants for File System Object
Const ForReading = 1 
Const ForWriting = 2 
Const ForAppending = 8 

'Starting text of the reason for request
Dim sReason
'InputBox("Please enter detailed reason for request:", _
'"Reason For Request", _
sReason = "The written/verbal order indicates a medical necessity of ..."

'OnBase Application object
Dim oApp
Set oApp = CreateObject("Onbase.Application")

'OnBase Document Object
Dim oDoc
Set oDoc = oApp.CurrentDocument

'OnBase Keyword Object
Dim oKeywords
Set oKeywords = oDoc.Keywords

'Get Document Handle
Dim sDocHandle
sDocHandle = oDoc.Handle

'Used to loop through document keywords
Dim iCount
iCount = 0

Dim sBeneficiary
Dim sDOS
Dim sCCN
Dim sHIC
Dim sDOD
Dim sHCPC
Dim sAcctNum
Dim sCarrier
Dim sDate
sDate = Date()

'Loop through keywords and set strings equal to keyword values
Do while iCount < oKeywords.Count
	Dim sKeyName
	Dim sKeyValue
	
	sKeyName = oKeywords(iCount).Name
	sKeyValue = oKeywords(iCount).Value
	
	Select Case sKeyName
		Case "AR - Carrier Name"
			sCarrier = sKeyvalue
		Case "AR - Account Name"
			sBeneficiary = sKeyValue
		Case "AR - Date of Service"
			sDOS = sKeyValue
		'Case "AR - Claim Number"
			'sCCN = sKeyValue
		Case "AR - CCN"
			sCCN = sKeyValue
		Case "AR - Patient ID"
			sHIC = sKeyValue
		Case "AR - Payment Date"
			sDOD = sKeyValue
		Case "AR - Procedure Code"
			sHCPC = sHCPC & sKeyValue & " "
		Case "AR - Account Number"
			sAcctNum = sKeyValue
	End Select 
	
	'Increment Counter	
	iCount = iCount + 1
Loop

'File System Object
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Choose correct template file based on carrier name
Dim sTemplateName
If InStr(1,UCase(sCarrier),"NATIONAL HERITAGE",1) > 0 Then
	sTemplateName = "reviewregiona.rtf"
ElseIf InStr(1,UCase(sCarrier),"NHIC",1) > 0 Then
	sTemplateName = "reviewregiona.rtf"
ElseIf InStr(1,UCase(sCarrier),"REGION A",1) > 0 Then
	sTemplateName = "reviewregiona.rtf"
ElseIf InStr(1,UCase(sCarrier),"ADMINASTAR",1) > 0 Then
	sTemplateName = "reviewregionb.rtf"
ElseIf InStr(1,UCase(sCarrier),"NATIONAL GOVERNMENT SERVICES",1) > 0 Then
	sTemplateName = "reviewregionb.rtf"
ElseIf InStr(1,UCase(sCarrier),"REGION B",1) > 0 Then
	sTemplateName = "reviewregionb.rtf"
ElseIf InStr(1,UCase(sCarrier),"NGS",1) > 0 Then
	sTemplateName = "reviewregionb.rtf"
ElseIf InStr(1,UCase(sCarrier),"NGS, INC. JURISDICTION B DME MAC",1) > 0 Then
	sTemplateName = "reviewregionb.rtf"
ElseIf InStr(1,UCase(sCarrier),"PALMETTO",1) > 0 Then
	sTemplateName = "reviewregionc.rtf" 
ElseIf InStr(1,UCase(sCarrier),"CIGNA                C",1) > 0 Then
	sTemplateName = "reviewregionc.rtf"
ElseIf InStr(1,UCase(sCarrier),"CIGNA GOVERNMENT SERVICES",1) > 0 Then
	sTemplateName = "reviewregionc.rtf"
ElseIf InStr(1,UCase(sCarrier),"REGION C",1) > 0 Then
	sTemplateName = "reviewregionc.rtf"
ElseIf InStr(1,UCase(sCarrier),"NAS",1) > 0 Then
        sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"NORIDIAN",1) > 0 Then
        sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"NORIDIAN ADMINISTRATIVE SERVICES",1) > 0 Then
        sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"REGION D",1) > 0 Then
        sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"CIGNA                D",1) > 0 Then
	sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"REGION D",1) > 0 Then
	sTemplateName = "reviewregiond.rtf"
ElseIf InStr(1,UCase(sCarrier),"CIGNA - MEDICARE",1) > 0 Then
	sTemplateName = "reviewregiond.rtf"
Else
	Dim sMessage
	sMessage = _
		"Cannot find template letter for this insurance. Please type the letter of the correct Medicare region for this claim:" & vbcrlf & _
		"A - REGION A" & vbcrlf & _
		"B - REGION B" & vbcrlf & _
		"C - REGION C" & vbcrlf & _
		"D - REGION D"
		
	Dim sInput
	sInput = InputBox(sMessage,"Cannot find template","A")
	Select Case sInput
		Case "A" 
			sTemplateName = "reviewregiona.rtf"
		Case "B" 
			sTemplateName = "reviewregionb.rtf"
		Case "C" 
			sTemplateName = "reviewregionc.rtf"
		Case "D" 
			sTemplateName = "reviewregiond.rtf"		
		Case Else
			sTemplateName = "reviewregiona.rtf"	
	End Select				
End If

'Open template
Dim oTemplate	
Set oTemplate = oFSO.OpenTextFile("\\3SG.net\applications\EC500\ONB\PRD_SYS\OnBaseFiles\Forms\MedicareReviews\" & sTemplateName,ForReading,False)

'Populate template string with text of word document
Dim sTemplate
sTemplate = oTemplate.ReadAll

'Close word document
oTemplate.Close

'Replace value placeholder with value
sTemplate = Replace(sTemplate,"#BENEFICIARY#",sBeneficiary)
sTemplate = Replace(sTemplate,"#DOS#",sDOS)
sTemplate = Replace(sTemplate,"#CCN#",sCCN)
sTemplate = Replace(sTemplate,"#HIC#",sHIC)
sTemplate = Replace(sTemplate,"#DOD#",sDOD)
sTemplate = Replace(sTemplate,"#HCPC#",sHCPC)
sTemplate = Replace(sTemplate,"#REASON#",sReason)
sTemplate = Replace(sTemplate,"#DATE#",sDate)

'Change name to first_middle_last
Dim aBeneficiary
aBeneficiary = Split(sBeneficiary,",")

If UBound(aBeneficiary) > 0 Then
	Dim sFirstName
	sFirstName = Trim(aBeneficiary(1))
End If

Dim sLastName
sLastName = Trim(aBeneficiary(0))

sBeneficary = sFirstName & "_" & sLastName
sBeneficiary = Replace(sBeneficary,",","")
sBeneficiary = Replace(sBeneficiary," ","_")
sBeneficiary = Replace(sBeneficiary,"'","")

'Filename of file to output
Dim sFileName
sFileName = "ReviewLetter_" & sBeneficiary & "_" & sDocHandle & ".rtf"

'Full path of output file
Dim sFilePath
sFilePath = "\\CORP.3SG.COM\ONB\PRD_SRC\Billing\MedicareReviews\" & sFileName

'Declare oFile
Dim oFile

If oFSO.FileExists(sFilePath) Then
	iReturn = MsgBox("Letter already exists, would you like to overwrite?",VbYesNo,"Letter already exists")
	If iReturn = 6 Then

		'Create output file from full path
		Set oFile = oFSO.OpenTextFile(sFilePath,ForWriting,True)

		'Write populated template to output file
		Call oFile.Write(sTemplate)

		'Close output file
		oFile.Close

		iReturn = MsgBox("Letter has been re-created at: " & vbCrLf & sFilePath & vbCrLf & "Would you like to open the file now?",vbYesNo,"Open file?")
	Else
		iReturn = MsgBox("Existing letter is located at: " & vbCrLf & sFilePath & vbCrLf & "Would you like to open the existing letter for review?",vbYesNo,"Open file?")
	End If
Else
	'Create output file from full path
		Set oFile = oFSO.OpenTextFile(sFilePath,ForWriting,True)

		'Write populated template to output file
		Call oFile.Write(sTemplate)

		'Close output file
		oFile.Close

		iReturn = MsgBox("Letter has been generated at: " & vbCrLf & sFilePath & vbCrLf & "Would you like to open the file now?",vbYesNo,"Open file?")
End If

'Open file if it was created...
If oFSO.FileExists(sFilePath) Then
	If iReturn = 6 Then
		Dim sCommand
		sCommand =  "winword " & chr(34) & sFilePath & chr(34)
		
		Dim oShell
		Set oShell = CreateObject("WScript.Shell")
		
		Dim iReturn
		iReturn = oShell.Run(sCommand,,False)
		
		If iReturn <> 0 Then
			sCommand = "wordpad " & chr(34) & sFilePath & chr(34)
			iReturn = oShell.Run(sCommand,,False)
			If iReturn <> 0 Then
				MsgBox "There was an error opening file, cannot find a suitable program to open it with.",,"Error opening file."
			End If
		End If
	End If
	
	'Write index file for DIP process
	Dim sImportPath
	sImportPath = "\\CORP.3SG.COM\ONB\PRD_SRC\BKOIMP\MedicareReviews\" & sFileName
	
	'Dim sTimeStamp
	'sTimeStamp = Now()
	'sTimeStamp = Replace(sTimeStamp,"/","")
	'sTimeStamp = Replace(sTimeStamp,":","")
	'sTimeStamp = Replace(sTimeStamp," ","")
		
	'See if account number was picked up, if not do not add to index file
	If sAcctNum <> "" Then
		Dim oIndexFile
		Set oIndexFile = oFSO.OpenTextFile("\\CORP.3SG.COM\ONB\PRD_SRC\Billing\MedicareReviews\MedicareReviewIndex.txt",ForAppending,True)	
		Call oIndexFile.WriteLine(sAcctNum & "," & sImportPath & "," & "AR - EP - Medicare Reviews")
		oIndexFile.Close
	Else
		MsgBox "Account Number is not set on this document, this document will not be imported into OnBase."
	End If
'...otherwise, show cannot be found message
Else
	MsgBox sFilePath & " cannot be found.",,"File not found"
End If
		
		


'Destroy Objects
'Removed oKeys - No reference to it in script except in the destroy objects
'Added oKeywords, changed order
Set oIndexFile = Nothing
Set oShell = Nothing
Set oFile = Nothing
Set oTemplate = Nothing
Set oFSO = Nothing
Set oKeywords = Nothing
Set oDoc = Nothing
Set oApp = Nothing
