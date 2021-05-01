Dim myApp
Dim myDocument
Dim docKeywords
Dim iKeywordCount
Dim iCount

Dim sBarcode
Dim sDocTypeNum

Set myApp = CreateObject("OnBase.Application")

Set myDocument = myApp.CurrentDocument

Set docKeywords = myDocument.Keywords

iKeywordCount = docKeywords.Count

iCount = 0
Do While iCount < iKeywordCount
	If docKeywords(iCount).Name = "AR - Unparsed Barcode" then
		Dim sAcctNum
		Dim sOrderNum
		Dim sStartDate
		Dim sScriptType
		Dim sDateOfService
		Dim sCreationDate
		Dim sMiscDocType
		Dim sPage
		Dim sPageO

		sBarcode = docKeywords(iCount).Value
		if len(sBarcode) > 0 then
			sDocTypeNum = mid(sBarcode,2,4)
			Select Case sDocTypeNum
				Case "0102"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0104"
					sAcctNum = right(sBarcode,len(sBarcode)-15)
					sCreationDate = mid(sBarcode,8,2) & "-" & mid(sBarcode,10,2) & "-" & mid(sBarcode,12,4)
					If Not IsDate(sCreationDate) Then
						sCreationDate = ""
					End If
				Case "0106"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0107"
					sAcctNum = right(sBarcode,len(sBarcode)-16)
					sStartDate = mid(sBarcode,9,2) & "-" & mid(sBarcode,11,2) & "-" & mid(sBarcode,13,4)
					If Not IsDate(sStartDate) Then
						sStartDate = ""
					End If
					sScriptType = mid(sBarcode,8,1)
				Case "0108"
					sAcctNum = right(sBarcode,len(sBarcode)-15)
					sStartDate = mid(sBarcode,8,2) & "-" & mid(sBarcode,10,2) & "-" & mid(sBarcode,12,4)
					If Not IsDate(sStartDate) Then
						sStartDate = ""
					End If
				Case "0117"
					sAcctNum = right(sBarcode,len(sBarcode)-15)
					sDateOfService = mid(sBarcode,8,2) & "-" & mid(sBarcode,10,2) & "-" & mid(sBarcode,12,4)
                        	        
                                         If Not IsDate(sDateOfService) Then
					        sDateOfService = ""
					End If  
                                        if left(sDateOfService,1) = "2" then
                                           sDateOfService = mid(sBarcode,12,2) & "-" & mid(sBarcode,14,2) & "-" & mid(sBarcode,8,4)      
                                        end if                                    
                        	Case "0118"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0119"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0121"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0122"
					sAcctNum = right(sBarcode,len(sBarcode)-7)
				Case "0139"
					sMiscDocType = mid(sBarcode,8,1)
					sPage = Mid(sBarcode,6,1)
					sPageO = Mid(sBarcode,7,1)
					sAcctNum = right(sBarcode,len(sBarcode)-8)
				Case "0286"
					sAcctNum = right(sBarcode,len(sBarcode)-22)
					sOrderNum = mid(sBarcode,16,7)
					sDateOfService = mid(sBarcode,12,2) & "-" & mid(sBarcode,14,2) & "-" & mid(sBarcode,8,4)
                        	        If Not IsDate(sDateOfService) Then
					        sDateOfService = ""
					End If                                    

			End Select
		
			If sAcctNum <> "" Then
				Call docKeywords.AddKeyword("AR - Account Number",sAcctNum)
			End If
		
			If sOrderNum <> "" Then
				Call docKeywords.AddKeyword("AR - Order Number",sOrderNum)
			End If

			If sStartDate <> "" Then
				Call docKeywords.AddKeyword("AR - Start Date",sStartDate)
			End If
		
			If sScriptType <> "" Then
				Call docKeywords.AddKeyword("AR - Script Type",sScriptType)
			End If
		
			If sDateOfService <> "" Then
				Call docKeywords.AddKeyword("AR - Date of Service",sDateOfService)
			End If

			If sCreationDate <> "" Then
				Call docKeywords.AddKeyword("AR - Creation Date",sCreationDate)
			End If     

			If sMiscDocType <> "" Then
				Dim sParsedType
				Select Case sMiscDocType
					Case "5"
						sParsedType = "5 - MEDICAL INFORMATION"
				End Select
			
				
				Call docKeywords.AddKeyword("AR - Misc Doc Type",sParsedType)
			End If     
		End If
	End If
	iCount = iCount + 1
Loop

Call myDocument.StoreKeywords()

'Destroy Objects
Set docKeywords = Nothing
Set myDocument = Nothing
Set myApp = Nothing
