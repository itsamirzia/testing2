
'
'---	This VBScript will retrieve the nessasary document data and then run a query against the DW
'---	returning the AR Balance. It will then populate the new key field AR - RGH AR Balance' and return
'---	to the workflow. If the AR Balance is equal to 0.00 the document will then exit the workflow. If 
'---	the AR Balance is not equal to 0.00 this script will load the AR Balance into the new keyword field
'---	'AR - RGH AR Balance'. Then the document will follow the existing rules
'
'---	Vaiable declarations
'
Dim strDBDatabase
Dim strDBPassword
Dim strDBServer
Dim strDBUsername
Dim strConn
Dim objConn
Dim objRS
Dim sql
'
Dim strAccountNumber
Dim strClaimNumber
Dim strDocId
Dim strDocumentDate
Dim strDOS
Dim strItemNum
'
Dim strMonth
Dim strDay
Dim strYear	
Dim strTempDate
Dim strTempDate2
Dim strToday
Dim Month
Dim Day
Dim Year
'
Dim AR_Balance
Dim Count
Dim DaysDiff
Dim Queue
'
Dim bARBalance
'
'---	Get current doc and keywords
'
Dim myApp
Set myApp = CreateObject("Onbase.Application")
'
Dim myDoc
Set myDoc = myApp.CurrentDocument
'
Dim myKeys
Set myKeys = myDoc.Keywords
'
Dim myUsername
myUsername = myApp.username
'
'Initalize variables
'
on error resume next
count = 0
'
'---	Need to save certain fields retrieved from the document for later use
'
Do while count < myKeys.Count
	If myKeys.item(count).name = "AR - Account Number" Then
		If Len(strAccountNumber) = 0 Then
			strAccountNumber = RTrim(myKeys.item(count).value)
		End If
	End If
'
 	If myKeys.item(count).name = "AR - Date of Service" Then

		If Len(strDOS) = 0 Then
			strTempDate = myKeys.item(count).value		
'			
			strMonth = Mid(strTempDate,1,2)
			Month = strMonth
'		
			strDay = Mid(strTempDate,4,2)
			Day = strDay
'			
			strYear	= Mid(strTempDate,7,4)
			Year = strYear
'			
			strDOS =  Year & "-" & Month & "-" & Day & " 00:00:00.000"	
		End If
	End If
'
	Count = count + 1
Loop

'MsgBox "Account Number: " & strAccountNumber
'MsgBox "Date of Service: " & strDOS

'
'---	If there is no account number on the document then do not process it.
'
If strAccountNumber <> "" Then
'
'
'---	Using the customer number and date of service from the document, this script will access the Edgepark Data Warehouse 
'---	and query the table rghdm.dbo.ARFact, sum up the A/R Balance for the date of service and return it to populate the 
'---	key field 'AR - RGH AR Balance' on the document
'
	strDBServer = "EPDW" 
	strDBUsername = "dw"	
	strDBPassword = "dw"	
	strDBDatabase = "EPDW"
'
	strConn = "driver={SQL Server};server=" & strDBServer & ";UID=" & strDBUsername & ";PWD=" & strDBPassword & ";database=" & strDBDatabase
	Set objConn = CreateObject("ADODB.Connection")
	Set objRS = CreateObject("ADODB.Recordset")
'
	objConn.open strConn
	objRS.ActiveConnection = objConn
'
	sql = 	"select SUM(ISNULL(AR_Amount,0)) as ARBalance from dbo.EdgAR (nolock) where AR_Company = '1' and AR_CustomerNo = '" & strAccountNumber & _ 
		"' and AR_ApplyToDos = '" & strDOS  & "' and AR_InsCode IN (5001,5002,5003,5004,5011,5012,5013,5014)"
'
	objRS.Open sql
	AR_Balance = objRS.fields("ARBalance").value
'	msgbox sql
'
'	msgbox "AR Balance: " & AR_Balance
'
'---	Add the newly calculated A/R Balance to the document key. Alway do this regardless of the amount.
'---	The workflow will check the AR Balance and decide what to do with the document.	
'
	If AR_Balance <> "" Then
		bARBalance = FALSE
		Count = 0
		Do while count < myKeys.Count
			If myKeys.item(count).name = "AR - RGH AR Balance" Then
				myKeys(count).value = AR_Balance
				bARBalance = TRUE
			End If
			Count = Count + 1 
		Loop
		'msgbox bARBalance
'
		If (bARBalance = FALSE) Then
			Call myKeys.AddKeyword("AR - RGH AR Balance",AR_Balance)
		End If
'
		Call myDoc.StoreKeywords()
'
	End If
'
	If objRS.State = 1 Then
		objRS.Close
	End If
'
'---	Close connection (if open)
'
	If objConn.State = 1 Then
		objConn.Close
	End If

        'Moved database object destroy to within if where it was created
       Set strConn = Nothing
       Set objRS = Nothing
       Set objConn = Nothing

End If
'

Set myKeys = Nothing
Set myDoc = Nothing
Set myApp = Nothing
