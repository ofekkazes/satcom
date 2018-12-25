Set objFSO=CreateObject("Scripting.FileSystemObject")
Dim connectionID, triedConnection, idArr, dmaID, DataminerIDArr, terminalArray
triedConnection = FALSE
DataminerIDArr = Array()
terminalArray = Array()
dmaID = 0

getTerminalList() 'Get list of terminals from the REST API | API -> idArr
getConnFromFile() 'Get latest connection id
getTerminalIDs 0 'Get terminal dmaIDs from the terminal list | idArr -> DataminerIDArr(JointTerminal)

'Place list of parameters to an array
Set serviceRequesting = New ServiceRequest 
serviceRequesting.SetSoapAction("GetParameters")
For index = 0 To UBound(DataminerIDArr) Step 1
	getAllParameters serviceRequesting, DataminerIDArr(index) ' | DataminerIDArr(JointTerminal) -> terminalArray(TerminalInfo)
Next
serviceRequesting.Close()

writeFile()


'Place the parameter array in a file
Public Function writeFile
	fileOut = "newtec_parameters.csv"
	Set fileObj = objFSO.CreateTextFile(fileOut, True)
	csvString = "Terminal Name" & "," & "Is it up (0 = up; 1 = down)" & "," & "HRC Es/No" & "," & "RTN Throughput" & "," & "HRC Allocated Bitrate" & "," & "FWD Es/No" & "," & "FWD Throughput"
	fileObj.WriteLine csvString
	For index = 0 To UBound(DataminerIDArr) Step 1
			csvString = DataminerIDArr(index).dmaName & "," & terminalArray(index).logged & "," & terminalArray(index).hrcEsNo & "," & terminalArray(index).rtnThroughput & "," & terminalArray(index).hrcBitrate & "," & terminalArray(index).fwdEsNo & "," & terminalArray(index).fwdThroughput
			fileObj.WriteLine csvString
	Next
	fileObj.Close
End Function

'Retrieve the terminals from the REST API
Public Function getTerminalList
	Dim url, userName, password, restArr, i, temp
	Set restReq = CreateObject("Microsoft.XMLHTTP")
	'Rest API and request URL
	url = "http://10.0.38.14/rest/modem/"

	'HTTP Authentication
	userName = "NEWTEC HUB LOGIN"
	password = "NEWTEC HUB PASSWORD"

	restReq.open "GET", url, false, userName, password
	restReq.send
	
	restArr = Split(restReq.responseText, "id")
	idArr = Array()
	
	For i = 1 To UBound(restArr) Step 1
		temp = Split(Split(Split(restArr(i), "name")(1), ":"&chr(34))(1), chr(34))(0) 'Remove all the unnecessary info | chr(34) = " sign
		
		'Filters unneeded are not added to the array
		If Not InStr(temp, "best-effort-only") > 0 Then
			If Not InStr(temp, "ICMP_PRIORITY") > 0 Then
				temp = "VNO-1." & temp
				Additem idArr, temp
			End If
		End If
	Next
End Function

'Add a value to a dynamic array
Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

'Add an object to a dynamic array
Function AddTerminal(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    Set arr(UBound(arr)) = val
    AddTerminal = arr
End Function

'Get all the parameters needed from the SOAP API
Public Function getAllParameters(serviceRequest, terminalId)
	'Query build
	infoQuery = "<?xml version='1.0' encoding='utf-8'?>"
	infoQuery = infoQuery & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v1=""http://www.skyline.be/api/v1"">"
		infoQuery = infoQuery & " <soapenv:Header/>"
		infoQuery = infoQuery & " <soapenv:Body>"
			infoQuery = infoQuery & "<v1:GetParameters>"
				infoQuery = infoQuery &"<v1:connection>" & connectionID & "</v1:connection>"
				infoQuery = infoQuery &"<v1:dmaID>" & dmaID & "</v1:dmaID>"
				infoQuery = infoQuery &"<v1:elementID>" & terminalId.dmaID & "</v1:elementID>"
				infoQuery = infoQuery &"<v1:parameters>"
					infoQuery = infoQuery &"<v1:string>Modem Total RTN Throughput</v1:string>"
					infoQuery = infoQuery &"<v1:string>HRC Allocated Bitrate</v1:string>"
					infoQuery = infoQuery &"<v1:string>Modem FWD Es/No</v1:string>"
					infoQuery = infoQuery &"<v1:string>Modem Total FWD Throughput</v1:string>"
					infoQuery = infoQuery &"<v1:string>HRC Es/No</v1:string>"
					infoQuery = infoQuery &"<v1:string>Modem Operational State</v1:string>"
				infoQuery = infoQuery &"</v1:parameters>"
			infoQuery = infoQuery &"</v1:GetParameters>"
		infoQuery = infoQuery & " </soapenv:Body>"
	infoQuery = infoQuery & " </soapenv:Envelope>"
	
	serviceRequest.sSOAPRequest = infoQuery
	serviceRequest.SendRequest
	
	'If the connection has timed out, create a new one (only once to prevent StackOverflowException) and check the paramete needed
	If serviceRequest.sStatus = 500 Then
		If triedConnection = FALSE Then
			triedConnection = TRUE
			setNewConnFile()
			getAllParameters serviceRequest, terminalId
			Exit Function
		Else
			WScript.Echo "Connection to the Hub failed, check connectionID manually"
			WScript.Quit 1
		End If
	Else
		Set terminal = New TerminalInfo
		terminal.dmaID = terminalId.dmaID
		terminal.logged = Split(Split(serviceRequest.sResponse, "<Value>")(1), "</Value>")(0)
		terminal.fwdEsNo = Split(Split(serviceRequest.sResponse, "<Value>")(2), "</Value>")(0)
		terminal.hrcBitrate = Split(Split(serviceRequest.sResponse, "<Value>")(3), "</Value>")(0)
		terminal.hrcEsNo = Split(Split(serviceRequest.sResponse, "<Value>")(4), "</Value>")(0)
		terminal.fwdThroughput = Split(Split(serviceRequest.sResponse, "<Value>")(5), "</Value>")(0)
		terminal.rtnThroughput = Split(Split(serviceRequest.sResponse, "<Value>")(6), "</Value>")(0)
		AddTerminal terminalArray, terminal
	End If
End Function

'Get the dmaIDs from the API
Public Function getTerminalIDs(startingIndex)
	Set serviceReq = New ServiceRequest 
	serviceReq.SetSoapAction("GetElementByName")
	
	For index = startingIndex To UBound(idArr) Step 1: Do
		'Query build
		infoQuery = "<?xml version='1.0' encoding='utf-8'?>"
		infoQuery = infoQuery & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v1=""http://www.skyline.be/api/v1"">"
			infoQuery = infoQuery & " <soapenv:Header/>"
			infoQuery = infoQuery & " <soapenv:Body>"
				infoQuery = infoQuery & "<v1:GetElementByName>"
					infoQuery = infoQuery &"<v1:connection>" & connectionID & "</v1:connection>"
					infoQuery = infoQuery &"<v1:elementName>" & idArr(index) & "</v1:elementName>"
				infoQuery = infoQuery &"</v1:GetElementByName>"
			infoQuery = infoQuery & " </soapenv:Body>"
		infoQuery = infoQuery & " </soapenv:Envelope>"
		
		serviceReq.sSOAPRequest = infoQuery
		serviceReq.SendRequest
		
		'If the connection has timed out, create a new one (only once to prevent StackOverflowException) and check the paramete needed
		If serviceReq.sStatus = 500 Then
			If InStr(serviceReq.sResponse, "No such element") > 0 Then
				Exit Do
			Else
				If triedConnection = FALSE Then
					triedConnection = TRUE
					setNewConnFile()
					getTerminalIDs(index)
					Exit Function
				Else
					WScript.Echo "Connection to the Hub failed, check connectionID manually"
					WScript.Quit 1
				End If
			End If
		Else
			If dmaID = 0 Then
				dmaID = Split(Split(serviceReq.sResponse, "<DataMinerID>")(1), "</DataMinerID>")(0)
			End If
			Set jointTerminal = New JointTerminalID
			jointTerminal.dmaID = Split(Split(serviceReq.sResponse, "<ID>")(1), "</ID>")(0)
			jointTerminal.dmaName = idArr(index)
			AddTerminal DataminerIDArr, jointTerminal
		End If
	Loop While False: Next
	
End Function

'Read the latest connection id from a file
Public Function getConnFromFile
	strFile = "connection.txt"
	If Not objFSO.FileExists(strFile) Then
		setNewConnFile()
	End If
	
	Set objFile = objFSO.OpenTextFile(strFile)
	Do Until objFile.AtEndOfStream
		strLine= objFile.ReadLine
		connectionID = strLine
	Loop
	objFile.Close
End Function

'Write the new connection id to a file
Public Function setNewConnFile
	connectQuery = "<?xml version='1.0' encoding='utf-8'?>"
	connectQuery = connectQuery & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v1=""http://www.skyline.be/api/v1"">"
		connectQuery = connectQuery & " <soapenv:Header/>"
		connectQuery = connectQuery & " <soapenv:Body>"
			connectQuery = connectQuery & "<v1:ConnectApp>"
				connectQuery = connectQuery &"<v1:host>NEWTEC HUB IP</v1:host>"
				connectQuery = connectQuery &"<v1:login>NEWTEC HUB LOGIN</v1:login>"
				connectQuery = connectQuery &"<v1:password>NEWTEC HUB PASSWORD</v1:password>"
				connectQuery = connectQuery &"<v1:clientAppName>?</v1:clientAppName>"
			connectQuery = connectQuery &"</v1:ConnectApp>"
		connectQuery = connectQuery & " </soapenv:Body>"
	connectQuery = connectQuery & " </soapenv:Envelope>"
	
	Set objSOAP = New ServiceRequest 
	objSOAP.SetSoapAction("ConnectApp")
	objSOAP.sSOAPRequest = connectQuery
	objSOAP.SendRequest
	
	
	connID = Split(Split(objSOAP.sResponse, "<ConnectAppResult>")(1), "</ConnectAppResult>")
	
	outFile="connection.txt"
	Set objFile = objFSO.CreateTextFile(outFile,True)
	objFile.Write connID(0)
	objFile.Close
	
	connectionID = connID(0)
End Function

Class TerminalInfo
	Public dmaId, rtnThroughput, fwdThroughput, hrcBitrate, hrcEsNo, fwdEsNo, logged
	Private Sub Class_Initialize 

	End Sub 
End Class

Class JointTerminalID
	Public dmaID, dmaName
End Class

'Soap connection request
Class ServiceRequest 
 Private oWinHttp,sContentType 
 Public sWebServiceURL, sSOAPRequest,sResponse, sStatus,wantedFunction,sHost, sStatusText
  
Private Sub Class_Initialize 
 Set oWinHttp = CreateObject("Microsoft.xmlhttp") 
  
'Web Service Content Type 
 sContentType ="text/xml; charset=utf-8" 
  
End Sub 
  
Public Function SetSoapAction(servicename) 
	sWebServiceURL = "http://10.0.38.14/API/v1/soap.asmx"
	wantedFunction = servicename
	sHost = "10.0.38.14"
End Function 
  
Public Function SendRequest 
	 
	'Open HTTP connection  
	oWinHttp.Open "POST", sWebServiceURL, False 
	  
	'Setting request headers  
	oWinHttp.setRequestHeader "Content-Type", sContentType 
	oWinHttp.setRequestHeader "Content-Length", len(sSOAPRequest)
	oWinHttp.setRequestHeader "SOAPAction", "http://www.skyline.be/api/v1/" & wantedFunction
	oWinHttp.setRequestHeader "Accept-Encoding", "gzip,deflate"
	oWinHttp.setRequestHeader "Host", sHost
	oWinHttp.setRequestHeader "Connection", "Keep-Alive"
	  
	'Send SOAP request 
	 oWinHttp.Send  sSOAPRequest 
	  
	'Get XML Response 
	sResponse = oWinHttp.responsetext
	sStatus = oWinHttp.status
	sStatusText = oWinHttp.statusText
End Function 
  
Public Function Close 
	Set oWinHttp = Nothing 
End Function 
  
 
End Class 



