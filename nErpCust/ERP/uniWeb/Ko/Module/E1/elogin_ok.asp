
<%	
'===========================================================================================
'*  10.Version Label			: 대사우 로그인프로세스 개선작업 
'*  10.Compatible uniConnecotor : 1,0,0,440
'*  10.Compatible PLProcess     : 2.0.0.2 (한솔전자 지원 PS.2003.02.14)
'===========================================================================================
' Const Variable 정의 

	Const C_ADODBConnString		= 0
	Const C_ComproxyConnString	= 1
	Const C_DBAgentDBConnString = 2
	Const C_DBLoginPassword		= 3
	Const C_DBSAPwd				= 4
	Const C_DsnNo				= 5
	Const C_DefaultDBServer		= 6
	Const C_DefaultDBServerIP	= 7
	Const C_HTTPXMLFileNm		= 8
	Const C_STADBConnString		= 9
	Const C_LogInTryCnt			= 10
	Const C_LogInUsrId			= 11
	Const C_LogInURLLANG		= 12	
	Const C_Version				= 13
	Const C_XMLFileNm			= 14
	Const C_EmpNo			    = 15
	Const C_DeptAuth			= 16
	Const C_ProAuth			    = 17	
			
		
	Dim ReadReg
	Dim pExecuteStatus	
	Dim strURLLang
	Dim strURLLangUserID
	Dim strProperties
	Dim StrGlobalCollection
	Dim ArrParam
	Dim APDateFormat
	Dim APDateSeperator

    On Error Resume Next

    Set ReadReg = Server.CreateObject("PLProcess.LCProcess")	 		
	
	strURLLang =  Request("txtURL") & "/" & Request("txtLangCdForURL") 
	strURLLangUserID = "[" & strURLLang & "][UID:" & Request("txtUID")  & "]"

'	ReadReg.DebugMode = false  '보안관련 : connection  string 숨김 

	'=====================================================================================
	' Get AP Server Date Format
	'=====================================================================================   

	APDateFormat = DateAdd("D",-1,"2004-01-01")

	APDateFormat = Replace(APDateFormat,"2003","YYYY")
	APDateFormat = Replace(APDateFormat,"03","YY")
	APDateFormat = Replace(APDateFormat,"Dec","MMM")
	APDateFormat = Replace(APDateFormat,"12","MM")
	APDateFormat = Replace(APDateFormat,"31","DD")

	For I = 1 To Len(APDateFormat)
	    APDateSeperator = Mid(APDateFormat,I,1) 
	    If Not (APDateSeperator = "Y" Or APDateSeperator = "M" Or APDateSeperator = "D" ) Then
	       Exit For
	    End If
	Next
	
	redim ArrParam(21)

	ArrParam(0) = Request("txtVD")
	ArrParam(1) = UCase(Request("txtLang"))
	ArrParam(2) = Request("txtUID")
	ArrParam(3) = Request("txtpassword")
	ArrParam(4) = Request("txtCnt")
	ArrParam(5) = Request.ServerVariables("APPL_PHYSICAL_PATH")
	ArrParam(6) = Request.ServerVariables("REMOTE_ADDR")
	ArrParam(7) = Request.ServerVariables("REMOTE_HOST")
	ArrParam(8) = Request("txtLangCdForURL")
	ArrParam(9) = Request("txtURL")
	ArrParam(10) = strURLLang
	ArrParam(11) = strURLLangUserID
	ArrParam(12) = Request("txtClientNum1000")
	ArrParam(13) = Request("txtClientNumDec")
	ArrParam(14) = Request("txtClientDateFormat")
	ArrParam(15) = Request("txtClientDateSeperator")	
	ArrParam(16) = Request("txtHttpWebSvrIPURL")
	ArrParam(17) = APDateFormat
	ArrParam(18) = APDateSeperator
	
	ArrParam(19) = "1"                                     '2003-08-07 leejinsoo
	ArrParam(20) = "MSSQL"                                 '2003-08-07 leejinsoo 
	ArrParam(21) = "F"                                     'RDSUse 2003-12-13 leejinsoo

	pExecuteStatus        = ReadReg.ESSLoginProcess(ArrParam)	

	If Err.number <> 0 Then		
		pExecuteStatus = "CreateObject error - PLProcess" & Chr(11) & Err.Number & Chr(11) & Err.Description		
	End If 	

	strProperties = Split(ReadReg.GetProperties(True), chr(11))

	Response.Cookies("unierp")("ESS_gXMLFileNm")	 = strProperties(C_XMLFileNm)
	Response.Cookies("unierp")("ESS_gHTTPXMLFileNm") = strProperties(C_HTTPXMLFileNm)	
	gURLLangUserID		                             = strURLLangUserID						
	gHTTPXMLFileNm		                             = strProperties(C_HTTPXMLFileNm)	
	gDsnNo				                             = strProperties(C_DsnNo)
	gXMLMODE                                         = strProperties(C_Version)
	
	Response.Cookies("unierp")("gServerIP")             = "HTTP://" & Request.ServerVariables("SERVER_NAME")
	
    Response.Cookies("unierp")("gCanBeDebug")           = Right("0" & Request("txtCanBeDebug"),1)
    Response.Cookies("unierp")("giSOLVL")               = Request.Form("txtILVL")

%>
<!-- #Include file="../../inc/incServer.asp" -->

<SCRIPT LANGUAGE="VBScript">
	
	Dim cExecuteStatus 
	Dim XMLFileNm			
	Dim RetV
	Dim arrParam(2)

	XMLFileNm  = "<%=strProperties(C_XMLFileNm)%>"
		
	cExecuteStatus = Trim("<%=Replace(pExecuteStatus,Chr(13) & Chr(10),"")%>")

	On Error Resume Next
	
'	Msgbox "cExecuteStatus:" & cExecuteStatus
	
	If cExecuteStatus <> "" Then
		cExecuteStatus = Split(cExecuteStatus,Chr(11))

		If cExecuteStatus(0) = "M" Then
	
			MsgBox cExecuteStatus(4), vbExclamation , "<%=gLogoName%>"   
			      
			Select Case cExecuteStatus(1)
 
   			   Case "210110"      '     'Please enter new password. 
			               RetV = window.showModalDialog("EchangePWFirst.asp?txtUID=" & "<%strProperties(C_LogInUsrId)%>", Array(arrParam), "dialogWidth=390px; dialogHeight=190px; center: Yes; help: No; resizable: No; status: No;")
							If RetV <> "C" then
							 	Top.location.href = "emenu.asp"
							End if       
'			   Case "210007"           'The valid period of password has been expired.  Please register a new password. 
'			               RetV = window.showModalDialog("EchangePW.asp", Array(arrParam),"dialogWidth=400px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
'							If RetV <> "C" then  'not change password and close the window
'							 	Top.location.href = "emenu.asp"
'							End if			               
			   Case "210005"          'input invalid password
			               parent.window.login.txtpassword.select 
			               parent.login.txtCnt.value = 0	'ReadReg.LogInTryCnt
			End Select 

		ElseIf cExecuteStatus(0) = "S" Then                         'Show port message

			MsgBox cExecuteStatus(1), vbExclamation , "<%=gLogoName%>"
		Else
			MsgBox "Step : " & cExecuteStatus(0) & vbCrLf & "Code : " & cExecuteStatus(1) & vbCrLf & "Desc : " & cExecuteStatus(2), vbExclamation , "<%=gLogoName%>"
		End If
	
	else	
		if XMLFileNm = "" then 
			MsgBox "XML File has not been created", vbExclamation , "<%=gLogoName%>"        
		end if
		Top.location.href = "emenu.asp"
	end if

</SCRIPT>

<%
Set ReadReg = Nothing
%>


