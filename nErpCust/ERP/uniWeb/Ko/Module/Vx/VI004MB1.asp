<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()

On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strPrintOpt
Dim strYyyyMm, strPlantCd
Dim strMsgCd, strMsgValue, strSpId


Dim IntRetCD
Dim lgObjComm
Dim lgErrorStatus 

strMode		= Request("txtMode")												'☜ : 현재 상태를 받음 
strPrintOpt	= Trim(Request("strPrintOpt"))

strYyyyMm	= Trim(Request("strYyyyMm"))
strPlantCd	= Trim(Request("strPlantCd"))

if isnull(strPlantCd) or strPlantCd = "" then
	strPlantCd = "%"
end if

Select Case strMode

Case CStr(UID_M0002)														'☜: 현재 조회/Prev/Next 요청을 받음 
	'********************************************************  
	'                        Execution
	'********************************************************  

	Err.Clear
	Call SubCreateCommandObject(lgObjComm)
 
	With lgObjComm

		.CommandText = "usp_v_vI0040t"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter("RETURN_VALUE",	adInteger,	adParamReturnValue)

		.Parameters.Append .CreateParameter("@strYyyyMm",	adVarChar,	adParamInput,	6,	strYyyyMm)
		.Parameters.Append .CreateParameter("@strUsrId",	adVarChar,	adParamInput,	13,	gUsrID)
		.Parameters.Append .CreateParameter("@strPlantCd",	adVarChar,	adParamInput,	4,	strPlantCD)
		.Parameters.Append .CreateParameter("@strMsgCd",	adVarChar,	adParamOutput,	6)
		.Parameters.Append .CreateParameter("@strMsgValue",	adVarChar,	adParamOutput,	255)
		.Parameters.Append .CreateParameter("@strSpId",		adVarChar,	adParamOutput,	13)

		.Execute ,, adExecuteNoRecords

	End With

	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		if  IntRetCD <> 1 then		'정상실행일때 1 을 return 한다.
			strMsgCd	= lgObjComm.Parameters("@strMsgCd").Value
			strMsgValue	= lgObjComm.Parameters("@strMsgValue").Value

			Call DisplayMsgBox(strMsgCd, vbInformation, strMsgValue, "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
		end if

		strSpId	= lgObjComm.Parameters("@strSpId").Value

		Else
			lgErrorStatus     = "YES"                                                         '☜: Set error status
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if
	    
	Call SubCloseCommandObject(lgObjComm)

End Select

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)

	On Error Resume Next                                                              '☜: Protect system from crashing
	Err.Clear                                                                         '☜: Clear Error status

	If CheckSYSTEMError(pErr,True) = True Then
		ObjectContext.SetAbort
		Call SetErrorStatus
	Else
		If CheckSQLError(pConn,True) = True Then
			ObjectContext.SetAbort
			Call SetErrorStatus
		End If
	End If

End Sub   

%>	

<Script Language="VBScript">
	With parent
		IF "<%=lgErrorStatus%>"	<> "YES" Then																    '☜: 화면 처리 ASP 를 지칭함 


			.txtSpId.Value = "<%=ConvSPChars(strSpId)%>"
			IF "<%=ConvSPChars(strPrintOpt)%>" = "Preview" Then
				.FncBtnPreview
			ElseIF "<%=ConvSPChars(strPrintOpt)%>" = "Print" Then
				.FncBtnPrint
			Else
			End If			

		End If
	End With
</Script>
