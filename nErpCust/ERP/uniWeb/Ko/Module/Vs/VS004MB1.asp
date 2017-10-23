<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()

On Error Resume Next														'��: 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strPrintOpt
Dim strYyyyMm, strBpCd
Dim strMsgCd, strMsgValue, strSpId,strPlantCd


Dim IntRetCD
Dim lgObjComm
Dim lgErrorStatus 

strMode		= Request("txtMode")												'�� : ���� ���¸� ���� 
strPrintOpt	= Trim(Request("strPrintOpt"))
strPlantCd  = Trim(Request("strPlantCd"))
strYyyyMm	= Trim(Request("strYyyyMm"))
strBpCd		= Trim(Request("strBpCd"))

Select Case strMode

Case CStr(UID_M0002)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	'********************************************************  
	'                        Execution
	'********************************************************  

	Err.Clear
	Call SubCreateCommandObject(lgObjComm)
	    
	With lgObjComm

		.CommandText = "usp_v_vs0040t"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter("RETURN_VALUE",	adInteger,	adParamReturnValue)

		.Parameters.Append .CreateParameter("@strYyyyMm",	adVarXChar,	adParamInput,	6,	strYyyyMm)
		.Parameters.Append .CreateParameter("@strBpCd",		adVarXChar,	adParamInput,	10,	strBpCd)
		.Parameters.Append .CreateParameter("@strPlantCd",	adVarXChar,	adParamInput,	10,	strPlantCd)

		.Parameters.Append .CreateParameter("@strUsrId",	adVarXChar,	adParamInput,	13,	gUsrID)
		.Parameters.Append .CreateParameter("@strLang",	    adVarXChar,	adParamInput,	5,	gLang)
		.Parameters.Append .CreateParameter("@strMsgCd",	adVarXChar,	adParamOutput,	6)
		.Parameters.Append .CreateParameter("@strMsgValue",	adVarXChar,	adParamOutput,	255)
		.Parameters.Append .CreateParameter("@strSpId",		adVarXChar,	adParamOutput,	13)

		.Execute ,, adExecuteNoRecords

	End With

	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		if  IntRetCD <> 1 then		'��������϶� 1 �� return �Ѵ�.
			strMsgCd	= lgObjComm.Parameters("@strMsgCd").Value
			strMsgValue	= lgObjComm.Parameters("@strMsgValue").Value

			Call DisplayMsgBox(strMsgCd, vbInformation, strMsgValue, "", I_MKSCRIPT )                                                              '��: Protect system from crashing   
			Response.end
		end if

		strSpId	= lgObjComm.Parameters("@strSpId").Value

		Else
			lgErrorStatus     = "YES"                                                         '��: Set error status
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if
	    
	Call SubCloseCommandObject(lgObjComm)

End Select

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)

	On Error Resume Next                                                              '��: Protect system from crashing
	Err.Clear                                                                         '��: Clear Error status

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
		IF "<%=lgErrorStatus%>"	<> "YES" Then																    '��: ȭ�� ó�� ASP �� ��Ī�� 


			.txtSpId.Value = "<%=ConvSPChars(strSpId)%>"
			IF "<%=ConvSPChars(strPrintOpt)%>" = "Preview" Then
				.FncBtnPreview
			ElseIF "<%=ConvSPChars(strPrintOpt)%>" = "Print" Then
				.FncBtnPrint
			Else
			End If			

		Else
		    Call parent.LayerShowHide(0)	
		End If
	End With
</Script>
