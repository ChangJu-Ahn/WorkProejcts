<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%'======================================================================================================
'*  1. Module Name          : BA
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B2403mbA1
'*  4. Program Name         : 부서개편내역등록 
'*  5. Program Desc         : 홍익인간 조직정보를 ERP상에 등록한다. 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2005/10/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True												'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%													

On Error Resume Next
Err.Clear 

Call LoadBasisGlobalInf() 

Dim ADF																'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg														'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0						'DBAgent Parameter 선언 

Const C_CLOSE_GB            = 0
Const C_TAGET_WORKING_MNTH  = 1
Const C_CLOSE               = 2
Const C_CANCEL              = 3

'---------------------------------------------------------------------------------------------------------

Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""  
    lgOpModeCRUD      = Request("txtMode")												'☜: Read Operation Mode (CRUD)   



Call SubBizQuery()
        

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

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : 
'============================================================================================================
Sub SubBizQuery()
'	On Error Resume Next
'	Err.Clear 
    Dim IntRetCD
	Dim strMsg_cd
	Dim iOrgChangeId,iWorkType,iYnfg
	Dim strMsg_text
	Dim strSp    


    Call SubCreateCommandObject(lgObjComm)	 

	'iOrgChangeId = Trim(Request("txtOrgChangeID"))
	'iWorkType = Trim(Request("txtWorkType"))
	'iYnfg = Trim(Request("txtYnFg"))
'Call DisplayMsgBox("124200", vbInformation, "", "", I_MKSCRIPT)
    With lgObjComm
        .CommandText = "usp_create_org_dept"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@msg_cd",adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If Err.number = 0 Then
	   IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
	   If IntRetCD <> 1 then
	      strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
	      If strMsg_Cd <> "" Then
		       Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
		  End If
	      Response.end
		Else
			'lgSp_Id = lgObjComm.Parameters("@sp_id").Value
	   End If

	Else
	  lgErrorStatus     = "YES"
	  Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
	  Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if
End Sub

%>

