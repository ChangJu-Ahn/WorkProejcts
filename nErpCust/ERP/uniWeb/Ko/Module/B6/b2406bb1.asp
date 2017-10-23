<% Option Explicit %>
<%'======================================================================================================
'*  1. Module Name          : BA
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B2406BA1
'*  4. Program Name         : 부서개편진행현황 
'*  5. Program Desc         : 부서개편과정 통제를 위한 조회 및 확정처리 화면 
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

'On Error Resume Next
'Err.Clear 

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

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)															'☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)															'☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)															'☜: Delete
             'Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizSave
' Desc : 
'============================================================================================================
Sub SubBizSave()
	Dim IntRetCD
	Dim strMsg_cd
	Dim iOrgChangeId,iWorkType,iYnfg


    Call SubCreateCommandObject(lgObjComm)	 

	iOrgChangeId = Trim(Request("txtOrgChangeID"))
	iWorkType = Trim(Request("txtWorkType"))
	iYnfg = Trim(Request("txtYnFg"))

    With lgObjComm
        .CommandText = "usp_dept_renewal_process"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@org_change_id"    ,adVarChar,adParamInput,5, iOrgChangeId)
		.Parameters.Append .CreateParameter("@work_type"		,adVarChar,adParamInput,1, iWorkType)
		.Parameters.Append .CreateParameter("@yn_fg"			,adVarChar,adParamInput,1, iYnfg)
	    .Parameters.Append .CreateParameter("@usr_id"			,adVarChar,adParamInput,13, gUsrID)
        .Parameters.Append .CreateParameter("@msg_cd"			,adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        If  IntRetCD <> 1 Then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
        End If
    Else
        lgErrorStatus = "YES"                                                         '☜: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If

	Call SubCloseCommandObject(lgObjComm)
	
	Response.Write " <Script Language=vbscript>                      " & vbCr
	Response.Write " With parent                                     " & vbCr
	Response.Write " 	If """ & lgErrorStatus & """ <> ""YES"" Then " & vbCr															    '☜: 화면 처리 ASP 를 지칭함 
	Response.Write "		.DbSaveOk                                " & vbCr
	Response.Write "	End If                                       " & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write " </Script>	                                     " & vbCr	
		
End Sub

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

	Dim strData
	Dim LngRow
	Dim iStr

	Redim UNISqlId(0)
	Redim UNIValue(0,0)

	UNISqlId(0) = "B2406MA101"
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set	ADF = Nothing

	iStr = Split(strRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(strRetMsg , vbInformation, I_MKSCRIPT)
		Response.End 	
    End If    

	If rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)	
		rs0.Close
		Set rs0 = Nothing
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If

    For LngRow = 0 To rs0.RecordCount - 1
		strData = strData & Chr(11) & ConvSPChars(rs0("ORG_CHANGE_ID"))
		strData = strData & Chr(11) & ConvSPChars(rs0("ORGNM"))
		strData = strData & Chr(11) & ConvSPChars(rs0("WORK_FLAG"))
		strData = strData & Chr(11) & ConvSPChars(rs0("WORK_FLAG_NM"))
		strData = strData & Chr(11) & ConvSPChars(rs0("WORK_DT"))
		strData = strData & Chr(11) & ConvSPChars(rs0("OK_OPEN_FG"))
		strData = strData & Chr(11) & ConvSPChars(rs0("CANCEL_OPEN_FG"))
		strData = strData & Chr(11) & ConvSPChars(rs0("WORK_OK"))
		strData = strData & Chr(11) & ConvSPChars(rs0("WORK_CANCEL"))
		strData = strData & Chr(11) & LngRow
		strData = strData & Chr(11) & Chr(12)

		rs0.MoveNext
	Next

	rs0.Close
	Set rs0 = Nothing

	Response.Write " <Script Language=vbscript>               " & vbCr
	Response.Write " With Parent                              " & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData       " & vbCr
	Response.Write " .ggoSpread.SSShowData """ & strData & """" & vbCr
	Response.Write " .DbQueryOk                               " & vbCr
	Response.Write " End With                                 " & vbCr
	Response.Write " </Script>	                              " & vbCr

	Set ADF = Nothing																	'☜: ActiveX Data Factory Object Nothing
End Sub

%>

