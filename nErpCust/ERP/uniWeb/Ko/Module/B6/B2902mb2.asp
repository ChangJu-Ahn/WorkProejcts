<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(내부부서코드반영 Transaction)
'*  3. Program ID           : B2902mb2.asp
'*  4. Program Name         : B2902mb2.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B29022BatchTableReflection
'*  7. Modified date(First) : 2000/10/04
'*  8. Modified date(Last)  : 2002/11/25
'*  9. Modifier (First)     : Hwnag Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%


    
Dim intRetCD
Dim iTotal, iSuccess, iConfirm

Call LoadBasisGlobalInf()

Call HideStatusWnd                                                             

lgErrorStatus     = "NO"
lgErrorPos        = ""																'☜: Set to space

	Call SubCreateCommandObject(lgObjComm)
	Call SubBizBatch()
	Call SubCloseCommandObject(lgObjComm)

Sub SubBizBatch()
    Dim Module_CD
    Dim Change_DT
    Dim OrgID
    Dim strMsg_cd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    OrgID = Request("txtOrgChangeId")    
	Module_CD = Request("txtModuleCd")
	Change_DT = Request("txtChangeDt")
	iConfirm = Request("txtConfirm")

	If OrgID="" Then
	    Response.End 
	End if

	With lgObjComm
        .CommandText = "USP_TABLE_REFLECTION"
        .CommandType = adCmdStoredProc

        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"   ,adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_change_id" ,adVarXChar,adParamInput, 5, OrgID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@module_cd"     ,adVarXChar,adParamInput, 1, Module_CD)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@org_change_dt" ,adVarXChar,adParamInput, 8, Change_DT)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@chg_fg"        ,adVarXChar,adParamInput, 1, iConfirm)	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@userid"        ,adVarXChar,adParamInput, 13, gUsrId)

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"        ,adVarXChar,adParamOutput, 6)	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@itotal"        ,adInteger,adParamOutput, 9)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@isuccess"      ,adInteger,adParamOutput, 9)
	    
	    lgObjComm.Execute ,, adExecuteNoRecords
	End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        iTotal   = lgObjComm.Parameters("@itotal").Value
        iSuccess = lgObjComm.Parameters("@isuccess").Value
        
        If  IntRetCD < 0 Then								'SP는 정상 수행했으나 Return 값이 -1인 경우 
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
'			If strMsg_cd <> "" Then
				Call DisplayMsgBox(strMsg_cd, vbInformation,strMsg_cd , "", I_MKSCRIPT)
				IntRetCD = -1				
'			Else
'				Call DisplayMsgBox(iTotal, vbInformation, iSuccess, "", I_MKSCRIPT)
'				IntRetCD = -1
 '           End If
            Exit Sub
        Else
            IntRetCD = 1									'SP 정상 수행 
        End If
    Else													'SP 수행 도중 ERROR 발생시 처리         
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End If
    
End Sub


Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
End Sub

Sub CommonOnTransactionCommit()

End Sub

Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"

End Sub

Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       Call DisplayMsgBox(Err.number & " : " & Err.Description, vbInformation, "", "", I_MKSCRIPT)
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          Call DisplayMsgBox(Err.number & " : " & Err.Description, vbInformation, "", "", I_MKSCRIPT)             'Can not create(Demo code)
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If
End Sub
%>

<Script Language="VBScript">

   If Trim("<%=lgErrorStatus%>") = "NO" Then
      With Parent
           IF  "<%=CInt(intRetCD)%>" >= 0 Then
           	    .frm1.hTotal.value = "<%=iTotal%>"
	            .frm1.hSuccess.value = "<%=iSuccess%>"
            
				.Batch_OK
           End If
      End with
   End If   
</Script>	