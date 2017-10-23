<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%

    Dim intRetCD

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    'lgErrorPos        = ""                                                           '¢Ð: Set to space

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
     
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim Maxprice, LngRecs
    Dim strRadio
    Dim strCAL
    Dim strFryymm
    Dim strToYymm
    Dim strFrAcctcd
    Dim strToAcctcd    
    Dim strFrAsstno
    Dim strToAsstno    
    Dim strMsg_cd
    Dim strMsg_text
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strRadio		= Request("txtRadio")
    strCAL			= Request("txtCAL")
    strFryymm		= Request("txtFryymm")
    strToYymm		= Request("txtToYymm")
    strFrAcctcd     = FilterVar(Trim(Request("txtFrAcctCd")),"","SNM")
    strToAcctcd		= FilterVar(Trim(Request("txtToAcctCd")),"","SNM")
    strFrAsstno     = FilterVar(Trim(Request("txtFrAsstCd")),"","SNM")
    strToAsstno		= FilterVar(Trim(Request("txtToAsstCd")),"","SNM")

	If strFrAcctcd  = "" Then strFrAcctcd = ""
	If strToAcctcd  = "" Then strToAcctcd = ""
	If strFrAsstno  = "" Then strFrAsstno = ""
	If strToAsstno  = "" Then strToAsstno = ""

	
	With lgObjComm
        .CommandText = "usp_a_compute_depr_asset"
        .CommandType = adCmdStoredProc
        .CommandTimeOut = 0
        
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE"	,adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fryyyymm"		,adVarWChar,adParamInput,6, strFryymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@toyyyymm"		,adVarWChar,adParamInput,6, strToYymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_acct_cd"		,adVarWChar,adParamInput,20, strFrAcctcd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_acct_cd"		,adVarWChar,adParamInput,20, strToAcctcd)	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@fr_asst_no"		,adVarWChar,adParamInput,18, strFrAsstno)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_asst_no"		,adVarWChar,adParamInput,18, strToAsstno)	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@depr_kind"		,adVarWChar,adParamInput,2, strRadio)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@wk_flag"		,adVarWChar,adParamInput,1, strCAL)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"			,adVarWChar,adParamInput,13, gUsrId)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"			,adVarWChar,adParamOutput,6)	    		
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"		,adVarWChar,adParamOutput,256)	    		

        lgObjComm.Execute  ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value       

		If  CInt(IntRetCD) <>  0 Then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value            
            strMsg_text = lgObjComm.Parameters("@msg_text").Value            
            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)           
 
 			IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End If    
    Else           
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End If
End Sub	

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status
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

   If Trim("<%=lgErrorStatus%>") = "NO" Then
      With Parent
           IF  "<%=CInt(intRetCD)%>" >= 0 Then
               .ExeReflectOk
           Else
               .ExeReflectNo
           End If
      End with
   End If   
       
</Script>	


