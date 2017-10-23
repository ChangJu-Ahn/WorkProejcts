<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Dim intRetCD
    Dim lgSvrDateTime

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgOpModeCRUD      = Request("txtMode") 
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    'Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0006)                                                         '¢Ð: Query
			 Call SubBizBatch()			 
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizBatchDelete()
    End Select
         
    Call SubCloseDB(lgObjConn)                                                      '¢Ð: Close DB Connection
    'Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strSend_dt
    Dim strGubun
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	strGubun	     = Request("strGubun")
    strSend_dt   = Request("txtSend_dt")
	
'call svrmsgbox(lgOpModeCRUD &"/"& gUsrId &"/"&  stryear_yymm &"/"& stryear_type & "/"& strallow_cd &"/"& strEmp_no , vbinformation,i_mkscript) 

	Select Case strGubun
		   Case "S"

                '2008-06-27 3:36¿ÀÈÄ :: hanc
				lgStrSQL = "DELETE FROM H_IF_HORG_MAS_KO441"
				lgObjConn.BeginTrans
				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If


				lgStrSQL = "INSERT INTO H_IF_HORG_MAS_KO441"
				lgStrSQL = lgStrSQL & " SELECT DEPT,"
				lgStrSQL = lgStrSQL & "        PDEPT,"
				lgStrSQL = lgStrSQL & "        BUILDID,"
				lgStrSQL = lgStrSQL & "        SDEPTNM,"
				lgStrSQL = lgStrSQL & "        LVL,"
				lgStrSQL = lgStrSQL & "        SEQ,"
				lgStrSQL = lgStrSQL & "        ENDDEPTYN,"
				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S")  & ", "
				lgStrSQL = lgStrSQL & "        ORGID "
				lgStrSQL = lgStrSQL & " FROM   HORG_MAS (NOLOCK)"

				lgObjConn.BeginTrans
				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If

'				lgStrSQL = ""
'				lgStrSQL = "INSERT INTO H_IF_HAA010T_KO441"
'				lgStrSQL = lgStrSQL & " SELECT EMP_NO,"
'				lgStrSQL = lgStrSQL & "        NAME,"
'				lgStrSQL = lgStrSQL & "        ENG_NAME,"
'				lgStrSQL = lgStrSQL & "        DEPT_CD,"
'				lgStrSQL = lgStrSQL & "        DEPT_NM,"
'				lgStrSQL = lgStrSQL & "        RES_NO,"
'				lgStrSQL = lgStrSQL & "        ENTR_DT,"
'				lgStrSQL = lgStrSQL & "        RETIRE_DT,"
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") 
'				lgStrSQL = lgStrSQL & " FROM   HAA010T (NOLOCK) "
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If

'				lgStrSQL = ""
'				lgStrSQL = "INSERT INTO H_IF_HBA010T_KO441 "
'				lgStrSQL = lgStrSQL & " SELECT EMP_NO,"
'				lgStrSQL = lgStrSQL & "        GAZET_DT,"
'				lgStrSQL = lgStrSQL & "        GAZET_CD,"
'				lgStrSQL = lgStrSQL & "        GAZET_RESN,"
'				lgStrSQL = lgStrSQL & "        DEPT_CD,"
'				lgStrSQL = lgStrSQL & "        PAY_GRD1,"
'				lgStrSQL = lgStrSQL & "        PAY_GRD2,"
'				lgStrSQL = lgStrSQL & "        ROLE_CD,"
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") 
'				lgStrSQL = lgStrSQL & " FROM   HBA010T (NOLOCK) "
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If

				lgObjConn.CommitTrans
			Case "D"
				lgStrSQL = "DELETE FROM H_IF_HORG_MAS_KO441"
				lgObjConn.BeginTrans
				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If

				lgStrSQL = "INSERT INTO H_IF_HORG_MAS_KO441"
				lgStrSQL = lgStrSQL & " SELECT DEPT,"
				lgStrSQL = lgStrSQL & "        PDEPT,"
				lgStrSQL = lgStrSQL & "        BUILDID,"
				lgStrSQL = lgStrSQL & "        SDEPTNM,"
				lgStrSQL = lgStrSQL & "        LVL,"
				lgStrSQL = lgStrSQL & "        SEQ,"
				lgStrSQL = lgStrSQL & "        ENDDEPTYN,"
				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S")  & ", "
				lgStrSQL = lgStrSQL & "        ORGID " 
				lgStrSQL = lgStrSQL & " FROM   HORG_MAS (NOLOCK)"

				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If

'				lgStrSQL = "DELETE FROM H_IF_HAA010T_KO441"
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If
'
'				lgStrSQL = "INSERT INTO H_IF_HAA010T_KO441"
'				lgStrSQL = lgStrSQL & " SELECT EMP_NO,"
'				lgStrSQL = lgStrSQL & "        NAME,"
'				lgStrSQL = lgStrSQL & "        ENG_NAME,"
'				lgStrSQL = lgStrSQL & "        DEPT_CD,"
'				lgStrSQL = lgStrSQL & "        DEPT_NM,"
'				lgStrSQL = lgStrSQL & "        RES_NO,"
'				lgStrSQL = lgStrSQL & "        ENTR_DT,"
'				lgStrSQL = lgStrSQL & "        RETIRE_DT,"
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") 
'				lgStrSQL = lgStrSQL & " FROM   HAA010T (NOLOCK)"
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If


'				lgStrSQL = "DELETE FROM H_IF_HBA010T_KO441"
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If
'
'				lgStrSQL = ""
'				lgStrSQL = "INSERT INTO H_IF_HBA010T_KO441 "
'				lgStrSQL = lgStrSQL & " SELECT EMP_NO,"
'				lgStrSQL = lgStrSQL & "        GAZET_DT,"
'				lgStrSQL = lgStrSQL & "        GAZET_CD,"
'				lgStrSQL = lgStrSQL & "        GAZET_RESN,"
'				lgStrSQL = lgStrSQL & "        DEPT_CD,"
'				lgStrSQL = lgStrSQL & "        PAY_GRD1,"
'				lgStrSQL = lgStrSQL & "        PAY_GRD2,"
'				lgStrSQL = lgStrSQL & "        ROLE_CD,"
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") & "," 
'				lgStrSQL = lgStrSQL &			   FilterVar(gUsrId, "''", "S") & ", "
'				lgStrSQL = lgStrSQL &           FilterVar(strSend_dt, "''", "S") 
'				lgStrSQL = lgStrSQL & " FROM   HBA010T (NOLOCK) "
'
'				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'				If lgObjConn.errors.count <> 0 Then
'					lgObjConn.RollbackTrans
'					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
'					Exit Sub
'				End If


				lgObjConn.CommitTrans
	End Select

'    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

        
End Sub	

'============================================================================================================
'''' Name : SubBizBatchDelete
'''' Desc : Delete Data from Db
''''============================================================================================================
'''Sub SubBizBatchDelete()
'''
'''    Dim strpay_yymm
'''    Dim strprov_dt
'''    Dim strProv_cd
'''    Dim strPay_cd
'''    Dim strChk_pay_cd
'''    Dim strMsg_cd
'''    Dim strMsg_text
'''    Dim strEmp_no
'''    On Error Resume Next                                                             '¢Ð: Protect system from crashing
'''    Err.Clear                                                                        '¢Ð: Clear Error status
'''
'''    strpay_yymm = Request("txtpay_yymm_dt")
'''    strprov_dt  = Request("txtprov_dt")
'''    strProv_cd  = Request("txtProv_cd")
'''
'''
'''    strPay_cd = Request("txtPay_cd")
'''    strChk_pay_cd = Request("txtPay_cd")
'''
'''    If  strpay_cd = "" Then
'''        strpay_cd = "%"
'''    else
'''        strpay_cd = strpay_cd
'''    End If
'''
'''    strEmp_no = Request("txtEmp_no")
'''    If  strEmp_no = "" Then
'''        strEmp_no = "%"
'''    End If
'''    
''''call svrmsgbox(strpay_cd & "/" &strchk_pay_cd , vbinformation,i_mkscript)
'''
'''       			   			       
'''    With lgObjComm
'''        .CommandText = "usp_main_pay_delete"
'''        .CommandType = adCmdStoredProc
'''
'''        lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
'''	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adXChar,adParamInput, Len(strPay_yymm), strPay_yymm)
'''	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adXChar,adParamInput, Len(strProv_cd),  strProv_cd)
'''	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_cd"     ,adXChar,adParamInput, Len(strPay_cd),   strPay_cd)
'''	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no"     ,adXChar,adParamInput, Len(strEmp_no),   strEmp_no)
'''	    
'''        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
'''        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)
'''
'''        lgObjComm.Execute ,, adExecuteNoRecords
'''
'''    End With
'''
'''        
'''    If  Err.number = 0 Then
'''        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value        
'''        if  IntRetCD < 0 then
'''            strMsg_cd = Trim(lgObjComm.Parameters("@msg_cd").Value)
'''            strMsg_text = Trim(lgObjComm.Parameters("@msg_text").Value)
'''            
'''            ObjectContext.SetAbort
'''            
'''            Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
'''            IntRetCD = -1
'''            Exit Sub
'''        else
'''            IntRetCD = 1
'''        end if
'''    Else           
'''        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
'''        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
'''        IntRetCD = -1
'''    End if
'''        
'''End Sub	

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
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
               .ExeSendOk
           Else
               .ExeSendNo
           End If
      End with
   End If
       
</Script>	
