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

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD      = Request("txtMode") 
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    'Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0006)                                                         '☜: Query
			 Call SubBizBatch()			 
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizBatchDelete()
    End Select
         
    Call SubCloseDB(lgObjConn)                                                      '☜: Close DB Connection
    'Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim strBasYY, strProvCd
    Dim strGubun
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	strGubun	    = Request("strGubun")
    strBasYY		= Trim(Request("txtBas_yy"))
    strProvCd	= Trim(Request("txtProv_Cd"))
	
'call svrmsgbox(lgOpModeCRUD &"/"& gUsrId &"/"&  stryear_yymm &"/"& stryear_type & "/"& strallow_cd &"/"& strEmp_no , vbinformation,i_mkscript) 

	Select Case strGubun
		   Case "S"

				'*************************************************************************************
				' 지우고 Insert 할지, 바로 Insert 할 지, 결정필요
				'*************************************************************************************
				lgStrSQL = "        DELETE FROM HBA040T "
				lgStrSQL = lgStrSQL & " WHERE EVAL_YY     = " & FilterVar(strBasYY, "''", "S")

				If not IsNull(strProvCd) And strProvCd <> "" Then
					lgStrSQL = lgStrSQL & "   AND EVAL_TYPE = " & FilterVar(strProvCd, "''", "S")
				End If

				lgObjConn.BeginTrans
				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
				'Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If
				'*************************************************************************************

				lgStrSQL = ""
				lgStrSQL = lgStrSQL & " INSERT INTO HBA040T "
				lgStrSQL = lgStrSQL & " 		                  (EMP_NO,EVAL_YY,EVAL_TYPE,VALUE_GRADE,VALUE_SCORE,VALUE_EMP_NO,TOT_VALU, "
				lgStrSQL = lgStrSQL & "                        ISRT_EMP_NO,ISRT_DT,UPDT_EMP_NO,UPDT_DT) "
				lgStrSQL = lgStrSQL & "          SELECT EMP_NO,EVAL_YY,EVAL_TYPE,VALUE_GRADE,VALUE_SCORE,VALUE_EMP_NO,TOT_VALU, "
				lgStrSQL = lgStrSQL & "                        ISRT_USER_ID,ISRT_DT,UPDT_USER_ID,UPDT_DT "
				lgStrSQL = lgStrSQL & "             FROM NEPES_TEST.DBO.H_IF_HBA040T_KO441 "		'[DB] : NEPES_TEST --> inbus
				lgStrSQL = lgStrSQL & "          WHERE EVAL_YY     = " & FilterVar(strBasYY, "''", "S")

				If not IsNull(strProvCd) And strProvCd <> "" Then
					lgStrSQL = lgStrSQL & "            AND EVAL_TYPE = " & FilterVar(strProvCd, "''", "S")
				End If

				lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

				If lgObjConn.errors.count <> 0 Then
					lgObjConn.RollbackTrans
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					Exit Sub
				End If

				lgObjConn.CommitTrans

	End Select

'    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
'	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

        
End Sub	

'============================================================================================================
' Name : SubBizBatchDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizBatchDelete()

    Dim strBasYY, strProvCd
    Dim strGubun
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	strGubun	    = Request("strGubun")
    strBasYY		= Trim(Request("txtBas_yy"))
    strProvCd	= Trim(Request("txtProv_Cd"))

	Select Case strGubun
		Case "D"
			lgStrSQL = "        DELETE FROM HBA040T "
			lgStrSQL = lgStrSQL & " WHERE EVAL_YY     = " & FilterVar(strBasYY, "''", "S")

			If not IsNull(strProvCd) And strProvCd <> "" Then
				lgStrSQL = lgStrSQL & "   AND EVAL_TYPE = " & FilterVar(strProvCd, "''", "S")
			End If

			lgObjConn.BeginTrans
			lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
			Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
			If lgObjConn.errors.count <> 0 Then
				lgObjConn.RollbackTrans
				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				Exit Sub
			End If

			lgObjConn.CommitTrans
	End Select
        
End Sub	

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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
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
