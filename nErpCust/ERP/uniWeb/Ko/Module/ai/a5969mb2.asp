<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
    On Error Resume Next
    Err.Clear

	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")

	Dim lgOption

    Call HideStatusWnd																	'☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus = "NO"
    lgErrorPos    = ""																	'☜: Set to space
    lgOption      = Request("txtMode")													'☜: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)															'☜: Make a DB Connection
	Call SubBizBatch(lgOption)
	Call SubCloseDB(lgObjConn)															'☜: Close DB Connection

'============================================================================================================
' Name : SubBizBatch
' Desc : Date data 
'============================================================================================================
Sub SubBizBatch(iOption)
    Dim Maxprice, LngRecs
	Dim IntRetCD
    Dim gl_date
    Dim insure_cd

	On Error Resume Next																'☜: Protect system from crashing
	Err.Clear																			'☜: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
     
	gl_date =  Trim(Request("txtGlDt"))
	insure_cd = Trim(Request("txtInsuerCd1"))
'//	Call ServerMesgBox("SP전||" & iOption &  "||" & insure_cd &"||" &  gl_date &"||" & gUsrID  , vbInformation, I_MKSCRIPT)
	With lgObjComm
	    .CommandText = "A_USP_A5955BA1_INSURE_POSTING"
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@iOption"   ,adVarWChar,adParamInput, 1, iOption)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@insure_cd"   ,adVarWChar,adParamInput, 20, insure_cd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@temp_gl_date"   ,adVarWChar,adParamInput, 8, gl_date)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",adVarWChar,adParamInput,13, gUsrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",adVarWChar,adParamOutput,6)

		lgObjComm.Execute ,, adExecuteNoRecords
	End With
	'Call ServerMesgBox("SP후||" & iOption &  "||" & insure_cd &"||" &  gl_date &"||" & gUsrID  , vbInformation, I_MKSCRIPT)

	If  Err.number = 0 Then
		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD <> 1 Then
	        strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
			ObjectContext.SetAbort		'//추가 
            Response.End
            Exit Sub
       End If
    Else
	    lgErrorStatus     = "YES"
        ObjectContext.SetAbort
		Call DisplayMsgBox("118114", vbInformation, "", "", I_MKSCRIPT)
       'Call SubHandleError("Batch", lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End If
    
    Call SubCloseCommandObject(lgObjComm)
End Sub	

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"
    ObjectContext.SetAbort
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "batch"
            If CheckSYSTEMError(pErr,True) = True Then
				ObjectContext.SetAbort
				Call SetErrorStatus
            Else
				If CheckSQLError(pConn,True) = True Then
					ObjectContext.SetAbort
					Call SetErrorStatus
				End If
            End If
        Case Else
			If CheckSYSTEMError(pErr,True) = True Then
				ObjectContext.SetAbort
				Call SetErrorStatus
            Else
				If CheckSQLError(pConn,True) = True Then
					ObjectContext.SetAbort
					Call SetErrorStatus
				End If
            End If
	End Select
End Sub
%>
<Script Language="VBScript">
	parent.DBQueryOk
</Script>	
