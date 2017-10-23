<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

	Dim sFISC_YEAR, sREP_TYPE, sUsrID
	Dim lgStrPrevKey, lgCurrGrid
	Dim lgUserID, lgPassword, lgURL
	
	lgErrorStatus		= "NO"
    lgOpModeCRUD		= Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    sFISC_YEAR			= Request("txtFISC_YEAR")
    sREP_TYPE			= Request("cboREP_TYPE")
	lgStrPrevKey		= UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0002)
             Call SubBizSave()
    End Select

    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
  
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE  TB_ERP_CONFIG WITH (ROWLOCK) "
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf  & vbCrLf 

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " INSERT INTO TB_ERP_CONFIG WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, W1, W2, W3, W4, W5, W6 "  & vbCrLf
    lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) "  & vbCrLf
    lgStrSQL = lgStrSQL & " VALUES ( "  & vbCrLf
        
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S")			& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S")			& "," & vbCrLf
	
    lgStrSQL = lgStrSQL & FilterVar(Request("cboW1"),"''","S") & ", --cboW1" & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(Request("cboW2"),"''","S") & ", --cboW2" & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW3"),"''","S") & ", --txtW3" & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW4"),"''","S") & ", --txtW4" & vbCrLf
    lgStrSQL = lgStrSQL & FilterVar(Request("txtW5"),"''","S") & ", --txtW5" & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtW6"),"''","S") & ", --txtW6" & vbCrLf

    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""
       
    lgStrSQL = lgStrSQL & "   ) " 

	PrintLog "SubBizSaveSingleCreate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_ERP_CONFIG WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       W1 = " & FilterVar(Request("cboW1"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W2 = " & FilterVar(Request("cboW2"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W3 = " & FilterVar(Request("txtW3"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W4 = " & FilterVar(Request("txtW4"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       W5 = " & FilterVar(Request("txtW5"),"''","S") & "," & vbCrLf
	lgStrSQL = lgStrSQL & "       W6 = " & FilterVar(Request("txtW6"),"''","S") & "," & vbCrLf

    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & "," & vbCrLf           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & vbCrLf        
        
	lgStrSQL = lgStrSQL & " WHERE CO_CD = " & FilterVar(Trim(UCase(wgCO_CD)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND FISC_YEAR = " & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") 	 & vbCrLf
	lgStrSQL = lgStrSQL & "		AND REP_TYPE = " & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") 	 & vbCrLf 
	
	PrintLog "SubBizSaveSingleUpdate = " & lgStrSQL
	
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)

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
    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "SD"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "SU"
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
    Select Case "<%=lgOpModeCRUD %>"

       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
  
    End Select    
       
</Script>