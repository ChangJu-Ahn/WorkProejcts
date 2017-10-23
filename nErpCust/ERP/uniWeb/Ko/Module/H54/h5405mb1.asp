<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgGetSvrDateTime

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
        
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '¢Ð: "M"(Spread #1) "S"(Spread #2)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection    
    
    lgGetSvrDateTime = GetSvrDateTime
    
        Select Case lgOpModeCRUD
            Case CStr(UID_M0001)                                                         '¢Ð: Query
                 Call SubBizQuery()
            Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
                 Call SubBizSaveMulti()
            Case CStr(UID_M0003)                                                         '¢Ð: Delete
                 Call SubBizDelete()
        End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub	
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strFr_acq_dt
    Dim strTo_acq_dt
    Dim strRprt_dt
    Dim strWhere
    Dim stryymm
    Dim lgSelframeFlg

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    strFr_acq_dt = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(1)),NULL),"NULL","S")
    strTo_acq_dt = FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL),"NULL","S")
    
    lgSelframeFlg = lgKeyStream(4)                                                       'Active Tab 
    stryymm       = FilterVar(UNIConvDateToYYYYMM(lgKeyStream(3),gServerDateFormat,""),"''" ,"S")

    If lgSelframeFlg="1" Then
        lgCurrentSpd = "M"
        strWhere = stryymm
        strWhere = strWhere & " And ( hdf020t.emp_no = hdf070t.emp_no ) "
        strWhere = strWhere & " And ( hdf070t.prov_type = " & FilterVar("1", "''", "S") & "  ) "
        strWhere = strWhere & " And ( IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt) >= " & strFr_acq_dt & ") "
        strWhere = strWhere & " And ( IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt) <= " & strTo_acq_dt & ") "
        Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                              '¡Ù: Make sql statements
    Else
        lgCurrentSpd = "S"
        strWhere = strFr_acq_dt
        strWhere = strWhere & " And ( hdf020t.anut_loss_dt <= " & strTo_acq_dt & ") "
        Call SubMakeSQLStatements("MR",strWhere,"X",">=")                              '¡Ù: Make sql statements
    End If 
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)    
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
            Select Case lgCurrentSpd
               Case "M"
                    lgstrData = lgstrData & Chr(11) & "KI04"
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_cd"))
                    lgstrData = lgstrData & Chr(11) & (String((5-Len(Cstr(lgStrPrevKey+iDx))),"0") & (lgStrPrevKey+iDx))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("anut_no"))
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("pay_tot_amt"), ggAmtOfMoney.DecPoint,0)
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("anut_grade"))
                    lgstrData = lgstrData & Chr(11) & "1"
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("anut_acq_dt"))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Rprt_dt"))
              Case Else
                    lgstrData = lgstrData & Chr(11) & "KI05"
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("comp_cd"))
                    lgstrData = lgstrData & Chr(11) & (String((5-Len(Cstr(lgStrPrevKey+iDx))),"0") & (lgStrPrevKey+iDx))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("anut_no")) 
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & "3"
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("anut_loss_dt"))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Rprt_dt"))
               End Select      
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
               
        Loop 
    End If
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO B_MAJOR("
    lgStrSQL = lgStrSQL & " MAJOR_CD     ," 
    lgStrSQL = lgStrSQL & " MINOR_CD     ," 
    lgStrSQL = lgStrSQL & " MINOR_NM     ," 
    lgStrSQL = lgStrSQL & " MINOR_TYPE   ," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "", "D")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  B_MINOR"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " MAJOR_NM   = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " MINOR_TYPE = " & FilterVar(Trim(UCase(arrColVal(4))), "", "D")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  B_MINOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     MAJOR_CD   = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND MINOR_CD   = " & FilterVar(arrColVal(2),""  ,"S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    
    Select Case Mid(pDataType,1,1)
       Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       If lgCurrentSpd = "M" Then
                          lgStrSQL = "Select top " & iSelCount & " hdf020t.name,  anut_no, hdf070t.pay_tot_amt, hdf020t.anut_grade, " 
                          lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(8), IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt), 12) anut_acq_dt , "
                          lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & " AS comp_cd, " 
                          lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(10), CONVERT(DATETIME, " & FilterVar(lgKeyStream(3), "''", "S") & "),12)" & "  As Rprt_dt "
                          lgStrSQL = lgStrSQL & " From  hdf020t, hdf070t "
                          lgStrSQL = lgStrSQL & " Where hdf070t.pay_yymm " & pComp & pCode
                       Else
                          lgStrSQL = "Select  top " & iSelCount & " hdf020t.name,   anut_no, hdf020t.anut_grade, CONVERT(VARCHAR(8), hdf020t.anut_loss_dt, 12)  anut_loss_dt, " & FilterVar(lgKeyStream(0), "''", "S") & " AS comp_cd, " & " CONVERT(VARCHAR(10), CONVERT(DATETIME," & FilterVar(lgKeyStream(3), "''", "S") & "),12)" & "  As Rprt_dt "
                          lgStrSQL = lgStrSQL & " From  hdf020t   "
                          lgStrSQL = lgStrSQL & " Where anut_loss_dt " & pComp & pCode
                       End If                       
           End Select             
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '¢Ð: Protect system from crashing
    Err.Clear                                                                         '¢Ð: Clear Error status

    Select Case pOpCode
        Case "MC"
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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                If Trim("<%=lgCurrentSpd%>") = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                Else
                   .ggoSpread.Source     = .frm1.vspdData1
                   .lgStrPrevKey1    = "<%=lgStrPrevKey%>"
                End If  
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
