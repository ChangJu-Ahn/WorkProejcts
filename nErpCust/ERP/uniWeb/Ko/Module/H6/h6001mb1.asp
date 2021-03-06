<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
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
    Call HideStatusWnd                                                               '��: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
    
    Dim lgSvrDateTime
    lgSvrDateTime = GetSvrDateTime

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strPay_grd1
    Dim strApply_strt_dt
    Dim strWhere
    
    iKey1 = FilterVar(UNIConvDateCompanyToDB(lgKeyStream(0), NULL), "''", "S")                             '������ 
    
    strPay_grd1 = FilterVar(lgKeyStream(1), "''", "S")
    
    strWhere = " apply_strt_dt = (Select Max(hdf010t.apply_strt_dt) From hdf010t "
    strWhere = strWhere & " Where hdf010t.apply_strt_dt <= " & iKey1 & ")  "
    strWhere = strWhere & " And pay_grd1 >= " & strPay_grd1
    strWhere = strWhere & " Order by pay_grd1, pay_grd2 "

    Call SubMakeSQLStatements("MR",strWhere,"X","X")                                 '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '�� : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_grd1_nm") )            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD2"))
            lgstrData = lgstrData & Chr(11) & uniNumClientFormat(lgObjRs("ALLOW1"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & uniNumClientFormat(lgObjRs("ALLOW2"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & uniNumClientFormat(lgObjRs("ALLOW3"), ggAmtOfMoney.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & uniNumClientFormat(CDbl(lgObjRs("ALLOW1"))+CDbl(lgObjRs("ALLOW2"))+CDbl(lgObjRs("ALLOW3")), ggAmtOfMoney.DecPoint,0)

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

            strApply_strt_dt =   UNIConvDateDBToCompany(lgObjRs("apply_strt_dt"),Null)
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   

		    lgObjRs.MoveNext
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   
%>
<Script Language=vbscript>
    If "<%=strApply_strt_dt%>"<>"" Then
        Parent.Frm1.txtStandardDt.Text = "<%=strApply_strt_dt%>"            
    End If     
</Script>       
<%
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
    
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data

'        Response.Write "[" & iDx & "]" & arrColVal(0) & vbCrLf
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
            Case "U"
'					Response.Write "In[" & iDx & "]" & arrColVal(0) & vbCrLf
                    Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
    

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "INSERT INTO Hdf010T         ("
    lgStrSQL = lgStrSQL & " apply_strt_dt         ," 
    lgStrSQL = lgStrSQL & " allow1_cd     ," 
    lgStrSQL = lgStrSQL & " allow2_cd     ," 
    lgStrSQL = lgStrSQL & " allow3_cd     ," 
    lgStrSQL = lgStrSQL & " pay_grd1     ," 
    lgStrSQL = lgStrSQL & " pay_grd2     ," 
    lgStrSQL = lgStrSQL & " allow1     ," 
    lgStrSQL = lgStrSQL & " allow2     ," 
    lgStrSQL = lgStrSQL & " allow3     ," 
    lgStrSQL = lgStrSQL & " Isrt_emp_no         ," 
    lgStrSQL = lgStrSQL & " isrt_dt         ," 
    lgStrSQL = lgStrSQL & " updt_emp_no     ," 
    lgStrSQL = lgStrSQL & " updt_dt         )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(2),NULL),"NULL","S")    & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8), 0)  & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9), 0)  & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10), 0)  & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "UPDATE  Hdf010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " allow1_cd       = " & FilterVar(UCase(arrColVal(3)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " allow2_cd       = " & FilterVar(UCase(arrColVal(4)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " allow3_cd       = " & FilterVar(UCase(arrColVal(5)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " allow1          = " & UNIConvNum(arrColVal(8), 0)   & ","
    lgStrSQL = lgStrSQL & " allow2          = " & UNIConvNum(arrColVal(9), 0)   & ","
    lgStrSQL = lgStrSQL & " allow3          = " & UNIConvNum(arrColVal(10), 0)   & ","

    lgStrSQL = lgStrSQL & " updt_emp_no     = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " updt_dt         = " & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " apply_strt_dt   = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(2),NULL),NULL,"S")
    lgStrSQL = lgStrSQL & " And pay_grd1    = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    lgStrSQL = lgStrSQL & " And pay_grd2    = " & FilterVar(UCase(arrColVal(7)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgStrSQL = "DELETE  hdf010T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " apply_strt_dt   = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(2),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & " And pay_grd1    = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " And pay_grd2    = " & FilterVar(UCase(arrColVal(4)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select  dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",pay_grd1) pay_grd1_nm, " 
                       lgStrSQL = lgStrSQL & " pay_grd1, pay_grd2, allow1_cd, allow2_cd, "
                       lgStrSQL = lgStrSQL & " allow3_cd, apply_strt_dt, allow1, allow2, allow3 "
                       lgStrSQL = lgStrSQL & " From  hdf010t "
                       lgStrSQL = lgStrSQL & " Where " & pCode                       
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

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
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
