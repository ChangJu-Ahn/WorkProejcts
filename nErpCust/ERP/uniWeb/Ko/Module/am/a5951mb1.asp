<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
    On Error Resume Next
    Err.Clear    

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A","NOCOOKIE", "MB")
    Call HideStatusWnd
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet
	Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

    lgPageNo = UNICInt(Trim(Request("lgPageNo")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update             
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
            ' Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, strSql, var1

    'On Error Resume Next                                                             '☜: Protect system from crashing
    'Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    strSql =		  " SELECT distinct MINOR_CD "
    strSql = strSql & " FROM B_MINOR "
    strSql = strSql & " WHERE MAJOR_CD = " & FilterVar("A1029", "''", "S") & "  AND MINOR_CD NOT IN (SELECT MINOR_CD FROM B_MINOR M, A_MONTHLY_BASE T WHERE M.MINOR_CD = T.REG_CD AND M.MAJOR_CD = " & FilterVar("A1029", "''", "S") & " ) "
    strSql = strSql & " AND MINOR_CD IN (" & FilterVar("01", "''", "S") & " , " & FilterVar("02", "''", "S") & " , " & FilterVar("03", "''", "S") & " , " & FilterVar("04", "''", "S") & " ," & FilterVar("05", "''", "S") & " , " & FilterVar("06", "''", "S") & " , " & FilterVar("07", "''", "S") & " , " & FilterVar("08", "''", "S") & " , " & FilterVar("09", "''", "S") & " , " & FilterVar("10", "''", "S") & " ) "
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,strSql,"X","X") = False Then
    
    Else
        Do While Not lgObjRs.EOF
            var1 = lgObjRs("MINOR_CD")
            Call SubBizSaveMultiCreate(var1)
            lgObjRs.MoveNext 
        Loop
    End If
  
    
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQGT)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgPageNo = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgPageNo)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REG_CD"))  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USEYN"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CRATE"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))  
            lgstrData = lgstrData & Chr(11) & ""		 'button
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_NM"))
                        
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgPageNo = lgPageNo + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgPageNo = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To C_SHEETMAXROWS_D
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
           ' Case "C"
                  '  Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update                    
          '  Case "D"
                    'Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
Sub SubBizSaveMultiCreate(var1)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO A_MONTHLY_BASE("
    lgStrSQL = lgStrSQL & " REG_CD     ," 
    lgStrSQL = lgStrSQL & " USE_YN     ," 
    lgStrSQL = lgStrSQL & " RATE       ," 
    lgStrSQL = lgStrSQL & " ACCT_CD    ," 
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(var1, "''", "S")					      & ","
    lgStrSQL = lgStrSQL & "  " & FilterVar("N", "''", "S") & " , "
    lgStrSQL = lgStrSQL & "   0,  "
    lgStrSQL = lgStrSQL & "   '', "
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords      
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  A_MONTHLY_BASE"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " USE_YN			= " & FilterVar(UCase(arrColVal(3)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " RATE			= " & UNIConvNum(arrColVal(4),0)					  & ","   
    lgStrSQL = lgStrSQL & " ACCT_CD			= " & FilterVar(UCase(arrColVal(5)), "''", "S")  & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID    = " & FilterVar(gUsrId, "''", "S")						  & ","    
    lgStrSQL = lgStrSQL & " UPDT_DT		    = "	& FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " REG_CD		    = " & FilterVar(UCase(arrColVal(2)), "''", "S")
	 'Response.Write lgStrSQL
	 'Response.End 
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  B_MAJOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgPageNo + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "INSERT INTO B_MAJOR  .......... " 
               Case "D"
                       lgStrSQL = "DELETE B_MAJOR  .......... " 
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  & " A.REG_CD, B.MINOR_NM, UPPER(A.USE_YN)AS USEYN, A.RATE AS CRATE, A.ACCT_CD AS ACCT_CD "
                       lgStrSQL = lgStrSQL & " ,(SELECT ACCT_NM From A_ACCT WHERE ACCT_CD = A.ACCT_CD) AS ACCT_NM "
                       lgStrSQL = lgStrSQL & " ,A.TRANS_TYPE, C.TRANS_NM "
                       lgStrSQL = lgStrSQL & " From  A_MONTHLY_BASE A, B_MINOR B, A_ACCT_TRANS_TYPE C "
                       lgStrSQL = lgStrSQL & " Where A.REG_CD = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("A1029", "''", "S") & " " 
                       lgStrSQL = lgStrSQL & " AND A.TRANS_TYPE *= C.TRANS_TYPE " 
					   lgStrSQL = lgStrSQL & " Order by A.REG_CD "					   
					   
               Case "U"
                       lgStrSQL = "UPDATE B_MAJOR  .......... " 
           End Select             
           
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgPageNo    = "<%=lgPageNo%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
