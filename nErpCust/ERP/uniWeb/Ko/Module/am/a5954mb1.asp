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
    Call HideStatusWnd
	Const C_SHEETMAXROWS_D  = 100
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    'gUsrId			  = Request("txtUpdtUserId")	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update             
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case CStr(UID_M0005)                                                         '☜: 변동환율 처리 함수 
             Call MoneyRateMove()
        Case CStr(UID_M0006)                                                         '☜: 고정환율 처리 함수 
             Call MoneyRatefix()
             
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : MoneyRateMove
' Desc : Delete DB data
'============================================================================================================

Sub MoneyRateMove()
    Dim iLoopMax
    dim pKey1
    Dim idxx
    Dim str
    Dim strWhere, strWhere1
        
    idxx = 1
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------    
  ' 조건의 년월, REG_CD 에 해당하는 데이터가 존재시에 중복되는 데이터들을 DELETE후 INSERT한다.
  '====================================================================================================================
    strWhere	= FilterVar(lgKeyStream(0), "''", "S")
    strWhere1	= FilterVar(lgKeyStream(2), "''", "S")
	
	
    Call CommonQueryRs("count(*)", "A_EXCHANGE_RATE ",  " yyyymm = " & strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    
    If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
		Call SubMakeSQLStatements("MD",strWhere,strWhere1,C_EQ)                           
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		    Call SetErrorStatus()
		End if
    END IF
    
	Call SubMakeSQLStatements("MG",strWhere,strWhere1,C_EQ)
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	    Call SetErrorStatus()
	End if
			
			    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub


'============================================================================================================
' Name : MoneyRateFix
' Desc : Delete DB data
'============================================================================================================

Sub MoneyRateFix()
    Dim iLoopMax
    dim pKey1
    Dim idxx
    Dim str
    Dim strWhere, strWhere1, IntRetCD
    
    idxx = 1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
  ' 조건의 년월 해당하는 데이터가 존재시에 중복되는 데이터들을 완전 DELETE후 INSERT한다.
  '====================================================================================================================
    
    strWhere = FilterVar(lgKeyStream(0), "''", "S")
    strWhere1 = FilterVar(lgKeyStream(1), "''", "S")

    Call CommonQueryRs("count(*)", "A_EXCHANGE_RATE",  " yyyymm = " & strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       
    If Trim(Replace(lgF0,Chr(11),"")) <> 0 then       
       
        Call SubMakeSQLStatements("MD",strWhere,strWhere1,C_EQ)
	       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		        Call SetErrorStatus()
			End if

	end if
	
	Call SubMakeSQLStatements("MC",strWhere,strWhere1,C_EQ)
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	    Call SetErrorStatus()
	End if
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
  '  Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    strWhere = FilterVar(lgKeyStream(0), "''", "S")
    
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_CUR"))            
            lgstrData = lgstrData & Chr(11) & ""		 'button
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RATE"), ggExchRate.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNIConvYYYYMMDDToDate(gDateFormat,Mid(lgObjRs("YYYYMMDD"),1,4),Mid(lgObjRs("YYYYMMDD"),5,2),Mid(lgObjRs("YYYYMMDD"),7,2))                        
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
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

    On Error Resume Next
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
	lgStrSQL = "DELETE  a_exchange_rate"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYYMM   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DOC_CUR  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
	lgStrSQL = lgStrSQL & vbCr
    lgStrSQL = lgStrSQL & "INSERT INTO a_exchange_rate("
    lgStrSQL = lgStrSQL & " YYYYMM       ," 
    lgStrSQL = lgStrSQL & " DOC_CUR      ," 
    lgStrSQL = lgStrSQL & " XCH_RATE     ," 
    lgStrSQL = lgStrSQL & " YYYYMMDD     ,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4),"0","N")				  & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")	  & ","        
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")			  & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords      
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
  '  Response.Write lgstrsql
   ' Response.end
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next
    Err.Clear
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  a_exchange_rate"
    lgStrSQL = lgStrSQL & " SET "     
    lgStrSQL = lgStrSQL & " XCH_RATE       = " & FilterVar(arrColVal(4),"0","N")		 & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId, "''", "S")			 & ","	
    lgStrSQL = lgStrSQL & " UPDT_DT        = " & FilterVar(GetSvrDateTime,NULL,"S")    
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & " DOC_CUR    = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND YYYYMM  = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  a_exchange_rate"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYYMM   = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DOC_CUR  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
   
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
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "INSERT INTO A_EXCHANGE_RATE  (YYYYMM, DOC_CUR, XCH_RATE, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT) "
					   lgStrSQL = lgStrSQL & " (SELECT "
					   lgStrSQL = lgStrSQL & pCode
					   lgStrSQL = lgStrSQL & " ,A.FROM_CURRENCY, A.STD_RATE, " 
		               lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
					   lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")			     & "," 
					   lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
					   lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
					   lgStrSQL = lgStrSQL & " FROM B_MONTHLY_EXCHANGE_RATE A, B_COMPANY B"
                       lgStrSQL = lgStrSQL & " Where A.APPRL_YRMNTH = " & pCode1  & " AND A.TO_CURRENCY = B.LOC_CUR)"                               
               Case "G"
                       lgStrSQL = "INSERT INTO A_EXCHANGE_RATE  (YYYYMM, DOC_CUR, XCH_RATE, INSRT_USER_ID, INSRT_DT, UPDT_USER_ID, UPDT_DT) "
                       lgStrSQL = lgStrSQL & " (SELECT "
                       lgStrSQL = lgStrSQL & pCode											 & "," 	
                       lgStrSQL = lgStrSQL & " A.FROM_CURRENCY, A.STD_RATE, " 
		               lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
					   lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")			     & "," 
					   lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
					   lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
					   lgStrSQL = lgStrSQL & " FROM b_daily_exchange_rate A, B_COMPANY B"
                       lgStrSQL = lgStrSQL & " Where APPRL_DT = " & pCode1  & " AND A.TO_CURRENCY = B.LOC_CUR)"                       
                        
               Case "D"
                       lgStrSQL = "DELETE A_EXCHANGE_RATE "
                       lgStrSQL = lgStrSQL & " WHERE YYYYMM = " & pCode
 '                     lgStrSQL = lgStrSQL & " AND DOC_CUR IN (SELECT TO_CURRENCY FROM B_MONTHLY_EXCHANGE_RATE WHERE APPRL_YRMNTH = " & pCode1
 '                      lgStrSQL = lgStrSQL & ")"

               Case "I"
                   
                       
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  & " DOC_CUR,  ISNULL(XCH_RATE,0) RATE, YYYYMMDD "
                       lgStrSQL = lgStrSQL & " From  A_EXCHANGE_RATE "
                       lgStrSQL = lgStrSQL & " Where YYYYMM = " & pCode
                       
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
    On Error Resume Next
    Err.Clear

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
               '  If CheckSYSTEMError(pErr,True) = True Then
               '     ObjectContext.SetAbort
               '     Call SetErrorStatus
               '  Else
               '     If CheckSQLError(pConn,True) = True Then
               '        ObjectContext.SetAbort
               '        Call SetErrorStatus
               '     End If
               '  End If
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
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
       Case "<%=UID_M0005%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOK
          End If
       Case "<%=UID_M0006%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOK
          End If
    End Select
</Script>
