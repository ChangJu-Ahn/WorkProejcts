<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    'On Error Resume Next 
    'Err.Clear                                                              '☜: Protect system from crashing
	Call LoadBasisGlobalInf()                                                                       '☜: Clear Error status
	
	Server.ScriptTimeOut = 10000
                                                                      '☜: Clear Error status
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim startDate
    Dim endDate
    Dim prevStartDate
    Dim prevEndDate
    Dim txtRegnm
    Dim amt1,amt2,amt3,amt4
    Dim txtCurrency
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D  = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    dim lginRate
    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)


    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


	'------ Developer Coding part (Start ) ------------------------------------------------------------------

   	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
           '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0006)
             'Call SubBizSaveMulti()
             ' CALL SubBizSaveMultiDelete()
             Call bulk_disposal()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtGlNo
    Dim iLcNo

	'On Error Resume Next
	'Err.Clear

                                                        '☜ : Release RecordSSet
    Call SubBizQueryMulti()

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data
'============================================================================================================
Sub SubBizSave()

	'On Error Resume Next
	'Err.Clear

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere

    On Error Resume Next 
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	strWhere = FilterVar(lgKeyStream(0), "''", "S") 
	If Trim(lgKeyStream(1)) <> "" Then
  	    strWhere = strWhere & "  and b.biz_area_cd LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S")
        Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtCurrency = ""
		else
		  txtCurrency = Trim(Replace(lgF0,Chr(11),""))
		end if
   Else
		txtCurrency = ""
   End If    

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                                 '☆: Make sql statements
    
    
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()

    Else
        
        
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF
        
        
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(5), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(6), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(7), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(8), ggAmtOfMoney.DecPoint, 0)       
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))
                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
		            lgstrData = lgstrData & Chr(11) & Chr(12)

        	    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
     Call SubMakeSQLStatements("MK",strWhere,"X",C_EQ)                                   '☆: Make sql statements   
     If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    	Else
    	     amt1 = lgObjRs(0)
    	     amt2 = lgObjRs(1)
    	     amt3 = lgObjRs(2)
    	     amt4 = lgObjRs(3)
      End If      
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent 
	  .frm1.txtCurrency.value = "<%=ConvSPChars(txtCurrency)%>"
	END With
</SCRIPT>
<%
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
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next 
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next 
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

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
	'데이터 중복 체크 
	lgStrSQL =			  " select  count(DEPT_CD) as Cnt "
	lgStrSQL = lgStrSQL & " from	A_ASSIGN			  "
	lgStrSQL = lgStrSQL & " where	YYYYMM	=	"	& FilterVar(arrColVal(2), "''", "S")
	lgStrSQL = lgStrSQL & " and		DEPT_CD =	"	& FilterVar(arrColVal(3), "''", "S")
	lgStrSQL = lgStrSQL & " and		ACCT_TYPE =	"	& FilterVar(arrColVal(4), "''", "S")
	lgStrSQL = lgStrSQL & " and		BIZ_AREA_CD =	"	& FilterVar(arrColVal(10), "''", "S")
	lgStrSQL = lgStrSQL & " and		ORG_CHANGE_ID =	"	& FilterVar(arrColVal(8), "''", "S")
	lgStrSQL = lgStrSQL & " and		INTERNAL_CD =	"	& FilterVar(arrColVal(9), "''", "S")
													
			
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		If cdbl(lgObjRs("Cnt")) <> 0  Then
			Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)
			Call SetErrorStatus
			Exit Sub
	   End If
	End if
	 
    lgStrSQL = "INSERT INTO A_ASSIGN("
    lgStrSQL = lgStrSQL & " YYYYMM     ,"
    lgStrSQL = lgStrSQL & " DEPT_CD     ,"
    lgStrSQL = lgStrSQL & " ACCT_TYPE    ,"
    lgStrSQL = lgStrSQL & " AMT1    ,"
    lgStrSQL = lgStrSQL & " AMT2         ,"
    lgStrSQL = lgStrSQL & " AMT3,"
    lgStrSQL = lgStrSQL & " org_change_id,"
    lgStrSQL = lgStrSQL & " internal_cd,"
    lgStrSQL = lgStrSQL & " biz_area_cd,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID     ,"
    lgStrSQL = lgStrSQL & " INSRT_DT ,"
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,  "
    lgStrSQL = lgStrSQL & " UPDT_DT    )"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(8)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(9)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(10)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                      & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,Null,"S")
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

    On Error Resume Next 
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  A_ASSIGN"
    lgStrSQL = lgStrSQL & " SET "   
    lgStrSQL = lgStrSQL & " AMT1   	= " & UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & " AMT2 	= " & UNIConvNum(arrColVal(6),0) & ","
    lgStrSQL = lgStrSQL & " AMT3    = " & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID 	= " & FilterVar(gUsrId, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " UPDT_DT    = " & FilterVar(GetSvrDateTime,Null,"S") 
    lgStrSQL = lgStrSQL & " WHERE  "
    lgStrSQL = lgStrSQL & " YYYYMM       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DEPT_CD = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ACCT_TYPE   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
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
    lgStrSQL = "DELETE  A_ASSIGN"
    lgStrSQL = lgStrSQL & " WHERE  "
    lgStrSQL = lgStrSQL & " YYYYMM           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " AND  DEPT_CD     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND  ACCT_TYPE   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
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
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext
                             Case ""

                             Case "P"

                             Case "N"

                        End Select
               Case "D"

               Case "U"

               Case "C"
            end select
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case ""

               Case "K"
					   lgStrSQL = "Select Top " & iSelCount & " sum(B.AMT1),sum(B.AMT2),sum(B.AMT3),sum((B.AMT1-(B.AMT2-B.AMT3)))"
                       lgStrSQL = lgStrSQL & " From  B_ACCT_DEPT A, A_ASSIGN B ,b_major c, b_minor d"
                       lgStrSQL = lgStrSQL & " WHERE A.DEPT_CD = B.DEPT_CD"
                       lgStrSQL = lgStrSQL & "  and  c.major_cd = d.major_cd"
                       lgStrSQL = lgStrSQL & "  and  d.minor_cd = b.acct_type"
                       lgStrSQL = lgStrSQL & "  and  b.org_change_id = a.org_change_id"
                       lgStrSQL = lgStrSQL & "  and  c.major_cd = " & FilterVar("h0071", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  AND  B.YYYYMM "&"=" &pCode
               Case "R" '1
                       lgStrSQL = "Select Top " & iSelCount & " B.DEPT_CD, B.INTERNAL_CD, A.DEPT_NM ,B.ACCT_type,d.MINOR_NM,B.AMT1,B.AMT2,B.AMT3,(B.AMT1-(B.AMT2-B.AMT3)),b.org_change_id, b.biz_area_cd"
                       lgStrSQL = lgStrSQL & " From  B_ACCT_DEPT A, A_ASSIGN B ,b_major c, b_minor d"
                       lgStrSQL = lgStrSQL & " WHERE A.DEPT_CD = B.DEPT_CD"
                       lgStrSQL = lgStrSQL & "  and  c.major_cd = d.major_cd"
                       lgStrSQL = lgStrSQL & "  and  d.minor_cd = b.acct_type"
                       lgStrSQL = lgStrSQL & "  and  b.org_change_id = a.org_change_id"
                       lgStrSQL = lgStrSQL & "  and  c.major_cd = " & FilterVar("h0071", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  AND  B.YYYYMM "&"=" &pCode
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"                          '☜ : Next next data tag
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .frm1.txtAmt1.text = "<%=amt1%>"
                .frm1.txtAmt2.text = "<%=amt2%>"   
                .frm1.txtAmt3.text = "<%=amt3%>"
                .frm1.txtAmt4.text = "<%=amt4%>"   
                .DBQueryOk
	         End with
          End If
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If
       Case "<%=UID_M0006%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOK
          End If
    End Select

</Script>
