<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
	Server.ScriptTimeOut = 10000
    Err.Clear                                                                        '¢Ð: Clear Error status
    Dim startDate
    Dim endDate
    Dim prevStartDate
    Dim prevEndDate
    Dim txtMinor
    Dim txtCost
    Dim txtCurrency
    Dim txtRegnm
    Dim dr_amt
    Dim cr_amt

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
 	Const C_SHEETMAXROWS_D  = 100
   '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    dim lginRate
    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '¢Ð: "P"(Prev search) "N"(Next search)

    Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)


	'------ Developer Coding part (Start ) ------------------------------------------------------------------

   	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
           '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0006)
             'Call SubBizSaveMulti()
             ' CALL SubBizSaveMultiDelete()
             Call bulk_disposal()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtGlNo
    Dim iLcNo

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

                                                        '¢Ð : Release RecordSSet
    Call SubBizQueryMulti()

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '¢Ð: Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
    Dim iLoopMax
    dim pKey1
    Dim strWhere
    Dim Currency_code
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    strWhere = FilterVar(lgKeyStream(0), "''", "S")
    
    If Trim(lgKeyStream(1)) <> "" Then
  	    strWhere = strWhere & " and A.BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S") 
        Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtCurrency = ""
		else
		  txtCurrency = Trim(Replace(lgF0,Chr(11),""))
		end if
   Else
		txtCurrency = ""
   End If    
   
   If Trim(lgKeyStream(2)) <> "" Then
  	    strWhere = strWhere & " and a.reg_cd = " & FilterVar(lgKeyStream(2), "''", "S") 
        Call CommonQueryRs("A.MINOR_NM","B_MINOR A, B_MAJOR B","MINOR_TYPE = " & FilterVar("S", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("A1029", "''", "S") & "  AND A.MINOR_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtRegnm = ""
		else
		  txtRegnm = Trim(Replace(lgF0,Chr(11),""))
		end if
   Else
		txtRegnm = ""
   End If    
   
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                   '¡Ù: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found.
        Call SetErrorStatus()

    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF

                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(4), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(5), ggAmtOfMoney.DecPoint, 0)
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
        			lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
     Call SubMakeSQLStatements("MS",strWhere,"X",C_EQ)                                   '¡Ù: Make sql statements   
     If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    	Else
    	     dr_amt = lgObjRs(0)
      End If   
    	
     Call SubMakeSQLStatements("MK",strWhere,"X",C_EQ)                                   '¡Ù: Make sql statements   
     If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    	Else
        		cr_amt = lgObjRs(0)
    	End If   	
        
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If
	
    Call SubHandleError("MR",lgObjRs,Err)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
   
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
   
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
                        lgStrSQL = " select "
						lgStrSQL = lgStrSQL & "       case when a.dr_cr_fg = " & FilterVar("CR", "''", "S") & "  then isnull(sum(a.item_loc_amt),0) else 0 end dr_amt"
						lgStrSQL = lgStrSQL & "       from 	a_monthly_gl_item a,b_acct_dept B"
						lgStrSQL = lgStrSQL & "       where 	a.dept_Cd = b.dept_cd"
						lgStrSQL = lgStrSQL & "       and a.org_change_id =b.org_change_id and a.item_loc_amt <> 0 and a.dr_cr_fg=" & FilterVar("CR", "''", "S") & "  "
						lgStrSQL = lgStrSQL & "       and a.yyyymm " & "=" & pCode
						lgStrSQL = lgStrSQL & "       group by  a.dr_cr_fg"
						
               Case "S"
                        lgStrSQL = " select "
						lgStrSQL = lgStrSQL & "       case when a.dr_cr_fg = " & FilterVar("DR", "''", "S") & "  then isnull(sum(a.item_loc_amt),0) else 0 end dr_amt"
						lgStrSQL = lgStrSQL & "       from 	a_monthly_gl_item a,b_acct_dept B "
						lgStrSQL = lgStrSQL & "       where 	 a.dept_Cd = b.dept_cd"
						lgStrSQL = lgStrSQL & "       and a.org_change_id =b.org_change_id and a.item_loc_amt <> 0 and a.dr_cr_fg=" & FilterVar("DR", "''", "S") & "  "
						lgStrSQL = lgStrSQL & "       and a.yyyymm " & "=" & pCode
						lgStrSQL = lgStrSQL & "       group by  a.dr_cr_fg"
						
               Case "R" '1
                        lgStrSQL = " select Top " &iSelCount& " a.dept_cd, b.dept_nm,"
						lgStrSQL = lgStrSQL & "       a.acct_cd ,"
						lgStrSQL = lgStrSQL & "       c.acct_nm ,"
						lgStrSQL = lgStrSQL & "       case when a.dr_cr_fg = " & FilterVar("DR", "''", "S") & "  then a.item_loc_amt else 0 end dr_amt,"
						lgStrSQL = lgStrSQL & "       case when a.dr_cr_fg = " & FilterVar("CR", "''", "S") & "  then a.item_loc_amt else 0 end cr_amt,"
						lgStrSQL = lgStrSQL & "       d.temp_gl_no ,"
						lgStrSQL = lgStrSQL & "       d.gl_no"
						lgStrSQL = lgStrSQL & "       from 	a_monthly_gl_item a,b_acct_dept B , a_acct c, a_monthly_gl d "
						lgStrSQL = lgStrSQL & "       where 	 a.dept_Cd = b.dept_cd and c.acct_cd = a.acct_cd "
						lgStrSQL = lgStrSQL & "       and a.org_change_id =b.org_change_id and a.item_loc_amt <> 0"
						lgStrSQL = lgStrSQL & "       and a.yyyymm " & "=" & pCode
						lgStrSQL = lgStrSQL & "       and a.yyyymm = d.yyyymm "
						lgStrSQL = lgStrSQL & "       and a.reg_cd = d.reg_cd"
						lgStrSQL = lgStrSQL & "       and a.biz_area_cd = d.biz_area_cd "
						lgStrSQL = lgStrSQL & "       and a.org_change_id= d.org_change_id"
						lgStrSQL = lgStrSQL & "       order by a.seq asc"
				
				End Select
    End Select
	'Response.Write lgStrSQL

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent 
		.frm1.txtRegnm.value = "<%=ConvSPChars(txtRegnm)%>"
		.frm1.txtCurrency.value = "<%=ConvSPChars(txtCurrency)%>"
	END With
</SCRIPT>
<%
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
                .ggoSpread.SSShowData "<%=lgstrData%>"                               '¢Ð : Display data
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"                          '¢Ð : Next next data tag
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .frm1.txtdrAmt.text = "<%=UNINumClientFormat(dr_amt, ggAmtOfMoney.DecPoint, 0)%>"
                .frm1.txtcrAmt.text = "<%=UNINumClientFormat(cr_amt, ggAmtOfMoney.DecPoint, 0)%>"   
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
