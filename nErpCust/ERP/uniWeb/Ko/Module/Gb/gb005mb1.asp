<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GB005MB1
*  4. Program Name         : 경영손익 본사공통비 배부경로 등록 
*  5. Program Desc         : 경영손익 본사공통비 배부경로 등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/28
*  8. Modified date(Last)  : 2001/12/31
*  9. Modifier (First)     : Song Sang Min
* 10. Modifier (Last)      : Lee Tae Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->

<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    Call LoadBasisGlobalInf()                                                                         '☜: Clear Error status
    Dim prevStartDate
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)
                         
    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

'    Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'   lgMaxCount        = UniConvNumStringToDouble(Request("lgMaxCount"),0)                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UniConvNumStringToDouble(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Const C_SHEETMAXROWS_D  = 100                      									
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
	Dim lginRate

    Dim txtMinor
    Dim txtCost

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
           '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0006)                                                         '☜: Delete
             'Call SubBizSaveMulti()
             'CALL SubBizSaveMultiDelete()
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

                                                        '☜ : Release RecordSSet
    Call SubBizQueryMulti()

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

Sub bulk_disposal()
    Dim iLoopMax
    dim pKey1
    Dim idxx
    Dim str
    Dim strWhere_in
    Dim Currency_code
    Dim strWhere_in1
    idxx = 1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    prevStartDate = Trim(lgKeyStream(4))
    
    Currency_code = FilterVar(lgkeyStream(5), "''", "D")

    Call LayerShowHide(1)
    

  ' 조건의 달에 해당하는 데이터가 존재시에 전달과 중복되는 데이터들을 삭제한다.
  '====================================================================================================================
    If Currency_code = "3" Then
        strWhere_in = "and acct_gp <>" & FilterVar("*", "''", "S") & "  AND from_alloc <> " & FilterVar("*", "''", "S") & "  "

        Call CommonQueryRs("count(*)","g_alloc_course","yyyymm = " & FilterVar(lgKeyStream(0), "''", "S") & " and alloc_kinds = " & FilterVar("1", "''", "S") & "  and from_alloc <>" & FilterVar("*", "''", "S") & "  and acct_gp <>" & FilterVar("*", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
            Call SubMakeSQLStatements("MD",strWhere_in,"X",C_EQ)
	        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	            Call SetErrorStatus()
	        End if
	    end if

    ElseIf Currency_code = "1" Then
        strWhere_in = "and acct_gp <>" & FilterVar("*", "''", "S") & "  and from_alloc = " & FilterVar("*", "''", "S") & "  "
        'strWhere_in1 = " and acct_gp <>'*'"
        Call CommonQueryRs("count(*)","g_alloc_course","yyyymm = " & FilterVar(lgKeyStream(0), "''", "S") & " and alloc_kinds = " & FilterVar("1", "''", "S") & "  and from_alloc = " & FilterVar("*", "''", "S") & "  and acct_gp <>" & FilterVar("*", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
            Call SubMakeSQLStatements("MD",strWhere_in,"X",C_EQ)
	        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	            Call SetErrorStatus()
	        End if
	    end if

    Else
        strWhere_in = "and from_alloc <> " & FilterVar("*", "''", "S") & "  and acct_gp = " & FilterVar("*", "''", "S") & " "

        Call CommonQueryRs("count(*)","g_alloc_course","yyyymm = " & FilterVar(lgKeyStream(0), "''", "S") & " and alloc_kinds = " & FilterVar("1", "''", "S") & "  and from_alloc <>" & FilterVar("*", "''", "S") & "  and acct_gp = " & FilterVar("*", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If Trim(Replace(lgF0,Chr(11),"")) <> 0 then
            Call SubMakeSQLStatements("MD",strWhere_in,"X",C_EQ)
	        If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	            Call SetErrorStatus()
	        End if
	    end if
    End If

    Call SubMakeSQLStatements("MB",strWhere_in,"X",C_EQ)                                   '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
         Call SetErrorStatus()

    End If

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim pKey1
    Dim Currency_code
    Dim strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    strWhere = FilterVar(lgKeyStream(0), "''", "S")
    Currency_code = FilterVar(lgkeyStream(5), "''", "D")

   If Trim(lgKeyStream(1)) <> "" Then
  	    strWhere = strWhere & " and gab.acct_gp LIKE  " & FilterVar(Trim(lgKeyStream(1)) & "%", "''", "S") & ""
        Call CommonQueryRs("gp_nm","a_acct_gp","gp_cd = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtMinor = ""
		else
		  txtMinor = Trim(Replace(lgF0,Chr(11),""))
		end if
   else
		txtMinor = ""
   End If

  If Trim(lgKeyStream(3)) <> "" Then
  	    strWhere = strWhere & "  and gab.from_alloc  LIKE  " & FilterVar(Trim(lgKeyStream(3)) & "%", "''", "S") & ""
       Call CommonQueryRs("cost_nm","b_cost_center","cost_cd = "&FilterVar(lgKeyStream(3), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
       if Trim(Replace(lgF0,Chr(11),"")) = "X" then
	       txtCost= ""
	    else
	       txtCost= Trim(Replace(lgF0,Chr(11),""))
	    end if
   else
	     txtCost= ""
   End If

   If Currency_code = "3" Then
       Call SubMakeSQLStatements("MU",strWhere,"X",C_EQ)                                   '☆: Make sql statements
   ElseIf Currency_code = "1" Then
       Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                   '☆: Make sql statements
   Else
       Call SubMakeSQLStatements("MK",strWhere,"X",C_EQ)                                   '☆: Make sql statements
   End If

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
             lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtCodeh.value = "<%=ConvSPChars(txtMinor)%>"
		.frm1.txtCosth.value = "<%=ConvSPChars(txtCost)%>"
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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

   dim strAll_from
   dim strAcct_cd
   Dim str
   Dim strAll_GP

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    str = FilterVar(lgKeyStream(0), "''", "S")

    if arrColVal(4)="" then
      strAll_from = "*"
    else
      strAll_from = arrColVal(4)
    end if

    if arrColVal(5)="" then
      strAcct_cd = "*"
    else
      strAcct_cd = arrColVal(5)
    end if

    if arrColVal(6)="" then
      strAll_GP = "*"
    else
      strAll_GP = arrColVal(6)
    end if

    lgStrSQL = "INSERT INTO G_ALLOC_COURSE("
    lgStrSQL = lgStrSQL & " YYYYMM     ,"
    lgStrSQL = lgStrSQL & " ALLOC_KINDS     ,"
    lgStrSQL = lgStrSQL & " FROM_ALLOC    ,"
    lgStrSQL = lgStrSQL & " ACCT_GP   ,"
    lgStrSQL = lgStrSQL & " ACCT_CD         ,"
    lgStrSQL = lgStrSQL & " TO_ALLOC,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID     ,"
    lgStrSQL = lgStrSQL & " INSRT_DT ,"
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,  "
    lgStrSQL = lgStrSQL & " UPDT_DT    )"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2),"*","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"*","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(strAll_from)),"*","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(strAll_GP)),"*","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(strAcct_cd)),"*","S")     & ","'lgStrSQL = lgStrSQL & strAcct_cd & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(7)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & "getdate(),"
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & "getdate())"

    lginRate = lginRate + UNIConvNum(arrColVal(7),0)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
Dim str
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    str = FilterVar(lgKeyStream(0), "''", "S")
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  G_ALLOC_COURSE"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " TO_ALLOC      = " & FilterVar(UCase(arrColVal(7)), "''", "S")   & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  = " & FilterVar(gUsrId, "''", "S")                & ","
    lgStrSQL = lgStrSQL & " UPDT_DT       = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE  "
    lgStrSQL = lgStrSQL & " YYYYMM   = " & str
    lgStrSQL = lgStrSQL & " AND ALLOC_KINDS  = " & FilterVar("1", "''", "S") & "  "
    lgStrSQL = lgStrSQL & " AND FROM_ALLOC   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ACCT_GP      = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    lgStrSQL = lgStrSQL & " AND ACCT_CD      = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " AND TO_ALLOC     = " & FilterVar(UCase(arrColVal(8)), "''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
  Dim str
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

     str = FilterVar(lgKeyStream(0), "''", "S")
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  G_ALLOC_COURSE"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYYMM        = " & str
    lgStrSQL = lgStrSQL & " and ALLOC_KINDS   = " & FilterVar("1", "''", "S") & "  "
    lgStrSQL = lgStrSQL & " and FROM_ALLOC    = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " and ACCT_GP       = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    lgStrSQL = lgStrSQL & " and ACCT_CD       = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " and TO_ALLOC      = " & FilterVar(UCase(arrColVal(7)), "''", "S")
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

           End Select
        Case "M"
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "C"

               Case "D"
                     lgStrSQL = " delete from g_alloc_course   "
 					 lgStrSQL = lgStrSQL & "   where yyyymm = "&FilterVar(UCase(lgKeyStream(0)), "''", "S")
 					 lgStrSQL = lgStrSQL & "     and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL & "     and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                        from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                        where yyyymm = "&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "	                    and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL  	                        &pCode
					 lgStrSQL = lgStrSQL & "	                    )"
					 lgStrSQL = lgStrSQL & "     and acct_gp  in (select acct_gp"
					 lgStrSQL = lgStrSQL & "                      from  g_alloc_course"
					 lgStrSQL = lgStrSQL & "                      where yyyymm= "&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                      and alloc_kinds = " & FilterVar("1", "''", "S") & " "
					 lgStrSQL = lgStrSQL & "                      and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                         from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                         where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                         and alloc_kinds = " & FilterVar("1", "''", "S") & "   "
					 lgStrSQL = lgStrSQL  	                                         &pCode
					 lgStrSQL = lgStrSQL & "	                                     )"
					 lgStrSQL = lgStrSQL & "     and acct_cd in (select acct_cd "
					 lgStrSQL = lgStrSQL & "                     from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                     where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                     and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL & "                     and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                        from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                        where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                        and alloc_kinds = " & FilterVar("1", "''", "S") & "   "
					 lgStrSQL = lgStrSQL  	                                         &pCode
					 lgStrSQL = lgStrSQL & "	                                     )"
					 lgStrSQL = lgStrSQL & "                     and acct_gp  in (select acct_gp"
					 lgStrSQL = lgStrSQL & "                                       from  g_alloc_course"
					 lgStrSQL = lgStrSQL & "                                       where yyyymm= "&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                       and alloc_kinds = " & FilterVar("1", "''", "S") & " "
					 lgStrSQL = lgStrSQL & "                                       and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                                          from g_alloc_course"
					 lgStrSQL = lgStrSQL & "                                                          where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                                          and alloc_kinds = " & FilterVar("1", "''", "S") & "   "
					 lgStrSQL = lgStrSQL  	                                                          &pCode
					 lgStrSQL = lgStrSQL & "	                                                      )"
					 lgStrSQL = lgStrSQL & "     and to_alloc in (select to_alloc "
					 lgStrSQL = lgStrSQL & "                        from g_alloc_course"
					 lgStrSQL = lgStrSQL & "                        where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                        and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL & "                        and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                           from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                           where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                           and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL  	                                           &pCode
					 lgStrSQL = lgStrSQL & "	                                       )"
					 lgStrSQL = lgStrSQL & "                        and acct_gp  in (select acct_gp"
					 lgStrSQL = lgStrSQL & "                                          from  g_alloc_course"
					 lgStrSQL = lgStrSQL & "                                          where yyyymm= "&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                          and alloc_kinds = " & FilterVar("1", "''", "S") & " "
					 lgStrSQL = lgStrSQL & "                                          and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                                             from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                                             where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                                             and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL  	                                                             &pCode
					 lgStrSQL = lgStrSQL & "	                                                         )"
					 lgStrSQL = lgStrSQL & "                        and acct_Cd in  (select acct_cd "
					 lgStrSQL = lgStrSQL & "                                         from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                         where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                         and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL & "                                         and from_alloc in (select from_alloc "
					 lgStrSQL = lgStrSQL & "                                                            from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                                            where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                                            and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL  	                                                            &pCode
					 lgStrSQL = lgStrSQL & "	                                                        )"
					 lgStrSQL = lgStrSQL & "                                         and acct_gp  in (select acct_gp"
					 lgStrSQL = lgStrSQL & "                                                          from  g_alloc_course"
					 lgStrSQL = lgStrSQL & "                                                          where yyyymm= "&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                                          and alloc_kinds = " & FilterVar("1", "''", "S") & " "
 					 lgStrSQL = lgStrSQL & "                                                          and from_alloc in (select from_alloc "
 					 lgStrSQL = lgStrSQL & "                                                                             from g_alloc_course "
					 lgStrSQL = lgStrSQL & "                                                                             where yyyymm="&FilterVar(UCase(prevStartDate), "''", "S")
					 lgStrSQL = lgStrSQL & "                                                                             and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
					 lgStrSQL = lgStrSQL  	                                                                             &pCode
					 lgStrSQL = lgStrSQL & "	                                                                         ))))))))"

               Case "R" '1
                       lgStrSQL = " Select  top  " & iSelCount&" gab.from_alloc, bcc.cost_nm,gab.acct_gp, b.gp_nm, gab.acct_cd, ac.acct_nm,gab.to_alloc,a.cost_nm"
                       lgStrSQL = lgStrSQL & " From  g_alloc_course  gab,b_cost_center bcc,a_acct ac,( select cost_nm ,cost_cd from b_cost_center)a ,a_acct_gp b"
                       lgStrSQL = lgStrSQL & " WHERE gab.from_alloc *= bcc.cost_cd"
                       lgStrSQL = lgStrSQL & "   and gab.acct_cd     *= ac.acct_cd and gab.acct_gp *= b.gp_cd"
                       lgStrSQL = lgStrSQL & "  and a.cost_cd = gab.to_alloc and gab.acct_gp <> " & FilterVar("*", "''", "S") & "  and gab.from_alloc = " & FilterVar("*", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  and gab.alloc_kinds = " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  AND gab.yyyymm  " & pComp & pCode
                       lgStrSQL = lgStrSQL & "  order by gab.acct_gp "
               Case "U" '3
                       lgStrSQL = " Select  top  " & iSelCount&" gab.from_alloc, bcc.cost_nm,gab.acct_gp, b.gp_nm, gab.acct_cd, ac.acct_nm,gab.to_alloc,a.cost_nm"
                       lgStrSQL = lgStrSQL & " From  g_alloc_course  gab,b_cost_center bcc,a_acct ac,( select cost_nm ,cost_cd from b_cost_center)a ,a_acct_gp b"
                       lgStrSQL = lgStrSQL & " WHERE gab.from_alloc *= bcc.cost_cd"
                       lgStrSQL = lgStrSQL & "   and gab.acct_cd     *= ac.acct_cd and gab.acct_gp *= b.gp_cd"
                       lgStrSQL = lgStrSQL & "  and a.cost_cd = gab.to_alloc and gab.acct_gp <> " & FilterVar("*", "''", "S") & "  and gab.from_alloc <>" & FilterVar("*", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  and gab.alloc_kinds = " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  AND gab.yyyymm  " & pComp & pCode
                       lgStrSQL = lgStrSQL & "  order by gab.from_alloc "
               Case "K" '2
                       lgStrSQL = " Select  top  " & iSelCount&" gab.from_alloc, bcc.cost_nm,gab.acct_gp, b.gp_nm, gab.acct_cd, ac.acct_nm,gab.to_alloc,a.cost_nm"
                       lgStrSQL = lgStrSQL & " From  g_alloc_course  gab,b_cost_center bcc,a_acct ac,( select cost_nm ,cost_cd from b_cost_center)a ,a_acct_gp b"
                       lgStrSQL = lgStrSQL & " WHERE gab.from_alloc *= bcc.cost_cd"
                       lgStrSQL = lgStrSQL & "   and gab.acct_cd     *= ac.acct_cd and gab.acct_gp *= b.gp_cd"
                       lgStrSQL = lgStrSQL & "  and a.cost_cd = gab.to_alloc and gab.from_alloc <>" & FilterVar("*", "''", "S") & "  and gab.acct_gp = " & FilterVar("*", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  and gab.alloc_kinds = " & FilterVar("1", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "  AND gab.yyyymm  " & pComp & pCode
                       lgStrSQL = lgStrSQL & "  order by gab.from_alloc "
               Case "B"
                        lgStrSQL = "INSERT INTO G_ALLOC_COURSE("
						lgStrSQL = lgStrSQL & " YYYYMM,"
						lgStrSQL = lgStrSQL & " ALLOC_KINDS,"
						lgStrSQL = lgStrSQL & " FROM_ALLOC,"
						lgStrSQL = lgStrSQL & " ACCT_GP,"
						lgStrSQL = lgStrSQL & " ACCT_CD,"
						lgStrSQL = lgStrSQL & " TO_ALLOC,"
						lgStrSQL = lgStrSQL & " INSRT_USER_ID     ,"
						lgStrSQL = lgStrSQL & " INSRT_DT ,"
						lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,  "
						lgStrSQL = lgStrSQL & " UPDT_DT    )"
						lgStrSQL = lgStrSQL & " select "&FilterVar(UCase(lgKeyStream(0)), "''", "S")
						lgStrSQL = lgStrSQL & "        ," & FilterVar("1", "''", "S") & " "
						lgStrSQL = lgStrSQL & "        ,from_alloc "
                        lgStrSQL = lgStrSQL & "        ,ACCT_GP "
						lgStrSQL = lgStrSQL & "        ,acct_Cd "
						lgStrSQL = lgStrSQL & "        ,TO_alloc "
						lgStrSQL = lgStrSQL & "        ,"&FilterVar(gUsrId, "''", "S")
						lgStrSQL = lgStrSQL & "        ,getdate()"
						lgStrSQL = lgStrSQL & "        ,"&FilterVar(gUsrId, "''", "S")
						lgStrSQL = lgStrSQL & "        ,getdate()"
						lgStrSQL = lgStrSQL & "from g_alloc_COURSE  "
						lgStrSQL = lgStrSQL & "where yyyymm = "&FilterVar(UCase(prevStartDate), "''", "S")
						lgStrSQL = lgStrSQL & "and alloc_kinds = " & FilterVar("1", "''", "S") & "  "
						lgStrSQL = lgStrSQL & pCode
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData  "<%=lgstrData%>"                             '☜ : Display data
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
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
