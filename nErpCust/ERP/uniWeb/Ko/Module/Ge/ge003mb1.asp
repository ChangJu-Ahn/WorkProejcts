<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")     
                                                                           '☜: Clear Error status
	Dim lgstrDataTotal, data_amt
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = UniConvNumStringToDouble(Request("lgMaxCount"),0)                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UniConvNumStringToDouble(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
   	Const C_SHEETMAXROWS_D  = 100  
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time	                 


	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	MSGBOX(strVal)
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
'             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()
             CALL SubBizSaveMultiDelete()
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
    Dim iKey1
    Dim strWhere
    Dim txtCurrency
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    strWhere = strWhere & FilterVar(lgKeyStream(0), "''", "S")

    If Trim(lgKeyStream(1)) <> "" Then
    	strWhere = strWhere & "  and  g.alloc_base = " & FilterVar(lgKeyStream(1), "''", "S")

        Call CommonQueryRs(" b.minor_nm "," b_configuration a, b_minor b "," a.minor_cd = b.minor_cd and a.major_cd = b.major_cd and  a.seq_no =5  and reference = " & FilterVar("Y", "''", "S") & "  and b.major_cd = " & FilterVar("G1004", "''", "S") & " and a.minor_cd = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
            txtCurrencyCode = ""
            txtCurrency = ""
			txtYn = ""
           Call DisplayMsgBox("GB0701", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.		   	

        else
			txtCurrencyCode =  Trim(lgKeyStream(1))        
            txtCurrency = Trim(Replace(lgF0,Chr(11),""))

	        Call CommonQueryRs("reference","b_configuration", " seq_no =1  and major_cd = " & FilterVar("G1004", "''", "S") & " and minor_cd = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				if Trim(Replace(lgF0,Chr(11),"")) = "X" then
					txtYntag = ""
				else
					txtYntag = Trim(Replace(lgF0,Chr(11),""))
				end if

		        if  txtYntag = UCase("Y") then
				    txtYn    = "자동생성"
				else
					txtYn    = "수작업입력"
				end if
        
         end if
        
	else
        txtCurrency = ""
    End If

    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus()
    Else


        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
         lgstrDataTotal = 0
        iDx       = 1

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("to_alloc")) 'cost center
            lgstrData = lgstrData & Chr(11) & ""		 'button
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm")) 'cost center 명 
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A"), ggAmtOfMoney.DecPoint, 0) '배부 data
            lgstrDataTotal = CLng(lgstrDataTotal) + CLng(lgObjRs("A"))
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If

        Loop
        
          Call SubMakeSQLStatements("MS",strWhere,"X",C_EQ)                                   '☆: Make sql statements

    	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    			data_amt = 0
    	Else
        		data_amt = lgObjRs(0)        		
    	End If
        
        
        
    End If

    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
%>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtCurrencyCode.value = "<%=ConvSPChars(txtCurrencyCode)%>"		
		.frm1.txtCurrency.value = "<%=ConvSPChars(txtCurrency)%>"
		.frm1.txtYntag.value = "<%=ConvSPChars(txtYntag)%>"
		.frm1.txtYn.value = "<%=ConvSPChars(txtYn)%>"
	END With
	Call Parent.spreadflag()						' toolbar를 조정하기 위한 func
</SCRIPT>
<%
End Sub




'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx, strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status


  strWhere = strWhere & FilterVar(lgKeyStream(0), "''", "S")

    If Trim(lgKeyStream(1)) <> "" Then
    	strWhere = strWhere & "  and  g.alloc_base = " & FilterVar(lgKeyStream(1), "''", "S")

        Call CommonQueryRs(" b.minor_nm "," b_configuration a, b_minor b "," a.minor_cd = b.minor_cd and a.major_cd = b.major_cd and  a.seq_no =5  and reference = " & FilterVar("Y", "''", "S") & "  and b.major_cd = " & FilterVar("G1004", "''", "S") & " and a.minor_cd = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
            txtCurrencyCode = ""
            txtCurrency = ""
			txtYn = ""
         Exit Sub
        
         end if
    End If




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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO G_ALLOC_DATA("
    lgStrSQL = lgStrSQL & " YYYYMM        ,"
    lgStrSQL = lgStrSQL & " ALLOC_KINDS   ,"
    lgStrSQL = lgStrSQL & " ALLOC_BASE    ,"
    lgStrSQL = lgStrSQL & " TO_ALLOC      ,"
    lgStrSQL = lgStrSQL & " TO_DATA       ,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID ,"
    lgStrSQL = lgStrSQL & " INSRT_DT	  ,"
    lgStrSQL = lgStrSQL & " UPDT_USER_ID  ,"
    lgStrSQL = lgStrSQL & " UPDT_DT    )"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")  & ","
    lgStrSQL = lgStrSQL & "" & FilterVar("4", "''", "S") & ", "
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & Replace(UNIConvNum(arrColVal(4),0),",","") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & "getdate(),"
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & ","
    lgStrSQL = lgStrSQL & "getdate())"



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
    lgStrSQL = "UPDATE  G_ALLOC_DATA"
    lgStrSQL = lgStrSQL & " SET "
    lgStrSQL = lgStrSQL & " TO_DATA    = " & Replace(UNIConvNum(arrColVal(5),0),",","")  & ","	
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")				  & ","	
    lgStrSQL = lgStrSQL & " UPDT_DT      = getdate()" 
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYYMM     = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND ALLOC_KINDS = " & FilterVar("4", "''", "S") & " "
    lgStrSQL = lgStrSQL & " AND ALLOC_BASE  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND TO_ALLOC    = " & FilterVar(UCase(arrColVal(4)), "''", "S")

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
    lgStrSQL = "DELETE  G_ALLOC_DATA"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYYMM			  = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " and ALLOC_KINDS   = " & FilterVar("4", "''", "S") & " "
    lgStrSQL = lgStrSQL & " and ALLOC_BASE    = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " and TO_ALLOC      = " & FilterVar(UCase(arrColVal(4)), "''", "S")



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

           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1

           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "INSERT INTO B_MAJOR  .......... "
               Case "D"
                       lgStrSQL = "DELETE B_MAJOR  .......... "
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  & " g.to_alloc, b.item_nm, IsNull(SUM(g.to_data),0) as A"
                       lgStrSQL = lgStrSQL & " From  g_alloc_data g, b_item b"
                       lgStrSQL = lgStrSQL & " Where  g.to_alloc = b.item_cd "
                       lgStrSQL = lgStrSQL & " and  g.alloc_kinds = " & FilterVar(4, "''", "S")
                       lgStrSQL = lgStrSQL & " and  g.yyyymm = " & pCode
                       lgStrSQL = lgStrSQL & " group by g.to_alloc, b.item_nm"
                       lgStrSQL = lgStrSQL & " order by g.to_alloc "
						
               Case "S"
                       lgStrSQL = "Select IsNull(SUM(g.to_data),0) as A"
                       lgStrSQL = lgStrSQL & " From  g_alloc_data g, b_item b"
                       lgStrSQL = lgStrSQL & " Where  g.to_alloc = b.item_cd "
                       lgStrSQL = lgStrSQL & " and  g.alloc_kinds = " & FilterVar(4, "''", "S")
                       lgStrSQL = lgStrSQL & " and  g.yyyymm = " & pCode
                       
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

		  Parent.frm1.txtDataAmt.text = "<%=UNINumClientFormat(data_amt, ggAmtOfMoney.DecPoint, 0)%>"	
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
    End Select


</Script>
