<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    Dim emp_no
    Dim name
    Dim dept_nm
	Dim C_Pay_GRD1
	Dim gDecimal_day
	Dim gDecimal_time
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
   'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    ' mode = UID_M0001                                                        			 '☜: Query
    Call SubBizQuery()
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1,iRet
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	C_Pay_GRD1 = ""
	call get_decimal()
    Call SubMakeSQLStatements("1")                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then
%>
        <Script Language=vbscript>
            With parent.frm1
                .name.Value			= "<%=ConvSPChars(lgObjRs("name"))%>"
                .emp_no.Value		= "<%=ConvSPChars(lgObjRs("Emp_no"))%>"
                .grade.Value		= "<%=ConvSPChars(lgObjRs("PAY_GRD1_NM"))%>"
                .dept_cd.Value		= "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"
                .BONUS_RATE.Value	= "<%=UNINumClientFormat(lgObjRs("BONUS_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
                .ADD_RATE.Value		= "<%=UNINumClientFormat(lgObjRs("ADD_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
                .MINUS1_RATE.Value	= "<%=UNINumClientFormat(lgObjRs("MINUS1_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
                .MINUS2_RATE.Value	= "<%=UNINumClientFormat(lgObjRs("MINUS2_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
                .PROV_RATE.Value	= "<%=UNINumClientFormat(lgObjRs("PROV_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
                .BONUS_BAS.Value	= "<%=UNINumClientFormat(lgObjRs("BONUS_BAS"), ggAmtOfMoney.DecPoint, 0)%>"
                .pay_tot.Value		= "<%=UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
                .sub_tot.Value		= "<%=UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
                .real.Value			= "<%=UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint, 0)%>"

				if Cdbl(<%=lgObjRs("SPLENDOR_RATE")%>) <> 0 then
					.SPLENDOR_Text.value = "생산장려율"
					.SPLENDOR_RATE.value = "<%=UNINumClientFormat(lgObjRs("SPLENDOR_RATE"), ggAmtOfMoney.DecPoint, 0)%>"
				end if
            End With          
        </Script>       
<%
	End If
	C_Pay_GRD1 = ConvSPChars(lgObjRs("PAY_GRD1"))
	
    
    Call SubMakeSQLStatements("5")                                       '☜ : Make sql statements
	    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then

	    i = 0
        Do While Not lgObjRs.EOF
%>	       <Script Language=vbscript>
    	        With parent.frm1
      	            .Pay_nm<%=i%>.Value = "<%=ConvSPChars(lgObjRs("ALLOW_NM"))%>"
'        	        .Pay_amt<%=i%>.Value = "<%=FormatNumber(lgObjRs("ALLOW"),0)%>"
        	        .Pay_amt<%=i%>.Value = "<%=UNINumClientFormat(lgObjRs("ALLOW"), ggAmtOfMoney.DecPoint, 0)%>"
          	  End With          
	        </Script>       
<%				
		   lgObjRs.MoveNext
           i = i + 1
        Loop 
	End If

    Call SubMakeSQLStatements("4")                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then

        i = 0
        Do While Not lgObjRs.EOF
%>	       <Script Language=vbscript>
    	        With parent.frm1
     	            .sub_nm<%=i%>.Value = "<%=ConvSPChars(lgObjRs("ALLOW_NM"))%>"
'        	        .sub_amt<%=i%>.Value = "<%=FormatNumber(lgObjRs("SUB_AMT"),0)%>"
        	        .sub_amt<%=i%>.Value = "<%=UNINumClientFormat(lgObjRs("SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
          	  End With          
	        </Script>       
<%				
		    lgObjRs.MoveNext
            i = i + 1
        Loop 
	End If

    Call SubCloseRs(lgObjRs)

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
    Dim strRowBak
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
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
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
'    Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
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

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
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
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
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
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pSection)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case pSection 
			'----------------basic data--------------------
      Case "1"		
      			lsStrSql = ""
   					lgStrSQL = lgStrSQL & "SELECT "
   					lgStrSQL = lgStrSQL & "     a.NAME name, "
   					lgStrSQL = lgStrSQL & "     b.EMP_NO emp_no, "
   					lgStrSQL = lgStrSQL & "     b.PAY_GRD1,"
   					lgStrSQL = lgStrSQL & "     b.DEPT_NM, "
   					lgStrSQL = lgStrSQL & "     b.BONUS_RATE, "
   					lgStrSQL = lgStrSQL & "     b.ADD_RATE, "
   					lgStrSQL = lgStrSQL & "     b.MINUS1_RATE, "
   					lgStrSQL = lgStrSQL & "     b.MINUS2_RATE, "
   					lgStrSQL = lgStrSQL & "     b.PROV_RATE, "
   					lgStrSQL = lgStrSQL & "     b.SPLENDOR_RATE, "
   					lgStrSQL = lgStrSQL & "     b.BONUS_BAS, "
   					lgStrSQL = lgStrSQL & "     b.PROV_TOT_AMT," 
   					lgStrSQL = lgStrSQL & "     b.SUB_TOT_AMT,"
   					lgStrSQL = lgStrSQL & "     b.REAL_PROV_AMT"
   					'lgStrSQL = lgStrSQL & " ,dbo.ufn_GetCodeName('H0001',PAY_GRD1) PAY_GRD1_NM"
   					lgStrSQL = lgStrSQL & " ,dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",b.PAY_GRD1) PAY_GRD1_NM"
   					lgStrSQL = lgStrSQL & " FROM HAA010T a, HDF070T b, HCA090T c"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO LIKE  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PAY_YYMM  =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.EMP_NO  = a.EMP_NO "
   					lgStrSQL = lgStrSQL & "   AND b.EMP_NO  *= c.EMP_NO "
   					lgStrSQL = lgStrSQL & "	  AND c.WK_YYMM =* b.PAY_YYMM"
			'---------------근태내역
      Case "2"
   					lgStrSQL =            " SELECT "
   					lgStrSQL = lgStrSQL & "      a.DILIG_NM,"
   					lgStrSQL = lgStrSQL & "      b.DILIG_HH,"
   					lgStrSQL = lgStrSQL & "      b.DILIG_MM "
   					lgStrSQL = lgStrSQL & " FROM HCA010T a, HCA070T b"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO     =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.DILIG_YYMM =  " & FilterVar(lgKeyStream(1), "''", "S") & "" 
   					lgStrSQL = lgStrSQL & "   AND b.DILIG_CD   = a.DILIG_CD"
   					lgStrSQL = lgStrSQL & "   AND a.DILIG_TYPE=" & FilterVar("2", "''", "S") & ""
			'---------------수당내역
      Case "3"
   					lgStrSQL = "SELECT"
   					lgStrSQL = lgStrSQL & "		b.ALLOW, a.ALLOW_NM"
   					lgStrSQL = lgStrSQL & " FROM"
   					lgStrSQL = lgStrSQL & "		HDA010T a, HDF040T b"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO    =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PAY_YYMM  =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.ALLOW_CD = a.ALLOW_CD"
   					lgStrSQL = lgStrSQL & "   AND a.PAY_CD=" & FilterVar("*", "''", "S") & " "
   					lgStrSQL = lgStrSQL & "   AND a.CODE_TYPE=" & FilterVar("1", "''", "S") & ""
   					lgStrSQL = lgStrSQL & "	order by a.ALLOW_SEQ"
			'---------------공제내역
      Case "4"
   					lgStrSQL = "SELECT"
   					lgStrSQL = lgStrSQL & "		b.SUB_AMT, a.ALLOW_NM"
   					lgStrSQL = lgStrSQL & " FROM"
   					lgStrSQL = lgStrSQL & "		HDA010T a, HDF060T b"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO   =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_YYMM =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_CD = a.ALLOW_CD"
   					lgStrSQL = lgStrSQL & "   AND a.PAY_CD=" & FilterVar("*", "''", "S") & " "
   					lgStrSQL = lgStrSQL & "   AND a.CODE_TYPE=" & FilterVar("2", "''", "S") & ""
   					lgStrSQL = lgStrSQL & "	order by a.ALLOW_SEQ"
			'---------------수당내역 (상여)
      Case "5"
   					lgStrSQL = "SELECT"
   					lgStrSQL = lgStrSQL & "		b.ALLOW, a.ALLOW_NM"
   					lgStrSQL = lgStrSQL & " FROM"
   					lgStrSQL = lgStrSQL & "		HDA010T a, HDF041T b"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO    =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PAY_YYMM  =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.ALLOW_CD = a.ALLOW_CD"
   					lgStrSQL = lgStrSQL & "   AND a.PAY_CD=" & FilterVar("*", "''", "S") & " "
   					lgStrSQL = lgStrSQL & "   AND a.CODE_TYPE=" & FilterVar("1", "''", "S") & ""
   					lgStrSQL = lgStrSQL & "	order by a.ALLOW_SEQ"

   End Select
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
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

sub get_decimal()
	dim intRetCd
	gDecimal_day  = 0
	gDecimal_time = 0
	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("1", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_day  = Trim(Replace(lgF0,Chr(11),""))
	end if

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("2", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_time = Trim(Replace(lgF0,Chr(11),""))
	end if

end sub


%>


