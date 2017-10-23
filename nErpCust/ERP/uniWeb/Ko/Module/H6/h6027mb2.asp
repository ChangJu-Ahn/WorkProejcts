<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
 
	DIM lgGetSvrDateTime
	
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    lgGetSvrDateTime = GetSvrDateTime
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")     '사업장으로 조회 
    iKey1 = iKey1 & " AND prov_yymm   = " & FilterVar(lgKeyStream(1), "''", "S") '지급연월 
    iKey1 = iKey1 & " AND revert_yymm = " & FilterVar(lgKeyStream(2), "''", "S") '귀속연월 
	
    Call SubMakeSQLStatements("R",iKey1)                                     '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
    Else

%>

<Script Language=vbscript>
       With Parent.Frm1
			
			.txt_i_A011.text       = "<%=UNINumClientFormat(lgObjRs("A01_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A012.text       = "<%=UNINumClientFormat(lgObjRs("A01_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A013.text       = "<%=UNINumClientFormat(lgObjRs("A01_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A014.text       = 0
			.txt_i_A015.text       = 0			
	
			.txt_i_A021.text       = "<%=UNINumClientFormat(lgObjRs("A02_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A022.text       = "<%=UNINumClientFormat(lgObjRs("A02_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A023.text       = "<%=UNINumClientFormat(lgObjRs("A02_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A024.text       = 0
			.txt_i_A025.text       = 0
			.txt_i_A207.text       = "<%=UNINumClientFormat(lgObjRs("A20_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A031.text       = "<%=UNINumClientFormat(lgObjRs("A03_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A032.text       = "<%=UNINumClientFormat(lgObjRs("A03_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A033.text       = "<%=UNINumClientFormat(lgObjRs("A03_INCOME"), ggAmtOfMoney.DecPoint,0)%>"

			.txt_i_A041.text       = "<%=UNINumClientFormat(lgObjRs("A04_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A042.text       = "<%=UNINumClientFormat(lgObjRs("A04_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A043.text       = "<%=UNINumClientFormat(lgObjRs("A04_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A044.text       = 0
			.txt_i_A045.text       = 0
									
			.txt_i_A101.text       = "<%=UNINumClientFormat(lgObjRs("NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A102.text       = "<%=UNINumClientFormat(lgObjRs("TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A103.text       = "<%=UNINumClientFormat(lgObjRs("INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A104.text       = 0
			.txt_i_A105.text       = 0
			.txt_i_A106.text       = 0
			.txt_i_A107.text       = "<%=UNINumClientFormat(lgObjRs("INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A108.text       = 0
						
			.txt_i_A991.text       = "<%=UNINumClientFormat(lgObjRs("TOT_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A992.text       = "<%=UNINumClientFormat(lgObjRs("TOT_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A993.text       = "<%=UNINumClientFormat(lgObjRs("TOT_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A997.text       = "<%=UNINumClientFormat(lgObjRs("TOT_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
			
			.txt_i_A201.text       = "<%=UNINumClientFormat(lgObjRs("A20_NUM"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A202.text       = "<%=UNINumClientFormat(lgObjRs("A20_TOT"), ggAmtOfMoney.DecPoint,0)%>"
			.txt_i_A203.text       = "<%=UNINumClientFormat(lgObjRs("A20_INCOME"), ggAmtOfMoney.DecPoint,0)%>"
'******** 마이너스 입력을 받기위해 default로 '0'를 설정 할 수 없어서 셋팅을 함.
			.txt_i_A205.text       = 0
			.txt_i_A206.text       = 0
	
			.txt_i_A251.text       = 0
			.txt_i_A252.text       = 0
			.txt_i_A253.text       = 0
			.txt_i_A255.text       = 0
			
			.txt_i_A261.text       = 0
			.txt_i_A262.text       = 0
			.txt_i_A263.text       = 0
			.txt_i_A264.text       = 0
			.txt_i_A265.text       = 0

			.txt_i_A301.text       = 0
			.txt_i_A302.text       = 0
			.txt_i_A303.text       = 0
			.txt_i_A304.text       = 0
			.txt_i_A305.text       = 0
			.txt_i_A306.text       = 0
			.txt_i_A307.text       = 0
			.txt_i_A308.text       = 0


			.txt_i_A401.text       = 0
			.txt_i_A402.text       = 0
			.txt_i_A403.text       = 0
			.txt_i_A405.text       = 0
			.txt_i_A406.text       = 0
			.txt_i_A407.text       = 0


			.txt_i_A451.text       = 0
			.txt_i_A452.text       = 0
			.txt_i_A453.text       = 0
			.txt_i_A455.text       = 0
			.txt_i_A456.text       = 0
			.txt_i_A457.text       = 0
			
			
			.txt_i_A501.text       = 0
			.txt_i_A502.text       = 0
			.txt_i_A503.text       = 0
			.txt_i_A504.text       = 0
			.txt_i_A505.text       = 0
			.txt_i_A506.text       = 0
			.txt_i_A507.text       = 0
			.txt_i_A508.text       = 0
			
			
			.txt_i_A601.text       = 0
			.txt_i_A602.text       = 0
			.txt_i_A603.text       = 0
			.txt_i_A604.text       = 0
			.txt_i_A605.text       = 0
			.txt_i_A606.text       = 0
			.txt_i_A607.text       = 0
			.txt_i_A608.text       = 0


			.txt_i_A691.text       = 0
			.txt_i_A693.text       = 0
			.txt_i_A694.text       = 0
			.txt_i_A695.text       = 0
			.txt_i_A696.text       = 0
			.txt_i_A697.text       = 0
			.txt_i_A698.text       = 0


			.txt_i_A801.text       = 0
			.txt_i_A802.text       = 0
			.txt_i_A803.text       = 0
			.txt_i_A805.text       = 0
			.txt_i_A806.text       = 0
			.txt_i_A807.text       = 0


			.txt_i_A903.text       = 0
			.txt_i_A904.text       = 0
			.txt_i_A905.text       = 0
			.txt_i_A906.text       = 0
			.txt_i_A907.text       = 0
			.txt_i_A908.text       = 0
			
			.txt_ii_A001.text       = 0
			.txt_ii_A002.text       = 0
			.txt_ii_A003.text       = 0
			.txt_ii_A004.text       = 0
			.txt_ii_A005.text       = 0
			.txt_ii_A006.text       = 0
			.txt_ii_A007.text       = 0
			.txt_ii_A008.text       = 0
			.txt_ii_A009.text       = 0		
       End With          
</Script>       
<%     
    End If
'Response.End    
    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
    
	
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                 Case ""
						lgStrSQL = "SELECT "
						lgStrSQL = lgStrSQL & " A01.A01_NUM, A01.A01_TOT, A01.A01_INCOME, "
						lgStrSQL = lgStrSQL & " A02.A02_NUM, A02.A02_TOT, A02.A02_INCOME, "
						lgStrSQL = lgStrSQL & " A03.A03_NUM, A03.A03_TOT, A03.A03_INCOME, "
						lgStrSQL = lgStrSQL & " case when right(left( " & FilterVar(lgKeyStream(3), "''", "S") & ",6),2) = '02' then A04.A04_NUM else 0 end A04_NUM, "
						lgStrSQL = lgStrSQL & " case when right(left( " & FilterVar(lgKeyStream(3), "''", "S") & ",6),2) = '02' then A04.A04_TOT else 0 end A04_TOT, "	
						lgStrSQL = lgStrSQL & " case when right(left( " & FilterVar(lgKeyStream(3), "''", "S") & ",6),2) = '02' then A04.A04_INCOME else 0 end A04_INCOME, "
						
						lgStrSQL = lgStrSQL & " A20.A20_NUM, A20.A20_TOT, A20.A20_INCOME, "						
						lgStrSQL = lgStrSQL & " SUM(A01.A01_NUM + A02.A02_NUM + A03.A03_NUM + A04.A04_NUM) NUM, "
						lgStrSQL = lgStrSQL & " SUM(A01.A01_TOT + A02.A02_TOT + A03.A03_TOT + A04.A04_TOT) TOT, "
						lgStrSQL = lgStrSQL & " SUM(A01.A01_INCOME + A02.A02_INCOME + A03.A03_INCOME + A04.A04_INCOME) INCOME, "
						lgStrSQL = lgStrSQL & " SUM(A01.A01_NUM + A02.A02_NUM + A03.A03_NUM + A04.A04_NUM + A20.A20_NUM) TOT_NUM, "		'2005-11-08
						lgStrSQL = lgStrSQL & " SUM(A01.A01_TOT + A02.A02_TOT + A03.A03_TOT + A04.A04_TOT + A20.A20_TOT) TOT_TOT, "		'2005-11-08
						lgStrSQL = lgStrSQL & " SUM(A01.A01_INCOME + A02.A02_INCOME + A03.A03_INCOME + A04.A04_INCOME + A20.A20_INCOME ) TOT_INCOME "	'2005-11-08
						
						
						lgStrSQL = lgStrSQL & " FROM "
						lgStrSQL = lgStrSQL & " (SELECT COUNT(distinct hdf070t.emp_no) A01_NUM , "
						'lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.prov_tot_amt, 0)),0) A01_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.prov_tot_amt, 0) + ISNULL(hdf070t.emp_insur,0) + ISNULL(hdf070t.anut,0) + ISNULL(hdf070t.med_insur,0)),0) A01_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.income_tax,0)),0) A01_INCOME "
						lgStrSQL = lgStrSQL & "    FROM hdf070t , haa010t "
						lgStrSQL = lgStrSQL & "   WHERE hdf070t.emp_no = haa010t.emp_no "
						lgStrSQL = lgStrSQL & "     AND (hdf070t.prov_type BETWEEN '1' AND '9' or hdf070t.prov_type = 'Z' ) "
'						lgStrSQL = lgStrSQL & "     AND hdf070t.except_type <> " & FilterVar("4", "''", "S")  
						' 퇴사자 미포함 2007.03.02
						lgStrSQL = lgStrSQL & "     AND (haa010t.retire_dt  is null or ( haa010t.retire_dt < " & FilterVar(lgKeyStream(5), "''", "S") & " and  haa010t.retire_dt > " & FilterVar(lgKeyStream(6), "''", "S") &") ) "
						lgStrSQL = lgStrSQL & "     AND hdf070t.pay_yymm = " & FilterVar(lgKeyStream(2), "''", "S")
						lgStrSQL = lgStrSQL & "     AND haa010t.year_area_cd  = " & FilterVar(lgKeyStream(0), "''", "S")
						lgStrSQL = lgStrSQL & "     AND haa010t.ocpt_type <> " & FilterVar("30", "''", "S") & " "
						lgStrSQL = lgStrSQL & " ) A01 , "
						
						lgStrSQL = lgStrSQL & " (SELECT COUNT(distinct hdf070t.emp_no) A02_NUM , "
						'lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.prov_tot_amt, 0)),0) A02_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.prov_tot_amt, 0) + ISNULL(hdf070t.emp_insur,0) + ISNULL(hdf070t.anut,0) + ISNULL(hdf070t.med_insur,0)),0) A02_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hdf070t.income_tax,0)),0) A02_INCOME "
						lgStrSQL = lgStrSQL & "    FROM hdf070t , haa010t "
						lgStrSQL = lgStrSQL & "   WHERE hdf070t.emp_no = haa010t.emp_no "
						lgStrSQL = lgStrSQL & "     AND (hdf070t.prov_type BETWEEN '1' AND '9' or hdf070t.prov_type = 'Z' ) "
						lgStrSQL = lgStrSQL & "     AND haa010t.retire_dt between " & FilterVar(lgKeyStream(5), "''", "S") & " and "& FilterVar(lgKeyStream(6), "''", "S")  
						lgStrSQL = lgStrSQL & "     AND hdf070t.pay_yymm = " & FilterVar(lgKeyStream(2), "''", "S")
						lgStrSQL = lgStrSQL & "     AND haa010t.year_area_cd  = " & FilterVar(lgKeyStream(0), "''", "S")
						lgStrSQL = lgStrSQL & "     AND haa010t.ocpt_type <> " & FilterVar("30", "''", "S") & " "
						lgStrSQL = lgStrSQL & " ) A02 , "

						lgStrSQL = lgStrSQL & " (SELECT 0 A03_NUM, 0 A03_TOT , 0 A03_INCOME"
						lgStrSQL = lgStrSQL & "  ) A03, "
						
						lgStrSQL = lgStrSQL & " (SELECT COUNT(hfa050t.emp_no) A04_NUM, "
						'lgStrSQL = lgStrSQL & "         ISNULL(SUM(hfa050t.NON_TAX1 + hfa050t.NON_TAX2 + hfa050t.NON_TAX3 + hfa050t.NON_TAX4 + hfa050t.NON_TAX5 + hfa050t.INCOME_TOT_AMT),0) A04_TOT , "
						lgStrSQL = lgStrSQL & "         case when right(left(" & FilterVar(lgKeyStream(2), "''", "S") & ",6),2) = '02' then ISNULL(SUM(hfa050t.NON_TAX1 + hfa050t.NON_TAX2 + hfa050t.NON_TAX3 + hfa050t.NON_TAX4 + hfa050t.NON_TAX5 + hfa050t.NON_TAX6 + hfa050t.INCOME_TOT_AMT + ISNULL(hfa050t.emp_insur,0) + ISNULL(hfa050t.med_insur,0) + ISNULL(hfa050t.national_pension_sub_amt,0)),0) else 0 end A04_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hfa050t.income_tax,0)),0) A04_INCOME"
						lgStrSQL = lgStrSQL & "    FROM hfa050t, haa010t "
						lgStrSQL = lgStrSQL & "   WHERE hfa050t.emp_no = haa010t.emp_no "
						lgStrSQL = lgStrSQL & "     AND hfa050t.year_yy  = left( " & FilterVar(lgKeyStream(2), "''", "S") & ",4)-1 "
						lgStrSQL = lgStrSQL & "     AND right(" & FilterVar(lgKeyStream(2), "''", "S") & ",2) = '01' "
						lgStrSQL = lgStrSQL & "     AND haa010t.year_area_cd  = " & FilterVar(lgKeyStream(0), "''", "S")
						lgStrSQL = lgStrSQL & "  ) A04 ,"
						
						lgStrSQL = lgStrSQL & " (SELECT COUNT(hga070t.emp_no) A20_NUM , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hga070t.tot_prov_amt, 0)),0) A20_TOT , "
						lgStrSQL = lgStrSQL & "         ISNULL(SUM(ISNULL(hga070t.income_tax,0)),0) A20_INCOME "
						lgStrSQL = lgStrSQL & "    FROM haa010t, "
						lgStrSQL = lgStrSQL & "         hga040t, "
						lgStrSQL = lgStrSQL & "         hga070t "
						lgStrSQL = lgStrSQL & "   WHERE haa010t.emp_no = hga040t.emp_no "
						lgStrSQL = lgStrSQL & "     AND hga040t.emp_no = hga070t.emp_no "
						lgStrSQL = lgStrSQL & "     AND hga040t.retire_dt = hga070t.retire_dt "
						lgStrSQL = lgStrSQL & "     AND hga040t.retire_dt between " & FilterVar(lgKeyStream(5), "''", "S") & " and "& FilterVar(lgKeyStream(6), "''", "S")
						lgStrSQL = lgStrSQL & "     AND haa010t.year_area_cd  = " & FilterVar(lgKeyStream(0), "''", "S")
						lgStrSQL = lgStrSQL & "  ) A20 "
						
						lgStrSQL = lgStrSQL & " GROUP BY A01.A01_NUM, A01.A01_TOT, A01.A01_INCOME, "
						lgStrSQL = lgStrSQL & "          A02.A02_NUM, A02.A02_TOT, A02.A02_INCOME, "
						lgStrSQL = lgStrSQL & "          A03.A03_NUM, A03.A03_TOT, A03.A03_INCOME, "
						lgStrSQL = lgStrSQL & "          A04.A04_NUM, A04.A04_TOT, A04.A04_INCOME, "
						lgStrSQL = lgStrSQL & "          A20.A20_NUM, A20.A20_TOT, A20.A20_INCOME "						
'Response.Write lgStrSQL
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "SC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk2
          End If   
    End Select    
       
</Script>
