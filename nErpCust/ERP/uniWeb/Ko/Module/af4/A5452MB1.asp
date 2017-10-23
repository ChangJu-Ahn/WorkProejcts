<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
 	Dim strFrDt	
    Dim strToDt
    DIm strBizAreaCd
    DIm strAcctCd  
	Dim strLoanerFg 
    DIm strLoanerCd

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus  = ""

	Call TrimData()
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubBizQueryMulti()             '
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection


'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
     strFrDt		= FilterVar(UNIConvDate(Request("txtLoanFrDt")), "''", "S")
     strToDt		= FilterVar(UNIConvDate(Request("txtLoanToDt")), "''", "S")
     strBizAreaCd   = FilterVar(Request("txtBizAreaCd"), "''", "S") 
	 strAcctCd      = FilterVar(Request("txtAcctCd"), "''", "S") 
	 strLoanerFg    = FilterVar(Request("txtLoanerFg"), "''", "S") 
	 strLoanerCd    = FilterVar(Request("txtLoanerCd"), "''", "S") 
     
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL    
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    Dim lgStrSel1, lgStrSel2
    Dim lgStrGrpBy
    
    Const C_SHEETMAXROWS_D  = 100													'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 거래처, 차입처 모두 선택 
		lgStrSel1 = ", BIZ_AREA_CD, BIZ_AREA_NM, LOANER, LOANER_NM " & vbCr
		
		lgStrSel2 = ", A.BIZ_AREA_CD, J.BIZ_AREA_NM,   " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.LOAN_BANK_CD ELSE D.BP_CD END  LOANER, " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN E.BANK_NM ELSE F.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy = ", A.BIZ_AREA_CD, J.BIZ_AREA_NM, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.LOAN_BANK_CD ELSE D.BP_CD END, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN E.BANK_NM ELSE F.BP_NM END " & vbCr
		
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "N" Then
	' 거래처 선택	
		lgStrSel1 = ", BIZ_AREA_CD, BIZ_AREA_NM, '', '' " & vbCr
		lgStrSel2 = ", A.BIZ_AREA_CD, J.BIZ_AREA_NM " & vbCr
		lgStrGrpBy = ", A.BIZ_AREA_CD, J.BIZ_AREA_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 차입처 선택 
		lgStrSel1 = ", '', '', LOANER, LOANER_NM " & vbCr
		
		lgStrSel2 = ",			  CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.LOAN_BANK_CD ELSE D.BP_CD END  LOANER, " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN E.BANK_NM ELSE F.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy = ",				CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.LOAN_BANK_CD ELSE D.BP_CD END, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  D.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN E.BANK_NM ELSE F.BP_NM END " & vbCr
	Else 
	' 선택 없음 
		lgStrSel1 = ", '', '', '', '' " & vbCr
		lgStrSel2 = "" & vbCr
		lgStrGrpBy = ""			
	End If 		
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT	PAY_ITEM_ACCT_CD, ACCT_NM, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM_PAY_AMT, SUM_GL_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		ISNULL(SUM_PAY_AMT, 0) - ISNULL(SUM_GL_ITEM_AMT, 0), SUM_TEMP_ITEM_AMT, SUM_BATCH_ITEM_AMT " & vbCr
	lgStrSQL = lgStrSQL & lgStrSel1  & vbCr
	lgStrSQL = lgStrSQL & " FROM (	SELECT B.PAY_ITEM_ACCT_CD, C.ACCT_NM, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(B.PAY_LOC_AMT,0)) SUM_PAY_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_GL_ITEM_AMT,   " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(I.ITEM_LOC_AMT,0)) SUM_TEMP_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(H.ITEM_LOC_AMT,0)) SUM_BATCH_ITEM_AMT " & vbCr
	lgStrSQL = lgStrSQL & lgStrSel2  & vbCr
	lgStrSQL = lgStrSQL & " 	FROM F_LN_REPAY	A	LEFT JOIN F_LN_REPAY_ITEM B ON A.PAY_NO = B.PAY_NO " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN A_ACCT C ON B.PAY_ITEM_ACCT_CD = C.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN F_LN_INFO D ON B.LOAN_NO = D.LOAN_NO " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN B_BANK E ON D.LOAN_BANK_CD = E.BANK_CD " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN B_BIZ_PARTNER F ON D.BP_CD = F.BP_CD " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN (SELECT G1.REF_NO, G2.ITEM_LOC_AMT, G1.GL_DT, G2.ACCT_CD, G1.GL_INPUT_TYPE, G2.ITEM_SEQ " & vbCr
	lgStrSQL = lgStrSQL & " 						FROM A_GL 	G1, A_GL_ITEM	G2 " & vbCr
	lgStrSQL = lgStrSQL & " 						WHERE  G1.GL_NO = G2.GL_NO " & vbCr
	lgStrSQL = lgStrSQL & " 						AND G2.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " & vbCr		
	lgStrSQL = lgStrSQL & " 						AND G1.GL_DT >= " & strFrDt  & " AND G1.GL_DT <= " & strToDt   & vbCr
	lgStrSQL = lgStrSQL & " 					) G ON A.PAY_NO = G.REF_NO AND B.PAY_ITEM_ACCT_CD = G.ACCT_CD AND B.SEQ = G.ITEM_SEQ " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN (SELECT  H1.BATCH_NO,  H1.REF_NO,  H2.JNL_CD, " & vbCr
	lgStrSQL = lgStrSQL & " 							CASE WHEN  H2.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1) * H2.ITEM_LOC_AMT " & vbCr
	lgStrSQL = lgStrSQL & " 							ELSE  H2.ITEM_LOC_AMT END ITEM_LOC_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 							H2.ACCT_CD, H1.GL_INPUT_TYPE, H2.KEY_VAL1 " & vbCr
	lgStrSQL = lgStrSQL & " 						FROM A_BATCH	H1, A_BATCH_GL_ITEM	H2 " & vbCr
	lgStrSQL = lgStrSQL & " 						WHERE  H1.BATCH_NO = H2.BATCH_NO " & vbCr
	lgStrSQL = lgStrSQL & " 						AND H2.JNL_CD IN (SELECT DISTINCT(JNL_CD) " & vbCr
	lgStrSQL = lgStrSQL & " 								FROM A_JNL_ACCT_ASSN 	" & vbCr
	lgStrSQL = lgStrSQL & " 								WHERE JNL_CD IN (" & FilterVar("SL", "''", "S") & " , " & FilterVar("LL", "''", "S") & " , " & FilterVar("SN", "''", "S") & " , " & FilterVar("LN", "''", "S") & " , " & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " , " & FilterVar("CL", "''", "S") & " , " & FilterVar("CN", "''", "S") & " )) " & vbCr
	lgStrSQL = lgStrSQL & " 						AND (H1.GL_INPUT_TYPE = " & FilterVar("LE", "''", "S") & "  or  H1.GL_INPUT_TYPE = " & FilterVar("LM", "''", "S") & " )	" & vbCr
	lgStrSQL = lgStrSQL & " 					) H ON A.PAY_NO = H.REF_NO  AND B.PAY_ITEM_ACCT_CD = H.ACCT_CD AND B.PAY_OBJ = H.JNL_CD AND H.KEY_VAL1 = B.LOAN_NO  " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN (SELECT   I1.REF_NO, I1.TEMP_GL_NO,  I2.ITEM_LOC_AMT, I2.ACCT_CD, I1.GL_INPUT_TYPE, I2.ITEM_SEQ " & vbCr
	lgStrSQL = lgStrSQL & " 						FROM A_TEMP_GL 	I1, A_TEMP_GL_ITEM	I2 " & vbCr
	lgStrSQL = lgStrSQL & " 						WHERE  I1.TEMP_GL_NO = I2.TEMP_GL_NO " & vbCr
	lgStrSQL = lgStrSQL & " 						AND I2.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 						AND I1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 					) I ON  A.PAY_NO = I.REF_NO AND B.PAY_ITEM_ACCT_CD = I.ACCT_CD AND B.SEQ = I.ITEM_SEQ " & vbCr
	lgStrSQL = lgStrSQL & " 					LEFT JOIN B_BIZ_AREA J ON A.BIZ_AREA_CD = J.BIZ_AREA_CD "   & vbCr
	
	lgStrSQL = lgStrSQL & "		WHERE A.PAY_DT >= " & strFrDt  & " AND A.PAY_DT <= " & strToDt   & vbCr
	lgStrSQL = lgStrSQL & "		AND B.PAY_OBJ IN  (" & FilterVar("SL", "''", "S") & " , " & FilterVar("LL", "''", "S") & " , " & FilterVar("SN", "''", "S") & " , " & FilterVar("LN", "''", "S") & " , " & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " , " & FilterVar("CL", "''", "S") & " , " & FilterVar("CN", "''", "S") & " )" & vbCr
	lgStrSQL = lgStrSQL & "		AND ISNULL(B.LOAN_NO,'') <> '' " & vbCr
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.BIZ_AREA_CD = " & strBizAreaCd  & vbCr
	End If
	
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND B.PAY_ITEM_ACCT_CD = " & strAcctCd  & vbCr
	End If
	
	If Trim(Request("txtLoanerFg")) = "BK" Then
		lgStrSQL = lgStrSQL & " AND D.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr		
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND D.LOAN_BANK_CD = " & strLoanerCd  & vbCr
		End If 
		
	ElseIf Trim(Request("txtLoanerFg")) = "BP" Then
		lgStrSQL = lgStrSQL & " AND D.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND D.BP_CD = " & strLoanerCd  & vbCr
		End If 	
	End If		
	
	lgStrSQL = lgStrSQL & " GROUP BY B.PAY_ITEM_ACCT_CD, C.ACCT_NM " & vbCr
	lgStrSQL = lgStrSQL & lgStrGrpBy & ") A" & vbCr
	
	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " WHERE  ISNULL(SUM_PAY_AMT, 0) <> ISNULL(SUM_GL_ITEM_AMT, 0)" & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " ORDER BY PAY_ITEM_ACCT_CD ASC " & vbCr	

    'Response.write lgStrSQL
    'Response.End 
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else 
	
	
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
        
		Do While Not lgObjRs.EOF
		  lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))					'LOAN_ACCT_CD
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))					'ACCT_NM
		  lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(2), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")           'SUM_LOAN_AMT
          lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(3), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_GL_ITEM_AMT
          lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(4), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_TEMP_ITEM_AMT
          lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(5), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_BATCH_ITEM_AMT
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))					'BIZ_AREA_CD
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))                 'BIZ_AREA_NM   
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))					'LOANER
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))					'LOANER_NM      
          lgstrData = lgstrData & Chr(11) & Chr(12)
          
          lgObjRs.MoveNext

          iDx =  iDx + 1
'          If iDx > C_SHEETMAXROWS_D Then
'             Exit Do
'         End If   
		Loop 
	End If    

    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                            " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData1 " & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If
End Sub    

%>

