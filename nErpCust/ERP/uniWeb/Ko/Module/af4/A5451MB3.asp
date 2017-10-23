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
	DIm strLoanerFg
    DIm strLoanerCd
    Dim strInputType
    
    Dim lgStrPrevKey
    
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus  = ""
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))									 '☜: Next Key    

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
    Dim lgStrSQL,lgStrSQL2
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    Dim lgstrFrDt,lgstrToDt
	Dim lgTotLoanLocAmt3 , lgTotDiffLocAmt3 , lgTotGlLocAmt3 , lgTotTempGlLocAmt3
    
    Dim lgStrSel1, lgStrSel2
    Dim lgStrGrpBy
    Dim lgMaxCount
    
    Const C_SHEETMAXROWS_D = 100													'☆: Server에서 한번에 fetch할 최대 데이타 건수 
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time

   	If Len(Trim(Request("lgStrPrevKey")))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
           lgStrPrevKey = CInt(lgStrPrevKey)          
        End If   
    Else   
        lgStrPrevKey = 0
    End If

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status   
    
    If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 거래처, 차입처 모두 선택 
		lgStrSel1 = ", BIZ_AREA_CD, BIZ_AREA_NM, LOANER, LOANER_NM " & vbCr
		
		lgStrSel2 = ", A.BIZ_AREA_CD, I.BIZ_AREA_NM,   " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN A.LOAN_BANK_CD ELSE A.BP_CD END  LOANER, " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN C.BANK_NM ELSE D.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy = ", A.BIZ_AREA_CD, I.BIZ_AREA_NM, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN A.LOAN_BANK_CD ELSE A.BP_CD END, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN C.BANK_NM ELSE D.BP_NM END " & vbCr
		
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "N" Then
	' 거래처 선택	
		lgStrSel1 = ", BIZ_AREA_CD, BIZ_AREA_NM, '', '' " & vbCr
		lgStrSel2 = ", A.BIZ_AREA_CD, I.BIZ_AREA_NM " & vbCr
		lgStrGrpBy = ", A.BIZ_AREA_CD, I.BIZ_AREA_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 차입처 선택 
		lgStrSel1 = ", '', '', LOANER, LOANER_NM " & vbCr
		
		lgStrSel2 = ",			  CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN A.LOAN_BANK_CD ELSE A.BP_CD END  LOANER, " & vbCr
		lgStrSel2 = lgStrSel2 & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN C.BANK_NM ELSE D.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy = ",				CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN A.LOAN_BANK_CD ELSE A.BP_CD END, " & vbCr
		lgStrGrpBy = lgStrGrpBy & " CASE WHEN  A.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN C.BANK_NM ELSE D.BP_NM END " & vbCr
	Else 
	' 선택 없음 
		lgStrSel1 = ", '', '', '', '' " & vbCr
		lgStrSel2 = "" & vbCr
		lgStrGrpBy = ""			
	End If 	

	If lgStrPrevKey  =0 Then
		'----
		lgStrSQL = ""
		lgStrSQL = lgStrSQL & " SELECT " & vbCr
		lgStrSQL = lgStrSQL & " SUM(ISNULL(SUM_LOAN_AMT,0)) ALLC_TOT_LOC_AMT, SUM(ISNULL(SUM_GL_ITEM_AMT,0)) GL_TOT_ITEM_LOC_AMT, SUM(ISNULL(SUM_DIFF_AMT,0)) Diff_TOT_ITEM_AMT, SUM(ISNULL(SUM_TEMP_ITEM_AMT,0)) GL_TEMP_TOT_ITEM_LOC_AMT " & vbCr
		lgStrSQL = lgStrSQL & "  FROM (" & vbCr
		lgStrSQL = lgStrSQL & " 		SELECT A.LOAN_ACCT_CD, B.ACCT_NM, A.LOAN_NO, A.LOAN_DT, J.GL_DT, " & vbCr
		lgStrSQL = lgStrSQL & " 		SUM(ISNULL(A.LOAN_LOC_AMT,0)) - SUM(ISNULL(E.ITEM_LOC_AMT,0))  SUM_DIFF_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 		SUM(ISNULL(A.LOAN_LOC_AMT,0)) SUM_LOAN_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 		SUM(ISNULL(E.ITEM_LOC_AMT,0)) SUM_GL_ITEM_AMT, " & vbCr
		lgStrSQL = lgStrSQL & " 		SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_TEMP_ITEM_AMT  " & vbCr
		lgStrSQL = lgStrSQL & " 	FROM F_LN_INFO 	A	LEFT JOIN A_ACCT B ON A.LOAN_ACCT_CD = B.ACCT_CD " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BANK C ON A.LOAN_BANK_CD = C.BANK_CD " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BIZ_PARTNER D ON A.BP_CD = D.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT E1.REF_NO, E2.ITEM_LOC_AMT, E1.GL_NO, E1.GL_DT, E2.ACCT_CD, E1.GL_INPUT_TYPE " & vbCr
		lgStrSQL = lgStrSQL & " 								FROM A_GL 	E1, A_GL_ITEM	E2 " & vbCr
		lgStrSQL = lgStrSQL & " 								WHERE  E1.GL_NO = E2.GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 								AND E2.DR_CR_FG = " & FilterVar("CR", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 								AND E1.GL_DT >= " & strFrDt  & " AND E1.GL_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 						) E ON A.LOAN_NO = E.REF_NO AND A.LOAN_ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT  F1.BATCH_NO,  F1.REF_NO,  F2.JNL_CD, " & vbCr
		lgStrSQL = lgStrSQL & " 										CASE WHEN  F2.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1) * F2.ITEM_LOC_AMT " & vbCr
		lgStrSQL = lgStrSQL & " 										ELSE  F2.ITEM_LOC_AMT END ITEM_LOC_AMT, F2.ACCT_CD, F1.GL_INPUT_TYPE " & vbCr
		lgStrSQL = lgStrSQL & " 									FROM A_BATCH	F1, A_BATCH_GL_ITEM	F2 " & vbCr
		lgStrSQL = lgStrSQL & " 				 					WHERE  F1.BATCH_NO = F2.BATCH_NO   " & vbCr
		lgStrSQL = lgStrSQL & " 			 						AND F2.JNL_CD IN (SELECT DISTINCT(JNL_CD) " & vbCr
		lgStrSQL = lgStrSQL & " 				        		                    FROM A_JNL_ACCT_ASSN " & vbCr
		lgStrSQL = lgStrSQL & " 													WHERE JNL_CD IN (" & FilterVar("SL", "''", "S") & " , " & FilterVar("LL", "''", "S") & " , " & FilterVar("SN", "''", "S") & " , " & FilterVar("LN", "''", "S") & " , " & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " )) " & vbCr
		lgStrSQL = lgStrSQL & " 									AND (F1.GL_INPUT_TYPE = " & FilterVar("LN", "''", "S") & "  or  (F1.GL_INPUT_TYPE = " & FilterVar("LO", "''", "S") & "  and  F2.JNL_CD IN (" & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " ))) " & vbCr
		lgStrSQL = lgStrSQL & " 									AND F1.GL_INPUT_TYPE IN ( SELECT DISTINCT(MINOR_CD) " & vbCr
		lgStrSQL = lgStrSQL & " 															   FROM B_MINOR " & vbCr
		lgStrSQL = lgStrSQL & " 																WHERE MINOR_CD IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LO", "''", "S") & " ) AND MAJOR_CD = " & FilterVar("A1001", "''", "S") & " ) " & vbCr
		lgStrSQL = lgStrSQL & " 						) F ON A.LOAN_NO = F.REF_NO AND A.LOAN_ACCT_CD = F.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT   G1.REF_NO, G1.TEMP_GL_NO,  G1.TEMP_GL_DT, G2.ITEM_LOC_AMT, G2.ACCT_CD, G1.GL_INPUT_TYPE " & vbCr
		lgStrSQL = lgStrSQL & " 									FROM A_TEMP_GL 	G1, A_TEMP_GL_ITEM	G2 " & vbCr
		lgStrSQL = lgStrSQL & " 									WHERE  G1.TEMP_GL_NO = G2.TEMP_GL_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 									AND G2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 									AND G1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 						) G ON A.LOAN_NO = G.REF_NO AND A.LOAN_ACCT_CD = G.ACCT_CD " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN B_MINOR H ON ((H.MINOR_CD = E.GL_INPUT_TYPE OR  H.MINOR_CD = G.GL_INPUT_TYPE ) AND MAJOR_CD = " & FilterVar("A1001", "''", "S") & " ) " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BIZ_AREA I ON A.BIZ_AREA_CD = I.BIZ_AREA_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN A_GL J ON A.GL_NO = J.GL_NO  " & vbCr
		lgStrSQL = lgStrSQL & "		WHERE  A.LOAN_DT >= " & strFrDt  & " AND A.LOAN_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & "		AND A.LOAN_BASIC_FG IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LR", "''", "S") & " ) " & vbCr
	
		If Trim(Request("txtBizAreaCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND A.BIZ_AREA_CD = " & strBizAreaCd  & vbCr
		End If
	
		If Trim(Request("txtAcctCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND A.LOAN_ACCT_CD = " & strAcctCd  & vbCr
		End If

		If Trim(Request("txtInputType")) <> "" Then
			lgStrSQL = lgStrSQL & " AND H.MINOR_CD = " & FilterVar(Request("txtInputType"), "''", "S")  & vbCr
		Else 
			lgStrSQL = lgStrSQL & " AND H.MINOR_CD IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LO", "''", "S") & " ) " & vbCr
		End If		
		lgStrSQL = lgStrSQL & " AND H.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
	
	
		If Trim(Request("txtLoanerFg")) = "BK" Then
			lgStrSQL = lgStrSQL & " AND A.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr
			If Trim(Request("txtLoanerCd")) <> "" Then 
				lgStrSQL = lgStrSQL & " AND A.LOAN_BANK_CD = " & strLoanerCd  & vbCr
			End If 
			
		ElseIf Trim(Request("txtLoanerFg")) = "BP" Then
			lgStrSQL = lgStrSQL & " AND A.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr
			If Trim(Request("txtLoanerCd")) <> "" Then 
				lgStrSQL = lgStrSQL & " AND A.BP_CD = " & strLoanerCd  & vbCr
			End If 	
		End If		
	
		lgStrSQL = lgStrSQL & " 	GROUP BY A.LOAN_ACCT_CD, B.ACCT_NM, A.LOAN_NO, A.LOAN_DT, J.GL_DT, " & vbCr
		lgStrSQL = lgStrSQL & " 	 E.GL_NO, G.TEMP_GL_NO,  F.BATCH_NO, F.GL_INPUT_TYPE, H.MINOR_NM " & vbCr
		lgStrSQL = lgStrSQL & lgStrGrpBy   & vbCr
		lgStrSQL = lgStrSQL & " 	,G.TEMP_GL_DT) A " & vbCr

		If UCase(Trim(Request("DispMeth"))) Then 
			lgStrSQL = lgStrSQL & " WHERE  ISNULL(SUM_LOAN_AMT, 0) <> ISNULL(SUM_GL_ITEM_AMT, 0)" & vbCr
		End If 	
	

	
	'----
	'*********************************
	'			합계찍기 
	'*********************************
							

	
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
			lgTotLoanLocAmt3 = UNIConvNumDBToCompanyByCurrency(lgObjRs("ALLC_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotDiffLocAmt3 = UNIConvNumDBToCompanyByCurrency(lgObjRs("Diff_TOT_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotGlLocAmt3  = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TOT_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotTempGlLocAmt3 = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TEMP_TOT_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotLoanLocAmt3.text     =  """ & lgTotLoanLocAmt3    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & lgTotDiffLocAmt3 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & lgTotGlLocAmt3    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & lgTotTempGlLocAmt3   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
	
		Else
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotLoanLocAmt3.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & 0 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & 0   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
		End If    
	End If












	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT  LOAN_ACCT_CD, ACCT_NM, LOAN_NO, LOAN_DT, GL_DT, " & vbCr
	lgStrSQL = lgStrSQL & " 	SUM_LOAN_AMT, SUM_GL_ITEM_AMT, SUM_DIFF_AMT, SUM_TEMP_ITEM_AMT, SUM_BATCH_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 	GL_NO, TEMP_GL_NO, BATCH_NO, GL_INPUT_TYPE, MINOR_NM " & vbCr
	lgStrSQL = lgStrSQL & lgStrSel1   & vbCr
	lgStrSQL = lgStrSQL & " 	,TEMP_GL_DT	" & vbCr
	lgStrSQL = lgStrSQL & " FROM (	SELECT A.LOAN_ACCT_CD, B.ACCT_NM, A.LOAN_NO, A.LOAN_DT, J.GL_DT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(A.LOAN_LOC_AMT,0)) - SUM(ISNULL(E.ITEM_LOC_AMT,0))  SUM_DIFF_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(A.LOAN_LOC_AMT,0)) SUM_LOAN_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(E.ITEM_LOC_AMT,0)) SUM_GL_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_TEMP_ITEM_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(F.ITEM_LOC_AMT,0)) SUM_BATCH_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		E.GL_NO,  G.TEMP_GL_NO,  F.BATCH_NO, F.GL_INPUT_TYPE, H.MINOR_NM " & vbCr
	lgStrSQL = lgStrSQL & lgStrSel2  & vbCr
	lgStrSQL = lgStrSQL & " 		,G.TEMP_GL_DT " & vbCr
	lgStrSQL = lgStrSQL & " 	FROM F_LN_INFO 	A	LEFT JOIN A_ACCT B ON A.LOAN_ACCT_CD = B.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BANK C ON A.LOAN_BANK_CD = C.BANK_CD " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BIZ_PARTNER D ON A.BP_CD = D.BP_CD " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT E1.REF_NO, E2.ITEM_LOC_AMT, E1.GL_NO, E1.GL_DT, E2.ACCT_CD, E1.GL_INPUT_TYPE " & vbCr
	lgStrSQL = lgStrSQL & " 								FROM A_GL 	E1, A_GL_ITEM	E2 " & vbCr
	lgStrSQL = lgStrSQL & " 								WHERE  E1.GL_NO = E2.GL_NO " & vbCr
	lgStrSQL = lgStrSQL & " 								AND E2.DR_CR_FG = " & FilterVar("CR", "''", "S") & " " & vbCr
	lgStrSQL = lgStrSQL & " 								AND E1.GL_DT >= " & strFrDt  & " AND E1.GL_DT <= " & strToDt  & vbCr
	lgStrSQL = lgStrSQL & " 						) E ON A.LOAN_NO = E.REF_NO AND A.LOAN_ACCT_CD = E.ACCT_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT  F1.BATCH_NO,  F1.REF_NO,  F2.JNL_CD, " & vbCr
	lgStrSQL = lgStrSQL & " 										CASE WHEN  F2.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1) * F2.ITEM_LOC_AMT " & vbCr
	lgStrSQL = lgStrSQL & " 										ELSE  F2.ITEM_LOC_AMT END ITEM_LOC_AMT, F2.ACCT_CD, F1.GL_INPUT_TYPE " & vbCr
	lgStrSQL = lgStrSQL & " 									FROM A_BATCH	F1, A_BATCH_GL_ITEM	F2 " & vbCr
	lgStrSQL = lgStrSQL & " 				 					WHERE  F1.BATCH_NO = F2.BATCH_NO   " & vbCr
	lgStrSQL = lgStrSQL & " 			 						AND F2.JNL_CD IN (SELECT DISTINCT(JNL_CD) " & vbCr
	lgStrSQL = lgStrSQL & " 				        		                    FROM A_JNL_ACCT_ASSN " & vbCr
	lgStrSQL = lgStrSQL & " 													WHERE JNL_CD IN (" & FilterVar("SL", "''", "S") & " , " & FilterVar("LL", "''", "S") & " , " & FilterVar("SN", "''", "S") & " , " & FilterVar("LN", "''", "S") & " , " & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " )) " & vbCr
	lgStrSQL = lgStrSQL & " 									AND (F1.GL_INPUT_TYPE = " & FilterVar("LN", "''", "S") & "  or  (F1.GL_INPUT_TYPE = " & FilterVar("LO", "''", "S") & "  and  F2.JNL_CD IN (" & FilterVar("SL_RO", "''", "S") & " , " & FilterVar("SN_RO", "''", "S") & " , " & FilterVar("LN_RO", "''", "S") & " , " & FilterVar("LL_RO", "''", "S") & " ))) " & vbCr
	lgStrSQL = lgStrSQL & " 									AND F1.GL_INPUT_TYPE IN ( SELECT DISTINCT(MINOR_CD) " & vbCr
	lgStrSQL = lgStrSQL & " 															   FROM B_MINOR " & vbCr
	lgStrSQL = lgStrSQL & " 																WHERE MINOR_CD IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LO", "''", "S") & " ) AND MAJOR_CD = " & FilterVar("A1001", "''", "S") & " ) " & vbCr
	lgStrSQL = lgStrSQL & " 						) F ON A.LOAN_NO = F.REF_NO AND A.LOAN_ACCT_CD = F.ACCT_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN (SELECT   G1.REF_NO, G1.TEMP_GL_NO,  G1.TEMP_GL_DT, G2.ITEM_LOC_AMT, G2.ACCT_CD, G1.GL_INPUT_TYPE " & vbCr
	lgStrSQL = lgStrSQL & " 									FROM A_TEMP_GL 	G1, A_TEMP_GL_ITEM	G2 " & vbCr
	lgStrSQL = lgStrSQL & " 									WHERE  G1.TEMP_GL_NO = G2.TEMP_GL_NO  " & vbCr
	lgStrSQL = lgStrSQL & " 									AND G2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 									AND G1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 						) G ON A.LOAN_NO = G.REF_NO AND A.LOAN_ACCT_CD = G.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN B_MINOR H ON ((H.MINOR_CD = E.GL_INPUT_TYPE OR  H.MINOR_CD = G.GL_INPUT_TYPE ) AND MAJOR_CD = " & FilterVar("A1001", "''", "S") & " ) " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN B_BIZ_AREA I ON A.BIZ_AREA_CD = I.BIZ_AREA_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 						LEFT JOIN A_GL J ON A.GL_NO = J.GL_NO  " & vbCr
	lgStrSQL = lgStrSQL & "		WHERE  A.LOAN_DT >= " & strFrDt  & " AND A.LOAN_DT <= " & strToDt  & vbCr
	lgStrSQL = lgStrSQL & "		AND A.LOAN_BASIC_FG IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LR", "''", "S") & " ) " & vbCr
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.BIZ_AREA_CD = " & strBizAreaCd  & vbCr
	End If
	
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.LOAN_ACCT_CD = " & strAcctCd  & vbCr
	End If

	If Trim(Request("txtInputType")) <> "" Then
		lgStrSQL = lgStrSQL & " AND H.MINOR_CD = " & FilterVar(Request("txtInputType"), "''", "S")  & vbCr
	Else 
		lgStrSQL = lgStrSQL & " AND H.MINOR_CD IN (" & FilterVar("LN", "''", "S") & " , " & FilterVar("LO", "''", "S") & " ) " & vbCr
	End If		
	lgStrSQL = lgStrSQL & " AND H.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
	
	
	If Trim(Request("txtLoanerFg")) = "BK" Then
		lgStrSQL = lgStrSQL & " AND A.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND A.LOAN_BANK_CD = " & strLoanerCd  & vbCr
		End If 
		
	ElseIf Trim(Request("txtLoanerFg")) = "BP" Then
		lgStrSQL = lgStrSQL & " AND A.LOAN_PLC_TYPE = " & strLoanerFg  & vbCr
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND A.BP_CD = " & strLoanerCd  & vbCr
		End If 	
	End If		
	
	lgStrSQL = lgStrSQL & " 	GROUP BY A.LOAN_ACCT_CD, B.ACCT_NM, A.LOAN_NO, A.LOAN_DT, J.GL_DT, " & vbCr
	lgStrSQL = lgStrSQL & " 	 E.GL_NO, G.TEMP_GL_NO,  F.BATCH_NO, F.GL_INPUT_TYPE, H.MINOR_NM " & vbCr
	lgStrSQL = lgStrSQL & lgStrGrpBy   & vbCr
	lgStrSQL = lgStrSQL & " 	,G.TEMP_GL_DT) A " & vbCr

	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " WHERE  ISNULL(SUM_LOAN_AMT, 0) <> ISNULL(SUM_GL_ITEM_AMT, 0)" & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " ORDER BY LOAN_ACCT_CD "	  & vbCr
		
	'Response.write lgStrSQL
    'Response.End 
    
		
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
   		lgStrPrevKey = "" 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
		lgStrPrevKey   = Trim(Request("lgStrPrevKey"))									'☜: Next Key
        lgErrorStatus  = "YES"
        Exit Sub
	Else
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = Request("txtMaxRows")												'☜: Read Operation Mode (CRUD)
		
		If CDbl(lgStrPrevKey) > 0 Then
          lgObjRs.Move     = CDbl(lgMaxCount) * CDbl(lgStrPrevKey)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
        End If

		Do While Not lgObjRs.EOF
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))				'LOAN_ACCT_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))				'ACCT_NM
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))				'LOAN_NO
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(3))		'LOAN_DT
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(4))		'GL_DT	    
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(5), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_DIFF_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(6), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_LOAN_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(7), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_BATCH_ITEM_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(8), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_GL_ITEM_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(9), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_TEMP_ITEM_AMT		
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))          	'GL_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(11))				'TEMP_GL_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))				'BATCH_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(13))				'GL_INPUT_TYPE
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(14))				'MINOR_NM	
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(15))				'BIZ_AREA_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(16))				'BIZ_AREA_NM
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(17))				'LOANER
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(18))    			'LOANER_NM
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(19))		'TEMP_GL_DT			  
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
          
			lgObjRs.MoveNext

			iDx =  iDx + 1

			If iDx > lgMaxCount Then
				lgStrPrevKey = lgStrPrevKey + 1

			Exit Do
			End If   
		Loop 
	End If
	If iDx <= lgMaxCount Then
	   lgStrPrevKey = ""
	End If   
	
	
	

	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                                             " & vbCr
		Response.Write  "    Parent.ggoSpread.Source              = Parent.frm1.vspdData3        " & vbCr
		Response.Write  "    Parent.lgStrPrevKey                  =  """ & lgStrPrevKey     & """" & vbCr       
		Response.Write  "    Parent.ggoSpread.SSShowData             """ & lgstrData        & """" & vbCr
		Response.Write  "    Parent.DBQueryOk												    " & vbCr      
		Response.Write  " </Script>															    " & vbCr
    End If
	
	Response.End
End Sub    


%>

