<%@ LANGUAGE="VBSCRIPT" %>
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
    Dim strLoanNo 
	DIm strLoanerFg
    DIm strLoanerCd
    Dim strInputType
    
    Dim lgStrPrevKey
    
' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "A","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus  = ""
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))									 '☜: Next Key    

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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
	 strLoanNo      = FilterVar(Request("txtLoanNo"), "''", "S")  	 
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
	Dim lgTotClsLocAmt2, lgTotDiffLocAmt2, lgTotGlLocAmt2, lgTotTempGlLocAmt2

    Dim lgStrSel0, lgStrSel1, lgStrSel2, lgStrSel3
    Dim lgStrGrpBy0, lgStrGrpBy1
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
		
    If Trim(Request("chkShowLoanNo")) = "Y"  Then		
	'차입번호 선택 	
		lgStrSel0 = ", LOAN_NO, SEQ, "
		lgStrSel1 = " A.LOAN_NO, A.SEQ, "
		lgStrGrpBy0 = " A.LOAN_NO, A.SEQ,"   
    Else
		lgStrSel0 = ", '', '', "
		lgStrSel1 = ""
		lgStrGrpBy0 = ""   
    End If 
    
    If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 거래처, 차입처 모두 선택 
		lgStrSel2 = ", BIZ_AREA_CD, BIZ_AREA_NM, LOANER, LOANER_NM " & vbCr
		
		lgStrSel3 = ", B.BIZ_AREA_CD, F.BIZ_AREA_NM,   " & vbCr
		lgStrSel3 = lgStrSel3 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN B.LOAN_BANK_CD ELSE B.BP_CD END  LOANER, " & vbCr
		lgStrSel3 = lgStrSel3 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.BANK_NM ELSE E.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy1 = ", B.BIZ_AREA_CD, F.BIZ_AREA_NM , "   & vbCr
		lgStrGrpBy1 = lgStrGrpBy1 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN B.LOAN_BANK_CD ELSE B.BP_CD END, " & vbCr
		lgStrGrpBy1 = lgStrGrpBy1 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.BANK_NM ELSE E.BP_NM END " & vbCr
		
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowLoaner")) = "N" Then
	' 거래처 선택	
		lgStrSel2 = ", BIZ_AREA_CD, BIZ_AREA_NM, '', '' " & vbCr
		
		lgStrSel3 = ", B.BIZ_AREA_CD, F.BIZ_AREA_NM " & vbCr
		lgStrGrpBy1 = ", B.BIZ_AREA_CD, F.BIZ_AREA_NM "	 & vbCr
		
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowLoaner")) = "Y" Then
	' 차입처 선택 
		lgStrSel2 = ", '', '', LOANER, LOANER_NM " & vbCr
		
		lgStrSel3 = ",			  CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN B.LOAN_BANK_CD ELSE B.BP_CD END  LOANER, " & vbCr
		lgStrSel3 = lgStrSel3 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.BANK_NM ELSE E.BP_NM END LOANER_NM " & vbCr
		
		lgStrGrpBy1 = ",			  CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN B.LOAN_BANK_CD ELSE B.BP_CD END, " & vbCr
		lgStrGrpBy1 = lgStrGrpBy1 & " CASE WHEN  B.LOAN_PLC_TYPE = " & FilterVar("BK", "''", "S") & "  THEN D.BANK_NM ELSE E.BP_NM END " & vbCr
	Else 
	' 선택 없음 
		lgStrSel2 = ", '', '', '', '' " & vbCr
		lgStrSel3 = "" & vbCr
		lgStrGrpBy1 = "" & vbCr		
	End If 	
	

	'-------------------------
	' Sum 구하기 시작 
	'------------------------
		
	If lgStrPrevKey  =0 Then
		'----
		lgStrSQL2 = ""
		lgStrSQL2 = lgStrSQL2 & " SELECT SUM(ISNULL(SUM_CLS_AMT,0)) SUM_CLS_AMT, SUM(ISNULL(SUM_GL_ITEM_AMT,0)) SUM_GL_ITEM_AMT, " & vbCr
		lgStrSQL2 = lgStrSQL2 & "		SUM(ISNULL(SUM_DIFF_AMT,0)) SUM_DIFF_AMT, SUM(ISNULL(SUM_TEMP_ITEM_AMT,0)) SUM_TEMP_ITEM_AMT " & vbCr

		lgStrSQL2 = lgStrSQL2 & " FROM (SELECT 	A.ADV_INT_ACCT_CD, C.ACCT_NM, "	 & vbCr
		lgStrSQL2 = lgStrSQL2 & "			" & lgStrSel1 & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			A.INT_CLS_DT, G.GL_DT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			SUM(ISNULL(A.INT_CLS_LOC_AMT, 0)) - SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_DIFF_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			SUM(ISNULL(A.INT_CLS_LOC_AMT, 0)) SUM_CLS_AMT,   " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_GL_ITEM_AMT,   " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			SUM(ISNULL(I.ITEM_LOC_AMT,0)) SUM_TEMP_ITEM_AMT, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			0 SUM_BATCH_ITEM_AMT ,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			G.GL_NO, I.TEMP_GL_NO, '' BATCH_NO, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.GL_INPUT_TYPE WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.GL_INPUT_TYPE ELSE '' END GL_INPUT_TYPE, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.minor_nm WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.minor_nm ELSE '' END minor_nm " & vbCr
		lgStrSQL2 = lgStrSQL2 & "			" & lgStrSel3 & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			,I.temp_gl_dt " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 		FROM F_LN_MON_ADV_INT A	" & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN F_LN_INFO B ON A.LOAN_NO = B.LOAN_NO " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN A_ACCT C ON A.ADV_INT_ACCT_CD = C.ACCT_CD " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN B_BANK D ON B.LOAN_BANK_CD = D.BANK_CD " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN B_BIZ_PARTNER E ON B.BP_CD = E.BP_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN B_BIZ_AREA F ON B.BIZ_AREA_CD = F.BIZ_AREA_CD " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN (SELECT G1.gl_no, G2.ITEM_LOC_AMT, G1.GL_DT, G2.ACCT_CD, G1.GL_INPUT_TYPE, G2.ITEM_SEQ, J.minor_nm " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								FROM A_GL G1 INNER JOIN A_GL_ITEM G2 ON G1.GL_NO = G2.GL_NO AND G2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 											LEFT JOIN B_MINOR J ON G1.GL_INPUT_TYPE = J.MINOR_CD AND J.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								WHERE G1.GL_INPUT_TYPE = " & FilterVar("LA", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						) G ON A.gl_no = G.gl_no AND A.ADV_INT_ACCT_CD = G.ACCT_CD " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					LEFT JOIN (SELECT I1.TEMP_GL_NO,  I2.ITEM_LOC_AMT, I1.temp_gl_dt, I2.ACCT_CD, I1.GL_INPUT_TYPE, I2.ITEM_SEQ, J.minor_nm " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								FROM A_TEMP_GL I1 INNER JOIN A_TEMP_GL_ITEM I2 ON I1.temp_gl_no = I2.temp_gl_no AND I2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 												LEFT JOIN B_MINOR J ON I1.GL_INPUT_TYPE = J.MINOR_CD AND J.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								WHERE I1.GL_INPUT_TYPE = " & FilterVar("LA", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						) I ON A.temp_gl_no = I.temp_gl_no AND A.ADV_INT_ACCT_CD = I.ACCT_CD " & vbCr
		lgStrSQL2 = lgStrSQL2 & "		WHERE  A.INT_CLS_DT >= " & strFrDt  & " AND A.INT_CLS_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & "			AND A.CLS_FG = " & FilterVar("Y", "''", "S") & "  AND ISNULL(A.LOAN_NO, '') <> '' AND ISNULL(A.PAY_NO, '') <> '' " & vbCr
		
		If Trim(Request("txtBizAreaCd")) <> "" Then
			lgStrSQL2 = lgStrSQL2 & " AND B.BIZ_AREA_CD = " & strBizAreaCd & vbCr
		End If
		
		If Trim(Request("txtAcctCd")) <> "" Then
			lgStrSQL2 = lgStrSQL2 & " AND A.ADV_INT_ACCT_CD = " & strAcctCd & vbCr
		End If	
		
		If Trim(Request("txtLoanNo")) <> "" Then
			lgStrSQL2 = lgStrSQL2 & " AND A.LOAN_NO = " & strLoanNo & vbCr
		End If			
		
		If Trim(Request("txtLoanerFg")) = "BK" Then
			lgStrSQL2 = lgStrSQL2 & " AND B.LOAN_PLC_TYPE = " & strLoanerFg & vbCr
			If Trim(Request("txtLoanerCd")) <> "" Then 
				lgStrSQL2 = lgStrSQL2 & " AND B.LOAN_BANK_CD = " & strLoanerCd & vbCr
			End If 
			
		ElseIf Trim(Request("txtLoanerFg")) = "BP" Then
			lgStrSQL2 = lgStrSQL2 & " AND B.LOAN_PLC_TYPE = " & strLoanerFg & vbCr
			If Trim(Request("txtLoanerCd")) <> "" Then 
				lgStrSQL2 = lgStrSQL2 & " AND B.BP_CD = " & strLoanerCd & vbCr
			End If 	
		End If


		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then
			lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
		End If
	
		If lgInternalCd <> "" Then
			lgInternalCdAuthSQL		= " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
		End If
	
		If lgSubInternalCd <> "" Then
			lgSubInternalCdAuthSQL	= " AND B.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
		End If
	
		If lgAuthUsrID <> "" Then
			lgAuthUsrIDAuthSQL		= " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
		End If

		lgStrSQL2	= lgStrSQL2	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

			
		lgStrSQL2 = lgStrSQL2 & " 	GROUP BY A.ADV_INT_ACCT_CD, C.ACCT_NM, " & vbCr
		lgStrSQL2 = lgStrSQL2 & "			" & lgStrGrpBy0 & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			 A.INT_CLS_DT, G.GL_DT, G.GL_NO, I.TEMP_GL_NO, " & vbCr
		lgStrSQL2 = lgStrSQL2 & "			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.GL_INPUT_TYPE WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.GL_INPUT_TYPE ELSE '' END,"	 & vbCr
		lgStrSQL2 = lgStrSQL2 & "			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.minor_nm WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.minor_nm ELSE '' END "	 & vbCr
		lgStrSQL2 = lgStrSQL2 & "			" & lgStrGrpBy1  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 			 ,I.temp_gl_dt ) A " & vbCr

		If UCase(Trim(Request("DispMeth"))) Then 
			lgStrSQL2 = lgStrSQL2 & " WHERE  ISNULL(SUM_CLS_AMT, 0) <> ISNULL(SUM_GL_ITEM_AMT, 0)" & vbCr
		End If 
	'----
	'*********************************
	'			합계찍기 
	'*********************************
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
			lgTotClsLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("SUM_CLS_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotDiffLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("SUM_DIFF_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotGlLocAmt2  = UNIConvNumDBToCompanyByCurrency(lgObjRs("SUM_GL_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotTempGlLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("SUM_TEMP_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotClsLocAmt2.text     =  """ & lgTotClsLocAmt2    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt2.text  =  """ & lgTotDiffLocAmt2 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt2.text     =  """ & lgTotGlLocAmt2    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt2.text =  """ & lgTotTempGlLocAmt2   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
	
		Else
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotClsLocAmt2.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt2.text  =  """ & 0 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt2.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt2.text =  """ & 0   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
		End If    
	End If
	'-------------------------
	' Sum 구하기 끝 
	'------------------------



	lgStrSQL = lgStrSQL & " SELECT ADV_INT_ACCT_CD, ACCT_NM " & vbCr
	lgStrSQL = lgStrSQL & "			" & lgStrSel0
	lgStrSQL = lgStrSQL & " 	INT_CLS_DT, GL_DT, " & vbCr
	lgStrSQL = lgStrSQL & " 	SUM_CLS_AMT, SUM_GL_ITEM_AMT,SUM_DIFF_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 	SUM_TEMP_ITEM_AMT, 0 SUM_BATCH_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 	GL_NO, TEMP_GL_NO, '' BATCH_NO,  " & vbCr
	lgStrSQL = lgStrSQL & " 	GL_INPUT_TYPE, MINOR_NM " & vbCr
	lgStrSQL = lgStrSQL & "			" & lgStrSel2  & vbCr
	lgStrSQL = lgStrSQL & " 	,temp_gl_dt " & vbCr

	lgStrSQL = lgStrSQL & " FROM (SELECT A.ADV_INT_ACCT_CD, C.ACCT_NM, "	 & vbCr
	lgStrSQL = lgStrSQL & "				" & lgStrSel1 & vbCr
	lgStrSQL = lgStrSQL & " 			A.INT_CLS_DT, G.GL_DT,  " & vbCr
	lgStrSQL = lgStrSQL & " 			SUM(ISNULL(A.INT_CLS_LOC_AMT, 0)) - SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_DIFF_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 			SUM(ISNULL(A.INT_CLS_LOC_AMT, 0)) SUM_CLS_AMT,   " & vbCr
	lgStrSQL = lgStrSQL & " 			SUM(ISNULL(G.ITEM_LOC_AMT,0)) SUM_GL_ITEM_AMT,   " & vbCr
	lgStrSQL = lgStrSQL & " 			SUM(ISNULL(I.ITEM_LOC_AMT,0)) SUM_TEMP_ITEM_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 			G.GL_NO, I.TEMP_GL_NO, " & vbCr
	lgStrSQL = lgStrSQL & " 			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.GL_INPUT_TYPE WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.GL_INPUT_TYPE ELSE '' END GL_INPUT_TYPE, " & vbCr
	lgStrSQL = lgStrSQL & " 			CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.minor_nm WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.minor_nm ELSE '' END minor_nm " & vbCr
	lgStrSQL = lgStrSQL & "				" & lgStrSel3 & vbCr
	lgStrSQL = lgStrSQL & " 			,I.temp_gl_dt " & vbCr
	lgStrSQL = lgStrSQL & " 		FROM F_LN_MON_ADV_INT A	" & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN F_LN_INFO B ON A.LOAN_NO = B.LOAN_NO " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN A_ACCT C ON A.ADV_INT_ACCT_CD = C.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN B_BANK D ON B.LOAN_BANK_CD = D.BANK_CD " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN B_BIZ_PARTNER E ON B.BP_CD = E.BP_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN B_BIZ_AREA F ON B.BIZ_AREA_CD = F.BIZ_AREA_CD " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT G1.gl_no, G2.ITEM_LOC_AMT, G1.GL_DT, G2.ACCT_CD, G1.GL_INPUT_TYPE, G2.ITEM_SEQ, J.minor_nm " & vbCr
	lgStrSQL = lgStrSQL & " 							FROM A_GL G1 INNER JOIN A_GL_ITEM G2 ON G1.GL_NO = G2.GL_NO AND G2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 											LEFT JOIN B_MINOR J ON G1.GL_INPUT_TYPE = J.MINOR_CD AND J.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 							WHERE G1.GL_INPUT_TYPE = " & FilterVar("LA", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 						) G ON A.gl_no = G.gl_no AND A.ADV_INT_ACCT_CD = G.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT I1.TEMP_GL_NO,  I2.ITEM_LOC_AMT, I1.temp_gl_dt, I2.ACCT_CD, I1.GL_INPUT_TYPE, I2.ITEM_SEQ, J.minor_nm " & vbCr
	lgStrSQL = lgStrSQL & " 							FROM A_TEMP_GL I1 INNER JOIN A_TEMP_GL_ITEM I2 ON I1.temp_gl_no = I2.temp_gl_no AND I2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 												LEFT JOIN B_MINOR J ON I1.GL_INPUT_TYPE = J.MINOR_CD AND J.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 							WHERE I1.GL_INPUT_TYPE = " & FilterVar("LA", "''", "S") & "  " & vbCr
	lgStrSQL = lgStrSQL & " 						) I ON A.temp_gl_no = I.temp_gl_no AND A.ADV_INT_ACCT_CD = I.ACCT_CD " & vbCr
	lgStrSQL = lgStrSQL & "			WHERE  A.INT_CLS_DT >= " & strFrDt  & " AND A.INT_CLS_DT <= " & strToDt  & vbCr
	lgStrSQL = lgStrSQL & "				AND A.CLS_FG = " & FilterVar("Y", "''", "S") & "  AND ISNULL(A.LOAN_NO, '') <> '' AND ISNULL(A.PAY_NO, '') <> '' " & vbCr
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND B.BIZ_AREA_CD = " & strBizAreaCd & vbCr
	End If
	
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ADV_INT_ACCT_CD = " & strAcctCd & vbCr
	End If	
	
	If Trim(Request("txtLoanNo")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.LOAN_NO = " & strLoanNo & vbCr
	End If			
	
	If Trim(Request("txtLoanerFg")) = "BK" Then
		lgStrSQL = lgStrSQL & " AND B.LOAN_PLC_TYPE = " & strLoanerFg & vbCr
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND B.LOAN_BANK_CD = " & strLoanerCd & vbCr
		End If 
		
	ElseIf Trim(Request("txtLoanerFg")) = "BP" Then
		lgStrSQL = lgStrSQL & " AND B.LOAN_PLC_TYPE = " & strLoanerFg & vbCr
		If Trim(Request("txtLoanerCd")) <> "" Then 
			lgStrSQL = lgStrSQL & " AND B.BP_CD = " & strLoanerCd & vbCr
		End If 	
	End If			


	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND B.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	lgStrSQL	= lgStrSQL	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

		
	lgStrSQL = lgStrSQL & " 	GROUP BY A.ADV_INT_ACCT_CD, C.ACCT_NM, " & vbCr
	lgStrSQL = lgStrSQL & "				" & lgStrGrpBy0 & vbCr
	lgStrSQL = lgStrSQL & " 			 A.INT_CLS_DT, G.GL_DT, G.GL_NO, I.TEMP_GL_NO, " & vbCr
	lgStrSQL = lgStrSQL & "				CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.GL_INPUT_TYPE WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.GL_INPUT_TYPE ELSE '' END,"	 & vbCr
	lgStrSQL = lgStrSQL & "				CASE WHEN LTrim(RTrim(ISNULL(G.GL_INPUT_TYPE, ''))) <> '' THEN G.minor_nm WHEN LTrim(RTrim(ISNULL(I.GL_INPUT_TYPE, ''))) <> '' THEN I.minor_nm ELSE '' END "	 & vbCr
	lgStrSQL = lgStrSQL & "				" & lgStrGrpBy1  & vbCr
	lgStrSQL = lgStrSQL & " 			,I.TEMP_GL_DT ) A " & vbCr

	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " WHERE  ISNULL(SUM_CLS_AMT, 0) <> ISNULL(SUM_GL_ITEM_AMT, 0)" & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " ORDER BY ADV_INT_ACCT_CD, INT_CLS_DT " & vbCr
    
		
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKey = "" 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
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
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))				'ADV_INT_ACCT_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))				'ACCT_NM
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))				'LOAN_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))				'SEQ
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(4))		'INT_CLS_DT
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(5))		'GL_DT	    
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(6), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_CLS_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(7), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_GL_ITEM_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(8), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_DIFF_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(9), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		'SUM_BATCH_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(10), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	'SUM_TEMP_ITEM_AMT		
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(11))          	'GL_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))				'TEMP_GL_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(13))				'BATCH_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(14))				'GL_INPUT_TYPE
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(15))				'MINOR_NM	
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(16))				'BIZ_AREA_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(17))				'BIZ_AREA_NM
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(18))				'LOANER
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(19))    			'LOANER_NM
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(20))		'TEMP_GL_DT			  
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
       Response.Write  "    Parent.ggoSpread.Source              = Parent.frm1.vspdData2        " & vbCr
       Response.Write  "    Parent.lgStrPrevKey                  =  """ & lgStrPrevKey     & """" & vbCr       
       Response.Write  "    Parent.ggoSpread.SSShowData             """ & lgstrData        & """" & vbCr
       Response.Write  "    Parent.DBQueryOk												    " & vbCr      
       Response.Write  " </Script>															    " & vbCr
    End If
	
	Response.End
End Sub    


%>

