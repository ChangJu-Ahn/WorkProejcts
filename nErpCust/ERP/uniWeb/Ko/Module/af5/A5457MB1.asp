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
    DIm strBpCd
    
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
     strFrDt		= FilterVar(UNIConvDate(Request("txtSttlFrDt")), "''", "S")
     strToDt		= FilterVar(UNIConvDate(Request("txtSttlToDt")), "''", "S")
     strBizAreaCd   = FilterVar(Request("txtBizAreaCd"), "''", "S") 
	 strAcctCd      = FilterVar(Request("txtAcctCd"), "''", "S") 	 
	 strBpCd		= FilterVar(Request("txtBpCd"), "''", "S") 
     
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
    
	On Error Resume Next
    Err.Clear    
	
    Const C_SHEETMAXROWS_D  = 100													'☆: Server에서 한번에 fetch할 최대 데이타 건수 
    
    
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT ACCT_CD, ACCT_NM, STTL_AMT, GL_AMT,ISNULL(STTL_AMT, 0) - ISNULL(GL_AMT, 0),  TEMP_AMT, BATCH_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 	GL_INPUT_TYPE, MINOR_NM " & vbCr
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	, BIZ_AREA_CD, BIZ_AREA_NM, BP_CD, BP_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
	lgStrSQL = lgStrSQL & " 	, BIZ_AREA_CD, BIZ_AREA_NM, '', '' " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	, '', '', BP_CD, BP_NM " & vbCr
	Else 	
	lgStrSQL = lgStrSQL & " 	, '', '', '', '' " & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " FROM (	SELECT BT.ACCT_CD, AC.ACCT_NM,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(GL.SUM_STTL_AMT, 0)) +  SUM(ISNULL(TMP.SUM_STTL_AMT, 0)) STTL_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(GL.SUM_GL_AMT, 0)) GL_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(TMP.SUM_TEMP_AMT, 0)) TEMP_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(BT.SUM_BATCH_AMT, 0)) BATCH_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		BT.GL_INPUT_TYPE, MN.MINOR_NM " & vbCr
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	, BT.BIZ_AREA_CD, BA.BIZ_AREA_NM, BT.BP_CD, BP.BP_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
	lgStrSQL = lgStrSQL & " 	, BT.BIZ_AREA_CD, BA.BIZ_AREA_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	, BT.BP_CD, BP.BP_NM " & vbCr
	Else 	
	lgStrSQL = lgStrSQL & "" & vbCr
	End If 		
	lgStrSQL = lgStrSQL & " 		FROM " & vbCr
	
	If Trim(Request("cboNoteFg")) = "CR" Then 
	' 받을어음 
		lgStrSQL = lgStrSQL & " (		SELECT 	C.ACCT_CD, C.GL_INPUT_TYPE, " & vbCr	
		lgStrSQL = lgStrSQL & "			0 BATCH_AMT, " & vbCr
		lgStrSQL = lgStrSQL & "			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_BATCH_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & "			C.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & "		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN (SELECT C1.BATCH_NO, C1.REF_NO, C1.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & "						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, " & vbCr
		lgStrSQL = lgStrSQL & "						C1.BIZ_AREA_CD, C2.KEY_VAL1, C2.JNL_CD, C1.GL_NO, C1.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "					FROM A_BATCH C1, A_BATCH_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & "					WHERE C1.BATCH_NO = C2.BATCH_NO " & vbCr
		lgStrSQL = lgStrSQL & "					AND C2.JNL_CD = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND ISNULL(C2.EVENT_CD, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & "					AND C2.TRANS_TYPE = " & FilterVar("FN006", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt & vbCr			
		lgStrSQL = lgStrSQL & "					AND (ISNULL(C1.GL_NO, '') <> '' OR ISNULL(C1.TEMP_GL_NO, '') <> '' ) " & vbCr
		lgStrSQL = lgStrSQL & "				) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND A.NOTE_NO = C.KEY_VAL1  " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN A_ACCT  D ON C.ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & "		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "		AND B.SEQ = 1 " & vbCr
		lgStrSQL = lgStrSQL & " 	AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr		
		lgStrSQL = lgStrSQL & "		AND D.ACCT_TYPE = " & FilterVar("D1", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "		GROUP BY C.ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD ) BT " & vbCr
		lgStrSQL = lgStrSQL & "LEFT JOIN(	SELECT  B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & "			SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & "			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & "			C.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & "		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,   " & vbCr
		lgStrSQL = lgStrSQL & "						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, " & vbCr
		lgStrSQL = lgStrSQL & "						C1.BIZ_AREA_CD, C2.ITEM_DESC  " & vbCr
		lgStrSQL = lgStrSQL & "					FROM A_GL C1, A_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & "					WHERE C1.GL_NO = C2.GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "					AND C2.DR_CR_FG =  " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt & vbCr		
		lgStrSQL = lgStrSQL & "				) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_SEQ = C.ITEM_SEQ AND B.GL_NO = C.GL_NO AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & "		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 	AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr							
		lgStrSQL = lgStrSQL & "		AND ISNULL(C.ITEM_SEQ, '') <> ''  " & vbCr
		lgStrSQL = lgStrSQL & "		AND D.ACCT_TYPE = " & FilterVar("D1", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "		GROUP BY B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD ) GL " & vbCr
		lgStrSQL = lgStrSQL & "		ON BT.ACCT_CD = GL.NOTE_ACCT_CD AND BT.BIZ_AREA_CD = GL.BIZ_AREA_CD AND BT.BP_CD = GL.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & " LEFT JOIN( " & vbCr
		lgStrSQL = lgStrSQL & "		SELECT 	C.ACCT_CD, D.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & "			SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & "			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & "			D.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & "		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN (SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,   " & vbCr
		lgStrSQL = lgStrSQL & "						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ,  " & vbCr
		lgStrSQL = lgStrSQL & "						C1.BIZ_AREA_CD, C2.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & "					FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & "					WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "					AND C2.DR_CR_FG =  " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt & vbCr		
		lgStrSQL = lgStrSQL & "				) C ON B.TEMP_GL_NO = C.TEMP_GL_NO  " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN (SELECT D1.BATCH_NO, D1.REF_NO, D1.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & "						D2.ITEM_LOC_AMT, D2.ACCT_CD, D2.ITEM_SEQ, " & vbCr
		lgStrSQL = lgStrSQL & "						D1.BIZ_AREA_CD , D2.KEY_VAL1, D2.JNL_CD, D1.GL_NO, D1.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "					FROM A_BATCH D1, A_BATCH_GL_ITEM D2  " & vbCr
		lgStrSQL = lgStrSQL & "					WHERE D1.BATCH_NO = D2.BATCH_NO " & vbCr
		lgStrSQL = lgStrSQL & "					AND D2.JNL_CD = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "					AND ISNULL(D2.EVENT_CD, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & "					AND D1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND D1.GL_DT >= " & strFrDt  & " AND D1.GL_DT <= " & strToDt & vbCr 	
		lgStrSQL = lgStrSQL & "					AND ISNULL(D1.TEMP_GL_NO, '') <> ''  " & vbCr
		lgStrSQL = lgStrSQL & "					AND ISNULL(D1.GL_NO, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & "				) D ON A.NOTE_NO = D.KEY_VAL1  " & vbCr
		lgStrSQL = lgStrSQL & "					AND C.TEMP_GL_NO = D.TEMP_GL_NO AND C.ITEM_DESC = D.KEY_VAL1 " & vbCr
		lgStrSQL = lgStrSQL & "				LEFT JOIN A_ACCT E ON C.ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & "		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 	AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr		 		
		lgStrSQL = lgStrSQL & "		AND E.ACCT_TYPE = " & FilterVar("D1", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & "		AND ISNULL(D.BIZ_AREA_CD, '') <> '' " & vbCr
		lgStrSQL = lgStrSQL & "		GROUP BY C.ACCT_CD, D.GL_INPUT_TYPE, D.BIZ_AREA_CD, A.BP_CD ) TMP  " & vbCr
		lgStrSQL = lgStrSQL & "		ON BT.ACCT_CD = TMP.ACCT_CD AND BT.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND BT.BP_CD = TMP.BP_CD " & vbCr

	
	ElseIf Trim(Request("cboNoteFg")) = "CP" Then 
	' 지급어음	
		lgStrSQL = lgStrSQL & " (		SELECT 	C.ACCT_CD, C.GL_INPUT_TYPE, " & vbCr	
		lgStrSQL = lgStrSQL & " 			0 BATCH_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_BATCH_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			C.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & " 		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT C1.BATCH_NO, C1.REF_NO, C1.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, " & vbCr
		lgStrSQL = lgStrSQL & " 						C1.BIZ_AREA_CD, C2.KEY_VAL1, C2.JNL_CD, C1.GL_NO, C1.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					FROM A_BATCH C1, A_BATCH_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & " 					WHERE C1.BATCH_NO = C2.BATCH_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C2.JNL_CD = " & FilterVar("CP", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND ISNULL(C2.EVENT_CD, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C2.TRANS_TYPE = " & FilterVar("FN006", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr		
		lgStrSQL = lgStrSQL & " 					AND (ISNULL(C1.GL_NO, '') <> '' OR ISNULL(C1.TEMP_GL_NO, '') <> '' ) " & vbCr
		lgStrSQL = lgStrSQL & " 				) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND A.NOTE_NO = C.KEY_VAL1  " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN A_ACCT  D ON C.ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		AND B.SEQ = 1 " & vbCr
		lgStrSQL = lgStrSQL & " 		AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr					
		lgStrSQL = lgStrSQL & " 		AND D.ACCT_TYPE = " & FilterVar("D3", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		GROUP BY C.ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD ) BT " & vbCr
		lgStrSQL = lgStrSQL & " LEFT JOIN(	SELECT  B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 			SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			C.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & " 		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,   " & vbCr
		lgStrSQL = lgStrSQL & " 						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ,  " & vbCr
		lgStrSQL = lgStrSQL & " 						C1.BIZ_AREA_CD, C2.ITEM_DESC  " & vbCr
		lgStrSQL = lgStrSQL & " 					FROM A_GL C1, A_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & " 					WHERE C1.GL_NO = C2.GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C2.DR_CR_FG =  " & FilterVar("DR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr			
		lgStrSQL = lgStrSQL & " 				) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_SEQ = C.ITEM_SEQ AND B.GL_NO = C.GL_NO AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD " & vbCr
		lgStrSQL = lgStrSQL & " 		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr			
		lgStrSQL = lgStrSQL & " 		AND ISNULL(C.ITEM_SEQ, '') <> '' " & vbCr
		lgStrSQL = lgStrSQL & " 		AND D.ACCT_TYPE = " & FilterVar("D3", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		GROUP BY B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD ) GL " & vbCr
		lgStrSQL = lgStrSQL & " 		ON BT.ACCT_CD = GL.NOTE_ACCT_CD AND BT.BIZ_AREA_CD = GL.BIZ_AREA_CD AND BT.BP_CD = GL.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & "  LEFT JOIN( " & vbCr
		lgStrSQL = lgStrSQL & " 		SELECT 	C.ACCT_CD, D.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 			SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 			D.BIZ_AREA_CD, A.BP_CD " & vbCr
		lgStrSQL = lgStrSQL & " 		FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,   " & vbCr
		lgStrSQL = lgStrSQL & " 						C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ,  " & vbCr
		lgStrSQL = lgStrSQL & " 						C1.BIZ_AREA_CD, C2.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & " 					FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2 " & vbCr
		lgStrSQL = lgStrSQL & " 					WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C2.DR_CR_FG =  " & FilterVar("DR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt & vbCr	
		lgStrSQL = lgStrSQL & " 				) C ON B.TEMP_GL_NO = C.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN (SELECT D1.BATCH_NO, D1.REF_NO, D1.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 						D2.ITEM_LOC_AMT, D2.ACCT_CD, D2.ITEM_SEQ, " & vbCr
		lgStrSQL = lgStrSQL & " 						D1.BIZ_AREA_CD , D2.KEY_VAL1, D2.JNL_CD, D1.GL_NO, D1.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					FROM A_BATCH D1, A_BATCH_GL_ITEM D2 " & vbCr
		lgStrSQL = lgStrSQL & " 					WHERE D1.BATCH_NO = D2.BATCH_NO " & vbCr
		lgStrSQL = lgStrSQL & " 					AND D2.JNL_CD = " & FilterVar("CP", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND ISNULL(D2.EVENT_CD, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & " 					AND D1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND D1.GL_DT >= " & strFrDt  & " AND D1.GL_DT <= " & strToDt  & vbCr			
		lgStrSQL = lgStrSQL & " 					AND ISNULL(D1.TEMP_GL_NO, '') <> ''  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND ISNULL(D1.GL_NO, '') = '' " & vbCr
		lgStrSQL = lgStrSQL & " 				) D ON A.NOTE_NO = D.KEY_VAL1  " & vbCr
		lgStrSQL = lgStrSQL & " 					AND C.TEMP_GL_NO = D.TEMP_GL_NO AND C.ITEM_DESC = D.KEY_VAL1 " & vbCr
		lgStrSQL = lgStrSQL & " 				LEFT JOIN A_ACCT E ON C.ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 		WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr		
		lgStrSQL = lgStrSQL & " 		AND E.ACCT_TYPE = " & FilterVar("D3", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 		AND ISNULL(D.BIZ_AREA_CD, '') <> '' " & vbCr
		lgStrSQL = lgStrSQL & " 		GROUP BY C.ACCT_CD, D.GL_INPUT_TYPE, D.BIZ_AREA_CD, A.BP_CD ) TMP  " & vbCr
		lgStrSQL = lgStrSQL & " 		ON BT.ACCT_CD = TMP.ACCT_CD AND BT.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND BT.BP_CD = TMP.BP_CD " & vbCr		
	
	End If 
	
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	A_ACCT		AC ON AC.ACCT_CD = BT.ACCT_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_BIZ_AREA	BA ON BA.BIZ_AREA_CD = BT.BIZ_AREA_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_BIZ_PARTNER  BP ON BP.BP_CD = BT.BP_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_MINOR	MN ON MN.MINOR_CD = BT.GL_INPUT_TYPE AND MN.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
		
	'Where 조건 
	If Trim(Request("txtBizAreaCd")) <> "" And Trim(Request("txtAcctCd")) <> "" And Trim(Request("txtBpCd")) <> "" Then
		lgStrSQL = lgStrSQL & " WHERE BT.BIZ_AREA_CD = " & strBizAreaCd  & vbCr
		lgStrSQL = lgStrSQL & " AND BT.ACCT_CD = " & strAcctCd & vbCr
		lgStrSQL = lgStrSQL & " AND BT.BP_CD = " & strBpCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) <> "" And Trim(Request("txtAcctCd")) <> "" And Trim(Request("txtBpCd")) = ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.BIZ_AREA_CD = " & strBizAreaCd & vbCr
		lgStrSQL = lgStrSQL & " AND BT.ACCT_CD = " & strAcctCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) <> "" And Trim(Request("txtAcctCd")) = "" And Trim(Request("txtBpCd")) <> ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.BIZ_AREA_CD = " & strBizAreaCd	 & vbCr
		lgStrSQL = lgStrSQL & " AND BT.BP_CD = " & strBpCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) = "" And Trim(Request("txtAcctCd")) <> "" And Trim(Request("txtBpCd")) <> ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.ACCT_CD = " & strAcctCd & vbCr
		lgStrSQL = lgStrSQL & " AND BT.BP_CD = " & strBpCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) <> "" And Trim(Request("txtAcctCd")) = "" And Trim(Request("txtBpCd")) = ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.BIZ_AREA_CD = " & strBizAreaCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) = "" And Trim(Request("txtAcctCd")) <> "" And Trim(Request("txtBpCd")) = ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.ACCT_CD = " & strAcctCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) = "" And Trim(Request("txtAcctCd")) = "" And Trim(Request("txtBpCd")) <> ""	Then
		lgStrSQL = lgStrSQL & " WHERE BT.BP_CD = " & strBpCd & vbCr
		
	ElseIf Trim(Request("txtBizAreaCd")) = "" And Trim(Request("txtAcctCd")) = "" And Trim(Request("txtBpCd")) = ""	Then
		lgStrSQL = lgStrSQL & "" & vbCr
	End If
	
	'Group By 조건 
	lgStrSQL = lgStrSQL & "	GROUP BY BT.ACCT_CD, AC.ACCT_NM,  BT.GL_INPUT_TYPE, MN.MINOR_NM " & vbCr
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " 	, BT.BIZ_AREA_CD, BA.BIZ_AREA_NM, BT.BP_CD, BP.BP_NM " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
		lgStrSQL = lgStrSQL & " 	, BT.BIZ_AREA_CD, BA.BIZ_AREA_NM "	 & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " 	, BT.BP_CD, BP.BP_NM " & vbCr
	Else 	
		lgStrSQL = lgStrSQL & "" & vbCr
	End If
	lgStrSQL = lgStrSQL & ") A"	 & vbCr
	
	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " WHERE  ISNULL(STTL_AMT, 0) <> ISNULL(GL_AMT, 0)" & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " ORDER BY ACCT_CD" & vbCr

	'Response.write lgStrSQL
    'Response.End 
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else 
		iDx         = 1
		lgstrData   = ""
		lgLngMaxRow = Request("txtMaxRows")												'☜: Read Operation Mode (CRUD)

		Do While Not lgObjRs.EOF			
	
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))				'NOTE_ACCT_CD
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))				'ACCT_NM
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(2), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	        'SUMSTTL_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(3), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_GL_ITEM_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(4), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_TEMP_ITEM_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(5), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_BATCH_ITEM_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(6), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'SUM_BATCH_ITEM_AMT
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))				'GL_INPUT_TYPE
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))				'MINOR_NM
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))				'BIZ_AREA_CD
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))				'BIZ_AREA_NM
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(11))				'BP_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))				'BP_NM
	        lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)
	          
	        lgObjRs.MoveNext

	        iDx =  iDx + 1
		Loop 
	End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>									" & vbCr
       Response.Write  "    Parent.ggoSpread.Source			= Parent.frm1.vspdData1 " & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If
End Sub    

%>

