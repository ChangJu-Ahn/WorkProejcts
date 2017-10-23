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
    Dim strNoteNoFr
	Dim strNoteNoTo
    
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
	 strNoteNoFr	= FilterVar(Request("txtNoteNoFr"), "''", "S") 
	 strNoteNoTo	= FilterVar(Request("txtNoteNoTo"), "''", "S") 
     
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
	Dim lgStrSel1, lgStrSel2, lgStrSQL2
    Dim lgStrGrpBy
	Dim lgMaxCount
	Dim lgStrPrevKeyIndex
    Dim lgTotRepayLocAmt2, lgTotBatchLocAmt2, lgTotGlLocAmt2, lgTotTempGlLocAmt2
	
	On Error Resume Next
    Err.Clear    
	
    Const C_SHEETMAXROWS_D  = 100													'☆: Server에서 한번에 fetch할 최대 데이타 건수 
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
   	If Len(Trim(Request("lgStrPrevKeyIndex")))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKeyIndex) Then
           lgStrPrevKeyIndex = CInt(lgStrPrevKeyIndex)          
        End If   
    Else   
        lgStrPrevKeyIndex = 0
    End If
	
	
	
	'-------------------------
	' Sum 구하기 시작 
	'------------------------
		
	If lgStrPrevKeyIndex  =0 Then
		'----
		
	lgStrSQL2 = ""
	lgStrSQL2 = lgStrSQL2 & " Select SUM(ISNULL(STTL_AMT, 0)) STTL_AMT,SUM(ISNULL(GL_AMT,0)) GL_AMT,SUM(ISNULL(DIFF_AMT,0)) DIFF_AMT, SUM(ISNULL(TEMP_GL_AMT,0)) TEMP_GL_AMT "
	lgStrSQL2 = lgStrSQL2 & " FROM (	SELECT GL.NOTE_ACCT_CD, AC.ACCT_NM, GL.NOTE_NO, GL.STS_DT, GL.GL_DT,  " & vbCr
	lgStrSQL2 = lgStrSQL2 & "			SUM(ISNULL(GL.SUM_STTL_AMT, 0)) +  SUM(ISNULL(TMP.SUM_STTL_AMT, 0)) - SUM(ISNULL(GL.SUM_GL_AMT, 0)) DIFF_AMT," & vbCr
	lgStrSQL2 = lgStrSQL2 & " 		SUM(ISNULL(GL.SUM_STTL_AMT, 0)) +  SUM(ISNULL(TMP.SUM_STTL_AMT, 0)) STTL_AMT, " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 		SUM(ISNULL(GL.SUM_GL_AMT, 0)) GL_AMT,  " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 		SUM(ISNULL(TMP.SUM_TEMP_AMT, 0)) TEMP_GL_AMT, " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 		GL.GL_NO, TMP.TEMP_GL_NO,  GL.GL_INPUT_TYPE, MN.MINOR_NM, " & vbCr
	
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL2 = lgStrSQL2 & " 		GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, GL.BP_CD, BP.BP_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
	lgStrSQL2 = lgStrSQL2 & " 		GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL2 = lgStrSQL2 & " 		GL.BP_CD, BP.BP_NM, " & vbCr
	Else 	
	lgStrSQL2 = lgStrSQL2 & "" & vbCr
	End If 		
	lgStrSQL2 = lgStrSQL2 & "		TMP.TEMP_GL_DT " & vbCr
	
	
	
	'내부 TABLE 
	If Trim(Request("cboNoteFg")) = "CR" Then 
	' 수취구매카드 
		lgStrSQL2 = lgStrSQL2 & " 		FROM  (  SELECT  A.NOTE_NO, B.STS_DT, C.GL_NO,  C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,    " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 											C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 											C1.BIZ_AREA_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										FROM A_GL C1, A_GL_ITEM C2  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										WHERE C1.GL_NO = C2.GL_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C2.DR_CR_FG = " & FilterVar("CR", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_NO = C.GL_NO    " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND B.GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND ISNULL(C.ITEM_SEQ, '') <> ''  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND A.NOTE_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				GROUP BY A.NOTE_NO, C.GL_NO, C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT) GL " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 		 LEFT JOIN (  SELECT 	A.NOTE_NO, B.STS_DT, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						LEFT JOIN ( SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								C1.BIZ_AREA_CD    " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						) C ON  B.NOTE_ACCT_CD = C.ACCT_CD and B.TEMP_GL_NO = C.TEMP_GL_NO " & vbCr
		lgStrSQL2 = lgStrSQL2 & "									AND B.TEMP_GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						LEFT JOIN A_ACCT E ON B.NOTE_ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND A.NOTE_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND ISNULL(C.ITEM_SEQ, '') <> '' " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				GROUP BY A.NOTE_NO, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT ) TMP " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				ON GL.NOTE_ACCT_CD = TMP.NOTE_ACCT_CD AND GL.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND GL.BP_CD = TMP.BP_CD AND GL.NOTE_NO = TMP.NOTE_NO " & vbCr
	
		ElseIf Trim(Request("cboNoteFg")) = "CP" Then 
	' 지불구매카드		
		lgStrSQL2 = lgStrSQL2 & " 		FROM  (  SELECT  A.NOTE_NO, B.STS_DT, C.GL_NO,  C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,    " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 											C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 											C1.BIZ_AREA_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										FROM A_GL C1, A_GL_ITEM C2  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										WHERE C1.GL_NO = C2.GL_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C2.DR_CR_FG = " & FilterVar("DR", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_NO = C.GL_NO    " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 										AND B.GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND ISNULL(C.ITEM_SEQ, '') <> ''  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND A.NOTE_FG = " & FilterVar("DR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				GROUP BY A.NOTE_NO, C.GL_NO, C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT) GL " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 		 LEFT JOIN (  SELECT 	A.NOTE_NO, B.STS_DT, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						LEFT JOIN ( SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 								C1.BIZ_AREA_CD, dbo.ufn_GL_DTL_by_table(C2.TEMP_gl_no, c2.item_seq, " & FilterVar("F_NOTE", "''", "S") & " , " & FilterVar("NOTE_NO", "''", "S") & " )  NOTE_NO   " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C2.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 							AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt  & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						) C ON  B.NOTE_ACCT_CD = C.ACCT_CD and B.TEMP_GL_NO = C.TEMP_GL_NO " & vbCr
		lgStrSQL2 = lgStrSQL2 & "									AND B.TEMP_GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 						LEFT JOIN A_ACCT E ON B.NOTE_ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND A.NOTE_FG = " & FilterVar("CP", "''", "S") & "  " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				AND ISNULL(C.ITEM_SEQ, '') <> '' " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				GROUP BY A.NOTE_NO, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT ) TMP " & vbCr
		lgStrSQL2 = lgStrSQL2 & " 				ON GL.NOTE_ACCT_CD = TMP.NOTE_ACCT_CD AND GL.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND GL.BP_CD = TMP.BP_CD AND GL.NOTE_NO = TMP.NOTE_NO " & vbCr
	
	End If 
	
	lgStrSQL2 = lgStrSQL2 & " 	LEFT JOIN	A_ACCT		AC ON AC.ACCT_CD = GL.NOTE_ACCT_CD  " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 	LEFT JOIN	B_BIZ_AREA	BA ON BA.BIZ_AREA_CD = GL.BIZ_AREA_CD  " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 	LEFT JOIN	B_BIZ_PARTNER  BP ON BP.BP_CD = GL.BP_CD  " & vbCr
	lgStrSQL2 = lgStrSQL2 & " 	LEFT JOIN	B_MINOR	MN ON MN.MINOR_CD = GL.GL_INPUT_TYPE AND MN.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
		
	'내부 TABLE Where 조건 
	lgStrSQL2 = lgStrSQL2 & " WHERE GL.GL_DT  >= " & strFrDt  & " AND GL.GL_DT <= " & strToDt & vbCr
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND GL.BIZ_AREA_CD = " & strBizAreaCd & vbCr
	End If 
	
	If Trim(Request("txtAcctCd")) <> "" Then		
		lgStrSQL2 = lgStrSQL2 & " AND GL.NOTE_ACCT_CD = " & strAcctCd & vbCr
	End If
		
	If Trim(Request("txtBpCd")) <> ""	Then		
		lgStrSQL2 = lgStrSQL2 & " AND GL.BP_CD = " & strBpCd  & vbCr
	End If		
	
	If Trim(Request("txtNoteNoFr")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND GL.NOTE_NO >= " & strNoteNoFr  & vbCr
	End If
	
	If Trim(Request("txtNoteNoTo")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND GL.NOTE_NO <= " & strNoteNoTo & vbCr
	End If		
	
	'내부 TABLE Group By 조건 
	lgStrSQL2 = lgStrSQL2 & "	GROUP BY  GL.NOTE_ACCT_CD, AC.ACCT_NM,  GL.NOTE_NO, GL.STS_DT, GL.GL_DT,  " & vbCr
	lgStrSQL2 = lgStrSQL2 & "				GL.GL_NO, TMP.TEMP_GL_NO,  " & vbCr
	lgStrSQL2 = lgStrSQL2 & "				GL.GL_INPUT_TYPE, MN.MINOR_NM, " & vbCr
	
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL2 = lgStrSQL2 & " 	GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, GL.BP_CD, BP.BP_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
		lgStrSQL2 = lgStrSQL2 & " 	GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL2 = lgStrSQL2 & " 	GL.BP_CD, BP.BP_NM, " & vbCr
	Else 	
		lgStrSQL2 = lgStrSQL2 & "" & vbCr
	End If
	lgStrSQL2 = lgStrSQL2 & " TEMP_GL_DT) A"	 & vbCr
	
	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL2 = lgStrSQL2 & " WHERE  ISNULL(STTL_AMT, 0) <> ISNULL(GL_AMT, 0)" & vbCr
	End If 	

	
	'----
	'*********************************
	'			합계찍기 
	'*********************************

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
			lgTotRepayLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("STTL_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotBatchLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("DIFF_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotGlLocAmt2  = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTotTempGlLocAmt2 = UNIConvNumDBToCompanyByCurrency(lgObjRs("TEMP_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotClsLocAmt2.text     =  """ & lgTotRepayLocAmt2    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotBatchLocAmt2.text  =  """ & lgTotBatchLocAmt2 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt2.text     =  """ & lgTotGlLocAmt2    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt2.text =  """ & lgTotTempGlLocAmt2   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
	
		Else
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotClsLocAmt2.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotBatchLocAmt2.text  =  """ & 0 & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt2.text     =  """ & 0    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt2.text =  """ & 0   & """" & vbCr
			Response.Write  " </Script>															    " & vbCr
		End If    
	End If
	'-------------------------
	' Sum 구하기 끝 
	'------------------------

	'SELECT 
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT NOTE_ACCT_CD, ACCT_NM, NOTE_NO, STS_DT, GL_DT,  " & vbCr
	lgStrSQL = lgStrSQL & "		   STTL_AMT, GL_AMT, DIFF_AMT, TEMP_GL_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 	   GL_NO, TEMP_GL_NO, GL_INPUT_TYPE, MINOR_NM,  " & vbCr
	
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	   BIZ_AREA_CD, BIZ_AREA_NM, BP_CD, BP_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
	lgStrSQL = lgStrSQL & " 	   BIZ_AREA_CD, BIZ_AREA_NM, '', '', "	 & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 	   '', '', BP_CD, BP_NM, " & vbCr
	Else 	
	lgStrSQL = lgStrSQL & " 	   '', '', '', '', " & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & "	       TEMP_GL_DT  " & vbCr
	
	'FROM		
 	
	lgStrSQL = lgStrSQL & " FROM (	SELECT GL.NOTE_ACCT_CD, AC.ACCT_NM, GL.NOTE_NO, GL.STS_DT, GL.GL_DT,  " & vbCr
	lgStrSQL = lgStrSQL & "			SUM(ISNULL(GL.SUM_STTL_AMT, 0)) +  SUM(ISNULL(TMP.SUM_STTL_AMT, 0)) - SUM(ISNULL(GL.SUM_GL_AMT, 0)) DIFF_AMT," & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(GL.SUM_STTL_AMT, 0)) +  SUM(ISNULL(TMP.SUM_STTL_AMT, 0)) STTL_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(GL.SUM_GL_AMT, 0)) GL_AMT,  " & vbCr
	lgStrSQL = lgStrSQL & " 		SUM(ISNULL(TMP.SUM_TEMP_AMT, 0)) TEMP_GL_AMT, " & vbCr
	lgStrSQL = lgStrSQL & " 		GL.GL_NO, TMP.TEMP_GL_NO,  GL.GL_INPUT_TYPE, MN.MINOR_NM,  " & vbCr
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 		GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, GL.BP_CD, BP.BP_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
	lgStrSQL = lgStrSQL & " 		GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
	lgStrSQL = lgStrSQL & " 		GL.BP_CD, BP.BP_NM, " & vbCr
	Else 	
	lgStrSQL = lgStrSQL & "" & vbCr
	End If 		
	lgStrSQL = lgStrSQL & "			TMP.TEMP_GL_DT " & vbCr


	
	'내부 TABLE 
	If Trim(Request("cboNoteFg")) = "CR" Then 
	' 수취구매카드 
		lgStrSQL = lgStrSQL & " 		FROM  (  SELECT  A.NOTE_NO, B.STS_DT, C.GL_NO,  C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT, " & vbCr
		lgStrSQL = lgStrSQL & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 								LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,    " & vbCr
		lgStrSQL = lgStrSQL & " 											C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC,  " & vbCr
		lgStrSQL = lgStrSQL & " 											C1.BIZ_AREA_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 										FROM A_GL C1, A_GL_ITEM C2  " & vbCr
		lgStrSQL = lgStrSQL & " 										WHERE C1.GL_NO = C2.GL_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C2.DR_CR_FG = " & FilterVar("CR", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 										) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_NO = C.GL_NO    " & vbCr
		lgStrSQL = lgStrSQL & " 										AND B.GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & " 								LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 				AND ISNULL(C.ITEM_SEQ, '') <> ''  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND D.ACCT_TYPE = " & FilterVar("D1", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 				GROUP BY A.NOTE_NO, C.GL_NO, C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT) GL " & vbCr
		lgStrSQL = lgStrSQL & " 		 LEFT JOIN (  SELECT 	A.NOTE_NO, B.STS_DT, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN ( SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL = lgStrSQL & " 								C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL = lgStrSQL & " 								C1.BIZ_AREA_CD    " & vbCr
		lgStrSQL = lgStrSQL & " 							FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2  " & vbCr
		lgStrSQL = lgStrSQL & " 							WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C2.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 						) C ON  B.NOTE_ACCT_CD = C.ACCT_CD and B.TEMP_GL_NO = C.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "									AND B.TEMP_GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN A_ACCT E ON B.NOTE_ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr
		lgStrSQL = lgStrSQL & " 				AND A.NOTE_FG = " & FilterVar("CR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND ISNULL(C.ITEM_SEQ, '') <> '' " & vbCr
		lgStrSQL = lgStrSQL & " 				GROUP BY A.NOTE_NO, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT ) TMP " & vbCr
		lgStrSQL = lgStrSQL & " 				ON GL.NOTE_ACCT_CD = TMP.NOTE_ACCT_CD AND GL.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND GL.BP_CD = TMP.BP_CD AND GL.NOTE_NO = TMP.NOTE_NO " & vbCr
	
	ElseIf Trim(Request("cboNoteFg")) = "CP" Then 
	' 지불구매카드		
		lgStrSQL = lgStrSQL & " 		FROM  (  SELECT  A.NOTE_NO, B.STS_DT, C.GL_NO,  C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_GL_AMT, " & vbCr
		lgStrSQL = lgStrSQL & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 								LEFT JOIN (SELECT C1.GL_NO,  C1.GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,    " & vbCr
		lgStrSQL = lgStrSQL & " 											C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL = lgStrSQL & " 											C1.BIZ_AREA_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 										FROM A_GL C1, A_GL_ITEM C2  " & vbCr
		lgStrSQL = lgStrSQL & " 										WHERE C1.GL_NO = C2.GL_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C2.DR_CR_FG = " & FilterVar("DR", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.REF_NO LIKE " & FilterVar("FC%", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 										AND C1.GL_DT >= " & strFrDt  & " AND C1.GL_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 										) C ON B.NOTE_ACCT_CD = C.ACCT_CD AND B.GL_NO = C.GL_NO    " & vbCr
		lgStrSQL = lgStrSQL & " 										AND B.GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC " & vbCr
		lgStrSQL = lgStrSQL & " 								LEFT JOIN A_ACCT D ON B.NOTE_ACCT_CD = D.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 				AND ISNULL(C.ITEM_SEQ, '') <> ''  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND A.NOTE_FG = " & FilterVar("CP", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				GROUP BY A.NOTE_NO, C.GL_NO, C.GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT) GL " & vbCr
		lgStrSQL = lgStrSQL & " 		 LEFT JOIN (  SELECT 	A.NOTE_NO, B.STS_DT, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(A.STTL_AMT, 0)) SUM_STTL_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					SUM(ISNULL(C.ITEM_LOC_AMT, 0)) SUM_TEMP_AMT,  " & vbCr
		lgStrSQL = lgStrSQL & " 					C.BIZ_AREA_CD, A.BP_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				FROM F_NOTE A	LEFT JOIN F_NOTE_ITEM B ON A.NOTE_NO = B.NOTE_NO  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN ( SELECT C1.TEMP_GL_NO,  C1.TEMP_GL_DT,  C1.REF_NO, C1.GL_INPUT_TYPE,  " & vbCr
		lgStrSQL = lgStrSQL & " 								C2.ITEM_LOC_AMT, C2.ACCT_CD, C2.ITEM_SEQ, C2.ITEM_DESC, " & vbCr
		lgStrSQL = lgStrSQL & " 								C1.BIZ_AREA_CD    " & vbCr
		lgStrSQL = lgStrSQL & " 							FROM A_TEMP_GL C1, A_TEMP_GL_ITEM C2  " & vbCr
		lgStrSQL = lgStrSQL & " 							WHERE C1.TEMP_GL_NO = C2.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C2.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.CONF_FG <> " & FilterVar("C", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.GL_INPUT_TYPE = " & FilterVar("FN", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 							AND C1.TEMP_GL_DT >= " & strFrDt  & " AND C1.TEMP_GL_DT <= " & strToDt  & vbCr
		lgStrSQL = lgStrSQL & " 						) C ON  B.NOTE_ACCT_CD = C.ACCT_CD and B.TEMP_GL_NO = C.TEMP_GL_NO " & vbCr
		lgStrSQL = lgStrSQL & "									AND B.TEMP_GL_SEQ = C.ITEM_SEQ AND B.NOTE_NO = C.ITEM_DESC  " & vbCr
		lgStrSQL = lgStrSQL & " 						LEFT JOIN A_ACCT E ON B.NOTE_ACCT_CD = E.ACCT_CD  " & vbCr
		lgStrSQL = lgStrSQL & " 				WHERE A.NOTE_STS = " & FilterVar("SM", "''", "S") & " " & vbCr
		lgStrSQL = lgStrSQL & " 				AND B.STS_DT >= " & strFrDt  & " AND B.STS_DT <= " & strToDt & vbCr
		lgStrSQL = lgStrSQL & " 				AND A.NOTE_FG = " & FilterVar("CP", "''", "S") & "  " & vbCr
		lgStrSQL = lgStrSQL & " 				AND ISNULL(C.ITEM_SEQ, '') <> '' " & vbCr
		lgStrSQL = lgStrSQL & " 				GROUP BY A.NOTE_NO, C.TEMP_GL_NO, C.TEMP_GL_DT, B.NOTE_ACCT_CD, C.GL_INPUT_TYPE, C.BIZ_AREA_CD, A.BP_CD, B.STS_DT ) TMP " & vbCr
		lgStrSQL = lgStrSQL & " 				ON GL.NOTE_ACCT_CD = TMP.NOTE_ACCT_CD AND GL.BIZ_AREA_CD = TMP.BIZ_AREA_CD  AND GL.BP_CD = TMP.BP_CD AND GL.NOTE_NO = TMP.NOTE_NO " & vbCr
	
	End If 

	lgStrSQL = lgStrSQL & " 	LEFT JOIN	A_ACCT		AC ON AC.ACCT_CD = GL.NOTE_ACCT_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_BIZ_AREA	BA ON BA.BIZ_AREA_CD = GL.BIZ_AREA_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_BIZ_PARTNER  BP ON BP.BP_CD = GL.BP_CD  " & vbCr
	lgStrSQL = lgStrSQL & " 	LEFT JOIN	B_MINOR	MN ON MN.MINOR_CD = GL.GL_INPUT_TYPE AND MN.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  " & vbCr
		
	'내부 TABLE Where 조건 
	lgStrSQL = lgStrSQL & " WHERE GL.GL_DT  >= " & strFrDt  & " AND GL.GL_DT <= " & strToDt & vbCr
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND GL.BIZ_AREA_CD = " & strBizAreaCd & vbCr
	End If 
	
	If Trim(Request("txtAcctCd")) <> "" Then		
		lgStrSQL = lgStrSQL & " AND GL.NOTE_ACCT_CD = " & strAcctCd & vbCr
	End If
		
	If Trim(Request("txtBpCd")) <> ""	Then		
		lgStrSQL = lgStrSQL & " AND GL.BP_CD = " & strBpCd  & vbCr
	End If		
	
	If Trim(Request("txtNoteNoFr")) <> "" Then
		lgStrSQL = lgStrSQL & " AND GL.NOTE_NO >= " & strNoteNoFr  & vbCr
	End If
	
	If Trim(Request("txtNoteNoTo")) <> "" Then
		lgStrSQL = lgStrSQL & " AND GL.NOTE_NO <= " & strNoteNoTo & vbCr
	End If		
	
	'내부 TABLE Group By 조건 
	lgStrSQL = lgStrSQL & "	GROUP BY  GL.NOTE_ACCT_CD, AC.ACCT_NM,  GL.NOTE_NO, GL.STS_DT, GL.GL_DT,  " & vbCr
	lgStrSQL = lgStrSQL & "				GL.GL_NO, TMP.TEMP_GL_NO,  " & vbCr
	lgStrSQL = lgStrSQL & "				GL.GL_INPUT_TYPE, MN.MINOR_NM, " & vbCr
	
	If Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " 	GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, GL.BP_CD, BP.BP_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "Y" and Trim(Request("txtShowBp")) = "N" Then	
		lgStrSQL = lgStrSQL & " 	GL.BIZ_AREA_CD, BA.BIZ_AREA_NM, " & vbCr
	ElseIf Trim(Request("txtShowBiz")) = "N" and Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " 	GL.BP_CD, BP.BP_NM, " & vbCr
	Else 	
		lgStrSQL = lgStrSQL & "" & vbCr
	End If
	lgStrSQL = lgStrSQL & " TEMP_GL_DT) A"	 & vbCr
	
	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " WHERE  ISNULL(STTL_AMT, 0) <> ISNULL(GL_AMT, 0)" & vbCr
	End If 	
	
	lgStrSQL = lgStrSQL & " ORDER BY NOTE_ACCT_CD, NOTE_NO"	 & vbCr

	'Response.write lgStrSQL
    'Response.End 
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKeyIndex = "" 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else 
		iDx         = 1
		lgstrData   = ""
		lgLngMaxRow = Request("txtMaxRows")												'☜: Read Operation Mode (CRUD)
		
		If CDbl(lgStrPrevKeyIndex) > 0 Then
          lgObjRs.Move     = CDbl(lgMaxCount) * CDbl(lgStrPrevKeyIndex)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
        End If

		Do While Not lgObjRs.EOF			
	
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))				'NOTE_ACCT_CD
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))				'ACCT_NM
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))				'NOTE_NO
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(3))		'STS_DT
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(4))		'GL_DT	   
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(5), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	        'STTL_AMT
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(6), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	        'GL_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(7), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'DIFF_AMT
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(8), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")			'TEMP_GL_AMT
			lgstrData = lgstrData & Chr(11) & ""
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))          	'GL_NO
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))				'TEMP_GL_NO
			lgstrData = lgstrData & Chr(11) & ""
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(11))				'GL_INPUT_TYPE
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))				'MINOR_NM
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(13))				'BIZ_AREA_CD
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(14))				'BIZ_AREA_NM
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(15))				'BP_CD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(16))				'BP_NM
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(17))		'TEMP_GL_DT			  
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
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                            " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData2 " & vbCr
       Response.Write  "    Parent.lgStrPrevKey                  =  """ & lgStrPrevKey     & """" & vbCr       
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If
End Sub    

%>

