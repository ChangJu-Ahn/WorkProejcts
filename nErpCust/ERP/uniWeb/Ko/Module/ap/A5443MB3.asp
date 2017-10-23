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
    Dim strGLFrDt
    Dim strGLToDt
    DIm strBizAreaCd
    DIm strAcctCd   
    DIm strDealBpCd
    Dim strInputType
    Dim ShowBiz
    Dim ShowBp
	Dim lgStrpage
    
    Dim lgStrPrevKey
    
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus  = ""
    lgStrPrevKey   = Trim(Request("lgStrPrevKey"))									 '☜: Next Key    
    lgStrpage   = lgStrPrevKey
	Call TrimData()
    Call SubOpenDB2_(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubBizQueryMulti()             '
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

Sub SubOpenDB2_(pObjConn)
    On Error Resume Next
    Err.Clear

	Set pObjConn = Server.CreateObject("ADODB.Connection")

	pObjConn.ConnectionString  = gADODBConnString
	pObjConn.commandtimeout  = 600
	pObjConn.Open

    If CheckSYSTEMError(Err,True) = True Then
    End If

End Sub


'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
    strFrDt	= FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S")
    strToDt	= FilterVar(UNIConvDate(Request("txtToDt")), "''", "S")
	strBizAreaCd = FilterVar(Request("txtBizAreaCd"), "''", "S") 
	strAcctCd = FilterVar(Request("txtAcctCd"), "''", "S") 
	strDealBpCd = FilterVar(Request("txtDealBpCd"), "''", "S") 
	strInputType = FilterVar(Request("txtGlInputType"), "''", "S") 
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
    Dim lgAllcTotLocAmt , lgDiffTotLocAmt , lgGlTotLocAmt , lgTempGlLocAmt
    
    Const C_SHEETMAXROWS_D = 100															'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT TOP " & C_SHEETMAXROWS_D + 1
	lgStrSQL = lgStrSQL & " A.ACCT_CD, B.ACCT_NM, A.ALLC_NO "
	lgStrSQL = lgStrSQL & " , CONVERT(CHAR(10),A.ALLC_DT,20) ALLC_DT, CONVERT(CHAR(10),I.GL_DT,20) GL_DT"
	lgStrSQL = lgStrSQL & " , ISNULL(A.ALLC_LOC_AMT,0)-ISNULL(D.ITEM_LOC_AMT,0) DIFF_AMT"
	lgStrSQL = lgStrSQL & " , ISNULL(A.ALLC_LOC_AMT,0) ALLC_LOC_AMT, ISNULL(D.ITEM_LOC_AMT,0) GL_ITEM_LOC_AMT"
	lgStrSQL = lgStrSQL & " , ISNULL(F.ITEM_LOC_AMT,0) GL_TEMP_ITEM_LOC_AMT, ISNULL(E.ITEM_LOC_AMT,0) BATCH_ITEM_AMT"
	lgStrSQL = lgStrSQL & " , A.GL_NO,A.TEMP_GL_NO,E.BATCH_NO,F.TEMP_GL_DT,A.GL_INPUT_TYPE,C.MINOR_NM  "
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " , E.BIZ_AREA_CD,G.BIZ_AREA_NM" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " , A.PAY_BP_CD,H.BP_NM" 

	lgStrSQL = lgStrSQL & " FROM ("
	lgStrSQL = lgStrSQL & "		SELECT A.CLS_AP_NO ALLC_NO, A.ACCT_CD, A.CLS_DT ALLC_DT,  SUM(A.CLS_LOC_AMT)  ALLC_LOC_AMT,A.GL_NO, A.TEMP_GL_NO, A.GL_INPUT_TYPE"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " , A.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & "	, A.PAY_BP_CD"
	
	lgStrSQL = lgStrSQL & "		FROM (  SELECT A.AP_NO, A.CLS_AP_NO, A.CLS_DT, A.ACCT_CD, A.CLS_LOC_AMT, A.GL_NO, A.TEMP_GL_NO, A.GL_INPUT_TYPE"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " , B.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " , B.PAY_BP_CD"
	
	lgStrSQL = lgStrSQL & "				FROM (SELECT A.AP_NO, CASE WHEN B.ALLC_TYPE = " & FilterVar("B", "''", "S") & "  THEN B.REF_NO ELSE CLS_AP_NO END CLS_AP_NO, A.CLS_DT, A.ACCT_CD ACCT_CD, A.CLS_AMT+A.DC_AMT CLS_AMT, A.CLS_LOC_AMT+A.DC_LOC_AMT CLS_LOC_AMT, B.GL_NO, B.TEMP_GL_NO,B.GL_INPUT_TYPE"						
	lgStrSQL = lgStrSQL & "					  FROM A_CLS_AP A JOIN (SELECT PAYM_NO, GL_NO,	TEMP_GL_NO, A_ALLC_PAYM.REF_NO, ALLC_TYPE,"
	lgStrSQL = lgStrSQL & "												CASE WHEN ALLC_TYPE = " & FilterVar("X", "''", "S") & "  THEN " & FilterVar("PX", "''", "S") & "  WHEN ALLC_TYPE = " & FilterVar("P", "''", "S") & "  THEN " & FilterVar("LP", "''", "S") & "  ELSE " & FilterVar("LR", "''", "S") & "   END GL_INPUT_TYPE "	
	lgStrSQL = lgStrSQL & "											FROM A_ALLC_PAYM UNION "
	lgStrSQL = lgStrSQL & "											SELECT ALLC_NO, GL_NO, TEMP_GL_NO, '' , '' , " & FilterVar("CR", "''", "S") & "  GL_INPUT_TYPE  FROM A_ALLC_RCPT UNION"
	lgStrSQL = lgStrSQL & "											SELECT CLEAR_NO, GL_NO, '', '', '', " & FilterVar("CL", "''", "S") & "   GL_INPUT_TYPE FROM A_CLEAR_AP_AR"	
	lgStrSQL = lgStrSQL & "											) B ON B.PAYM_NO = A.CLS_AP_NO"
	lgStrSQL = lgStrSQL & "					 UNION ALL"
	lgStrSQL = lgStrSQL & "				SELECT A.AP_NO, A.ADJUST_NO, A.ADJUST_DT, B.ACCT_CD,A.ADJUST_AMT, A.ADJUST_LOC_AMT,  A.GL_NO, A.TEMP_GL_NO, " & FilterVar("JP", "''", "S") & "  GL_INPUT_TYPE"
	lgStrSQL = lgStrSQL & "				FROM A_AP_ADJUST A JOIN A_GL_ITEM B ON B.GL_NO = A.GL_NO AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & " "
	lgStrSQL = lgStrSQL & "			)A INNER JOIN A_OPEN_AP B ON A.AP_NO=B.AP_NO"
	lgStrSQL = lgStrSQL & "		)A"
	
	If Trim(Request("txtShowBp")) = "Y" Then  lgStrSQL = lgStrSQL & " LEFT JOIN  B_BIZ_PARTNER E ON A.PAY_BP_CD = E.BP_CD"
	
	lgStrSQL = lgStrSQL & "		WHERE A.CLS_DT <= " & strToDt & " AND A.CLS_DT  >= " & strFrDt & " AND (A.GL_NO <> '' OR A.TEMP_GL_NO <> '')"
	lgStrSQL = lgStrSQL & "			AND A.ACCT_CD IN (SELECT ACCT_CD FROM A_ACCT WHERE ACCT_TYPE LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL = lgStrSQL & "		GROUP BY A.CLS_AP_NO, A.ACCT_CD,A.CLS_DT,A.GL_INPUT_TYPE,A.GL_NO, A.TEMP_GL_NO"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " , A.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " , A.PAY_BP_CD"

	lgStrSQL = lgStrSQL & ") A "

	lgStrSQL = lgStrSQL & "	LEFT JOIN (	SELECT  A_BATCH.BATCH_NO,  A_BATCH.REF_NO,A_BATCH_GL_ITEM.ACCT_CD,"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " C.BIZ_AREA_CD,"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " C.PAY_BP_CD, " 
	
	lgStrSQL = lgStrSQL & "					SUM(CASE WHEN A_BATCH_GL_ITEM.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*A_BATCH_GL_ITEM.ITEM_LOC_AMT"
	lgStrSQL = lgStrSQL & "							 ELSE A_BATCH_GL_ITEM.ITEM_LOC_AMT  END) ITEM_LOC_AMT"
	lgStrSQL = lgStrSQL & "				FROM A_BATCH , A_BATCH_GL_ITEM"
	
	lgStrSQL = lgStrSQL & "				LEFT JOIN A_OPEN_AP  C ON C.AP_NO=A_BATCH_GL_ITEM.KEY_VAL1"
	
	lgStrSQL = lgStrSQL & "				WHERE  A_BATCH.BATCH_NO=A_BATCH_GL_ITEM.BATCH_NO" 	
	lgStrSQL = lgStrSQL & "					AND A_BATCH_GL_ITEM.JNL_CD IN (select distinct(jnl_cd) from a_jnl_acct_assn"
	lgStrSQL = lgStrSQL & "												where acct_cd in(select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " ))"
	lgStrSQL = lgStrSQL & "					AND  A_BATCH.GL_DT >=" & strFrDt & "AND A_BATCH.GL_DT <= " & strToDt
	lgStrSQL = lgStrSQL & "				GROUP BY A_BATCH.BATCH_NO,  A_BATCH.REF_NO,A_BATCH_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & ", C.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " , C.PAY_BP_CD " 
	
	lgStrSQL = lgStrSQL & "			) E ON  A.ALLC_NO=E.REF_NO  AND A.ACCT_CD= E.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " AND E.BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " AND E.PAY_BP_CD= A.PAY_BP_CD"
	
	lgStrSQL = lgStrSQL & "	LEFT JOIN (SELECT A_GL.REF_NO,SUM(A_GL_ITEM.ITEM_LOC_AMT) ITEM_LOC_AMT,A_GL.GL_DT,A_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " , G.CTRL_VAL BP_CD"
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL = lgStrSQL & " ,A_GL.BIZ_AREA_CD"
	
	lgStrSQL = lgStrSQL & "			   FROM A_GL,A_GL_ITEM 	LEFT JOIN A_GL_DTL G ON A_GL_ITEM.GL_NO=G.GL_NO AND A_GL_ITEM.ITEM_SEQ=G.ITEM_SEQ AND G.CTRL_CD IN ( Select CTRL_CD From A_CTRL_ITEM Where  TBL_ID = " & FilterVar("B_BIZ_PARTNER", "''", "S") & " )"
	lgStrSQL = lgStrSQL & "			   WHERE  A_GL.GL_NO=A_GL_ITEM.GL_NO AND A_GL_ITEM.ACCT_CD IN (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL = lgStrSQL & "					AND  A_GL.GL_DT >= " & strFrDt & "  AND  A_GL.GL_DT <= " & strToDt
	lgStrSQL = lgStrSQL & "			   GROUP BY A_GL.REF_NO,A_GL.GL_DT,A_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL = lgStrSQL & " ,A_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " ,G.CTRL_VAL"
	
	lgStrSQL = lgStrSQL & "			) D ON A.ALLC_NO=D.REF_NO  AND A.ACCT_CD= D.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL = lgStrSQL & " AND D.BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " AND D.BP_CD= A.PAY_BP_CD"
	
	lgStrSQL = lgStrSQL & " LEFT JOIN (	SELECT A_TEMP_GL.REF_NO, SUM(A_TEMP_GL_ITEM.ITEM_LOC_AMT) ITEM_LOC_AMT, A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A_TEMP_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " , G.CTRL_VAL"
	
	lgStrSQL = lgStrSQL & "				FROM A_TEMP_GL, A_TEMP_GL_ITEM "
	
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " LEFT JOIN A_TEMP_GL_DTL G ON A_TEMP_GL_ITEM.TEMP_GL_NO=G.TEMP_GL_NO AND A_TEMP_GL_ITEM.ITEM_SEQ=G.ITEM_SEQ AND G.CTRL_CD=" & FilterVar("BP", "''", "S") & " "
	
	lgStrSQL = lgStrSQL & "				WHERE  A_TEMP_GL.TEMP_GL_NO=A_TEMP_GL_ITEM.TEMP_GL_NO"
	lgStrSQL = lgStrSQL & "					AND  A_TEMP_GL.TEMP_GL_DT >= " & strFrDt & " AND  A_TEMP_GL.TEMP_GL_DT <=" & strToDt
	lgStrSQL = lgStrSQL & "					AND A_TEMP_GL_ITEM.ACCT_CD IN (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL = lgStrSQL & "					AND A_TEMP_GL.CONF_FG<>" & FilterVar("C", "''", "S") & " "
	lgStrSQL = lgStrSQL & "				GROUP BY A_TEMP_GL.REF_NO,  A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A_TEMP_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " , G.CTRL_VAL"

	lgStrSQL = lgStrSQL & "			) F ON A.ALLC_NO=F.REF_NO AND A.ACCT_CD=F.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " AND F. BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL = lgStrSQL & " AND F. CTRL_VAL= A.PAY_BP_CD"
	
	lgStrSQL = lgStrSQL & " LEFT JOIN A_ACCT B ON A.ACCT_CD=B.ACCT_CD"
	lgStrSQL = lgStrSQL & " LEFT JOIN B_MINOR C ON (A.GL_INPUT_TYPE=C.MINOR_CD  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & " )"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & "	LEFT JOIN B_BIZ_AREA G ON E.BIZ_AREA_CD=G.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & "	LEFT JOIN B_BIZ_PARTNER H ON A.PAY_BP_CD=H.BP_CD"
	
	lgStrSQL = lgStrSQL & "	LEFT JOIN A_GL I ON A.GL_NO=I.GL_NO"
	lgStrSQL = lgStrSQL & "	WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '')"
	lgStrSQL = lgStrSQL & "	 AND  A.ALLC_DT <=  " & strToDt & "  AND  A.ALLC_DT >= " & strFrDt 
	lgStrSQL = lgStrSQL & "  AND  A.ALLC_NO >= " & FilterVar(lgStrPrevKey, "''", "S")
	
	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL = lgStrSQL & " AND  a.GL_INPUT_TYPE  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND E.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.PAY_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL = lgStrSQL & " AND ISNULL(A.ALLC_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "
	lgStrSQL = lgStrSQL & " ORDER BY A.ALLC_NO "
	
	
	
	
	lgStrSQL2 = ""
	lgStrSQL2 =	lgStrSQL2 & " SELECT SUM( ISNULL(A.ALLC_LOC_AMT,0)) ALLC_TOT_LOC_AMT,SUM( ISNULL(D.ITEM_LOC_AMT,0)) GL_TOT_ITEM_LOC_AMT"
	lgStrSQL2 = lgStrSQL2 & ",SUM(ISNULL(F.ITEM_LOC_AMT,0)) GL_TEMP_TOT_ITEM_LOC_AMT, SUM(ISNULL(E.ITEM_LOC_AMT,0)) BATCH_TOT_ITEM_AMT"
	lgStrSQL2 = lgStrSQL2 & ",SUM( ISNULL(A.ALLC_LOC_AMT,0)) - SUM( ISNULL(D.ITEM_LOC_AMT,0))  Diff_TOT_ITEM_AMT"
	lgStrSQL2 = lgStrSQL2 & " FROM ("
	lgStrSQL2 = lgStrSQL2 & "		SELECT A.CLS_AP_NO ALLC_NO, A.ACCT_CD, A.CLS_DT ALLC_DT,  SUM(A.CLS_LOC_AMT)  ALLC_LOC_AMT,A.GL_NO, A.TEMP_GL_NO, A.GL_INPUT_TYPE"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , A.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & "	, A.PAY_BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & "		FROM (  SELECT A.AP_NO, A.CLS_AP_NO, A.CLS_DT, A.ACCT_CD, A.CLS_LOC_AMT, A.GL_NO, A.TEMP_GL_NO, A.GL_INPUT_TYPE"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , B.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , B.PAY_BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & "				FROM (SELECT A.AP_NO, CASE WHEN B.ALLC_TYPE = " & FilterVar("B", "''", "S") & "  THEN B.REF_NO ELSE CLS_AP_NO END CLS_AP_NO, A.CLS_DT, A.ACCT_CD ACCT_CD, A.CLS_AMT+A.DC_AMT CLS_AMT, A.CLS_LOC_AMT+A.DC_LOC_AMT CLS_LOC_AMT, B.GL_NO, B.TEMP_GL_NO,B.GL_INPUT_TYPE"						
	lgStrSQL2 = lgStrSQL2 & "					  FROM A_CLS_AP A JOIN (SELECT PAYM_NO, GL_NO,	TEMP_GL_NO, A_ALLC_PAYM.REF_NO, ALLC_TYPE,"
	lgStrSQL2 = lgStrSQL2 & "												CASE WHEN ALLC_TYPE = " & FilterVar("X", "''", "S") & "  THEN " & FilterVar("PX", "''", "S") & "  WHEN ALLC_TYPE = " & FilterVar("P", "''", "S") & "  THEN " & FilterVar("LP", "''", "S") & "  ELSE " & FilterVar("LR", "''", "S") & "   END GL_INPUT_TYPE "	
	lgStrSQL2 = lgStrSQL2 & "											FROM A_ALLC_PAYM UNION "
	lgStrSQL2 = lgStrSQL2 & "											SELECT ALLC_NO, GL_NO, TEMP_GL_NO, '' , '' , " & FilterVar("CR", "''", "S") & "  GL_INPUT_TYPE  FROM A_ALLC_RCPT UNION"
	lgStrSQL2 = lgStrSQL2 & "											SELECT CLEAR_NO, GL_NO, '', '', '', " & FilterVar("CL", "''", "S") & "   GL_INPUT_TYPE FROM A_CLEAR_AP_AR"	
	lgStrSQL2 = lgStrSQL2 & "											) B ON B.PAYM_NO = A.CLS_AP_NO"
	lgStrSQL2 = lgStrSQL2 & "					 UNION ALL"
	lgStrSQL2 = lgStrSQL2 & "				SELECT A.AP_NO, A.ADJUST_NO, A.ADJUST_DT, B.ACCT_CD,A.ADJUST_AMT, A.ADJUST_LOC_AMT,  A.GL_NO, A.TEMP_GL_NO, " & FilterVar("JP", "''", "S") & "  GL_INPUT_TYPE"
	lgStrSQL2 = lgStrSQL2 & "				FROM A_AP_ADJUST A JOIN A_GL_ITEM B ON B.GL_NO = A.GL_NO AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & " "
	lgStrSQL2 = lgStrSQL2 & "			)A INNER JOIN A_OPEN_AP B ON A.AP_NO=B.AP_NO"
	lgStrSQL2 = lgStrSQL2 & "		)A"
	
	If Trim(Request("txtShowBp")) = "Y" Then  lgStrSQL2 = lgStrSQL2 & " LEFT JOIN  B_BIZ_PARTNER E ON A.PAY_BP_CD = E.BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & "		WHERE A.CLS_DT <= " & strToDt & " AND A.CLS_DT  >= " & strFrDt & " AND (A.GL_NO <> '' OR A.TEMP_GL_NO <> '')"
	lgStrSQL2 = lgStrSQL2 & "			AND A.ACCT_CD IN (SELECT ACCT_CD FROM A_ACCT WHERE ACCT_TYPE LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL2 = lgStrSQL2 & "		GROUP BY A.CLS_AP_NO, A.ACCT_CD,A.CLS_DT,A.GL_INPUT_TYPE,A.GL_NO, A.TEMP_GL_NO"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , A.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , A.PAY_BP_CD" 

	lgStrSQL2 = lgStrSQL2 & ") A "

	lgStrSQL2 = lgStrSQL2 & "	LEFT JOIN (	SELECT  A_BATCH.BATCH_NO,  A_BATCH.REF_NO,A_BATCH_GL_ITEM.ACCT_CD,"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " C.BIZ_AREA_CD,"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " C.PAY_BP_CD, " 
	
	lgStrSQL2 = lgStrSQL2 & "					SUM(CASE WHEN A_BATCH_GL_ITEM.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*A_BATCH_GL_ITEM.ITEM_LOC_AMT"
	lgStrSQL2 = lgStrSQL2 & "							 ELSE A_BATCH_GL_ITEM.ITEM_LOC_AMT  END) ITEM_LOC_AMT"
	lgStrSQL2 = lgStrSQL2 & "				FROM A_BATCH , A_BATCH_GL_ITEM"
	
	lgStrSQL2 = lgStrSQL2 & "				LEFT JOIN A_OPEN_AP  C ON C.AP_NO=A_BATCH_GL_ITEM.KEY_VAL1"
	
	lgStrSQL2 = lgStrSQL2 & "				WHERE  A_BATCH.BATCH_NO=A_BATCH_GL_ITEM.BATCH_NO" 	
	lgStrSQL2 = lgStrSQL2 & "					AND A_BATCH_GL_ITEM.JNL_CD IN (select distinct(jnl_cd) from a_jnl_acct_assn"
	lgStrSQL2 = lgStrSQL2 & "												where acct_cd in(select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " ))"
	lgStrSQL2 = lgStrSQL2 & "					AND  A_BATCH.GL_DT >=" & strFrDt & "AND A_BATCH.GL_DT <= " & strToDt
	lgStrSQL2 = lgStrSQL2 & "				GROUP BY A_BATCH.BATCH_NO,  A_BATCH.REF_NO,A_BATCH_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & ", C.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " , C.PAY_BP_CD " 
	
	lgStrSQL2 = lgStrSQL2 & "			) E ON  A.ALLC_NO=E.REF_NO  AND A.ACCT_CD= E.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " AND E.BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " AND E.PAY_BP_CD= A.PAY_BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & "	LEFT JOIN (SELECT A_GL.REF_NO,SUM(A_GL_ITEM.ITEM_LOC_AMT) ITEM_LOC_AMT,A_GL.GL_DT,A_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " , G.CTRL_VAL BP_CD"
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " ,A_GL.BIZ_AREA_CD"
	
	lgStrSQL2 = lgStrSQL2 & "			   FROM A_GL,A_GL_ITEM 	LEFT JOIN A_GL_DTL G ON A_GL_ITEM.GL_NO=G.GL_NO AND A_GL_ITEM.ITEM_SEQ=G.ITEM_SEQ AND G.CTRL_CD IN ( Select CTRL_CD From A_CTRL_ITEM Where  TBL_ID = " & FilterVar("B_BIZ_PARTNER", "''", "S") & " )"
	lgStrSQL2 = lgStrSQL2 & "			   WHERE  A_GL.GL_NO=A_GL_ITEM.GL_NO AND A_GL_ITEM.ACCT_CD IN (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL2 = lgStrSQL2 & "					AND  A_GL.GL_DT >= " & strFrDt & "  AND  A_GL.GL_DT <= " & strToDt
	lgStrSQL2 = lgStrSQL2 & "			   GROUP BY A_GL.REF_NO,A_GL.GL_DT,A_GL_ITEM.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " ,A_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " ,G.CTRL_VAL"
	
	lgStrSQL2 = lgStrSQL2 & "			) D ON A.ALLC_NO=D.REF_NO  AND A.ACCT_CD= D.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " AND D.BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " AND D.BP_CD= A.PAY_BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (	SELECT A_TEMP_GL.REF_NO, SUM(A_TEMP_GL_ITEM.ITEM_LOC_AMT) ITEM_LOC_AMT, A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " ,A_TEMP_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " , G.CTRL_VAL"
	
	lgStrSQL2 = lgStrSQL2 & "				FROM A_TEMP_GL, A_TEMP_GL_ITEM "
	
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " LEFT JOIN A_TEMP_GL_DTL G ON A_TEMP_GL_ITEM.TEMP_GL_NO=G.TEMP_GL_NO AND A_TEMP_GL_ITEM.ITEM_SEQ=G.ITEM_SEQ AND G.CTRL_CD in ( Select CTRL_CD From A_CTRL_ITEM Where  TBL_ID = " & FilterVar("B_BIZ_PARTNER", "''", "S") & " )"
	
	lgStrSQL2 = lgStrSQL2 & "				WHERE  A_TEMP_GL.TEMP_GL_NO=A_TEMP_GL_ITEM.TEMP_GL_NO"
	lgStrSQL2 = lgStrSQL2 & "					AND  A_TEMP_GL.TEMP_GL_DT >= " & strFrDt & " AND  A_TEMP_GL.TEMP_GL_DT <=" & strToDt
	lgStrSQL2 = lgStrSQL2 & "					AND A_TEMP_GL_ITEM.ACCT_CD IN (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%J%", "''", "S") & " )"
	lgStrSQL2 = lgStrSQL2 & "					AND A_TEMP_GL.CONF_FG<>" & FilterVar("C", "''", "S") & " "
	lgStrSQL2 = lgStrSQL2 & "				GROUP BY A_TEMP_GL.REF_NO,  A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " ,A_TEMP_GL.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " , G.CTRL_VAL"

	lgStrSQL2 = lgStrSQL2 & "			) F ON A.ALLC_NO=F.REF_NO AND A.ACCT_CD=F.ACCT_CD"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & " AND F. BIZ_AREA_CD= A.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then lgStrSQL2 = lgStrSQL2 & " AND F. CTRL_VAL= A.PAY_BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN A_ACCT B ON A.ACCT_CD=B.ACCT_CD"
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN B_MINOR C ON (A.GL_INPUT_TYPE=C.MINOR_CD  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & " )"
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & "	LEFT JOIN B_BIZ_AREA G ON E.BIZ_AREA_CD=G.BIZ_AREA_CD"
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL2 = lgStrSQL2 & "	LEFT JOIN B_BIZ_PARTNER H ON A.PAY_BP_CD=H.BP_CD"
	
	lgStrSQL2 = lgStrSQL2 & "	LEFT JOIN A_GL I ON A.GL_NO=I.GL_NO"
	lgStrSQL2 = lgStrSQL2 & "	WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '')"
	lgStrSQL2 = lgStrSQL2 & "	 AND  A.ALLC_DT <=  " & strToDt & "  AND  A.ALLC_DT >= " & strFrDt 
	lgStrSQL2 = lgStrSQL2 & "  AND  A.ALLC_NO >= " & FilterVar(lgStrPrevKey, "''", "S")
	
	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL2 = lgStrSQL2 & " AND  a.GL_INPUT_TYPE  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND E.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.PAY_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL2 = lgStrSQL2 & " AND ISNULL(A.ALLC_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "
	  

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'☜ : No data is found.
		lgStrPrevKey   = Trim(Request("lgStrPrevKey"))									'☜: Next Key
        lgErrorStatus  = "YES"
        Exit Sub
	Else
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = Request("txtMaxRows")												'☜: Read Operation Mode (CRUD)
       
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))

			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ALLC_DT"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))      
			
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("ALLC_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DIFF_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TEMP_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("BATCH_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BATCH_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TEMP_GL_DT")) 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_INPUT_TYPE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLC_NO"))
			 
			
			If Trim(Request("txtShowBiz")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & ""
			End If
			If Trim(Request("txtShowBp")) = "Y" Then			
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_BP_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & ""
			End If
    			  
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
          
			lgObjRs.MoveNext

			iDx =  iDx + 1

			If iDx > C_SHEETMAXROWS_D Then
			    Exit Do
			End If   
		Loop 
		
		If Not lgObjRs.EOF Then
		   lgStrPrevKey = lgObjRs("ALLC_NO")
		Else
		   lgStrPrevKey = ""
		End If
	End If

	If lgStrpage = "" Then
		'*********************************
		'			합계찍기 
		'*********************************
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
			lgAllcTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("ALLC_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgDiffTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("Diff_TOT_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgGlTotLocAmt  = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TOT_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgTempGlLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TEMP_TOT_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotAllcLocAmt3.text     =  """ & lgAllcTotLocAmt    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & lgDiffTotLocAmt & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & lgGlTotLocAmt    & """" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & lgTempGlLocAmt   & """" & vbCr                     
			Response.Write  " </Script>															    " & vbCr
		Else
			Response.Write  " <Script Language=vbscript>                                             " & vbCr
			Response.Write  "    Parent.frm1.txtTotAllcLocAmt3.text     =  0" & vbCr
			Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  0" & vbCr
			Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     = 0" & vbCr
			Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  0" & vbCr                     
			Response.Write  " </Script>															    " & vbCr
		End If    
	End If
	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                                             " & vbCr
       Response.Write  "    Parent.ggoSpread.Source              = Parent.frm1.vspdData3        " & vbCr
       Response.Write  "    Parent.lgStrPrevKey                  =  """ & lgStrPrevKey     & """" & vbCr       
       Response.Write  "    Parent.ggoSpread.SSShowData             """ & lgstrData        & """" & vbCr
'       Response.Write  "    Parent.frm1.txtTotAllcLocAmt3.text     =  """ & lgAllcTotLocAmt    & """" & vbCr
'       Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & lgDiffTotLocAmt & """" & vbCr
'       Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & lgGlTotLocAmt    & """" & vbCr
'       Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & lgTempGlLocAmt   & """" & vbCr                     
       Response.Write  "    Parent.DBQueryOk												    " & vbCr      
       Response.Write  " </Script>															    " & vbCr
    End If
	
	Response.End
End Sub    


%>

