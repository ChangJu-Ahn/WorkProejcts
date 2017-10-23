<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
	Dim strRadio
 	Dim strFrDt
    Dim strToDt
    DIm strBizAreaCd
    DIm strFrAcctCd
    DIm strToAcctCd
    Dim strFrAsstNo
    Dim strToAsstNo
    Dim lgStrColorFlag

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus  = ""

	Call TrimData()
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubBizQueryMulti()															 '
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : TrimData
' Desc : 
'============================================================================================================
Sub  TrimData()
     strRadio		= Trim(Request("txtRadio"))
     strFrDt		= Trim(Request("txtFrDt"))
     strToDt		= Trim(Request("txtToDt"))
     strBizAreaCd	= Trim(Request("txtBizAreaCd"))
	 strFrAcctCd	= Trim(Request("txtFrAcctCd"))
	 strToAcctCd	= Trim(Request("txtToAcctCd"))
	 strFrAsstNo	= Trim(Request("txtFrAsstCd"))
	 strToAsstNo	= Trim(Request("txtToAsstCd"))
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
    Dim lgStrClsAr
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " " & vbcr

	lgStrSQL = lgStrSQL & " SELECT	AA.ROW_NO," & vbcr
	lgStrSQL = lgStrSQL & " CASE AA.ROW_NO WHEN 3 THEN " & FilterVar("총합계", "''", "S") & " " & vbcr
	lgStrSQL = lgStrSQL & " 			WHEN 2 THEN " & FilterVar("계정합계", "''", "S") & " " & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE   AA.ASST_ACCT_CD END ASST_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 3 THEN NULL" & vbcr
	lgStrSQL = lgStrSQL & " 			WHEN 2 THEN NULL" & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE   B.ACCT_NM END ASST_ACCT_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 1 THEN " & FilterVar("자산합계", "''", "S") & " " & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE AA.ASST_NO END ASST_NO," & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 1 THEN NULL" & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE  C.ASST_NM END ASST_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.DUR_YRS_FG," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_SEQ," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.HIS_FG," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.HIS_FG_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.HIS_DT," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.GL_NO," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.GL_DT," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.TEMP_GL_NO," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.TEMP_GL_DT," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.DEPT_CD," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.DEPT_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.ORG_CHANGE_ID," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_INV_QTY_INC," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_INV_QTY_DEC," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_COST_INC," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_COST_DEC," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.ACCU_DEPR_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.ACCU_DEPR_ACCT_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_ACCU_DEPR_INC," & vbcr
	lgStrSQL = lgStrSQL & " 	AA.HIS_ACCU_DEPR_DEC," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.REF_NO," & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 0 THEN AA.HIS_DUR_YRS" & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE 0 END HIS_DUR_YRS, " & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 0 THEN AA.HIS_DUR_MNTH" & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE 0 END HIS_DUR_MNTH, " & vbcr
	lgStrSQL = lgStrSQL & " 	CASE AA.ROW_NO WHEN 0 THEN AA.HIS_RES_AMT" & vbcr
	lgStrSQL = lgStrSQL & " 			ELSE 0 END HIS_RES_AMT, " & vbcr
	lgStrSQL = lgStrSQL & " 	BB.DEPR_EXP_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.DEPR_EXP_ACCT_NM," & vbcr
	lgStrSQL = lgStrSQL & " 	BB.HIS_DESC" & vbcr
	lgStrSQL = lgStrSQL & "FROM	(" & vbcr
	lgStrSQL = lgStrSQL & "	SELECT	GROUPING(CC.ASST_ACCT_CD) + GROUPING(CC.ASST_NO) + GROUPING(CC.HIS_SEQ) ROW_NO," & vbcr
	lgStrSQL = lgStrSQL & "		CC.ASST_ACCT_CD ASST_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & "		CC.ASST_NO ASST_NO," & vbcr
	lgStrSQL = lgStrSQL & "		MAX(CC.DUR_YRS_FG)	DUR_YRS_FG," & vbcr
	lgStrSQL = lgStrSQL & "		CC.HIS_SEQ HIS_SEQ," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_DUR_YRS) HIS_DUR_YRS," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_DUR_MNTH) HIS_DUR_MNTH," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_RES_AMT) HIS_RES_AMT," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_INV_QTY_INC) HIS_INV_QTY_INC," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_INV_QTY_DEC) HIS_INV_QTY_DEC," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_COST_INC) HIS_COST_INC," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_COST_DEC) HIS_COST_DEC," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_ACCU_DEPR_INC) HIS_ACCU_DEPR_INC," & vbcr
	lgStrSQL = lgStrSQL & "		SUM(CC.HIS_ACCU_DEPR_DEC) HIS_ACCU_DEPR_DEC" & vbcr
	lgStrSQL = lgStrSQL & "	FROM	A_ASSET_HIS_DETAIL_DD CC JOIN A_ASSET_MASTER DD ON CC.ASST_NO = DD.ASST_NO " & vbcr
	lgStrSQL = lgStrSQL & "									 LEFT OUTER JOIN A_GL EE ON CC.GL_NO = EE.GL_NO " & vbcr
	if strRadio = "01" then		'gl_dt
		lgStrSQL = lgStrSQL & "	WHERE	EE.GL_DT >=  " & FilterVar(strFrDt , "''", "S") & "" & vbcr
		lgStrSQL = lgStrSQL & "	AND		EE.GL_DT <=  " & FilterVar(strToDt , "''", "S") & "" & vbcr
    else
		lgStrSQL = lgStrSQL & "	WHERE	CC.HIS_DT >=  " & FilterVar(strFrDt , "''", "S") & "" & vbcr
		lgStrSQL = lgStrSQL & "	AND		CC.HIS_DT <=  " & FilterVar(strToDt , "''", "S") & "" & vbcr
    end if
    
    if strBizAreaCd <> "" then
		lgStrSQL = lgStrSQL & "	AND		DD.BIZ_AREA_CD =  " & FilterVar(strBizAreaCd , "''", "S") & "" & vbcr
    end if
    if strFrAcctCd <> "" then
		lgStrSQL = lgStrSQL & "	AND		CC.ASST_ACCT_CD >=  " & FilterVar(strFrAcctCd , "''", "S") & "" & vbcr
    end if
    if strToAcctCd <> "" then
		lgStrSQL = lgStrSQL & "	AND		CC.ASST_ACCT_CD <=  " & FilterVar(strToAcctCd , "''", "S") & "" & vbcr
    end if
    if strFrAsstNo <> "" then
		lgStrSQL = lgStrSQL & "	AND		CC.ASST_NO >=  " & FilterVar(strFrAsstNo , "''", "S") & "" & vbcr
    end if
    if strToAsstNo <> "" then
		lgStrSQL = lgStrSQL & "	AND		CC.ASST_NO <=  " & FilterVar(strToAsstNo , "''", "S") & "" & vbcr
    end if
	lgStrSQL = lgStrSQL & "	GROUP BY CC.ASST_ACCT_CD, CC.ASST_NO, CC.HIS_SEQ" & vbcr
	lgStrSQL = lgStrSQL & "	WITH ROLLUP" & vbcr
	lgStrSQL = lgStrSQL & "	) AA " & vbcr
	lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN (" & vbcr
	lgStrSQL = lgStrSQL & "				SELECT	A.ASST_NO		ASST_NO," & vbcr
	lgStrSQL = lgStrSQL & "					A.DUR_YRS_FG		DUR_YRS_FG," & vbcr
	lgStrSQL = lgStrSQL & "					A.HIS_SEQ		HIS_SEQ," & vbcr
	lgStrSQL = lgStrSQL & "					A.HIS_FG		HIS_FG," & vbcr
	lgStrSQL = lgStrSQL & "					I.MINOR_NM		HIS_FG_NM," & vbcr
	lgStrSQL = lgStrSQL & "					A.HIS_DT		HIS_DT," & vbcr
	lgStrSQL = lgStrSQL & "					A.DEPT_CD		DEPT_CD," & vbcr
	lgStrSQL = lgStrSQL & "					D.DEPT_NM		DEPT_NM," & vbcr
	lgStrSQL = lgStrSQL & "					A.ORG_CHANGE_ID	ORG_CHANGE_ID," & vbcr
	lgStrSQL = lgStrSQL & "					A.HIS_DESC		HIS_DESC," & vbcr
	lgStrSQL = lgStrSQL & "					A.GL_NO			GL_NO," & vbcr
	lgStrSQL = lgStrSQL & "					E.GL_DT			GL_DT," & vbcr
	lgStrSQL = lgStrSQL & "					A.TEMP_GL_NO		TEMP_GL_NO," & vbcr
	lgStrSQL = lgStrSQL & "					F.TEMP_GL_DT		TEMP_GL_DT," & vbcr
	lgStrSQL = lgStrSQL & "					A.ACCU_DEPR_ACCT_CD	ACCU_DEPR_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & "					G.ACCT_NM		ACCU_DEPR_ACCT_NM," & vbcr
	lgStrSQL = lgStrSQL & "					A.DEPR_EXP_ACCT_CD	DEPR_EXP_ACCT_CD," & vbcr
	lgStrSQL = lgStrSQL & "					H.ACCT_NM		DEPR_EXP_ACCT_NM," & vbcr
	lgStrSQL = lgStrSQL & "					A.REF_NO		REF_NO" & vbcr
	lgStrSQL = lgStrSQL & "				FROM	A_ASSET_HIS_DETAIL_DD A 	JOIN B_ACCT_DEPT D ON A.DEPT_CD = D.DEPT_CD AND A.ORG_CHANGE_ID = D.ORG_CHANGE_ID" & vbcr
	lgStrSQL = lgStrSQL & "									LEFT OUTER JOIN A_GL E ON A.GL_NO = E.GL_NO" & vbcr
	lgStrSQL = lgStrSQL & "									LEFT OUTER JOIN A_TEMP_GL F ON A.TEMP_GL_NO = F.TEMP_GL_NO" & vbcr
	lgStrSQL = lgStrSQL & "									LEFT OUTER JOIN A_ACCT G ON A.ACCU_DEPR_ACCT_CD = G.ACCT_CD" & vbcr
	lgStrSQL = lgStrSQL & "									LEFT OUTER JOIN A_ACCT H ON A.DEPR_EXP_ACCT_CD = H.ACCT_CD" & vbcr
	lgStrSQL = lgStrSQL & "									JOIN B_MINOR I ON I.MAJOR_CD = " & FilterVar("A2011", "''", "S") & "  AND A.HIS_FG = I.MINOR_CD" & vbcr
	lgStrSQL = lgStrSQL & "			) BB ON AA.ASST_NO = BB.ASST_NO AND AA.DUR_YRS_FG = BB.DUR_YRS_FG AND AA.HIS_SEQ = BB.HIS_SEQ" & vbcr
	lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN A_ACCT B ON AA.ASST_ACCT_CD = B.ACCT_CD" & vbcr
	lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN A_ASSET_MASTER C ON AA.ASST_NO = C.ASST_NO" & vbcr

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else
		iDx = 1
		lgstrData = ""
		lgStrColorFlag = ""
		lgLngMaxRow = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
        
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ROW_NO"))
			
			lgStrColorFlag = lgStrColorFlag & CStr(iDx) & gColSep & lgObjRs("ROW_NO") & gRowSep
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DUR_YRS_FG"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HIS_FG"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HIS_FG_NM"))
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("HIS_COST_INC"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("HIS_COST_DEC"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCU_DEPR_ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCU_DEPR_ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("HIS_ACCU_DEPR_INC"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("HIS_ACCU_DEPR_DEC"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REF_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("HIS_DT"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TEMP_GL_DT"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPR_EXP_ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPR_EXP_ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ORG_CHANGE_ID"))
'			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_INV_QTY_INC"), ggQtyNo.DecPoint, 0)
'			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_INV_QTY_DEC"), ggQtyNo.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_INV_QTY_INC"), 0, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_INV_QTY_DEC"), 0, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_DUR_YRS"), 0, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_DUR_MNTH"), 0, 0)
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("HIS_RES_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("HIS_DESC"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("HIS_SEQ"), 0, 0)
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext

			iDx =  iDx + 1
		Loop
	End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
		Response.Write  " <Script Language=vbscript>                                 " & vbCr
		Response.Write  "   Parent.ggoSpread.Source     = Parent.frm1.vspdData     " & vbCr
		Response.Write  "   Parent.ggoSpread.SSShowDataByClip   """ & lgstrData      & """" & vbCr
		Response.Write	"	parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
		Response.Write  "   Parent.DBQueryOk										" & vbCr     
		Response.Write  " </Script>													" & vbCr
    End If
    
    Response.End 
End Sub

%>

 
