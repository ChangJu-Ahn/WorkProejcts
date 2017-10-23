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
    strFrDt	= FilterVar(UNIConvDate(Request("txtPpFrDt")), "''", "S")
    strToDt	= FilterVar(UNIConvDate(Request("txtPpToDt")), "''", "S")
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
    Dim lgPpTotLocAmt , lgDiffTotLocAmt , lgGlTotLocAmt , lgTempGlLocAmt
    
    Const C_SHEETMAXROWS_D = 100															'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT TOP " & C_SHEETMAXROWS_D + 1
	lgStrSQL = lgStrSQL & " A.DEBIT_ACCT_CD,B.ACCT_NM,  "
	lgStrSQL = lgStrSQL & " A.PRPAYM_NO, CONVERT(CHAR(10),A.PRPAYM_DT,20) PRPAYM_DT, CONVERT(CHAR(10),I.GL_DT,20) GL_DT,"
	lgStrSQL = lgStrSQL & " ISNULL(A.PRPAYM_LOC_AMT,0)-ISNULL(D.ITEM_LOC_AMT,0) DIFF_AMT, "
	lgStrSQL = lgStrSQL & " ISNULL(A.PRPAYM_LOC_AMT,0) PRPAYM_LOC_AMT,ISNULL(D.ITEM_LOC_AMT,0) GL_ITEM_LOC_AMT,"
	lgStrSQL = lgStrSQL & " ISNULL(F.ITEM_LOC_AMT,0) GL_TEMP_ITEM_LOC_AMT , ISNULL(E.ITEM_LOC_AMT,0) BATCH_ITEM_AMT , "
	lgStrSQL = lgStrSQL & " I.GL_NO,A.TEMP_GL_NO,E.BATCH_NO,F.TEMP_GL_DT "
	lgStrSQL = lgStrSQL & " ,A.PRPAYM_FG,C.MINOR_NM "
	lgStrSQL = lgStrSQL & " ,A.BIZ_AREA_CD,G.BIZ_AREA_NM "
	lgStrSQL = lgStrSQL & " ,A.BP_CD,H.BP_NM "
	lgStrSQL = lgStrSQL & " FROM F_PRPAYM A  "
	lgStrSQL = lgStrSQL & " LEFT JOIN A_ACCT B ON A.DEBIT_ACCT_CD=B.ACCT_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_MINOR C ON (A.PRPAYM_FG=C.MINOR_CD  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & "  ) "
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT  AA.BATCH_NO , AA.REF_NO , BB.JNL_CD , "
	lgStrSQL = lgStrSQL & "                    CASE WHEN BB.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*BB.ITEM_LOC_AMT  "
	lgStrSQL = lgStrSQL & "   				        ELSE BB.ITEM_LOC_AMT  END ITEM_LOC_AMT "
	lgStrSQL = lgStrSQL & "				FROM   A_BATCH AA "
	lgStrSQL = lgStrSQL & "				LEFT   JOIN A_BATCH_GL_ITEM BB ON AA.BATCH_NO=BB.BATCH_NO  "
	lgStrSQL = lgStrSQL & "				WHERE  AA.GL_DT >= " & strFrDt	
	lgStrSQL = lgStrSQL & "				 AND   AA.GL_DT <= " & strToDt	
	lgStrSQL = lgStrSQL & "			  ) E ON  (A.PRPAYM_NO=E.REF_NO AND A.PRPAYM_TYPE=E.JNL_CD) "
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT A_GL_ITEM.REF_NO,A_GL_ITEM.ITEM_LOC_AMT,A_GL.GL_DT,A_GL_ITEM.ACCT_CD,A_GL.GL_NO FROM A_GL,A_GL_ITEM "
	lgStrSQL = lgStrSQL & "				WHERE  A_GL.GL_NO=A_GL_ITEM.GL_NO "
	lgStrSQL = lgStrSQL & "				AND  A_GL.GL_DT >= " & strFrDt	
	lgStrSQL = lgStrSQL & "				AND  A_GL.GL_DT <= " & strToDt	
	lgStrSQL = lgStrSQL & " 			) D ON A.PRPAYM_NO=D.REF_NO AND A.DEBIT_ACCT_CD=D.ACCT_CD "  
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT A_TEMP_GL.REF_NO, A_TEMP_GL_ITEM.ITEM_LOC_AMT, A_TEMP_GL.TEMP_GL_NO,  "
	lgStrSQL = lgStrSQL & "                   A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT "
	lgStrSQL = lgStrSQL & "				FROM A_TEMP_GL , A_TEMP_GL_ITEM WHERE  A_TEMP_GL.TEMP_GL_NO=A_TEMP_GL_ITEM.TEMP_GL_NO "
	lgStrSQL = lgStrSQL & "				AND  A_TEMP_GL.CONF_FG=" & FilterVar("U", "''", "S") & "  " 	
	lgStrSQL = lgStrSQL & "				AND  A_TEMP_GL.TEMP_GL_DT >= " & strFrDt	
	lgStrSQL = lgStrSQL & "				AND  A_TEMP_GL.TEMP_GL_DT <= " & strToDt	
	lgStrSQL = lgStrSQL & "				) F ON A.PRPAYM_NO=F.REF_NO AND A.DEBIT_ACCT_CD=F.ACCT_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_AREA G ON A.BIZ_AREA_CD=G.BIZ_AREA_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_PARTNER H ON A.BP_CD=H.BP_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN A_GL I ON A.GL_NO=I.GL_NO "
	lgStrSQL = lgStrSQL & " WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '') "
    lgStrSQL = lgStrSQL & "  AND  A.PRPAYM_NO >= " & FilterVar(lgStrPrevKey, "''", "S")
	lgStrSQL = lgStrSQL & "  AND  A.PRPAYM_DT <= " & strToDt 
	lgStrSQL = lgStrSQL & "  AND  A.PRPAYM_DT >= " & strFrDt

	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL = lgStrSQL & " AND  A.PRPAYM_FG  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.DEAL_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL = lgStrSQL & " AND ISNULL(A.PRPAYM_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "
	lgStrSQL = lgStrSQL & " ORDER BY A.PRPAYM_NO "

	
	
	lgStrSQL2 = ""
	lgStrSQL2 = lgStrSQL2 & " SELECT "
	lgStrSQL2 = lgStrSQL2 & "  ISNULL(SUM(A.PRPAYM_LOC_AMT),0) AR_TOT_LOC_AMT, ISNULL(SUM(E.ITEM_LOC_AMT),0) BATCH_TOT_LOC_AMT ,"
	lgStrSQL2 = lgStrSQL2 & "  ISNULL(SUM(D.ITEM_LOC_AMT),0) GL_TOT_LOC_AMT, ISNULL(SUM(F.ITEM_LOC_AMT),0) TEMP_GL_TOT_LOC_AMT,  "
	lgStrSQL2 = lgStrSQL2 & "  ISNULL(SUM(A.PRPAYM_LOC_AMT),0)- ISNULL(SUM(D.ITEM_LOC_AMT),0) diff_TOT_LOC_AMT "
	lgStrSQL2 = lgStrSQL2 & " FROM F_PRPAYM A  "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN A_ACCT B ON A.DEBIT_ACCT_CD=B.ACCT_CD "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN B_MINOR C ON (A.PRPAYM_TYPE=C.MINOR_CD  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & "  ) "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT  AA.BATCH_NO , AA.REF_NO , BB.JNL_CD , "
	lgStrSQL2 = lgStrSQL2 & "                    CASE WHEN BB.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*BB.ITEM_LOC_AMT  "
	lgStrSQL2 = lgStrSQL2 & "   				        ELSE BB.ITEM_LOC_AMT  END ITEM_LOC_AMT "
	lgStrSQL2 = lgStrSQL2 & "				FROM   A_BATCH AA "
	lgStrSQL2 = lgStrSQL2 & "				LEFT   JOIN A_BATCH_GL_ITEM BB ON AA.BATCH_NO=BB.BATCH_NO  "
	lgStrSQL2 = lgStrSQL2 & "				WHERE  AA.GL_DT >= " & strFrDt	
	lgStrSQL2 = lgStrSQL2 & "				 AND   AA.GL_DT <= " & strToDt	
	lgStrSQL2 = lgStrSQL2 & "			  ) E ON  (A.PRPAYM_NO=E.REF_NO AND A.PRPAYM_TYPE=E.JNL_CD) "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT A_GL_ITEM.REF_NO,A_GL_ITEM.ITEM_LOC_AMT,A_GL.GL_DT,A_GL_ITEM.ACCT_CD,A_GL.GL_NO FROM A_GL,A_GL_ITEM "
	lgStrSQL2 = lgStrSQL2 & "				WHERE  A_GL.GL_NO=A_GL_ITEM.GL_NO "
	lgStrSQL2 = lgStrSQL2 & "				AND  A_GL.GL_DT >= " & strFrDt	
	lgStrSQL2 = lgStrSQL2 & "				AND  A_GL.GL_DT <= " & strToDt	
	lgStrSQL2 = lgStrSQL2 & " 			) D ON A.PRPAYM_NO=D.REF_NO AND A.DEBIT_ACCT_CD=D.ACCT_CD "  
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT A_TEMP_GL.REF_NO, A_TEMP_GL_ITEM.ITEM_LOC_AMT, A_TEMP_GL.TEMP_GL_NO,  "
	lgStrSQL2 = lgStrSQL2 & "                   A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT "
	lgStrSQL2 = lgStrSQL2 & "				FROM A_TEMP_GL , A_TEMP_GL_ITEM WHERE  A_TEMP_GL.TEMP_GL_NO=A_TEMP_GL_ITEM.TEMP_GL_NO "
	lgStrSQL2 = lgStrSQL2 & "				AND  A_TEMP_GL.CONF_FG=" & FilterVar("U", "''", "S") & "  " 
	lgStrSQL2 = lgStrSQL2 & "				AND  A_TEMP_GL.TEMP_GL_DT >= " & strFrDt	
	lgStrSQL2 = lgStrSQL2 & "				AND  A_TEMP_GL.TEMP_GL_DT <= " & strToDt	
	lgStrSQL2 = lgStrSQL2 & "				) F ON A.PRPAYM_NO=F.REF_NO AND A.DEBIT_ACCT_CD=F.ACCT_CD "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN B_BIZ_AREA G ON A.BIZ_AREA_CD=G.BIZ_AREA_CD "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN B_BIZ_PARTNER H ON A.BP_CD=H.BP_CD "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN A_GL I ON A.GL_NO=I.GL_NO "
	lgStrSQL2 = lgStrSQL2 & " WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '') "
	lgStrSQL2 = lgStrSQL2 & "  AND  A.PRPAYM_DT <= " & strToDt 
	lgStrSQL2 = lgStrSQL2 & "  AND  A.PRPAYM_DT >= " & strFrDt

	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL2 = lgStrSQL2 & " AND  A.PRPAYM_FG  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.DEAL_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL2 = lgStrSQL2 & " AND ISNULL(A.PRPAYM_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "   
	
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
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEBIT_ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRPAYM_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PRPAYM_DT"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))      
			
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PRPAYM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DIFF_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TEMP_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("BATCH_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BATCH_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TEMP_GL_DT")) 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRPAYM_FG"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))			 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))      			  
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
          
			lgObjRs.MoveNext

			iDx =  iDx + 1

			If iDx > C_SHEETMAXROWS_D Then
			    Exit Do
			End If   
		Loop 
		
		If Not lgObjRs.EOF Then
		   lgStrPrevKey = lgObjRs("PRPAYM_NO")
		Else
		   lgStrPrevKey = ""
		End If
	End If

	'*********************************
	'			합계찍기 
	'*********************************
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
		lgPpTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("AR_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgDiffTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("diff_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgGlTotLocAmt  = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgTempGlLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("TEMP_GL_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	Else
		lgPpTotLocAmt = 0
		lgDiffTotLocAmt = 0
		lgGlTotLocAmt  = 0
		lgTempGlLocAmt = 0
	End If    

	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                                             " & vbCr
       Response.Write  "    Parent.ggoSpread.Source              = Parent.frm1.vspdData3        " & vbCr
       Response.Write  "    Parent.lgStrPrevKey                  =  """ & lgStrPrevKey     & """" & vbCr       
       Response.Write  "    Parent.ggoSpread.SSShowData             """ & lgstrData        & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotPpLocAmt3.text     =  """ & lgPpTotLocAmt    & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & lgDiffTotLocAmt & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & lgGlTotLocAmt    & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & lgTempGlLocAmt   & """" & vbCr                     
       Response.Write  "    Parent.DBQueryOk												    " & vbCr      
       Response.Write  " </Script>															    " & vbCr
    End If
	
	Response.End
End Sub    


%>

