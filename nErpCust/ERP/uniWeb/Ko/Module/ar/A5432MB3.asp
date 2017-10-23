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
    strFrDt	= FilterVar(UNIConvDate(Request("txtArFrDt")), "''", "S")
    strToDt	= FilterVar(UNIConvDate(Request("txtArToDt")), "''", "S")
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
    Dim lgArTotLocAmt , lgDiffTotLocAmt , lgGlTotLocAmt , lgTempGlLocAmt
    
    Const C_SHEETMAXROWS_D = 100															'☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear																			'☜: Clear Error status

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT TOP " & C_SHEETMAXROWS_D + 1
	lgStrSQL = lgStrSQL & " A.ACCT_CD,B.ACCT_NM,  "
	lgStrSQL = lgStrSQL & " A.AR_NO, CONVERT(CHAR(10),A.AR_DT,20) AR_DT, CONVERT(CHAR(10),I.GL_DT,20) GL_DT,"
	lgStrSQL = lgStrSQL & " ISNULL(A.AR_LOC_AMT,0)-ISNULL(D.ITEM_LOC_AMT,0) DIFF_AMT, "
	lgStrSQL = lgStrSQL & " ISNULL(A.AR_LOC_AMT,0) AR_LOC_AMT,ISNULL(D.ITEM_LOC_AMT,0) GL_ITEM_LOC_AMT,"
	lgStrSQL = lgStrSQL & " ISNULL(F.ITEM_LOC_AMT,0) GL_TEMP_ITEM_LOC_AMT , ISNULL(E.ITEM_LOC_AMT,0) BATCH_ITEM_AMT , "
	lgStrSQL = lgStrSQL & " I.GL_NO,F.TEMP_GL_NO,E.BATCH_NO,F.TEMP_GL_DT "
	lgStrSQL = lgStrSQL & " ,A.AR_TYPE,C.MINOR_NM "
	lgStrSQL = lgStrSQL & " ,A.BIZ_AREA_CD,G.BIZ_AREA_NM "
	lgStrSQL = lgStrSQL & " ,A.DEAL_BP_CD,H.BP_NM "
	lgStrSQL = lgStrSQL & " FROM A_OPEN_AR A  "
	lgStrSQL = lgStrSQL & " LEFT JOIN A_ACCT B ON A.ACCT_CD=B.ACCT_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_MINOR C ON (A.AR_TYPE=C.MINOR_CD  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & "  ) "
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT  A_BATCH.BATCH_NO,  A_BATCH.REF_NO, "
	lgStrSQL = lgStrSQL & "            CASE WHEN A_BATCH_GL_ITEM.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*A_BATCH_GL_ITEM.ITEM_LOC_AMT  "
	lgStrSQL = lgStrSQL & "   				ELSE A_BATCH_GL_ITEM.ITEM_LOC_AMT  END ITEM_LOC_AMT "
	lgStrSQL = lgStrSQL & "			   FROM A_BATCH , A_BATCH_GL_ITEM WHERE  A_BATCH.BATCH_NO=A_BATCH_GL_ITEM.BATCH_NO "
	lgStrSQL = lgStrSQL & "				AND A_BATCH_GL_ITEM.JNL_CD=(select top 1 a.jnl_cd from (select distinct(jnl_cd) jnl_cd from a_jnl_acct_assn "
	lgStrSQL = lgStrSQL & "	                                         where acct_cd in (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%I%", "''", "S") & " )) a "
	lgStrSQL = lgStrSQL & "										      order by jnl_cd) "
	lgStrSQL = lgStrSQL & "				AND  A_BATCH.GL_DT >= " & strFrDt
	lgStrSQL = lgStrSQL & "				AND  A_BATCH.GL_DT <= " & strToDt
	lgStrSQL = lgStrSQL & "				) E ON  A.REF_NO=E.REF_NO "
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT AA.REF_NO,CASE WHEN BB.ITEM_LOC_AMT > 0 AND CC.AR_LOC_AMT < 0 THEN (-1)*BB.ITEM_LOC_AMT "
	lgStrSQL = lgStrSQL & "                                  ELSE BB.ITEM_LOC_AMT END ITEM_LOC_AMT ,AA.GL_NO,BB.ACCT_CD,AA.GL_DT " 
	lgStrSQL = lgStrSQL & "     	     FROM A_GL AA "
	lgStrSQL = lgStrSQL & "             LEFT JOIN A_GL_ITEM	BB ON AA.GL_NO=BB.GL_NO "
	lgStrSQL = lgStrSQL & "             INNER JOIN A_OPEN_AR CC ON AA.REF_NO=CC.REF_NO "	
	lgStrSQL = lgStrSQL & "			    WHERE AA.GL_DT >= " & strFrDt
	lgStrSQL = lgStrSQL & "				 AND   AA.GL_DT <= " & strToDt
	lgStrSQL = lgStrSQL & " 		  ) D ON A.REF_NO=D.REF_NO AND A.ACCT_CD=D.ACCT_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN (SELECT AA.REF_NO,CASE WHEN BB.ITEM_LOC_AMT > 0 AND CC.AR_LOC_AMT < 0 THEN (-1)*BB.ITEM_LOC_AMT "
	lgStrSQL = lgStrSQL & "                                  ELSE BB.ITEM_LOC_AMT END ITEM_LOC_AMT,AA.TEMP_GL_NO,BB.ACCT_CD,AA.TEMP_GL_DT "
	lgStrSQL = lgStrSQL & "				FROM A_TEMP_GL AA "
	lgStrSQL = lgStrSQL & "		        LEFT JOIN A_TEMP_GL_ITEM BB ON AA.TEMP_GL_NO=BB.TEMP_GL_NO "
	lgStrSQL = lgStrSQL & "             INNER JOIN A_OPEN_AR CC ON AA.REF_NO=CC.REF_NO "		
	lgStrSQL = lgStrSQL & "				WHERE AA.TEMP_GL_DT >= " & strFrDt
	lgStrSQL = lgStrSQL & "				 AND  AA.TEMP_GL_DT <= " & strToDt
	lgStrSQL = lgStrSQL & "              AND  AA.CONF_FG = " & FilterVar("U", "''", "S") & "  "
	lgStrSQL = lgStrSQL & "			  ) F ON A.REF_NO=F.REF_NO AND A.ACCT_CD=F.ACCT_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_AREA G ON A.BIZ_AREA_CD=G.BIZ_AREA_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_PARTNER H ON A.DEAL_BP_CD=H.BP_CD "
	lgStrSQL = lgStrSQL & " LEFT JOIN A_GL I ON A.GL_NO=I.GL_NO "
	lgStrSQL = lgStrSQL & " WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '') "
    lgStrSQL = lgStrSQL & "  AND  A.AR_NO >= " & FilterVar(lgStrPrevKey, "''", "S")
	lgStrSQL = lgStrSQL & "  AND  A.AR_DT <= " & strToDt
	lgStrSQL = lgStrSQL & "  AND  A.AR_DT >= " & strFrDt

	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL = lgStrSQL & " AND  A.AR_TYPE  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND A.DEAL_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL = lgStrSQL & " AND ISNULL(A.AR_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "
	lgStrSQL = lgStrSQL & " ORDER BY A.AR_NO "
	
	lgStrSQL2 = ""
	lgStrSQL2 = lgStrSQL2 & " SELECT "
	lgStrSQL2 = lgStrSQL2 & "  ISNULL(SUM(A.AR_LOC_AMT),0) AR_TOT_LOC_AMT, ISNULL(SUM(E.ITEM_LOC_AMT),0) BATCH_TOT_LOC_AMT ,"
	lgStrSQL2 = lgStrSQL2 & "  ISNULL(SUM(D.ITEM_LOC_AMT),0) GL_TOT_LOC_AMT, ISNULL(SUM(F.ITEM_LOC_AMT),0) TEMP_GL_TOT_LOC_AMT, "
	lgStrSQL2 = lgStrSQL2 & " ISNULL(SUM(A.AR_LOC_AMT),0) - ISNULL(SUM(D.ITEM_LOC_AMT),0) DIFF_TOT_LOC_AMT "
	lgStrSQL2 = lgStrSQL2 & " FROM A_OPEN_AR A  "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT  A_BATCH.BATCH_NO,  A_BATCH.REF_NO,  "
	lgStrSQL2 = lgStrSQL2 & "            CASE WHEN A_BATCH_GL_ITEM.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*A_BATCH_GL_ITEM.ITEM_LOC_AMT  "
	lgStrSQL2 = lgStrSQL2 & "   				ELSE A_BATCH_GL_ITEM.ITEM_LOC_AMT  END ITEM_LOC_AMT "	
	lgStrSQL2 = lgStrSQL2 & "				FROM A_BATCH , A_BATCH_GL_ITEM WHERE  A_BATCH.BATCH_NO=A_BATCH_GL_ITEM.BATCH_NO "
	lgStrSQL2 = lgStrSQL2 & "				AND A_BATCH_GL_ITEM.JNL_CD=(select top 1 a.jnl_cd from (select distinct(jnl_cd) jnl_cd from a_jnl_acct_assn "
	lgStrSQL2 = lgStrSQL2 & "	                                         where acct_cd in (select acct_cd from a_acct where acct_type LIKE " & FilterVar("%I%", "''", "S") & " )) a "
	lgStrSQL2 = lgStrSQL2 & "										      order by jnl_cd) "
	lgStrSQL2 = lgStrSQL2 & "				AND  A_BATCH.GL_DT >= " & strFrDt	
	lgStrSQL2 = lgStrSQL2 & "				AND  A_BATCH.GL_DT <= " & strToDt		
	lgStrSQL2 = lgStrSQL2 & "				) E ON  A.REF_NO=E.REF_NO "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT AA.REF_NO,CASE WHEN BB.ITEM_LOC_AMT > 0 AND CC.AR_LOC_AMT < 0 THEN (-1)*BB.ITEM_LOC_AMT "
	lgStrSQL2 = lgStrSQL2 & "                                  ELSE BB.ITEM_LOC_AMT END ITEM_LOC_AMT ,AA.GL_NO,BB.ACCT_CD,AA.GL_DT " 
	lgStrSQL2 = lgStrSQL2 & "     	     FROM A_GL AA "
	lgStrSQL2 = lgStrSQL2 & "             LEFT JOIN A_GL_ITEM	BB ON AA.GL_NO=BB.GL_NO "
	lgStrSQL2 = lgStrSQL2 & "             INNER JOIN A_OPEN_AR CC ON AA.REF_NO=CC.REF_NO "	
	lgStrSQL2 = lgStrSQL2 & "			    WHERE AA.GL_DT >= " & strFrDt
	lgStrSQL2 = lgStrSQL2 & "				 AND   AA.GL_DT <= " & strToDt
	lgStrSQL2 = lgStrSQL2 & " 		  ) D ON A.REF_NO=D.REF_NO AND A.ACCT_CD=D.ACCT_CD "
	lgStrSQL2 = lgStrSQL2 & " LEFT JOIN (SELECT A_TEMP_GL_ITEM.REF_NO, A_TEMP_GL_ITEM.ITEM_LOC_AMT, A_TEMP_GL.TEMP_GL_NO, A_TEMP_GL_ITEM.ACCT_CD, A_TEMP_GL.TEMP_GL_DT "
	lgStrSQL2 = lgStrSQL2 & "				FROM A_TEMP_GL , A_TEMP_GL_ITEM WHERE  A_TEMP_GL.TEMP_GL_NO=A_TEMP_GL_ITEM.TEMP_GL_NO "
	lgStrSQL2 = lgStrSQL2 & "				AND  A_TEMP_GL.TEMP_GL_DT >= " & strFrDt	
	lgStrSQL2 = lgStrSQL2 & "				AND  A_TEMP_GL.TEMP_GL_DT <= " & strToDt		
	lgStrSQL2 = lgStrSQL2 & "               AND  A_TEMP_GL.CONF_FG = " & FilterVar("U", "''", "S") & "  "		
	lgStrSQL2 = lgStrSQL2 & "				) F ON A.REF_NO=F.REF_NO AND A.ACCT_CD=F.ACCT_CD "
	lgStrSQL2 = lgStrSQL2 & " WHERE (A.GL_NO <> '' OR A.TEMP_GL_NO <> '') "
	lgStrSQL2 = lgStrSQL2 & "  AND  A.AR_DT <= " & strToDt 
	lgStrSQL2 = lgStrSQL2 & "  AND  A.AR_DT >= " & strFrDt

	If Trim(Request("txtGlInputType")) <> "" Then lgStrSQL2 = lgStrSQL2 & " AND  A.AR_TYPE  = " & strInputType
	
	If Trim(Request("txtBizAreaCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.BIZ_AREA_CD = " & strBizAreaCd
	End If
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtdealbpCd")) <> "" Then
		lgStrSQL2 = lgStrSQL2 & " AND A.DEAL_BP_CD = " & strDealBpCd
	End If		
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL2 = lgStrSQL2 & " AND ISNULL(A.AR_LOC_AMT,0) <> ISNULL(D.ITEM_LOC_AMT,0) "   
	
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
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AR_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("AR_DT"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("GL_DT"))      
			
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("AR_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DIFF_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TEMP_ITEM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("BATCH_ITEM_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))          
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BATCH_NO"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TEMP_GL_DT")) 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AR_TYPE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))			 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEAL_BP_CD"))
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
		   lgStrPrevKey = lgObjRs("AR_NO")
		Else
		   lgStrPrevKey = ""
		End If
	End If

	'*********************************
	'			합계찍기 
	'*********************************
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
		lgArTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("AR_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgDiffTotLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("DIFF_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgGlTotLocAmt  = UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
		lgTempGlLocAmt = UNIConvNumDBToCompanyByCurrency(lgObjRs("TEMP_GL_TOT_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	Else
		lgArTotLocAmt = 0
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
       Response.Write  "    Parent.frm1.txtTotArLocAmt3.text     =  """ & lgArTotLocAmt    & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotDiffLocAmt3.text  =  """ & lgDiffTotLocAmt & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotGlLocAmt3.text     =  """ & lgGlTotLocAmt    & """" & vbCr
       Response.Write  "    Parent.frm1.txtTotTempGlLocAmt3.text =  """ & lgTempGlLocAmt   & """" & vbCr                     
       Response.Write  "    Parent.DBQueryOk												    " & vbCr      
       Response.Write  " </Script>															    " & vbCr
    End If
	
	Response.End
End Sub    


%>

