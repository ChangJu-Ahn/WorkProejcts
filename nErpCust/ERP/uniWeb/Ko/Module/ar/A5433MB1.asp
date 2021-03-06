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
    DIm strDealBpCd

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
     strFrDt		= FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S")
     strToDt		= FilterVar(UNIConvDate(Request("txtToDt")), "''", "S")
     strBizAreaCd   = FilterVar(Request("txtBizAreaCd"), "''", "S") 
	 strAcctCd      = FilterVar(Request("txtAcctCd"), "''", "S") 
	 strDealBpCd    = FilterVar(Request("txtDealBpCd"), "''", "S") 
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

	lgStrClsAr =			 "	SELECT A.cls_ar_no, a.cls_dt" 
	lgStrClsAr = lgStrClsAr & "	FROM (SELECT A.AR_NO, a.cls_ar_no, a.cls_dt"
	lgStrClsAr = lgStrClsAr & "		  FROM A_CLS_AR A JOIN (SELECT ALLC_NO, GL_NO,	TEMP_GL_NO, REF_NO FROM A_ALLC_RCPT UNION"
	lgStrClsAr = lgStrClsAr & "								 SELECT PAYM_NO, GL_NO, TEMP_GL_NO, '' FROM A_ALLC_PAYM UNION"							
	lgStrClsAr = lgStrClsAr & "								 SELECT CLEAR_NO, GL_NO, '', '' FROM A_CLEAR_AP_AR) B ON B.ALLC_NO = A.CLS_AR_NO"
	lgStrClsAr = lgStrClsAr & "		  UNION ALL"
	lgStrClsAr = lgStrClsAr & "		  SELECT A.AR_NO, a.adjust_no, A.ADJUST_DT FROM A_AR_ADJUST A JOIN A_GL_ITEM B ON B.GL_NO = A.GL_NO AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & " "
	lgStrClsAr = lgStrClsAr & "		  )A left JOIN A_OPEN_AR B ON A.AR_NO=B.AR_NO" 
	lgStrClsAr = lgStrClsAr & "	WHERE A.CLS_DT <= " & strToDt & " AND  A.CLS_DT >= " & strFrDt
	lgStrClsAr = lgStrClsAr & "	group  by A.cls_ar_no,  a.cls_dt"


	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT " 
	lgStrSQL = lgStrSQL & " ACCT_CD,ACCT_NM,  " 
	lgStrSQL = lgStrSQL & " ALLC_LOC_AMT,GL_LOC_AMT,BATCH_LOC_AMT,Diff_LOC_AMT, TEMP_GL_LOC_AMT " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,BIZ_AREA_CD,BIZ_AREA_NM " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,DEAL_BP_CD,BP_NM " 
	
	lgStrSQL = lgStrSQL & " FROM (  " 
	lgStrSQL = lgStrSQL & " SELECT ACCT_CD,ACCT_NM, " 
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN AR_FG=" & FilterVar("AR", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) ALLC_LOC_AMT , " 
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN AR_FG=" & FilterVar("BATCH", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) BATCH_LOC_AMT, " 
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN AR_FG=" & FilterVar("GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) GL_LOC_AMT, " 
	lgStrSQL = lgStrSQL & " (SUM(CASE WHEN AR_FG=" & FilterVar("AR", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END)) - (SUM(CASE WHEN AR_FG=" & FilterVar("GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END)) Diff_LOC_AMT, "
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN AR_FG=" & FilterVar("TEMP_GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) TEMP_GL_LOC_AMT " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,BIZ_AREA_CD,BIZ_AREA_NM " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,DEAL_BP_CD,BP_NM " 
	
	lgStrSQL = lgStrSQL & " FROM ( " 
	'*****************************************A_cls_AP에서 가져오기 시작******************************************			
	lgStrSQL = lgStrSQL & " SELECT A.ACCT_CD,B.ACCT_NM, SUM(A.CLS_LOC_AMT)  ITEM_LOC_AMT," & FilterVar("AR", "''", "S") & "  AR_FG " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A.BIZ_AREA_CD,D.BIZ_AREA_NM " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A.DEAL_BP_CD,E.BP_NM " 
	
	lgStrSQL = lgStrSQL & " FROM (" 
	lgStrSQL = lgStrSQL & "			SELECT A.AR_NO, A.CLS_AR_NO, A.CLS_DT, A.ACCT_CD, A.CLS_AMT, A.CLS_LOC_AMT, A.GL_NO, A.TEMP_GL_NO, B.BIZ_AREA_CD, B.DEAL_BP_CD " 
	lgStrSQL = lgStrSQL & "			FROM (SELECT A.AR_NO, CLS_AR_NO, A.CLS_DT, A.ACCT_CD ACCT_CD, " 
	lgStrSQL = lgStrSQL & "					A.CLS_AMT+A.DC_AMT CLS_AMT, A.CLS_LOC_AMT+A.DC_LOC_AMT CLS_LOC_AMT, B.GL_NO, B.TEMP_GL_NO " 
	lgStrSQL = lgStrSQL & "				   FROM A_CLS_AR A JOIN (SELECT ALLC_NO, GL_NO,	TEMP_GL_NO, REF_NO FROM A_ALLC_RCPT UNION " 
	lgStrSQL = lgStrSQL & "										 SELECT PAYM_NO, GL_NO, TEMP_GL_NO, '' FROM A_ALLC_PAYM UNION " 
	lgStrSQL = lgStrSQL & "										 SELECT CLEAR_NO, GL_NO, '', '' FROM A_CLEAR_AP_AR) B ON B.ALLC_NO = A.CLS_AR_NO " 
	lgStrSQL = lgStrSQL & "				UNION ALL" 
	lgStrSQL = lgStrSQL & "				SELECT A.AR_NO, A.ADJUST_NO, A.ADJUST_DT, B.ACCT_CD,A.ADJUST_AMT, A.ADJUST_LOC_AMT,  A.GL_NO, A.TEMP_GL_NO" 
	lgStrSQL = lgStrSQL & "				FROM A_AR_ADJUST A JOIN A_GL_ITEM B ON B.GL_NO = A.GL_NO AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & " " 
	lgStrSQL = lgStrSQL & "				)A INNER JOIN A_OPEN_AR B ON A.AR_NO=B.AR_NO" 
	lgStrSQL = lgStrSQL & "			) A LEFT JOIN  A_ACCT B ON A.ACCT_CD=B.ACCT_CD" 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN  B_BIZ_AREA D ON A.BIZ_AREA_CD=D.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN  B_BIZ_PARTNER E ON A.DEAL_BP_CD = E.BP_CD" 
	
	lgStrSQL = lgStrSQL & " WHERE A.CLS_DT <= " & strToDt 
	lgStrSQL = lgStrSQL & "  AND A.CLS_DT >= " & strFrDt 
	lgStrSQL = lgStrSQL & "  AND (A.GL_NO <> '' OR A.TEMP_GL_NO <> '') " 
	lgStrSQL = lgStrSQL & "  AND A.ACCT_CD IN (SELECT ACCT_CD FROM A_ACCT WHERE ACCT_TYPE LIKE " & FilterVar("%I%", "''", "S") & " )" 
	lgStrSQL = lgStrSQL & " GROUP BY A.ACCT_CD,B.ACCT_NM" 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & "	,A.BIZ_AREA_CD ,D.BIZ_AREA_NM" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & "	,A.DEAL_BP_CD ,E.BP_NM" 
	
'*****************************************A_Cls_AR에서 가져오기 끝******************************************	
	lgStrSQL = lgStrSQL & " UNION " 
'************************A_BATCH_GL_ITEM중에서  가져오기 시작*******************		
	
	lgStrSQL = lgStrSQL & " SELECT A.ACCT_CD, F.ACCT_NM," 
	lgStrSQL = lgStrSQL & " ISNULL(SUM(CASE WHEN A.REVERSE_FG=" & FilterVar("Y", "''", "S") & "  THEN (-1)*A.ITEM_LOC_AMT" 
	lgStrSQL = lgStrSQL & "                 ELSE A.ITEM_LOC_AMT     END),0)  ITEM_LOC_AMT," 
	lgStrSQL = lgStrSQL & " " & FilterVar("BATCH", "''", "S") & "  AR_FG " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A.BIZ_AREA_CD,D.BIZ_AREA_NM" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,C.DEAL_BP_CD, E.BP_NM" 
	
	lgStrSQL = lgStrSQL & " FROM A_BATCH_GL_ITEM A  " 
	lgStrSQL = lgStrSQL & "			INNER JOIN  (SELECT A_BATCH.BATCH_NO, A_BATCH.BIZ_AREA_CD" 
	lgStrSQL = lgStrSQL & "						FROM A_BATCH left join (" & lgStrClsAr & " )A ON A.CLS_AR_NO=A_BATCH.REF_NO" 
	lgStrSQL = lgStrSQL & "						WHERE A_BATCH.GL_DT >= " & strFrDt & "AND A_BATCH.GL_DT <= " & strToDt   
	lgStrSQL = lgStrSQL & "						AND GL_INPUT_TYPE IN  (" & FilterVar("RX", "''", "S") & ", " & FilterVar("CP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " , " & FilterVar("LR", "''", "S") & " ," & FilterVar("CL", "''", "S") & "  ," & FilterVar("CR", "''", "S") & " ," & FilterVar("JR", "''", "S") & ")" 
	lgStrSQL = lgStrSQL & "						) B ON A.BATCH_NO=B.BATCH_NO" 
	lgStrSQL = lgStrSQL & "			LEFT JOIN A_OPEN_AR  C ON C.AR_NO=A.KEY_VAL1" 
	lgStrSQL = lgStrSQL & "			LEFT JOIN A_ACCT F ON A.ACCT_CD=F.ACCT_CD" 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_AREA D ON B.BIZ_AREA_CD=D.BIZ_AREA_CD" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_PARTNER E ON C.DEAL_BP_CD=E.BP_CD" 
	
	lgStrSQL = lgStrSQL & "	WHERE A.ACCT_CD  IN (SELECT ACCT_CD FROM A_ACCT  WHERE ACCT_TYPE LIKE " & FilterVar("%I%", "''", "S") & " )" 
	lgStrSQL = lgStrSQL & " GROUP BY A.ACCT_CD,  F.ACCT_NM " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,A.BIZ_AREA_CD,D.BIZ_AREA_NM  " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,C.DEAL_BP_CD,E.BP_NM " 
	
'************************A_BATCH_GL_ITEM중에서 가져오기 끝*******************************
	lgStrSQL = lgStrSQL & " UNION " 
'********************************************A_GL_ITEM중에서 가져오기 시작***********************************		
	
	lgStrSQL = lgStrSQL & " SELECT  A.ACCT_CD,F.ACCT_NM,ISNULL(SUM( A.ITEM_LOC_AMT),0) ITEM_LOC_AMT, " & FilterVar("GL", "''", "S") & "  AR_FG " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,B.BIZ_AREA_CD,D.BIZ_AREA_NM" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,H.CTRL_VAL,E.BP_NM" 
	
	lgStrSQL = lgStrSQL & " FROM A_GL_ITEM A " 
	lgStrSQL = lgStrSQL & "			INNER JOIN  (SELECT A_GL.GL_NO,A_GL.REF_NO, A_GL.BIZ_AREA_CD" 
	lgStrSQL = lgStrSQL & "						 FROM A_GL left join (" & lgStrClsAr & " )A ON A.CLS_AR_NO=A_GL.REF_NO" 
	lgStrSQL = lgStrSQL & "           			 WHERE GL_DT >= " & strFrDt  & " AND  GL_DT <= " & strToDt 
	lgStrSQL = lgStrSQL & "           			  AND GL_INPUT_TYPE IN  (" & FilterVar("RX", "''", "S") & ", " & FilterVar("CP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " , " & FilterVar("LR", "''", "S") & " ," & FilterVar("CL", "''", "S") & "  ," & FilterVar("CR", "''", "S") & " ," & FilterVar("JR", "''", "S") & ")" 
	lgStrSQL = lgStrSQL & "   					) B ON A.GL_NO=B.GL_NO" 
	lgStrSQL = lgStrSQL & "			LEFT JOIN A_ACCT F ON A.ACCT_CD=F.ACCT_CD" 
	
	If Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " LEFT JOIN A_GL_DTL H ON A.GL_NO=H.GL_NO AND A.ITEM_SEQ=H.ITEM_SEQ AND H.CTRL_CD IN ( Select CTRL_CD From A_CTRL_ITEM Where  TBL_ID = " & FilterVar("B_BIZ_PARTNER", "''", "S") & " )" 
		lgStrSQL = lgStrSQL & "	LEFT JOIN B_BIZ_PARTNER E ON H.CTRL_VAL=E.BP_CD"	 
	End if
		
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_AREA D ON B.BIZ_AREA_CD=D.BIZ_AREA_CD" 
	
	
	lgStrSQL = lgStrSQL & "	WHERE A.ACCT_CD  IN (SELECT ACCT_CD FROM A_ACCT  WHERE ACCT_TYPE LIKE " & FilterVar("%I%", "''", "S") & " )" 
	lgStrSQL = lgStrSQL & " GROUP BY A.ACCT_CD,  F.ACCT_NM " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,B.BIZ_AREA_CD,D.BIZ_AREA_NM  " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,H.CTRL_VAL,E.BP_NM " 

'************************A_GL_ITEM중에서 가져오기 끝*******************************
	lgStrSQL = lgStrSQL & " UNION " 
'********************************************A_Temp_GL_ITEM중에서 가져오기 시작***********************************
	
	lgStrSQL = lgStrSQL & " SELECT A.ACCT_CD,F.ACCT_NM,ISNULL(SUM( A.ITEM_LOC_AMT),0) ITEM_LOC_AMT, " & FilterVar("TEMP_GL", "''", "S") & "  AR_FG  " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,B.BIZ_AREA_CD,D.BIZ_AREA_NM" 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,G.CTRL_VAL,E.BP_NM" 
		
	lgStrSQL = lgStrSQL & " FROM A_TEMP_GL_ITEM A  " 
	lgStrSQL = lgStrSQL & "			INNER JOIN  (SELECT A_TEMP_GL.TEMP_GL_NO,A_TEMP_GL.REF_NO, A_TEMP_GL.BIZ_AREA_CD, CONF_FG" 
	lgStrSQL = lgStrSQL & "						 FROM A_TEMP_GL LEFT JOIN (" & lgStrClsAr & " )A ON A.CLS_AR_NO=A_TEMP_GL.REF_NO" 
	lgStrSQL = lgStrSQL & "           			 WHERE A_TEMP_GL.TEMP_GL_DT >= " & strFrDt  & " AND  A_TEMP_GL.TEMP_GL_DT  <= " & strToDt 
	lgStrSQL = lgStrSQL & "           			  AND GL_INPUT_TYPE IN  (" & FilterVar("RX", "''", "S") & ", " & FilterVar("CP", "''", "S") & " ," & FilterVar("CS", "''", "S") & " , " & FilterVar("LR", "''", "S") & " ," & FilterVar("CL", "''", "S") & "  ," & FilterVar("CR", "''", "S") & " ," & FilterVar("JR", "''", "S") & ")" 
	lgStrSQL = lgStrSQL & "   					) B ON A.TEMP_GL_NO=B.TEMP_GL_NO " 
	lgStrSQL = lgStrSQL & "			LEFT JOIN A_ACCT F ON A.ACCT_CD=F.ACCT_CD" 
		
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " LEFT JOIN B_BIZ_AREA D ON B.BIZ_AREA_CD=D.BIZ_AREA_CD" 
	
	If Trim(Request("txtShowBp")) = "Y" Then	
		lgStrSQL = lgStrSQL & " LEFT JOIN A_TEMP_GL_DTL G ON A.TEMP_GL_NO=G.TEMP_GL_NO AND A.ITEM_SEQ=G.ITEM_SEQ AND G.CTRL_CD in ( Select CTRL_CD From A_CTRL_ITEM Where  TBL_ID = " & FilterVar("B_BIZ_PARTNER", "''", "S") & " )" 
		lgStrSQL = lgStrSQL & "	LEFT JOIN B_BIZ_PARTNER E ON G.CTRL_VAL=E.BP_CD"	 
	End if
	
	lgStrSQL = lgStrSQL & "	WHERE A.ACCT_CD  IN (SELECT ACCT_CD FROM A_ACCT  WHERE ACCT_TYPE LIKE " & FilterVar("%I%", "''", "S") & " )" 
	lgStrSQL = lgStrSQL & " AND B.CONF_FG <>" & FilterVar("C", "''", "S") & " "
	lgStrSQL = lgStrSQL & " GROUP BY A.ACCT_CD,  F.ACCT_NM " 
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,B.BIZ_AREA_CD,D.BIZ_AREA_NM " 
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,G.CTRL_VAL,E.BP_NM " 
	
	
	lgStrSQL = lgStrSQL & " ) A " 
	lgStrSQL = lgStrSQL & " GROUP BY ACCT_CD,ACCT_NM "
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " ,BIZ_AREA_CD,BIZ_AREA_NM "
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " ,DEAL_BP_CD,BP_NM "
	
	
	lgStrSQL = lgStrSQL & " ) A " 
	lgStrSQL = lgStrSQL & " WHERE 1=1 "

	If Trim(Request("txtShowBiz")) = "Y" Then
		If Trim(Request("txtBizAreaCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND BIZ_AREA_CD = " & strBizAreaCd
		End If
	End If 
	
	If Trim(Request("txtShowBp")) = "Y" Then
		If Trim(Request("txtdealbpCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND DEAL_BP_CD = " & strDealBpCd
		End If
	End If
	
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND ACCT_CD = " & strAcctCd
	End If
	
	If UCase(Trim(Request("DispMeth"))) Then lgStrSQL = lgStrSQL & " AND ISNULL(ALLC_LOC_AMT,0) <> ISNULL(GL_LOC_AMT,0) "

	lgStrSQL = lgStrSQL & " ORDER BY ACCT_CD ASC "
	
	If Trim(Request("txtShowBiz")) = "Y" Then	lgStrSQL = lgStrSQL & " , BIZ_AREA_CD ASC "
	If Trim(Request("txtShowBp")) = "Y" Then	lgStrSQL = lgStrSQL & " , DEAL_BP_CD ASC "
	

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else 
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
        
		Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("ALLC_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("Diff_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("TEMP_GL_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("BATCH_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			
			If Trim(Request("txtShowBiz")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & ""
			End If
			
			If Trim(Request("txtShowBp")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEAL_BP_CD"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & ""
			End If
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext

			iDx =  iDx + 1
		Loop
	End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                                 " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData1     " & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
'       Response.Write  "    Parent.lgGlInputType        =  """"                     " & vbCr
'       Response.Write  "    Parent.lgGlInputTypeNm      =  """"                     " & vbCr       
       Response.Write  "    Parent.DBQueryOk										" & vbCr      
       Response.Write  " </Script>													" & vbCr
    End If
    
    Response.End 
End Sub

%>

 
