<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

 	Dim strFrDt	
    Dim strToDt
    DIm strBizAreaCd
    DIm strAcctCd   
    DIm strDealBpCd
	Dim strGlInputType
	Dim lgstrQueryId

	Call MakeQuryId()
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus  = ""

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
	Call TrimData()
	Call SubPreQueryBySp	
	Call SubBizQueryMulti()             '
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

Sub  MakeQuryId()
	Dim CurrentDt
	Dim CurrentTime
	Dim yyyy
	Dim mm
	Dim dd
	Dim hh
	Dim nn' ºÐ 
	Dim ss' ÃÊ 
	Dim rr
	
	CurrentDt =	Now 
	yyyy = Year(CurrentDt) 
	mm = Month (CurrentDt) 
	dd = Day  (CurrentDt) 	 
	hh = Hour (CurrentDt) 
	nn = Minute (CurrentDt) 
	ss = Second (CurrentDt) 
	Randomize
	rr = Int((98) * Rnd + 0)
	lgstrQueryId = yyyy & MakeTwoDigit(mm)  & MakeTwoDigit(dd) & MakeTwoDigit(hh) & MakeTwoDigit(nn) & MakeTwoDigit(ss) & MakeTwoDigit(rr)

End Sub

Function MakeTwoDigit ( ByVal strDt )
	If Len(strDt) = 1 Then
		strDt = "0" & strDt
	end if	
	MakeTwoDigit = strDt
End Function 	
'============================================================================================================
' Name : TrimData
' Desc : 
'============================================================================================================
Sub  TrimData()
     strFrDt		= Trim(UNIConvDateToYYYYMMDD(Request("txtPpFrDt"),gDateFormat,""))
     strToDt		= Trim(UNIConvDateToYYYYMMDD(Request("txtPpToDt"),gDateFormat,""))
     strBizAreaCd   = FilterVar(Request("txtBizAreaCd"), "''", "S") 
	 strAcctCd      = FilterVar(Request("txtAcctCd"), "''", "S") 
	 strDealBpCd    = FilterVar(Request("txtDealBpCd"), "''", "S")
	 strGlInputType = FilterVar(Request("txtGlinputtype"), "''", "S")
End Sub

'============================================================================================================
' Name : SubPreQueryBySp
' Desc : 
'============================================================================================================
Sub  SubPreQueryBySp()
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " DELETE FROM A_TMP_PP_TRACE_LIST WHERE QUERY_ID < CONVERT(VARCHAR(8),  GETDATE(),112)"
	
    If 	FncOpenRs("D",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("800407", vbInformation, "", "", I_MKSCRIPT)              '¢Ð : An Error Occur
        lgErrorStatus     = "YES"
	    Exit Sub
	End If
	
    CONST CALLSPNAME = "A_USP_FROM_PP_TO_GL_TRACE"
    
    Call SubCreateCommandObject(lgObjComm)

    With lgObjComm
        .CommandTimeout = 360
        .CommandText = CALLSPNAME			
        .CommandType = adCmdStoredProc
		.Parameters.Append lgObjComm.CreateParameter("@FROMDT", adVarWChar,adParamInput,Len(Trim(strFrDt)), strFrDt)
		.Parameters.Append lgObjComm.CreateParameter("@TODT", adVarWChar,adParamInput,Len(Trim(strToDt)), strToDt)
		.Parameters.Append lgObjComm.CreateParameter("@QUERY_ID", adVarWChar,adParamInput,16 , lgstrQueryId)				
		.Execute ,, adExecuteNoRecords
    End With
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

	On Error Resume Next
    Err.Clear 
    
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT ACCT_CD,ACCT_NM,"
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN PP_FG=" & FilterVar("PP", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) PRPAYM_LOC_AMT , "
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN PP_FG=" & FilterVar("BATCH", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) BATCH_LOC_AMT, "  
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN PP_FG=" & FilterVar("GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) GL_LOC_AMT, "
	lgStrSQL = lgStrSQL & " (SUM(CASE WHEN PP_FG=" & FilterVar("PP", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) - SUM(CASE WHEN PP_FG=" & FilterVar("GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END)) Diff_LOC_AMT, "
	lgStrSQL = lgStrSQL & " SUM(CASE WHEN PP_FG=" & FilterVar("TEMP_GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END) TEMP_GL_LOC_AMT, "
	lgStrSQL = lgStrSQL & " GL_INPUT_TYPE, GL_INPUT_TYPE_NM "	
	If Trim(Request("txtShowBiz")) = "Y" Then
		lgStrSQL = lgStrSQL & " ,BIZ_AREA_CD,BIZ_AREA_NM "
	End If
	If Trim(Request("txtShowBp")) = "Y" Then
		lgStrSQL = lgStrSQL & " ,DEAL_BP_CD,BP_NM "
	End If
	lgStrSQL = lgStrSQL & " FROM A_TMP_PP_TRACE_LIST "
	lgStrSQL = lgStrSQL & " WHERE QUERY_ID = " & FilterVar(lgstrQueryId, "''", "S") & ""		
	If Trim(Request("txtShowBiz")) = "Y" Then
		If Trim(Request("txtBizAreaCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND BIZ_AREA_CD = " & strBizAreaCd
		End If
	End If
	
	If Trim(Request("txtAcctCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND ACCT_CD = " & strAcctCd
	End If
	
	If Trim(Request("txtGlinputtype")) <> "" Then lgStrSQL = lgStrSQL & " AND  GL_INPUT_TYPE  = " & strGlInputType
	
	If Trim(Request("txtShowBp")) = "Y" Then
		If Trim(Request("txtdealbpCd")) <> "" Then
			lgStrSQL = lgStrSQL & " AND DEAL_BP_CD = " & strDealBpCd
		End If
	End If

	lgStrSQL = lgStrSQL & " GROUP BY ACCT_CD , ACCT_NM , GL_INPUT_TYPE , GL_INPUT_TYPE_NM "    
	If Trim(Request("txtShowBiz")) = "Y" Then
		lgStrSQL = lgStrSQL & " ,BIZ_AREA_CD,BIZ_AREA_NM "
	End If
	If Trim(Request("txtShowBp")) = "Y" Then
		lgStrSQL = lgStrSQL & " ,DEAL_BP_CD,BP_NM "
	End If
	
	If UCase(Trim(Request("DispMeth"))) Then 
		lgStrSQL = lgStrSQL & " HAVING ISNULL(SUM(CASE WHEN PP_FG=" & FilterVar("PP", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END),0) "
		lgStrSQL = lgStrSQL & "  <> ISNULL(SUM(CASE WHEN PP_FG=" & FilterVar("GL", "''", "S") & "  THEN ITEM_LOC_AMT ELSE 0 END),0) "
	End If		

	lgStrSQL = lgStrSQL & " ORDER BY ACCT_CD ASC ,GL_INPUT_TYPE ASC "
	If Trim(Request("txtShowBiz")) = "Y" Then
		lgStrSQL = lgStrSQL & " , BIZ_AREA_CD ASC "
	End If
	If Trim(Request("txtShowBp")) = "Y" Then
		lgStrSQL = lgStrSQL & " , DEAL_BP_CD ASC "
	End If

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					'¢Ð : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
	Else 
		iDx         = 1
		lgstrData   = ""
		lgLngMaxRow = Request("txtMaxRows")												'¢Ð: Read Operation Mode (CRUD)
       
		Do While Not lgObjRs.EOF
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
			lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PRPAYM_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("GL_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("Diff_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("TEMP_GL_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	        lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("BATCH_LOC_AMT"), gCurrency,  ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_INPUT_TYPE"))
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_INPUT_TYPE_NM"))
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
    
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

    If lgErrorStatus  = "" Then
       Response.Write  " <Script Language=vbscript>                             " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData2 " & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData  & """" & vbCr
       Response.Write  "    Parent.DBQueryOk									" & vbCr      
       Response.Write  " </Script>												" & vbCr
    End If
    
    Response.End
End Sub    

%>

