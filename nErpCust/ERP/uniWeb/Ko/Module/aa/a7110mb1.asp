<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%
    Dim lgPageNo
	Dim lgPageCnt

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					
    
    On Error Resume Next
    Err.Clear

    Call LoadBasisGlobalInf()    
    Call LoadInfTB19029B("Q","A","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","MB") 

	lgPageNo = 0
    Call HideStatusWnd																	'☜: Hide Processing message
	
	lgStrSQL = ""
    lgErrorStatus		= "NO"
    lgErrorPos			= ""															'☜: Set to space
    lgKeyStream			= Split(Request("txtKeyStream"),gColSep) 
    lgPageCnt			= Request("lgPageNo")											'☜ : Next key flag
    
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


    Const C_SHEETMAXROWS_D  = 100														'☆: Server에서 한번에 fetch할 최대 데이타 건수 

	If Len(Trim(lgPageCnt)) Then														'☜ : Chnage Nextkey str into int value
	   If Isnumeric(lgPageCnt) Then
		  lgPageNo = CInt(lgPageCnt)
	   End If
	End If

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    
    Call SubOpenDB(lgObjConn)															'☜: Make a DB Connection
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)															'☜: Query
             Call SubBizQuery()
    End Select
    Call SubCloseDB(lgObjConn)															'☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim lgStrSQL2
    
    On Error Resume Next
    Err.Clear
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    If lgPageNo = 0 Then
       lgStrSQL2 = "Select A.Asst_No, A.Asst_Nm, B.Acct_Nm" 
       lgStrSQL2 = lgStrSQL2 & " From A_Asset_Master A(NOLOCK), A_Acct B(NOLOCK)" 
       lgStrSQL2 = lgStrSQL2 & " WHERE A.Acct_Cd  = B.Acct_cd"
       lgStrSQL2 = lgStrSQL2 & " AND	A.Asst_No  = " & FilterVar(lgKeyStream(0), "''", "S")


		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then			
			lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
		End If			

		If lgInternalCd <> "" Then			
			lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
		End If			

		If lgSubInternalCd <> "" Then	
			lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If	

		If lgAuthUsrID <> "" Then	
			lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
		End If	

		lgStrSQL2	= lgStrSQL2	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	




		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then
		
			Response.Write  " <Script Language=vbscript>            " & vbCr
			Response.Write  "   Parent.Frm1.txtAsstNo.Value		= """ & lgObjRs("Asst_No") & """" & vbCr             ' Set condition area
			Response.Write  "   Parent.Frm1.htxtAsstNo.Value	= """ & lgObjRs("Asst_No") & """" & vbCr             ' Set condition area
			Response.Write  "   Parent.Frm1.txtAsstNm.Value		= """ & lgObjRs("Asst_Nm") & """" & vbCr 
			Response.Write  "   Parent.Frm1.txtAcctNm.Value		= """ & lgObjRs("Acct_Nm") & """" & vbCr             ' Set next key data
			Response.Write  " </Script> " & vbCr

			Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
			Call SubBizQueryMulti()
		ELSE
			Call DisplayMsgBox("117400", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
			lgPageNo  = 0
			lgErrorStatus = "YES"
			Exit Sub 
		End If
    Else
       Call SubBizQueryMulti()
    End If 
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    


'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
   ' Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
    
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

	Call TrimData()

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
		lgPageNo  = 0
		lgErrorStatus = "YES"
		Exit Sub 
		Else  
		  
		iCnt = 0

		If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
			If Isnumeric(lgPageNo) Then
				iCnt = CInt(lgPageNo)
			End If
		End If

		For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
			lgObjRs.MoveNext
		Next

		iDx = 0
		Do While Not lgObjRs.EOF
			iDx =  iDx + 1
			If iDx > C_SHEETMAXROWS_D Then
				Exit Do
			End If

			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("A_SEQ"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("A_DATE"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("A_TYPE"))
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A_QTY"),ggQty.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A_CAPITAL"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A_REVENUE"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A_DEPR"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("A_GL"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("A_TEMP_GL"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("A_DESC"))
			lgstrData = lgstrData & Chr(11) & (iCnt  *  C_SHEETMAXROWS_D) + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
			Loop 
		End If

		If  iDx < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
			lgPageNo = 0                                                '☜: 다음 데이타 없다.
		Else 
			lgPageNo = lgPageNo + 1
		End If

		Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

		If CheckSYSTEMError(Err,True) = True Then
			ObjectContext.SetAbort
		Exit Sub
	End If   

	If lgErrorStatus = "NO" Then
		Response.Write  " <Script Language=vbscript>								" & vbCr
		Response.Write  "    Parent.ggoSpread.Source	= Parent.frm1.vspdData		" & vbCr
		Response.Write  "    Parent.lgPageNo			= " & lgPageNo				& vbCr
		Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData			& """" & vbCr
		Response.Write  "    Parent.DBQueryOk			" & vbCr      
		Response.Write  " </Script>						" & vbCr
	End If

	IF iCnt = 0 Then
		Call SubBizQueryQty()
		Call SubBizQueryAmt()
		Call SubBizDeprAmt()
	End  If



End Sub    



'============================================================================================================
' Name : SubBizQueryQty
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryQty()
    Dim lgStrSQL3
    
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	lgStrSQL3 = " SELECT a.acq_qty - (select isnull(sum(chg_qty),0) from a_asset_chg where asst_cd = a.asst_no and chg_fg in (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " )) Qty_Sum" 
	lgStrSQL3 = lgStrSQL3 & " FROM a_asset_master a(NOLOCK)" 
	lgStrSQL3 = lgStrSQL3 & " WHERE	A.Asst_No  = " & FilterVar(lgKeyStream(0), "''", "S")

	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL3,"X","X") = True Then
%>
		<Script Language=vbscript>
		Parent.Frm1.txtQtySum.text  = "<%=UNINumClientFormat(lgObjRs("Qty_Sum"),ggQty.DecPoint, 0)%>"
		</Script>
<%
		Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	End If

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SubBizQueryAmt
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryAmt()
    Dim lgStrSQL4
    
    'On Error Resume Next
    'Err.Clear
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	lgStrSQL4 = " select a.acq_loc_amt + (select isnull(sum(net_loc_amt),0) from a_asset_chg where asst_cd = a.asst_no and chg_fg = " & FilterVar("01", "''", "S") & " )  - (select isnull(sum(decr_acq_loc_amt),0) from a_asset_chg where asst_cd = a.asst_no and chg_fg in (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " )) Amt_Sum1 " 
	lgStrSQL4 = lgStrSQL4 & "  from a_asset_master a(NOLOCK) "
	lgStrSQL4 = lgStrSQL4 & " WHERE	A.Asst_No  = " & FilterVar(lgKeyStream(0), "''", "S")

	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL4,"X","X") = True Then
%>
		 <Script Language=vbscript>
		   Parent.Frm1.txtAmtSum1.text  =  "<%=UNINumClientFormat(lgObjRs("Amt_Sum1"),ggAmtOfMoney.DecPoint, 0)%>"
		 </Script> 

<%
		Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	End If
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    


'============================================================================================================
' Name : SubBizQueryAmt
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDeprAmt()
    Dim lgStrSQL5
    
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	lgStrSQL5 = " select	case	when b.l_mnth_depr_tot_amt <> 0 "
	lgStrSQL5 = lgStrSQL5 & "		then b.l_mnth_depr_tot_amt  - (select  isnull(sum(depr_tot_loc_amt),0) from a_asset_chg(NOLOCK) where asst_cd = a.asst_no and chg_fg in (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " ))"
	lgStrSQL5 = lgStrSQL5 & " 		else a.cas_end_l_term_depr_tot_amt  - (select  isnull(sum(depr_tot_loc_amt),0) from a_asset_chg(NOLOCK) where asst_cd = a.asst_no and chg_fg in (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " )) end Depr_Amt"
	lgStrSQL5 = lgStrSQL5 & " from 	a_asset_master a(NOLOCK) LEFT OUTER JOIN (select isnull(sum(end_l_term_depr_tot_amt),0) end_l_term_depr_tot_amt, depr_yyyymm, asst_no, isnull(sum(l_mnth_depr_tot_amt),0) l_mnth_depr_tot_amt"
	lgStrSQL5 = lgStrSQL5 & " from a_asset_depr_master c(NOLOCK) where   dur_yrs_fg= " & FilterVar("C", "''", "S") & "  and depr_yyyymm = (select max(depr_yyyymm) from a_asset_depr_of_dept(NOLOCK) where asst_no = c.asst_no and dur_yrs_fg= " & FilterVar("C", "''", "S") & " )"
	lgStrSQL5 = lgStrSQL5 & " group by depr_yyyymm, asst_no ) b ON  a.asst_no  = b.asst_no"
	lgStrSQL5 = lgStrSQL5 & " group by a.asst_no, a.cas_end_l_term_depr_tot_amt, b.end_l_term_depr_tot_amt,b.l_mnth_depr_tot_amt,b.depr_yyyymm"
	lgStrSQL5 = lgStrSQL5 & " having	a.asst_no = " & FilterVar(lgKeyStream(0), "''", "S")



	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL5,"X","X") = True Then
%>
		<Script Language=vbscript>
			Parent.Frm1.txtAmtSum3.text  =  "<%=UNINumClientFormat(lgObjRs("Depr_Amt"),ggAmtOfMoney.DecPoint, 0)%>"
		</Script>
<%
		Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	End If
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    

Sub TrimData()

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    lgStrSQL = "Select A_SEQ, A.A_DATE, A.A_TYPE, A.A_QTY, A.A_CAPITAL, A.A_REVENUE, A.A_DEPR, A.A_GL, A.A_TEMP_GL, A.A_DESC" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & " FROM	(" & vbCr  
   
'최초취득 
	lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.REG_DT			A_DATE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					C.MINOR_NM			A_TYPE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A.ACQ_QTY			A_QTY," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A.ACQ_LOC_AMT		A_CAPITAL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0		 			A_REVENUE,"		 & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0					A_DEPR,"		 & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B.GL_NO				A_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B.TEMP_GL_NO		A_TEMP_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					''					A_DESC,"		 & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0					A_SEQ" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_MASTER	A," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A_ASSET_ACQ		B," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR			C" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	A.ACQ_NO	= B.ACQ_NO" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		C.MAJOR_CD	= " & FilterVar("A2005", "''", "S") & " " & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.ACQ_FG	= C.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_NO	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.ACQ_FG	IN (" & FilterVar("01", "''", "S") & " ," & FilterVar("02", "''", "S") & " )" & vbCr  
    
'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  
    
    
'자본적지출 
	lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.CHG_DT			A_DATE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B.MINOR_NM			A_TYPE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.CHG_QTY			A_QTY," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.NET_LOC_AMT		A_CAPITAL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0					A_REVENUE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.DEPR_TOT_LOC_AMT	A_DEPR," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.GL_NO				A_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.TEMP_GL_NO		A_TEMP_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					''					A_DESC," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					isnull(A.ASST_CHG_SEQ,1)		A_SEQ" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_CHG	A," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR		B" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	CHG_FG		= " & FilterVar("01", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.MAJOR_CD	= " & FilterVar("A2001", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.CHG_FG	= B.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_CD	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  

'매각/폐기 
	lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.CHG_DT			A_DATE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B.MINOR_NM			A_TYPE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.CHG_QTY			A_QTY," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.DECR_ACQ_LOC_AMT	A_CAPITAL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0					A_REVENUE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.DEPR_TOT_LOC_AMT	A_DEPR," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.GL_NO				A_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.TEMP_GL_NO		A_TEMP_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					''					A_DESC," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					isnull(A.ASST_CHG_SEQ,1)		A_SEQ" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_CHG	A," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR		B" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	CHG_FG		IN (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " )" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.MAJOR_CD	= " & FilterVar("A2001", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.CHG_FG	= B.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_CD	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  

'수익적지출 
	lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.CHG_DT			A_DATE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B.MINOR_NM			A_TYPE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.CHG_QTY			A_QTY," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0					A_CAPITAL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.NET_LOC_AMT		A_REVENUE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0					A_DEPR," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.GL_NO				A_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.TEMP_GL_NO		A_TEMP_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					''					A_DESC," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					isnull(A.ASST_CHG_SEQ,1)		A_SEQ" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_CHG	A," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR		B" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	CHG_FG		= " & FilterVar("02", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.MAJOR_CD	= " & FilterVar("A2001", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.CHG_FG	= B.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_CD	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  

'부서이동 
    lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.CHG_DT					A_DATE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B.MINOR_NM					A_TYPE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A.CHG_QTY					A_QTY," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0							A_CAPITAL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0							A_REVENUE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0							A_DEPR," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A.GL_NO						A_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					A.TEMP_GL_NO				A_TEMP_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					C.DEPT_NM  + " & FilterVar("->", "''", "S") & "  + D.DEPT_NM	A_DESC," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					isnull(A.ASST_CHG_SEQ,1)				A_SEQ" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_CHG	A," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR		B," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B_ACCT_DEPT	C," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					B_ACCT_DEPT	D" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	CHG_FG		= " & FilterVar("05", "''", "S") & " " & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.MAJOR_CD	= " & FilterVar("A2001", "''", "S") & " " & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.CHG_FG	= B.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.FROM_ORG_CHANGE_ID	= C.ORG_CHANGE_ID" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.FROM_DEPT_CD			= C.DEPT_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.TO_ORG_CHANGE_ID		= D.ORG_CHANGE_ID" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.TO_DEPT_CD			= D.DEPT_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_CD	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  

'감가상각액 
    lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	DATEADD(DAY,-1,DATEADD(MONTH,1,CONVERT( DATETIME,DEPR_YYYYMM + " & FilterVar("01", "''", "S") & " ,120)))		A_DATE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					" & FilterVar("감가상각", "''", "S") & "			A_TYPE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0						A_QTY," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0						A_CAPITAL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					0						A_REVENUE," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					DEPR_AMT				A_DEPR," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					GL_NO					A_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					TEMP_GL_NO				A_TEMP_GL," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					''						A_DESC," & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "					999						A_SEQ" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_DEPR_OF_DEPT" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	DUR_YRS_FG	= " & FilterVar("C", "''", "S") & " " & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		DEPR_AMT	<> 0" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		ASST_NO	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'UNION ALL
    lgStrSQL = lgStrSQL & vbCrLf & "			UNION ALL" & vbCr  

'기초자산 
	lgStrSQL = lgStrSQL & vbCrLf & "			SELECT	A.REG_DT				A_DATE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					C.MINOR_NM				A_TYPE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.ACQ_QTY				A_QTY," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A.ACQ_LOC_AMT			A_CAPITAL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0						A_REVENUE," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					CASE	WHEN	D.END_L_TERM_DEPR_TOT_AMT <> 0" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "							THEN	D.END_L_TERM_DEPR_TOT_AMT" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "							ELSE	A.CAS_END_L_TERM_DEPR_TOT_AMT" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					END						A_DEPR," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B.GL_NO					A_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B.TEMP_GL_NO			A_TEMP_GL," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					''						A_DESC,"	 & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					0						A_SEQ" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			FROM	A_ASSET_ACQ		B," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					B_MINOR			C," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					A_ASSET_MASTER	A" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					LEFT OUTER JOIN" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					(	SELECT	ISNULL(SUM(END_L_TERM_DEPR_TOT_AMT),0)" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "								END_L_TERM_DEPR_TOT_AMT," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "								DEPR_YYYYMM," & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "								ASST_NO" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "						FROM	A_ASSET_DEPR_MASTER" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "						WHERE	DUR_YRS_FG= " & FilterVar("C", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "						GROUP BY DEPR_YYYYMM, ASST_NO" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					)	D" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					ON	A.START_DEPR_YYMM =  D.DEPR_YYYYMM" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "					AND	A.ASST_NO  = D.ASST_NO" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			WHERE	A.ACQ_NO = B.ACQ_NO" & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.ACQ_FG  = " & FilterVar("03", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		C.MAJOR_CD = " & FilterVar("A2005", "''", "S") & " " & vbCr  
	lgStrSQL = lgStrSQL & vbCrLf & "			AND		B.ACQ_FG = C.MINOR_CD" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & "			AND		A.ASST_NO	= " & FilterVar(lgKeyStream(0), "''", "S") & vbCr  

'첫번째 SELECT 문의 FROM 절 완료 
    lgStrSQL = lgStrSQL & vbCrLf & "			) A" & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf & " ORDER BY A.A_SEQ, A.A_DATE " & vbCr  
    lgStrSQL = lgStrSQL & vbCrLf



'==================================================================================================

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>


