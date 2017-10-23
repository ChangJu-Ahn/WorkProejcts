<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strDueDt															
Dim strDocCur																
Dim strDueFg															
Dim strPayBpCd		
Dim strBatchAllcNo	
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1	
Dim GetApNo													
Dim iPayAmount
Dim strCond

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB") 
    
    Call HideStatusWnd() 
	lgErrorStatus = ""
 	lgIntFlgMode = Trim(Request("txtFlgMode"))	
 	
    Call MakeSQLWhere()
    
    Call SubOpenDB(lgObjConn)  

    If Cstr(lgIntFlgMode) = Cstr(OPMD_UMODE) Then
  		IF SubBizQuerySingle Then
  			Call SubBizQueryMulti()
  		End If
  	Else
  		Call SubBizQueryMulti()
	End If

    Call SubCloseDB(lgObjConn)  
    
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim lgStrSQL
	Dim lgLngMaxRow
	Dim iDx
	Dim strFromWhere
	Dim lgCurrency

	On Error Resume Next																	'☜: Protect system from crashing
	Err.Clear																				'☜: Clear Error status
	
    If Cstr(lgIntFlgMode) = Cstr(OPMD_UMODE) Then
		lgStrSQL = lgStrSQL & "SELECT " & FilterVar("1", "''", "S") & "  CHECKED, A.pay_bp_cd, B.bp_nm, A.doc_cur, SUM(A.ap_amt) ap_amt, SUM(A.cls_amt) cls_amt, SUM(A.bal_amt) bal_amt,'' Ap_no,a.note_no "
		lgStrSQL = lgStrSQL & " FROM ufn_A_ClsApByBatchPayment(" & FilterVar(strBatchAllcNo , "''", "S") & ") A "
		lgStrSQL = lgStrSQL & " LEFT JOIN b_biz_partner B ON A.pay_bp_cd = B.bp_cd"
		lgStrSQL = lgStrSQL & " GROUP BY A.pay_bp_cd, B.bp_nm, A.doc_cur , A.note_no"
	Else
		lgStrSQL = lgStrSQL & "SELECT " & FilterVar("1", "''", "S") & "  CHECKED, A.PAY_BP_CD, B.BP_NM, A.DOC_CUR,SUM(A.AP_AMT) AP_AMT,SUM(A.BAL_AMT) ClS_AMT, SUM(A.BAL_AMT) BAL_AMT,'' Ap_no,'' note_no "
		
		strFromWhere = " FROM a_open_ap A JOIN b_biz_partner B ON B.bp_cd = A.pay_bp_cd WHERE A.ap_sts = " & FilterVar("O", "''", "S") & " "
		strFromWhere = strFromWhere & " AND A.ap_amt > 0 AND A.conf_fg=" & FilterVar("C", "''", "S") & "  AND A.gl_no <>''"
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
			lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
		End If
	
		' 권한관리 추가 
		strFromWhere = strFromWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL			
		strFromWhere = strFromWhere & strCond
		
		lgStrSQL = lgStrSQL &	strFromWhere
	
		lgStrSQL = lgStrSQL & " GROUP BY A.pay_bp_cd, A.doc_cur, B.bp_nm ORDER BY A.PAY_BP_CD"
	End If	

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then						'R(Read) X(CursorType) X(LockType) 
		lgStrPrevKey = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)						'☜: No data is found. 
		lgErrorStatus = "YES"
		Exit Sub 
	Else
		iDx = 1
		lgstrData = ""
		lgLngMaxRow = Request("txtMaxRows")													'☜: Read Operation Mode (CRUD)
		
		iPayAmount = 0
		Do While Not lgObjRs.EOF
			lgCurrency = ConvSPChars(lgObjRs("DOC_CUR"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHECKED"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_BP_CD"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_CUR"))
									
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AP_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BAL_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Cls_AMT"),ggAmtOfMoney.DecPoint, 0)
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("note_no"))
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
			
			iPayAmount = Cdbl(iPayAmount) + Cdbl(lgObjRs("cls_amt"))

			Call BatchAllcNoToGetApno(lgObjRs("PAY_BP_CD"), strFromWhere,lgObjRs("DOC_CUR"), strBatchAllcNo)
			
			lgObjRs.MoveNext
			iDx =  iDx + 1
   		Loop 
	End If

	Call SubCloseRs(lgObjRs)																'☜: Release RecordSSet

	If CheckSYSTEMError(Err,True) = True Then
		ObjectContext.SetAbort
		Exit Sub
	End If   
End Sub
'============================================================================================================
' Name : BatchAllcNoToGetApno
' Desc : pay_bp_cd, AP_Due_Dt, Doc_Cur에 맞는 ap_ap_no 가져오기 
'============================================================================================================
Sub BatchAllcNoToGetApno(byval PAY_BP_CD, byval strFromWhere, byval Doc_Cur,byVal strBatchAllcNo)
	Dim ii 
	Dim lgStrSQL	 
	
	iPayAmount = 0
	Call SubOpenDB(lgObjConn)   
																							'☜: Make a DB Connection	 
	If Cstr(lgIntFlgMode) = Cstr(OPMD_UMODE) Then
		lgStrSQL	= "select ap_no from ufn_A_ClsApByBatchPayment(" & FilterVar(strBatchAllcNo , "''", "S") & ")"	
		lgStrSQL= lgStrSQL & " WHERE doc_cur = "	& FilterVar(Doc_Cur , "''", "S")
	    lgStrSQL= lgStrSQL & " And pay_bp_cd= "	& FilterVar(PAY_BP_CD , "''", "S") 
	Else
	    lgStrSQL = "SELECT ap_no " & strFromWhere & " And pay_bp_cd= "	& FilterVar(PAY_BP_CD , "''", "S") 
	End If

    Call FncOpenRs("R",lgObjConn,lgObjRs2,lgStrSQL,"X","X") 
    
    If  lgObjRs2.EOF And lgObjRs2.BOF Then  
		
    Else 
        Do While Not lgObjRs2.EOF
			GetApNo= GetApNo & Chr(11) & lgObjRs2("ap_no")
		    lgObjRs2.MoveNext
        Loop 
    End if

    GetApNo= GetApNo & Chr(12) 
    Call SubCloseRs(lgObjRs2)   
End sub

'============================================================================================================
' Name : SubBizQuerySingle
' Desc : Query Data from Db
'============================================================================================================
Function SubBizQuerySingle()
	Dim lgStrSQL

	On Error Resume Next                                                                 '☜: Protect system from crashing
	Err.Clear                                                                            '☜: Clear Error status
   
	lgStrSQL = lgStrSQL & "SELECT A.paym_dt,A.dept_cd, C.dept_nm, A.paym_type, D.minor_nm, A.bank_cd, E.bank_nm, a.note_no,"
	lgStrSQL = lgStrSQL & " A.bank_acct_no, A.gl_no, A.temp_gl_no,  A.doc_cur, A.xch_rate, a.paym_desc, A.acct_cd, F.acct_nm,"
	lgStrSQL = lgStrSQL & " SUM(paym_amt) paym_amt, SUM(paym_loc_amt) paym_loc_amt, g.due_dt,g.card_co_cd , h.card_co_nm  "
	lgStrSQL = lgStrSQL & " FROM a_allc_paym A LEFT JOIN b_acct_dept C	ON C.dept_cd = A.dept_cd AND C.org_change_id = A.org_change_id"
	lgStrSQL = lgStrSQL & " LEFT JOIN b_minor D	ON D.major_cd = " & FilterVar("A1006", "''", "S") & "  AND D.minor_cd = A.paym_type"
	lgStrSQL = lgStrSQL & " LEFT JOIN b_bank E	ON E.bank_cd = A.bank_cd"
	lgStrSQL = lgStrSQL & " JOIN a_acct F ON F.acct_cd = A.acct_cd"
	lgStrSQL = lgStrSQL & " left JOIN f_note g ON g.note_no = A.note_no"
	lgStrSQL = lgStrSQL & " left JOIN b_card_co h ON h.card_co_cd = g.card_co_cd"	
	lgStrSQL = lgStrSQL & " WHERE A.allc_type = " & FilterVar("B", "''", "S") & "  AND A.ref_no = " & FilterVar(strBatchAllcNo , "''", "S")
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
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' 권한관리 추가 
	lgStrSQL = lgStrSQL & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	lgStrSQL = lgStrSQL & " GROUP BY A.paym_dt, A.dept_cd, C.dept_nm,A.paym_type, D.minor_nm, A.bank_cd, E.bank_nm,a.note_no,"
	lgStrSQL = lgStrSQL & " A.bank_acct_no, A.gl_no, A.temp_gl_no,A.doc_cur,A.xch_rate,a.paym_desc,A.acct_cd,F.acct_nm,g.due_dt, "
	lgStrSQL = lgStrSQL & " g.card_co_cd , h.card_co_nm "	
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
		SubBizQuerySingle= False
		lgStrPrevKey = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
		lgErrorStatus = "YES"
		Exit Function 
	Else
		iPaymAmount = 0
		Do While Not lgObjRs.EOF
			Dim lgCurrency
			lgCurrency = ConvSPChars(lgObjRs("Doc_Cur"))

			Response.write "<Script Language=vbscript>													     " & vbCr
			Response.write " With parent																     " & vbCr
			
			Response.Write " .frm1.txtAllcDt.TEXT	    = """ & UNIDateClientFormat(lgObjRs("Paym_Dt")) & """" & vbcr	
			Response.Write " .frm1.txtDeptCd.Value		= """ & ConvSPChars(lgObjRs("Dept_Cd"))		    & """" & vbcr	
			Response.Write " .frm1.txtDeptNm.Value		= """ & ConvSPChars(lgObjRs("Dept_Nm"))		    & """" & vbcr	
			If UCase(Trim(lgObjRs("Paym_Type"))) = "CP" Then
				Response.Write " .frm1.txtBankCd.Value		= """"" & vbcr	
				Response.Write " .frm1.hBankCd.Value		= """"" & vbcr				
				Response.Write " .frm1.txtBankNm.Value		= """"" & vbcr
				Response.Write " .frm1.hBankNm.Value		= """"" & vbcr			    	
			Else
				Response.Write " .frm1.txtBankCd.Value		= """ & ConvSPChars(lgObjRs("Bank_Cd"))		    & """" & vbcr	
				Response.Write " .frm1.hBankCd.Value		= """ & ConvSPChars(lgObjRs("Bank_Cd"))		    & """" & vbcr				
				Response.Write " .frm1.txtBankNm.Value		= """ & ConvSPChars(lgObjRs("Bank_Nm"))		    & """" & vbcr
				Response.Write " .frm1.hBankNm.Value		= """ & ConvSPChars(lgObjRs("Bank_Nm"))		    & """" & vbcr	
			End If											
			Response.Write " .frm1.txtBankAcct.Value	= """ & ConvSPChars(lgObjRs("Bank_Acct_No"))	& """" & vbcr	
			If UCase(Trim(lgObjRs("Paym_Type"))) = "CP" Then			
				Response.Write " .frm1.txtCardCoCd.Value	= """ & ConvSPChars(lgObjRs("Card_Co_Cd"))		& """" & vbcr	
				Response.Write " .frm1.hCardCoCd.Value		= """ & ConvSPChars(lgObjRs("Card_Co_Cd"))		& """" & vbcr				
				Response.Write " .frm1.txtCardCoNm.Value	= """ & ConvSPChars(lgObjRs("Card_Co_Nm"))		& """" & vbcr
				Response.Write " .frm1.hCardCoNm.Value		= """ & ConvSPChars(lgObjRs("Card_Co_Nm"))		& """" & vbcr				
			Else
				Response.Write " .frm1.txtCardCoCd.Value	= """"" & vbcr	
				Response.Write " .frm1.hCardCoCd.Value		= """"" & vbcr				
				Response.Write " .frm1.txtCardCoNm.Value	= """"" & vbcr
				Response.Write " .frm1.hCardCoNm.Value		= """"" & vbcr	
			End If											
		    Response.Write " .frm1.txtAcctCd.Value		= """ & ConvSPChars(lgObjRs("acct_cd"))		    & """" & vbcr	
			Response.Write " .frm1.txtAcctnm.Value		= """ & ConvSPChars(lgObjRs("acct_nm"))         & """" & vbcr	
			Response.Write " .frm1.txtInputType.Value	= """ & ConvSPChars(lgObjRs("Paym_Type"))       & """" & vbcr	
			Response.Write " .frm1.txtInputTypeNm.Value	= """ & ConvSPChars(lgObjRs("Minor_Nm"))        & """" & vbcr	
			Response.Write " .frm1.txtNoteDueDt.Text	= """ & UNIDateClientFormat(lgObjRs("Due_Dt"))  & """" & vbcr	
			Response.Write " .frm1.txtDocCur.value		= """ & ConvSPChars(lgObjRs("Doc_Cur"))         & """" & vbcr	
			Response.Write " .frm1.txtGlNo.value		= """ & ConvSPChars(lgObjRs("Gl_No"))           & """" & vbcr	
			Response.Write " .frm1.txtTempGlNo.value	= """ & ConvSPChars(lgObjRs("Temp_Gl_No"))      & """" & vbcr	
			 			    
			Response.Write " .frm1.txtXchRate.Text		= """ & UNINumClientFormat(lgObjRs("Xch_Rate"), ggAmtOfMoney.DecPoint, 0)                  & """" & vbcr	
			iPaymAmount = iPaymAmount + Cdbl(lgObjRs("paym_amt"))
			
			Response.Write " .frm1.txtPaymAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("Paym_Amt"), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr	
'	 		Response.Write " .frm1.txtPaymLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("Paym_Loc_Amt"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbcr	
			Response.Write " .frm1.txtPaymDesc.value	= """ & ConvSPChars(lgObjRs("paym_desc"))	    & """" & vbCr
					      
			Response.Write " End With                                                                        " & vbCr
			Response.Write " </script>																	     " & vbCr
					
			lgObjRs.MoveNext
	   	Loop 
	End If
    
	Response.write "<Script Language=vbscript>															     " & vbCr    
    Response.Write " parent.frm1.txtPaymAmt.Text	    = """ & UNIConvNumDBToCompanyByCurrency(iPaymAmount, lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr	
	Response.Write " </script>																	             " & vbCr
    
	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	
	SubBizQuerySingle= true
	
	If CheckSYSTEMError(Err,True) = True Then
		ObjectContext.SetAbort
		Exit Function
	End If   
End Function    

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  MakeSQLWhere()
	Dim strTemp

	strDueDt		= UNIConvDate(Request("txtDueDt"))        
    strDocCur		= Trim(Request("txtDocCur"))          
	strDueFg		= Trim(Request("txtDueDtFg"))    
	strPayBpCd		= Trim(Request("txtPayBpCd"))
	strPaymType	    = Trim(Request("txtPaymType"))
	strBatchAllcNo  = Trim(Request("txtBatchAllcNo"))
	strBizAreaCd	  = Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1	  = Trim(UCase(Request("txtBizAreaCd1")))            '사업장To

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))
	
	If strDueFg Then
		strCond = " AND A.ap_due_dt = " & FilterVar(strDueDt , "''", "S") & " AND A.doc_cur = " & FilterVar(strDocCur , "''", "S")
	Else
		strCond = " AND A.ap_due_dt <=" & FilterVar(strDueDt , "''", "S") & " AND A.doc_cur = " & FilterVar(strDocCur , "''", "S")
	End if

	If strPayBpCd <> ""	Then strCond = strCond & " And A.pay_bp_cd  = " & FilterVar(strPayBpCd , "''", "S")

	If strPaymType <> "" Then strCond = strCond & " And A.paym_type = " & FilterVar(strPaymType , "''", "S")	
	
	If strBizAreaCd <> "" then
		strCond = strCond & " AND A.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	Else
		strCond = strCond & " AND A.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	End If
	
	If strBizAreaCd1 <> "" Then
		strCond = strCond & " AND A.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	Else
		strCond = strCond & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	End If

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
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' 권한관리 추가 
	UNIValue(0,7)	= UNIValue(0,7) & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
End Sub

%>
<Script Language=vbscript>
    If "<%=lgErrorStatus%>" = "" Then
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       parent.ggoSpread.ClearSpreadData()       
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"            '☜ : Display data
       parent.frm1.txtTempAPno.value = "<%=GetApNo%>" 
       Parent.ggoSpread.Source  = Parent.frm1.vspdData2
       parent.ggoSpread.ClearSpreadData()
       Parent.DbQueryOk1
       parent.DoSum()
    End If  
</Script>	

