<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f4234mb3
'*  4. Program 이름      : 만기일연장등록 
'*  5. Program 설명      : 만기일연장등록 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2002/05/11
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : 오수민 
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear											' ☜: 

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")

Dim lgADF                                                       ' ☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                 ' ☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0					' ☜ : DBAgent Parameter 선언 

Call HideStatusWnd												' ☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim LngMaxRow													' 현재 그리드의 최대Row
Dim LngRow

Dim ColSep, RowSep 
Dim Where01, Where02, Where03, Where04

Dim strMode														'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim lgStrPrevKey												' Note NO 이전 값 
Dim txtCommand

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


Const GroupCount = 30

	strMode			= Request("txtMode")									'☜ : 현재 상태를 받음 
	txtCommand		= Request("txtCommand")
'	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)             '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgStrPrevKey	= "" & UCase(Trim(Request("lgStrPrevKey")))	

	gChangeOrgId	= GetGlobalInf("gChangeOrgId")

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


Call FixUNISQLData()
Call QueryData()

'#########################################################################################################
'												2.1 FixUNISQLData()
'##########################################################################################################	
Sub FixUNISQLData()

	Where01 = ""    
	Where01 = Where01 & "RO.LOAN_NO,LN.LOAN_NO,RO.LOAN_BANK_CD,BN.BANK_NM,"
	Where01 = Where01 & "LN.LOAN_AMT,LN.LOAN_LOC_AMT,LI.PAY_AMT LN_INT_PAY_AMT,LI.PAY_LOC_AMT LN_INT_PAY_LOC_AMT,"
	Where01 = Where01 & "LP.PAY_AMT LN_RDP_AMT,LP.PAY_LOC_AMT LN_RDP_LOC_AMT,"
	Where01 = Where01 & "ISNULL(LN.LOAN_AMT,0)-ISNULL(LP.PAY_AMT,0) LN_BAL_AMT,ISNULL(LN.LOAN_LOC_AMT,0)-ISNULL(LP.PAY_LOC_AMT,0) LN_BAL_LOC_AMT,"
	Where01 = Where01 & "RO.LOAN_NM,RO.DEPT_CD,DP.DEPT_NM,RO.LOAN_DT,RO.DUE_DT," 
	Where01 = Where01 & "RO.LOAN_FG,AC1.ACCT_CD LOAN_ACCT_CD,AC1.ACCT_NM LOAN_ACCT_NM,"
	Where01 = Where01 & "RO.LOAN_TYPE,MN.MINOR_NM,RO.DOC_CUR,RO.XCH_RATE," 
	Where01 = Where01 & "RO.PR_RDP_COND,RO.LOAN_AMT,RO.LOAN_LOC_AMT,RO.ST_PR_RDP_DT,RO.PR_RDP_PERD," & vbcr
	Where01 = Where01 & "RO.INT_VOTL,RO.LOAN_INT_RATE,RO.INT_PAY_STND,AC2.ACCT_CD INT_ACCT_CD,AC2.ACCT_NM INT_ACCT_NM,"
	Where01 = Where01 & "RO.ST_INT_DUE_DT,RO.INT_PAY_PERD,RO.INT_PAY_PERD_BASE,RO.DAY_MTHD,RO.INT_BASE_MTHD,"
	Where01 = Where01 & "RO.ST_ADV_INT_PAY_AMT,RO.ST_ADV_INT_PAY_LOC_AMT,LF.PAY_AMT,LF.PAY_LOC_AMT,"
	Where01 = Where01 & "AC3.ACCT_CD ITEM_ACCT_CD,AC3.ACCT_NM ITEM_ACCT_NM,ME.MEAN_TYPE,MI.MINOR_NM,"
	Where01 = Where01 & "AC4.ACCT_CD MEAN_ACCT_CD,AC4.ACCT_NM MEAN_ACCT_NM,ME.BANK_ACCT_NO,ME.BANK_CD,"
	Where01 = Where01 & "RO.USER_FLD1,RO.USER_FLD2,RO.LOAN_DESC," 
	Where01 = Where01 & "ISNULL(RP.PAY_AMT,0) RO_RDP_AMT,ISNULL( RP.PAY_LOC_AMT,0) RO_RDP_LOC_AMT," & vbcr
	Where01 = Where01 & "ISNULL(RI.PAY_AMT,0) RO_INT_PAY_AMT,ISNULL(RI.PAY_LOC_AMT,0) RO_INT_PAY_LOC_AMT,"
	Where01 = Where01 & "ISNULL(RO.LOAN_AMT,0)-ISNULL(RP.PAY_AMT,0) RO_BAL_AMT," 
	Where01 = Where01 & "ISNULL(RO.LOAN_LOC_AMT,0)-ISNULL(RP.PAY_LOC_AMT,0) RO_BAL_LOC_AMT,"
	Where01 = Where01 & "RO.RDP_CLS_FG,RO.RDP_SPRD_FG,RO.CLS_RO_FG,RO.TEMP_GL_NO,RO.GL_NO,"
	Where01 = Where01 & "LN.LOAN_BANK_CD,ME.MEAN_TYPE,ME.PAY_MEAN_ACCT_CD,LN.BANK_ACCT_NO,LN.BANK_CD,RO.ORG_CHANGE_ID,"				'65~69
	Where01 = Where01 & "AC5.ACCT_CD ITEM_BP_ACCT_CD,AC5.ACCT_NM ITEM_BP_ACCT_NM,LB.PAY_AMT,LB.PAY_LOC_AMT,BK2.BANK_NM"				'65~69

	Where02 = ""
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj IN (" & FilterVar("SL", "''", "S") & " ," & FilterVar("SN", "''", "S") & " ," & FilterVar("LL", "''", "S") & " ," & FilterVar("LN", "''", "S") & " ) GROUP BY loan_no) RP ON HS.loan_no=RP.loan_no"
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj IN (" & FilterVar("PI", "''", "S") & " ," & FilterVar("IA", "''", "S") & " ," & FilterVar("DI", "''", "S") & " ," & FilterVar("AI", "''", "S") & " ) GROUP BY loan_no, PAY_ITEM_ACCT_CD) RI ON HS.loan_no=RI.loan_no"
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj IN (" & FilterVar("SL", "''", "S") & " ," & FilterVar("SN", "''", "S") & " ," & FilterVar("LL", "''", "S") & " ," & FilterVar("LN", "''", "S") & " ) GROUP BY loan_no) LP ON HS.ref_loan_no=LP.loan_no"
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj IN (" & FilterVar("PI", "''", "S") & " ," & FilterVar("IA", "''", "S") & " ," & FilterVar("DI", "''", "S") & " ," & FilterVar("AI", "''", "S") & " ) GROUP BY loan_no) LI ON HS.ref_loan_no=LI.loan_no "
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,PAY_ITEM_ACCT_CD,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj=" & FilterVar("BC", "''", "S") & "  GROUP BY loan_no,PAY_ITEM_ACCT_CD) LF ON HS.loan_no=LF.loan_no "
	Where02 = Where02 & " LEFT JOIN (SELECT loan_no,PAY_ITEM_ACCT_CD,SUM(ISNULL(pay_amt,0)) pay_amt,SUM(ISNULL(pay_loc_amt,0)) pay_loc_amt 	FROM f_ln_repay_item WHERE pay_obj=" & FilterVar("BP", "''", "S") & "  GROUP BY loan_no,PAY_ITEM_ACCT_CD) LB ON HS.loan_no=LB.loan_no "
	
	Where04 = ""
	Where04 = Where04 & " INNER JOIN B_ACCT_DEPT DP ON RO.DEPT_CD = DP.DEPT_CD"
	Where04 = Where04 & " LEFT JOIN B_BANK BK2 ON ME.BANK_CD = BK2.BANK_CD"
	Where04 = Where04 & "	 INNER JOIN B_MINOR MN ON RO.LOAN_TYPE = MN.MINOR_CD"
	Where04 = Where04 & " INNER JOIN A_ACCT AC1 ON RO.LOAN_ACCT_CD = AC1.ACCT_CD"
	Where04 = Where04 & " INNER JOIN A_ACCT AC2 ON RO.INT_ACCT_CD = AC2.ACCT_CD"
	Where04 = Where04 & " LEFT JOIN A_ACCT AC3 ON LF.PAY_ITEM_ACCT_CD = AC3.ACCT_CD"
	Where04 = Where04 & " LEFT JOIN A_ACCT AC4 ON ME.PAY_MEAN_ACCT_CD = AC4.ACCT_CD"
	Where04 = Where04 & " LEFT JOIN A_ACCT AC5 ON LB.PAY_ITEM_ACCT_CD = AC5.ACCT_CD"
	

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND RO.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND RO.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND RO.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND RO.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	Where03	= Where03	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL



    Select Case  txtCommand 
		Case "LOOKUP"
			Where03 = Where03 & " and RO.LOAN_NO = " & Filtervar(UCase(Trim(Request("txtLoanNo"))), "''", "S")						'입력조건 
        Case "PREV"
			Where03 = Where03 & " and RO.LOAN_NO < " & Filtervar(UCase(Trim(Request("txtLoanNo"))), "''", "S")
			Where03 = Where03 & " ORDER BY RO.LOAN_NO DESC "
		Case "NEXT"
			Where03 = Where03 & " and RO.LOAN_NO > " & Filtervar(UCase(Trim(Request("txtLoanNo"))), "''", "S")
			Where03 = Where03 & " ORDER BY RO.LOAN_NO ASC "
	End Select



    Redim UNISqlId(1)                                                      '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "F4234MA101"	'차입금마스터정보 
'   UNISqlId(1) = "F4234MA102"	'만기연장번호정보 

    Redim UNIValue(1,4)

	UNIValue(0,0) = Where01
	UNIValue(0,1) = Where02
	UNIValue(0,2) = Where04	
	UNIValue(0,3) = Where03
			
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode	    
    
End Sub

'#########################################################################################################
'												2.2 QueryData()
'##########################################################################################################	
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If

    If rs0.EOF And rs0.BOF Then
    
		Select Case  txtCommand 
			Case "LOOKUP"
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '☜ : No data is found. 
				rs0.Close:		Set rs0 = Nothing	
				Set lgADF = Nothing
			Case "PREV"
				Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the starting data. 
				rs0.Close:		Set rs0 = Nothing	
				Set lgADF = Nothing

				txtCommand = "LOOKUP"
				Call FixUNISQLData()
				Call QueryData()
			Case "NEXT"
				rs0.Close:		Set rs0 = Nothing	
				Set lgADF = Nothing

				Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)            '☜ : This is the ending data.
				txtCommand = "LOOKUP"

				Call FixUNISQLData()
				Call QueryData()
		End Select

	Else
		Call  MakeSpreadSheetData()
		rs0.close:			Set rs0 = Nothing:	                               '☜: ActiveX Data Factory Object Nothing	
		Set lgADF = Nothing
    End If						
		    
    												   '☜: 비지니스 로직 처리를 종료함 
End Sub

'#########################################################################################################
'												2.4. HTML 결과 생성부 
'##########################################################################################################		
Sub MakeSpreadSheetData()	
%>
<Script Language=vbscript>
Option Explicit

'<%	
'	Dim lgCurrency
'	lgCurrency = ConvSPChars(rs0("Doc_Cur"))
'%>	
	
	'-----------------------
	'Result data display area
	'-----------------------	
	With parent.frm1		
		.txtLoanRoNo.value			= "<%=ConvSpchars(rs0(0))%>"		
		.txtLoanNo.value			= "<%=ConvSPChars(rs0(1))%>"		
		.txtBankLoanRoCd.value		= "<%=ConvSPChars(rs0(2))%>"
		.txtBankLoanRoNm.value		= "<%=ConvSPChars(rs0(3))%>"
		.txtLoanAmt.Text				= "<%=UNINumClientFormat(rs0(4), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanLocAmt.Text			= "<%=UNINumClientFormat(rs0(5), ggAmtOfMoney.DecPoint, 0)%>"
		.txtIntPayAmt.Text			= "<%=UNINumClientFormat(rs0(6), ggAmtOfMoney.DecPoint, 0)%>"
		.txtIntPayLocAmt.Text		= "<%=UNINumClientFormat(rs0(7), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanBalAmt.Text			= "<%=UNINumClientFormat(rs0(10), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanBalLocAmt.Text		= "<%=UNINumClientFormat(rs0(11), ggAmtOfMoney.DecPoint, 0)%>"
		.txtRdpAmt.Text				= "<%=UNINumClientFormat(rs0(8), ggAmtOfMoney.DecPoint, 0)%>"
		.txtRdpLocAmt.Text			= "<%=UNINumClientFormat(rs0(9), ggAmtOfMoney.DecPoint, 0)%>"
		
		.txtLoanRoNm.value			= "<%=ConvSPChars(rs0(12))%>"						
		.txtDeptCd.value			= "<%=ConvSPChars(rs0(13))%>"
		.txtDeptNm.value			= "<%=ConvSPChars(rs0(14))%>"
		.txtLoanRoDt.text			= "<%=UNIDateClientFormat(rs0(15))%>"
		.txtDueDt.text				= "<%=UNIDateClientFormat(rs0(16))%>"
		.cboLoanFg.value			= "<%=ConvSPChars(rs0(17))%>"		
		.txtLoanAcctCd.Value		= "<%=ConvSPChars(rs0(18))%>"
		.txtLoanAcctNm.Value		= "<%=ConvSPChars(rs0(19))%>"						
		.txtLoanType.value			= "<%=ConvSPChars(rs0(20))%>"
		.txtLoanTypeNm.value		= "<%=ConvSPChars(rs0(21))%>"
		.txtDocCur.value				= "<%=ConvSPChars(rs0(22))%>"
		.txtXchRate.Text				= "<%=UNINumClientFormat(rs0(23), ggExchRate.DecPoint  , 0)%>"
		.cboPrRdpCond.value			= "<%=ConvSPChars(rs0(24))%>"		
		.txtLoanRoAmt.Text			= "<%=UNINumClientFormat(rs0(25), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanRoLocAmt.Text		= "<%=UNINumClientFormat(rs0(26), ggAmtOfMoney.DecPoint, 0)%>"
		.txt1StPrRdpDt.text			= "<%=UNIDateClientFormat(rs0(27))%>"
		.txtPrRdpPerd.Text			= "<%=ConvSPChars(rs0(28))%>"		
		
		If "<%=ConvSPChars(rs0(29))%>" = "X" Then		
			.Rb_IntVotl1.Checked	= True
		Else 
			.Rb_IntVotl2.Checked	= True
		End If
		
		.txtIntRate.Text				= "<%=UNINumClientFormat(rs0(30), ggExchRate.DecPoint  , 0)%>"
		.cboIntPayStnd.value		= "<%=ConvSPChars(rs0(31))%>"
		.txtIntAcctCd.Value			= "<%=ConvSPChars(rs0(32))%>"
		.txtIntAcctNm.Value			= "<%=ConvSPChars(rs0(33))%>"
		.txt1StIntDueDT.text		= "<%=UNIDateClientFormat(rs0(34))%>"				
		.txtIntPayPerd.Text			= "<%=ConvSPChars(rs0(35))%>"		

		If "<%=ConvSPChars(rs0(37))%>" = "YY" Then		
			.Rb_IntStart1.Checked =True																				'시작일포함여부 
			.Rb_IntEnd1.Checked   =True																				'만기일포함여부 
		ElseIf "<%=ConvSPChars(rs0(37))%>" = "YN" Then
			.Rb_IntStart1.Checked =True
			.Rb_IntEnd2.Checked   =True
		ElseIf "<%=ConvSPChars(rs0(37))%>" = "NY" Then
			.Rb_IntStart2.Checked =True
			.Rb_IntEnd1.Checked   =True
		Else 
			.Rb_IntStart2.Checked =True
			.Rb_IntEnd2.Checked   =True
		End If		
		
		.cboIntBaseMthd.value		= "<%=ConvSPChars(rs0(38))%>"	
		.txtStIntPayAmt.Text		= "<%=UNINumClientFormat(rs0(39), ggAmtOfMoney.DecPoint, 0)%>"		
		.txtStIntPayLocAmt.Text	= "<%=UNINumClientFormat(rs0(40), ggAmtOfMoney.DecPoint, 0)%>"
		.txtChargeAmt.Text			= "<%=UNINumClientFormat(rs0(41), ggAmtOfMoney.DecPoint, 0)%>"
		.txtChargeLocAmt.Text		= "<%=UNINumClientFormat(rs0(42), ggAmtOfMoney.DecPoint, 0)%>"
		.txtChargeAcctCd.Value		= "<%=ConvSPChars(rs0(43))%>"
		.txtChargeAcctNm.Value	= "<%=ConvSPChars(rs0(44))%>"				
		.txtRcptType.value			= "<%=ConvSPChars(rs0(45))%>"
		.txtRcptTypeNm.value		= "<%=ConvSPChars(rs0(46))%>"
		.txtRcptAcctCd.Value		= "<%=ConvSPChars(rs0(47))%>"
		.txtRcptAcctNM.Value		= "<%=ConvSPChars(rs0(48))%>"
		.txtBankAcctNo.value		= "<%=ConvSPChars(rs0(49))%>"
		.txtBankCd.value				= "<%=ConvSPChars(rs0(50))%>"
		.txtUserFld1.value			= "<%=ConvSPChars(rs0(51))%>"
		.txtUserFld2.value			= "<%=ConvSPChars(rs0(52))%>"
		.txtLoanDesc.value			= "<%=ConvSPChars(rs0(53))%>"
		
		.txtTotPrRdpRoAmt.Text	= "<%=UNINumClientFormat(rs0(54), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotPrRdpRoLocAmt.Text	= "<%=UNINumClientFormat(rs0(55), ggAmtOfMoney.DecPoint, 0)%>"
		.txtIntPayRoAmt.Text		= "<%=UNINumClientFormat(rs0(56), ggAmtOfMoney.DecPoint, 0)%>"
		.txtIntPayRoLocAmt.Text	= "<%=UNINumClientFormat(rs0(57), ggAmtOfMoney.DecPoint, 0)%>"  
		.txtLoanBalRoAmt.Text		= "<%=UNINumClientFormat(rs0(58), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanBalRoLocAmt.Text	= "<%=UNINumClientFormat(rs0(59), ggAmtOfMoney.DecPoint, 0)%>"  
		
		'.hClsRoFg.value				= "<%=ConvSPChars(rs0(25))%>"	
		.cboRdpClsFg.value			= "<%=ConvSPChars(rs0(60))%>" 		
		.hRdpSprdFg.value			= "<%=ConvSPChars(rs0(61))%>"	
		
		.htxtTempGlNo.value				= "<%=ConvSPChars(rs0(63))%>"
		.htxtGlNo.value					= "<%=ConvSPChars(rs0(64))%>"
		.htxtBankLoanCd.value			= "<%=ConvSPChars(rs0(65))%>"		
		.htxtOrgLoanRcptType.value	= "<%=ConvSPChars(rs0(66))%>"
		.htxtOrgLoanRcptAcctCd.value	= "<%=ConvSPChars(rs0(67))%>"
		.htxtOrgLoanBankAcctNo.value	= "<%=ConvSPChars(rs0(68))%>"
		.htxtOrgLoanBankCd.value			= "<%=ConvSPChars(rs0(69))%>"
		.hOrgChangeId.value			= "<%=ConvSPChars(rs0(70))%>"
		.txtBPAcctCd.value			= "<%=ConvSPChars(rs0(71))%>"
		.txtBPAcctNm.value			= "<%=ConvSPChars(rs0(72))%>"
		.txtBPamt.text				= "<%=UNINumClientFormat(rs0(73), ggAmtOfMoney.DecPoint, 0)%>"
		.txtBPLocamt.text			= "<%=UNINumClientFormat(rs0(74), ggAmtOfMoney.DecPoint, 0)%>"
		.txtBankNm.value			= "<%=ConvSPChars(rs0(75))%>"
<%	 
	
	Set lgADF = Nothing        		                                            '☜: ActiveX Data Factory Object Nothing    	    				

%>			
	End With

	With parent	
'		If .lgStrPrevKey <> "" Then
'			.DbQuery					
'		Else
'			.frm1.htxtLoanNo.value	= "<%=ConvSPChars(Request("txtLoanNo"))%>"									
			.DbQueryOK
'		End If
		
	End With

</script>
<%      
End Sub
		
Sub ReleaseObj()			
	Set rs0 = Nothing
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub			
		
%>		
