<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : A_RECEIPT
'*  3. Program ID        : f4223mb3
'*  4. Program 이름      : 차입금계획변경 
'*  5. Program 설명      : 차입금계획변경 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2002/04/27
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
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "B","NOCOOKIE","QB")
%>


<%					

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next											' ☜: 
Dim lgADF                                                       ' ☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                 ' ☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1				' ☜ : DBAgent Parameter 선언 

Call HideStatusWnd												' ☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim LngMaxRow													' 현재 그리드의 최대Row
Dim LngRow

Dim ColSep, RowSep 
Dim Where01

Dim strMode														'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim lgStrPrevKey												' Note NO 이전 값 

Dim lgPageNo

Const GroupCount = 30

strMode = Request("txtMode")									'☜ : 현재 상태를 받음 

    lgPageNo = UNICInt(Trim(Request("lgPageNo")),0)             '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgStrPrevKey = "" & UCase(Trim(Request("lgStrPrevKey")))	
		

Call FixUNISQLData()
Call QueryData()

'#########################################################################################################
'												2.1 FixUNISQLData()
'##########################################################################################################	
Sub FixUNISQLData()
    
    Where01 = ""				'상단 single
    Where01 = Where01 & " LN.LOAN_NO, LN.LOAN_NM, LN.LOAN_DT, LN.DUE_DT, LN.INT_PAY_STND,  "
    Where01 = Where01 & " LN.LOAN_FG,  MN1.MINOR_NM FG_MINOR_NM, LOAN_ACCT_CD, AC.ACCT_NM, "
    Where01 = Where01 & " LN.LOAN_TYPE, MN2.MINOR_NM TYPE_MINOR_NM, LN.LOAN_INT_RATE, "
    Where01 = Where01 & " LN.DOC_CUR, LN.XCH_RATE, LN.LOAN_AMT, LN.LOAN_LOC_AMT, "    
	Where01 = Where01 & " (LN.loan_amt - ISNULL(LN.bas_rdp_amt,0) - ISNULL(PR.pay_amt, 0)) LOAN_BAL_AMT, "
	Where01 = Where01 & " (LN.loan_loc_amt -  ISNULL(LN.bas_rdp_loc_amt,0) - ISNULL(PR.pay_loc_amt_FOR_LOAN, 0)) LOAN_BAL_LOC_AMT," 
	Where01 = Where01 & " ISNULL(PR.pay_amt, 0) TOT_PR_RDP_AMT, ISNULL(PR.pay_loc_amt, 0) TOT_PR_RDP_LOC_AMT, "
	Where01 = Where01 & " ISNULL(IT.pay_amt, 0) TOT_INT_PAY_AMT, ISNULL(IT.pay_loc_amt, 0) TOT_INT_PAY_LOC_AMT, "
	Where01 = Where01 & " ISNULL(PLPR.plan_amt, 0) TOT_PR_PLAN_AMT, ISNULL(PLPR.plan_loc_amt, 0) TOT_PR_PLAN_LOC_AMT, "
	Where01 = Where01 & " ISNULL(PLIT.plan_amt, 0) TOT_INT_PLAN_AMT, ISNULL(PLIT.plan_loc_amt, 0) TOT_INT_PLAN_LOC_AMT"			


    Redim UNISqlId(1)                                                      '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "F4223MA101"	'차입금마스터정보 
    UNISqlId(1) = "F4223MA102"	'차입금상환 및 계획 정보 

    Redim UNIValue(1,2)

	UNIValue(0,0) = Where01		
	UNIValue(0,1) = "" & Filtervar(UCase(Request("txtLoanNo"))	, "", "S")
	UNIValue(0,2) = "" '권한필드 
	
	UNIValue(1,0) = "" & Filtervar(UCase(Request("txtLoanNo"))	, "", "S")
	UNIValue(1,1) = ""'권한필드 
			
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode	    
    
End Sub

'#########################################################################################################
'												2.2 QueryData()
'##########################################################################################################	
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If

    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		rs1.Close:		Set rs1 = Nothing
		%><Script Language=vbscript>parent.frm1.txtLoanNo.Focus</Script><%
		Set lgADF = Nothing
	Else
		Call  MakeSpreadSheetData()
    End If						
		    
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing    	    
															'☜: 비지니스 로직 처리를 종료함 
End Sub

'#########################################################################################################
'												2.4. HTML 결과 생성부 
'##########################################################################################################		
Sub MakeSpreadSheetData()
	Dim intLoopCnt
%>
<Script Language=vbscript>
Option Explicit
	Dim LngMaxRow       
	Dim strData	
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	'-----------------------
	'Result data display area
	'-----------------------	

	With parent
		Call .CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & parent.FilterVar("F2040", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		lgF0 = Replace(lgF0, Chr(11), vbTab)
		.ggoSpread.SetCombo lgF0, .C_PAY_OBJ_CD
		lgF1 = Replace(lgF1, Chr(11), vbTab)
		.ggoSpread.SetCombo lgF1, .C_PAY_OBJ_NM
	End With

	With parent.frm1
		.txtLoanNm.value			= "<%=ConvSPChars(rs0("LOAN_NM"))%>"
		.txtLoanDt.value			= "<%=UNIDateClientFormat(rs0("LOAN_DT"))%>"
		.txtDueDt.value				= "<%=UNIDateClientFormat(rs0("DUE_DT"))%>"
		.cboLoanFg.value			= "<%=ConvSPChars(rs0("LOAN_FG"))%>"
		.htxtLoanFgNm.value			= "<%=ConvSPChars(rs0("FG_MINOR_NM"))%>"
		.txtLoanAcctCd.value		= "<%=ConvSPChars(rs0("LOAN_ACCT_CD"))%>"
		.txtLoanAcctNm.value		= "<%=ConvSPChars(rs0("ACCT_NM"))%>"		
		.txtLoanType.value			= "<%=ConvSPChars(rs0("LOAN_TYPE"))%>"
		.txtLoanTypeNm.value		= "<%=ConvSPChars(rs0("TYPE_MINOR_NM"))%>"
		.txtIntRate.Text			= "<%=UNINumClientFormat(rs0("LOAN_INT_RATE"), ggExchRate.DecPoint, 0)%>"
		.txtDocCur.Value			= "<%=ConvSPChars(rs0("DOC_CUR"))%>"								
        .txtXchrate.Text			= "<%=UNINumClientFormat(rs0("XCH_RATE"), ggExchRate.DecPoint, 0)%>"	
		.txtLoanAmt.Text			= "<%=UNINumClientFormat(rs0("LOAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanLocAmt.Text			= "<%=UNINumClientFormat(rs0("LOAN_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanBalAmt.Text			= "<%=UNINumClientFormat(rs0("LOAN_BAL_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLoanBalLocAmt.Text		= "<%=UNINumClientFormat(rs0("LOAN_BAL_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotPrRdpAmt.Text		= "<%=UNINumClientFormat(rs0("TOT_PR_RDP_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotPrRdpLocAmt.Text		= "<%=UNINumClientFormat(rs0("TOT_PR_RDP_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotIntPayAmt.Text		= "<%=UNINumClientFormat(rs0("TOT_INT_PAY_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotIntPayLocAmt.Text	= "<%=UNINumClientFormat(rs0("TOT_INT_PAY_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotPrPlanAmt.Text		= "<%=UNINumClientFormat(rs0("TOT_PR_PLAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotPrPlanLocAmt.Text	= "<%=UNINumClientFormat(rs0("TOT_PR_PLAN_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotIntPlanAmt.Text		= "<%=UNINumClientFormat(rs0("TOT_INT_PLAN_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.txtTotIntPlanLocAmt.Text	= "<%=UNINumClientFormat(rs0("TOT_INT_PLAN_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		.htxtIntPayStnd.value		= "<%=ConvSPChars(rs0("INT_PAY_STND"))%>"
 
<%	 
		intLoopCnt = rs1.recordcount

%>  
	End With
	LngMaxRow = parent.frm1.vspdData.MaxRows										'Save previous Maxrow
<%
    If cint(intLoopCnt) = 0 Then
		rs0.close:			Set rs0 = Nothing:	                                                    '☜: ActiveX Data Factory Object Nothing
		rs1.close:			Set rs1 = Nothing		
		Set lgADF = Nothing        		                                            '☜: ActiveX Data Factory Object Nothing
	Else

		For LngRow = 1 To intLoopCnt
%>		  
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs1("pay_plan_dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs1("pay_dt"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("pay_obj"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("plan_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("plan_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("pay_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("pay_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("resl_fg"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("doc_cur"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("xch_rate"), ggExchRate.DecPoint, 0)%>" 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("flt_conv_fg"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs1("plan_desc"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs1("pay_plan_dt"))%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("plan_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs1("plan_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
			strData = strData & Chr(11) & Chr(12)

<%  	
		rs1.MoveNext
		Next
%>

		With parent
			.ggoSpread.Source = .frm1.vspddata
			.frm1.vspdData.Redraw = False

			.ggoSpread.SSShowData strData , "F"
			Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,LngMaxRow + 1, LngMaxRow + <%=LngRow%> - 1 ,.frm1.txtDocCur.value,.C_PAY_PLAN_AMT,   "A" ,"I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,LngMaxRow + 1, LngMaxRow + <%=LngRow%> - 1 ,.frm1.txtDocCur.value,.C_PAY_AMT,   "A" ,"I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,LngMaxRow + 1, LngMaxRow + <%=LngRow%> - 1 ,.frm1.txtDocCur.value,.C_H_PAY_PLAN_AMT,   "A" ,"I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.Frm1.vspdData,LngMaxRow + 1, LngMaxRow + <%=LngRow%> - 1 ,.frm1.txtDocCur.value,.C_H_PAY_CHG_AMT,   "A" ,"I","X","X")

			If .lgStrPrevKey <> "" Then
				.DbQuery
			Else
				.frm1.htxtLoanNo.value	= "<%=ConvSPChars(Request("txtLoanNo"))%>"									
				.DbQueryOK
			End If
			.frm1.vspdData.Redraw = True
			
		End With

	With parent
		Call .CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & parent.FilterVar("F2040", "''", "S") & "  AND minor_cd IN (" & parent.FilterVar("<%=ConvSPChars(rs0("LOAN_FG"))%>", "''", "S") & " ," & parent.FilterVar("<%=ConvSPChars(rs0("INT_PAY_STND"))%>", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		lgF0 = Replace(lgF0, Chr(11), vbTab)
		.ggoSpread.SetCombo lgF0, .C_PAY_OBJ_CD
		lgF1 = Replace(lgF1, Chr(11), vbTab)
		.ggoSpread.SetCombo lgF1, .C_PAY_OBJ_NM
	End With

<%
		rs0.close:			Set rs0 = Nothing:	                                                    '☜: ActiveX Data Factory Object Nothing
		rs1.close:			Set rs1 = Nothing		
		Set lgADF = Nothing        		                                            '☜: ActiveX Data Factory Object Nothing

%>    

<%
    End If
%>
</script>
<%      
End Sub
		
%>		


