<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5114MA1
'*  4. Program Name         : 매출채권수금내역 
'*  5. Program Desc         :
'*  6. Comproxy List        : S51119LookupBillHdrSvr, S51158ListBillCollectingSvr, S51151MaintBillCollectingSvr
'*							  S51115PostOpenArSvr, Fr0019LookupPrrcptSvr
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'==========================================

Const BIZ_PGM_ID = "s5114mb1.asp"
Const BIZ_BillHdr_JUMP_ID = "s5111ma1"
Const BIZ_BillDtl_JUMP_ID = "s5111ma2"

'☆: Spread Sheet의 Column별 상수 
Dim C_BillTypeCd 			'수금유형 
Dim C_BillTypePop			'수금유형팝업 
Dim C_BillTypeNm 			'수금유형명 
Dim C_BillAmt 				'수금액 
Dim C_BillLocAmt 			'수금자국금액 
Dim C_BankCd 				'은행 
Dim C_BankPop 				'은행팝업 
Dim C_BankNm 				'은행명 
Dim C_BankAcct 				'은행계좌번호 
Dim C_BankAcctPop			'은행계좌번호팝업 
Dim C_Note 					'어음번호 
Dim C_NotePop 				'어음번호팝업 
Dim C_PreReceipt 			'선수금 
Dim C_PreReceiptPop 		'선수금팝업 
Dim C_Remark 				'비고 
Dim C_XchRate 				'환율 
Dim C_XchCalop				'환율연산자 
Dim C_BillSeq				'수금순번 

'=========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim arrCollectType		'수금유형 배열 

Dim IsOpenPop						' Popup
Const PostFlag = "PostFlag"

'========================================
Sub initSpreadPosVariables()  

	C_BillTypeCd = 1
	C_BillTypePop = 2
	C_BillTypeNm = 3
	C_BillAmt = 4	
	C_BillLocAmt = 5
	C_BankCd = 6	
	C_BankPop = 7	
	C_BankNm = 8	
	C_BankAcct = 9	
	C_BankAcctPop = 10
	C_Note = 11		
	C_NotePop = 12	
	C_PreReceipt = 13
	C_PreReceiptPop = 14
	C_Remark = 15		
	C_XchRate = 16		
	C_XchCalop = 17		
	C_BillSeq = 18		

End Sub

'========================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           

    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtConBillNo.focus
	Set gActiveElement = document.activeElement 
	
	frm1.btnPostFlag.disabled = True
	frm1.btnPostFlag.value = "확정"
	frm1.rdoExceptBillYes.checked = True
	frm1.rdoPostNo.checked = True
	frm1.btnGLView.disabled = True
	lgBlnFlgChgValue = False
End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    
		
		.ReDraw = False
				
	    .MaxRows = 0	: .MaxCols = 0
	    	
	    .MaxCols = C_BillSeq+1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    '.Col = .MaxCols														'☜: 공통콘트롤 사용 Hidden Column
	    '.ColHidden = True
	    
        Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit C_BillTypeCd, "수금유형", 10,,,5,2
	    ggoSpread.SSSetButton C_BillTypePop
		ggoSpread.SSSetEdit C_BillTypeNm, "수금유형명", 20
		ggoSpread.SSSetFloat C_BillAmt,"수금액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_BillLocAmt,"수금자국액",15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_BankCd, "은행", 10,,,10,2
	    ggoSpread.SSSetButton C_BankPop
		ggoSpread.SSSetEdit C_BankNm, "은행명", 20
		ggoSpread.SSSetEdit C_BankAcct, "은행계좌번호", 18,,,30,2
		ggoSpread.SSSetButton C_BankAcctPop
		ggoSpread.SSSetEdit C_Note, "어음번호", 18,,,30,2
		ggoSpread.SSSetButton C_NotePop
		ggoSpread.SSSetEdit C_PreReceipt, "선수금번호", 18,,,18,2
		ggoSpread.SSSetButton C_PreReceiptPop
		ggoSpread.SSSetEdit C_Remark, "비고", 50,,,200,1	
		ggoSpread.SSSetFloat C_XchRate,"환율",15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_XchCalop, "환율연산자", 15
		ggoSpread.SSSetFloat C_BillSeq,"수금순번" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"				

       call ggoSpread.MakePairsColumn(C_BillTypeCd,C_BillTypePop)
       call ggoSpread.MakePairsColumn(C_BankCd,C_BankPop)
       call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPop)
       call ggoSpread.MakePairsColumn(C_Note,C_NotePop)
       call ggoSpread.MakePairsColumn(C_PreReceipt,C_PreReceiptPop)

       Call ggoSpread.SSSetColHidden(C_BillSeq,C_BillSeq,True)
	   Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column

		.ReDraw = True
   
    End With
	
	If len(frm1.txtCurrency.value) <> 0 Then
		If len(frm1.txtLocCur.value) <> 0 Then
			If frm1.txtCurrency.value = frm1.txtLocCur.value Then
				Call SetInitSpreadSheet()	
			End If
		End If
	End If			
	    
End Sub

'========================================
Sub SetInitSpreadSheet()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		.ReDraw = False
   
       Call ggoSpread.SSSetColHidden(C_BillLocAmt,C_BillLocAmt,True)
       Call ggoSpread.SSSetColHidden(C_XchRate,C_XchRate,True)
       Call ggoSpread.SSSetColHidden(C_XchCalop,C_XchCalop,True)

		.ReDraw = True
   
    End With
    
End Sub

'========================================
Sub SetInitSpreadSheet2()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		
		.ReDraw = False
   
       Call ggoSpread.SSSetColHidden(C_BillLocAmt,C_BillLocAmt,False)
       Call ggoSpread.SSSetColHidden(C_XchRate,C_XchRate,False)
       Call ggoSpread.SSSetColHidden(C_XchCalop,C_XchCalop,False)

		.ReDraw = True
   
    End With
    
End Sub

'========================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	Dim i 
	
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.Source = frm1.vspdData

	ggoSpread.SSSetProtected	C_BillSeq, pvStartRow, pvEndRow    
	ggoSpread.SSSetRequired		C_BillTypeCd, pvStartRow, pvEndRow    
	ggoSpread.SSSetProtected	C_BillTypeNm, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_BillAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_BillLocAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_BankCd, pvStartRow, pvStartRow
	ggoSpread.SSSetProtected	C_BankPop, pvStartRow, pvEndRow		
	ggoSpread.SSSetProtected	C_BankNm, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_BankAcct, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_BankAcctPop, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_Note, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_NotePop, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_PreReceipt, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_PreReceiptPop, pvStartRow, pvEndRow												
	ggoSpread.SSSetProtected	C_XchRate, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_XchCalop, pvStartRow, pvEndRow	

	For i = pvStartRow to pvEndRow
		Call chgProtectOfBillType(i)
	Next

    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================
Sub InitCollectType()	'kek 11.14 
 Dim i
 Dim iCodeArr, iTypeArr

 Err.Clear

 Call CommonQueryRs(" MINOR.MINOR_CD, CONFIG.REFERENCE ", " B_MINOR MINOR, B_CONFIGURATION CONFIG ", " MINOR.MINOR_CD *= CONFIG.MINOR_CD AND MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND CONFIG.SEQ_NO = " & FilterVar("4", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iTypeArr = Split(lgF1, Chr(11))

 If Err.number <> 0 Then
  MsgBox Err.description 
  Err.Clear 
  Exit Sub
 End If

 Redim arrCollectType(UBound(iCodeArr) - 1, 2)

 For i = 0 to UBound(iCodeArr) - 1
  arrCollectType(i, 0) = iCodeArr(i)
  arrCollectType(i, 1) = iTypeArr(i)
 Next
End Sub

'========================================
Function OpenConBillDtl()
	Dim iCalledAspName
	Dim strRet
	Dim arrVal

	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
		
	arrVal = ""
	
	iCalledAspName = AskPRAspName("s5111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111pa1", "x")
		IsOpenPop = False
		exit Function
	end if

	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=A", Array(window.parent,arrVal), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	frm1.txtConBillNo.focus
	If strRet <> "" Then frm1.txtConBillNo.value = strRet 

End Function

'========================================
Function OpenBillTypePop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "수금유형"
	arrParam(1) = "B_CONFIGURATION Config, B_MINOR Minor"		
	arrParam(2) = Trim(strCode)									
	arrParam(3) = ""											
	arrParam(4) = "Config.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND Config.SEQ_NO = " & FilterVar("1", "''", "S") & "  " _
				& "AND Config.MINOR_CD = Minor.MINOR_CD AND Config.MAJOR_CD = Minor.MAJOR_CD " _
				& "AND Config.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )"			
	arrParam(5) = "수금유형"								

	arrField(0) = "Config.MINOR_CD"								
	arrField(1) = "Minor.MINOR_NM"								

	arrHeader(0) = "수금유형"								
	arrHeader(1) = "수금유형명"								

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then	Call SetBillAmtDtl(arrRet, C_BillTypePop)
	
End Function

'========================================
Function OpenBankPop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "입금은행"
	arrParam(1) = "B_BANK BK, F_DPST DP"					
	arrParam(2) = Trim(strCode)							
	arrParam(3) = ""									
	arrParam(4) = "BK.BANK_CD=DP.BANK_CD" 	
	arrParam(5) = "입금은행"						

	arrField(0) = "BK.BANK_CD"					
	arrField(1) = "BK.BANK_NM"						

	arrHeader(0) = "입금은행"						
	arrHeader(1) = "입금은행명"						

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then	Call SetBillAmtDtl(arrRet, C_BankPop)
	
End Function

'========================================
Function OpenBankAcctPop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		frm1.vspdData.Col = C_BankCd
		If Trim(frm1.vspdData.Text) = "" Then
			Call DisplayMsgBox("205152", "X", "은행", "X")
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "은행계좌번호"
		arrParam(1) = "B_BANK BK, F_DPST DP"					
		arrParam(2) = Trim(strCode)							
		arrParam(3) = ""									
		arrParam(4) = "BK.BANK_CD=DP.BANK_CD And BK.BANK_CD = " _
			& FilterVar(Trim(frm1.vspdData.Text), "" , "S") & ""				

		arrParam(5) = "은행계좌번호"						

		arrField(0) = "DP.BANK_ACCT_NO"			
		arrField(1) = "BK.BANK_NM"						
		arrField(2) = "BK.BANK_CD"					

		arrHeader(0) = "은행계좌번호"				
		arrHeader(1) = "은행명"
		arrHeader(2) = "은행"						

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then	Call SetBillAmtDtl(arrRet, C_BankAcctPop)
	
End Function

'========================================
Function OpenNotePop(strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strBpCd, strChargeDt, strChargeLocAmt, strVatLocAmt, strTotAmt, strBillTypeCd

	On Error Resume Next
	
	OpenNotePop = False

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	strBpCd = Trim(frm1.txtSoldtoParty.value)
	strChargeDt = Trim(frm1.txtBillDt.text)

	'# 수금유형(C_BillTypeCd)가 'CR'(수취구매카드)일때	구매카드번호팝업을 탄다.		'son
	'# 수금유형(C_BillTypeCd)가	'NR'(받을어음)일때		어음번호팝업을 탄다.		'son
	frm1.vspdData.Row = IRow
	frm1.vspdData.col = C_BillTypeCd						'수금유형 
	strBillTypeCd = frm1.vspdData.text
	
	If Trim(strBillTypeCd) = "CR" THEN
		arrParam(0) = "구매카드번호"
		arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d"
		arrParam(2) = Trim(strCode)
		arrParam(3) = ""			
		arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & " AND a.note_fg = " & FilterVar("CR", "''", "S") & " and a.bp_cd = b.bp_cd  " & _		
				 " and a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
		arrParam(5) = "구매카드번호"

		arrField(0) = "a.Note_no"
		arrField(1) = "F2" & parent.gColSep & "a.Note_amt"	
		arrField(2) = "DD" & parent.gColSep & "a.Issue_dt"	
		arrField(3) = "b.bp_nm"					
		arrField(4) = "d.card_co_nm"    	    

		arrHeader(0) = "구매카드번호"
		arrHeader(1) = "금액"
		arrHeader(2) = "발행일"
		arrHeader(3) = "거래처"
		arrHeader(4) = "카드사"

	Else
		arrParam(0) = "어음번호"
		arrParam(1) = "F_NOTE"								
		arrParam(2) = Trim(strCode)							
		arrParam(3) = ""									
		arrParam(4) = "NOTE_FG IN (" & FilterVar("D1", "''", "S") & ", " & FilterVar("CR", "''", "S") & ") AND NOTE_STS = " & FilterVar("BG", "''", "S") & " AND BP_CD =  " & FilterVar(strBpCd , "''", "S") & "" _
			& " AND (Convert(CHAR(10), ISSUE_DT, 112) <= '" & UniConvDateToYYYYMMDD(strChargeDt, parent.gDateFormat,"") & _
			"' And Convert(CHAR(10), DUE_DT, 112) >=  " & FilterVar(UniConvDateToYYYYMMDD(strChargeDt, parent.gDateFormat,""), "''", "S") & ")" 

		arrParam(5) = "어음번호"						

		arrField(0) = "NOTE_NO"								
		arrField(1) = "HH" & parent.gColSep & "NOTE_AMT"			
		arrField(2) = "F2" & parent.gColSep & "NOTE_AMT"							
		arrField(3) = "NOTE_FG"								
		arrField(4) = "NOTE_STS"							
		
		arrHeader(0) = "어음번호"						
		arrHeader(1) = "어음금액"						
		arrHeader(2) = "어음금액"						
		arrHeader(3) = "어음구분"						
		arrHeader(4) = "어음상태"						
	End If
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then	OpenNotePop = SetBillAmtDtl(arrRet, C_NotePop)
			
End Function

'===========================================================================
Function OpenPreReceiptPop(ByVal prStrCode)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(4)
	
	OpenPreReceiptPop = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	With frm1
		arrParam(0) = Trim(.txtBillDt.Text)				'매출채권일 
		arrParam(1) = Trim(.txtSoldToParty.value)		'주문처 
		arrParam(2) = Trim(.txtSoldToPartyNm.value)		'주문처 
		arrParam(3) = Trim(.txtCurrency.value)			'화폐 
		.vspddata.col = C_PreReceipt
		arrParam(4) = Trim(.vspddata.text)				'선수금번호 
	End With	
	
	iCalledAspName = AskPRAspName("s5111ra7")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111ra7", "x")
		IsOpenPop = False
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName & "?txtFlag=CH&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) <> "" Then	OpenPreReceiptPop = SetBillAmtDtl(arrRet, C_PreReceiptPop)
			
End Function

'===========================================================================
Function SetBillAmtDtl(Byval arrRet,ByVal iWhere)

	With frm1

		Select Case iWhere
		Case C_BillTypePop	'수금유형 
			.vspdData.Col = C_BillTypeCd	:	.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BillTypeNm	:	.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_BillTypeCd, .vspdData.Row)		<% ' 변경이 읽어났다고 알려줌 %>

		Case C_BankPop		'은행 
			.vspdData.Col = C_BankCd		:	.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BankNm		:	.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_BankCd, .vspdData.Row)
		
		Case C_BankAcctPop	'은행계좌번호 
			.vspdData.Col = C_BankAcct		:	.vspdData.Text = arrRet(0)
			Call vspdData_Change(C_BankAcct, .vspdData.Row)
			
		Case C_NotePop	'어음번호 
			.vspdData.Col = C_Note		:	.vspdData.Text = arrRet(0)
			'수금자국금액(어음은 Local Currency에 대해서만 등록가능)
			.vspdData.Col = C_BillAmt	:	.vspdData.Text = arrRet(1)
			.vspdData.Col = C_BillLocAmt	:	.vspdData.Text = arrRet(1)
			Call BillTotalSum(C_BillAmt)

		Case C_PreReceiptPop	'선수금번호 
			.vspdData.Col = C_PreReceipt		:	.vspdData.Text = arrRet(1)
			.vspdData.Col = C_XchRate			:	.vspdData.Text = UNIFormatNumber(arrRet(8), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
'			원화금액계산 
			Call ClientXchRateCalcu(.vspdData.Row)
		End Select

	End With

	lgBlnFlgChgValue = True
	
End Function

'====================================================
Function GetNoteInfo(IRow)
	Dim strSoldToParty, strNoteNO, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strNoteInfo
	
	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_Note						'품목코드 
		strNoteNo = .vspdData.text

		strSoldToParty = .txtSoldtoParty.value		'주문처 
		strValidDt = UniConvDateToYYYYMMDD(.txtBillDt.Text, parent.gDateFormat,"")
	End With

	if Trim(strNoteNo) = "" Then Exit Function
	
	strSelectList = " note_amt "
	strFromList  = " f_note "
	strWhereList = " note_no =  " & FilterVar(strNoteNo , "''", "S") & " AND bp_cd =  " & FilterVar(strSoldToParty , "''", "S") & " AND note_fg IN (" & FilterVar("D1", "''", "S") & ", " & FilterVar("CR", "''", "S") & ") AND note_sts = " & FilterVar("BG", "''", "S") & " " & _
					" AND issue_dt <=  " & FilterVar(strValidDt , "''", "S") & " AND due_dt >=  " & FilterVar(strValidDt , "''", "S") & ""
    Err.Clear
    
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		strNoteInfo = Split(strRs, Chr(11))
		frm1.vspdData.Col = C_BillAmt
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(strNoteInfo(1), parent.gCurrency, parent.ggAmtOfMoneyNo, "X" , "X")
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(strNoteInfo(1), parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		Call BillTotalSum(C_BillAmt)
		Exit Function
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		End If
		
		If Not OpenNotePop(strNoteNo) Then
			'취소한 경우 입력된 내용을 clear한다.
			frm1.vspdData.Col = C_Note
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_BillAmt
			frm1.vspdData.text = "0"
			frm1.vspdData.Col = C_BillLocAmt
			frm1.vspdData.text = "0"
			Call BillTotalSum(C_BillAmt)
		End if
	End if
End Function

'====================================================
Function GetPreReceiptInfo(byVal prIntRow)
	Dim iStrPreReceiptNo
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrPreReceiptInfo
	
	With frm1
		.vspdData.Row = prIntRow
		.vspdData.col = C_PreReceipt						'선수금 번호 
		iStrPreReceiptNo = .vspdData.text
	End With

	If Trim(iStrPreReceiptNo) = "" Then
		With frm1
			.vspdData.col = C_XchRate
			.vspdData.Text = .HXchRate.value
		End With
		Call ClientXchRateCalcu(prIntRow)		' 원화금액 재계산 
		Exit Function
	End If
		
	iStrSelectList = " FP.xch_rate "
	iStrFromList  = " f_prrcpt FP INNER JOIN a_jnl_item AJ ON (FP.prrcpt_type = AJ.jnl_cd) "
	With frm1
		iStrWhereList = " FP.bp_cd =  " & FilterVar(.txtSoldtoParty.value , "''", "S") & " AND FP.doc_cur =  " & FilterVar(.txtCurrency.value , "''", "S") & "" & _
					   " AND FP.prrcpt_dt <=  " & FilterVar(UniConvDateAToB(.txtBillDt.Text, parent.gDateFormat,parent.gAPDateFormat), "''", "S") & "" & _
					   " AND FP.bal_amt > 0 AND FP.conf_fg = " & FilterVar("C", "''", "S") & "  AND AJ.jnl_type = " & FilterVar("PR", "''", "S") & " AND FP.prrcpt_no =  " & FilterVar(iStrPreReceiptNo , "''", "S") & ""
	End With

    Err.Clear
 
 	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrPreReceiptInfo = Split(iStrRs, Chr(11))
		frm1.vspdData.Col = C_XchRate
		frm1.vspdData.text = UNIFormatNumber(iArrPreReceiptInfo(1), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		Call ClientXchRateCalcu(prIntRow)		' 원화금액 재계산 
		Exit Function
	Else
		If Not OpenPreReceiptPop(iStrPreReceiptNo) Then
			'취소한 경우 입력된 내용을 clear한다.
			With frm1
				.vspdData.Col = C_PreReceipt
				.vspdData.Text = ""
				.vspdData.col = C_XchRate
				.vspdData.Text = .HXchRate.value
			End With
			Call ClientXchRateCalcu(prIntRow)		' 원화금액 재계산 
		End if
	End if
End Function

'====================================================
Sub ClientXchRateCalcu(ByVal Row)

	Dim ldbBillAmt, ldbXchgRate

	frm1.vspdData.Row = Row
		
	frm1.vspdData.Col = C_BillAmt	:	ldbBillAmt = UNICDbl(Trim(frm1.vspdData.Text))
	frm1.vspdData.Col = C_XchRate	:	ldbXchgRate = UNICDbl(Trim(frm1.vspdData.Text))

	frm1.vspdData.Col = C_XchCalop
	Select Case Trim(frm1.vspdData.Text)
	Case "+"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt + ldbXchgRate, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
	Case "-"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt - ldbXchgRate, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
	Case "*"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text= UNIConvNumPCToCompanyByCurrency(ldbBillAmt * ldbXchgRate, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
	Case "/"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt / ldbXchgRate, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
	End Select

	Call BillTotalSum(C_BillAmt)

End Sub

'====================================================
Function chgProtectOfBillType(ByVal Row)

					'관련번호	은행 
	'-----------------------------------
	'예적금	 DP		Edit:O		Edit:O
	'받을어음NR		Edit:O		Edit:X
	'선수금	 PR		Edit:O		Edit:X
	'현금	 CS		Edit:X		Edit:X
	Dim iCnt
	Dim strRefVal
	
	With frm1

		'.vspdData.ReDraw = False

		ggoSpread.Source = frm1.vspdData
		.vspdData.Col = C_BillTypeCd	:	.vspdData.Row = Row
		
		strRefVal = GetCollectTypeRef(UCase(Trim(.vspdData.Text)))

		Select Case strRefVal
		Case "DP"	'예적금			
			ggoSpread.SpreadUnLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadUnLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row

			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_BankCd, Row, Row
				ggoSpread.SSSetRequired C_BankAcct, Row, Row
			End If
		Case "NO"	'받을어음(어음은 분할 될 수 없다. 따라서 금액도 수정할 수 없음)
		
			if frm1.txtCurrency.value <> frm1.txtLocCur.value Then
				.vspdData.Text = ""
				Call DisplayMsgBox("205628", "X", "X", "X")
				Exit Function
			End if
			ggoSpread.SpreadUnLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			If GetSetupMod(parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_Note, Row, Row
			Else
				ggoSpread.SSSetProtected C_NotePop, Row, Row
			End If
		Case "PR"	'선수금 
			ggoSpread.SpreadUnLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row

			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_PreReceipt, Row, Row
			Else
				ggoSpread.SSSetProtected C_PreReceiptPop, Row, Row
			End If			
		Case "CS"	'현금,수표 
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row		
			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

		Case Else
			.vspdData.Col = C_BillTypeNm : 	.vspdData.text = ""
			.vspdData.Col = C_BillAmt : 	.vspdData.text = ""		
			.vspdData.Col = C_BillLocAmt : 	.vspdData.text = ""		
			.vspdData.Col = C_BankCd : 	.vspdData.text = ""		
			.vspdData.Col = C_BankNm : 	.vspdData.text = ""
			.vspdData.Col = C_BankAcct : 	.vspdData.text = ""				
			.vspdData.Col = C_Note : 	.vspdData.text = ""				
			.vspdData.Col = C_PreReceipt : 	.vspdData.text = ""				
			.vspdData.Col = C_Remark : 	.vspdData.text = ""				
			.vspdData.Col = C_XchRate : 	.vspdData.text = ""
			.vspdData.Col = C_XchCalop : 	.vspdData.text = ""																
					
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadUnLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadUnLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadUnLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadUnLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(parent.gSetupMod, "A") <> "Y" Then
				ggoSpread.SSSetProtected C_NotePop, Row, Row
				ggoSpread.SSSetProtected C_PreReceiptPop, Row, Row
			End If				
						
		End Select

		.vspdData.Col = 0	:	.vspdData.Row = Row
		Select Case Trim(.vspdData.Text)
	    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
			.vspdData.Row = Row
			.vspdData.Col = C_BankCd		:	.vspdData.Text = ""
			.vspdData.Col = C_BankNm		:	.vspdData.Text = ""
			.vspdData.Col = C_BankAcct		:	.vspdData.Text = ""
			.vspdData.Col = C_Note			:	.vspdData.Text = ""
			.vspdData.Col = C_PreReceipt	:	.vspdData.Text = ""
			.vspdData.Col = C_XchCalop		: 	.vspdData.text = TRim(.HXchRateOp.value)
			.vspdData.Col = C_XchRate		:	.vspdData.Text = TRim(.HXchRate.value)
		End Select

	End With
	
End Function

'====================================================
Sub BillTotalSum(ByVal Col)

	Select Case Col
	Case C_BillAmt, C_BillLocAmt
	Case Else
		Exit Sub
	End Select

	Dim SumBillAmt, BillAmt, SumBillLocAmt, BillLocAmt
	Dim lRow

	SumBillAmt = 0
	SumBillLocAmt = 0

	ggoSpread.source = frm1.vspdData
	For lRow = 1 To frm1.vspdData.MaxRows 
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then

			Select Case Col 
			Case C_BillAmt
				frm1.vspdData.Col = C_BillAmt		'수금금액 
				BillAmt = UNICDbl(frm1.vspdData.Text)

				SumBillAmt = SumBillAmt + BillAmt

				frm1.vspdData.Col = C_BillLocAmt	'수금자국금액 
				BillLocAmt = UNICDbl(frm1.vspdData.Text)

				SumBillLocAmt = SumBillLocAmt + BillLocAmt

			Case C_BillLocAmt
				frm1.vspdData.Col = C_BillLocAmt	'수금자국금액 
				BillLocAmt = UNICDbl(frm1.vspdData.Text)

				SumBillLocAmt = SumBillLocAmt + BillLocAmt
			
			End Select

		End If
	Next

	frm1.txtSumBillAmt.Text = UNIConvNumPCToCompanyByCurrency(SumBillAmt, frm1.txtCurrency.value, parent.ggAmtOfMoneyNo, "X" , "X")
	frm1.txtSumLocBillAmt.Text = UNIConvNumPCToCompanyByCurrency(SumBillLocAmt, parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")

End Sub

'========================================
Sub SetPostYesSpreadColor(ByVal lRow)

    With frm1

		Call SetToolbar("11100000000111")
    
		.vspdData.ReDraw = False
    
		Dim GridCol
		For GridCol = 1 To .vspdData.MaxCols
			ggoSpread.SSSetProtected GridCol, 1, .vspdData.MaxRows
		Next
    
		.vspdData.ReDraw = True
    
    End With

End Sub

'========================================
Sub SetQuerySpreadColor(ByVal lRow)
  
    With frm1
    
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired		C_BillTypeCd, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected	C_BillTypeNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetRequired		C_BillAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetRequired		C_BillLocAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected	C_XchRate, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected	C_XchCalop, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected	C_BillSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected	C_BankNm, lRow, .vspdData.MaxRows

		Dim GridRow
		For GridRow = 1 To .vspdData.MaxRows
			Call chgProtectOfBillType(GridRow)
		Next

		.vspdData.ReDraw = True

    End With

End Sub

'========================================
Function CookiePage(Byval Kubun)

	On Error Resume Next

	Const CookieSplit = 4877
	
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtHBillNo.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtConBillNo.value =  arrVal(0)

		WriteCookie CookieSplit , ""
		
		Call DbQuery()
			
	End If
	
End Function

'====================================================
Function JumpChgCheck(DesID)

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)

	Select Case DesID
	Case BIZ_BillHdr_JUMP_ID
		Call PgmJump(BIZ_BillHdr_JUMP_ID)
	Case BIZ_BillDtl_JUMP_ID
		Call PgmJump(BIZ_BillDtl_JUMP_ID)
	End Select	

End Function

'====================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD
	ggoSpread.Source = frm1.vspdData	

	If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then Exit Function
	End If

	If ggoSpread.SSCheckChange = False Then
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function

'====================================================
Sub CurFormatNumericOCX()

	With frm1
		'총매출채권금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotBillAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		'총 수금액 
		ggoOper.FormatFieldByObjectOfCur .txtSumBillAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With

End Sub

'====================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData
		'수금액 
		ggoSpread.SSSetFloatByCellOfCur C_BillAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With

End Sub

'====================================================
Function GetCollectTypeRef(CollectType)
	Dim iCnt
	For iCnt = 0 to UBound(arrCollectType)
		If arrCollectType(iCnt,0) = CollectType Then
			GetCollectTypeRef = arrCollectType(iCnt,1)
			Exit Function
		End If
	Next
	GetCollectTypeRef = ""	
End Function

'========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_BillTypeCd	= iCurColumnPos(1)
			C_BillTypePop	= iCurColumnPos(2)
			C_BillTypeNm	= iCurColumnPos(3)
			C_BillAmt		= iCurColumnPos(4)	
			C_BillLocAmt	= iCurColumnPos(5)
			C_BankCd		= iCurColumnPos(6)	
			C_BankPop		= iCurColumnPos(7)	
			C_BankNm		= iCurColumnPos(8)	
			C_BankAcct		= iCurColumnPos(9)	
			C_BankAcctPop	= iCurColumnPos(10)
			C_Note			= iCurColumnPos(11)		
			C_NotePop		= iCurColumnPos(12)
			C_PreReceipt	= iCurColumnPos(13)
			C_PreReceiptPop = iCurColumnPos(14)
			C_Remark		= iCurColumnPos(15)
			C_XchRate		= iCurColumnPos(16)
			C_XchCalop		= iCurColumnPos(17)
			C_BillSeq		= iCurColumnPos(18)

    End Select    
End Sub

'====================================================
Sub Form_Load()
	
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call SetDefaultVal
	Call InitVariables														
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitSpreadSheet
	Call InitCollectType

    Call SetToolbar("11000000000011")										
	Call CookiePage(0)

End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	If Row <= 0 Then Exit Sub

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		.Row = Row
		
		Select Case Col
			Case C_BillTypePop			'수금유형팝업 
				.Col = Col - 1
				Call OpenBillTypePop(.Text)
				
			Case C_BankPop				'은행팝업 
				.Col = Col - 1
				Call OpenBankPop(.Text)
				
			Case C_BankAcctPop			'은행계좌번호팝업 
				.Col = Col - 1
				Call OpenBankAcctPop(.Text)
				
			Case C_NotePop				'어음번호 
				.Col = Col - 1
				Call OpenNotePop(.Text)

			Case C_PreReceiptPop		'선수금번호 
				.Col = Col - 1
				Call OpenPreReceiptPop(.Text)		    
		End Select		    

		Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")

	End With

End Sub

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==========================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim iCnt, strRefVal
	
	If Row < 0 Then Exit Sub

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

	With frm1

		.vspdData.Row = Row

		Select Case Col
		Case C_BillTypeCd
			.vspdData.ReDraw = False

			Call chgProtectOfBillType(Row)
			.vspdData.ReDraw = True

		Case C_BillAmt
			' 원화금액 계산 
			Call ClientXchRateCalcu(Row)
			
		Case C_BillLocAmt
			' 원화금액 계산 
			Call BillTotalSum(C_BillLocAmt)

		Case C_PreReceipt
			CALL GetPreReceiptInfo(Row)			' 관련 환율 Fetch
			
		Case C_Note
			Call GetNoteInfo(Row)

		End Select

	End With

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then Exit Sub
		    
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End If
End Sub

'==========================================
Sub btnPostFlag_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Dim strVal

	frm1.txtInsrtUserId.value = parent.gUsrID 

			
		If   LayerShowHide(1) = False Then
             Exit Sub
        End If


	strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag									
	strVal = strVal & "&txtHBillNo=" & Trim(frm1.txtHBillNo.value)						
	strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)

	Call RunMyBizASP(MyBizASP, strVal)												
	
End Sub

'==========================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	
	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True
	
	If Trim(frm1.txtGLNo.value) <> "" Then
		arrParam(0) = Trim(frm1.txtGLNo.value)	'회계전표번호 
		
		if arrParam(0) = "" THEN Exit Sub
		
		arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
		arrParam(0) = Trim(frm1.txtTempGLNo.value)	'결의전표번호 
		
		if arrParam(0) = "" THEN Exit Sub
		arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else	
		Call DisplayMsgBox("205154", "X", "X", "X")
	End If	
	IsOpenPop = False
End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkField(Document, "1") Then Exit Function

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables

    Call DbQuery																

    FncQuery = True																
        
End Function

'========================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call SetToolbar("11000000000011")										
    Call SetDefaultVal
    Call InitVariables														

    FncNew = True																

End Function

'========================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
	    Exit Function
    End If

    If ggoSpread.SSDefaultCheck = False Then Exit Function

   	If UNICDbl(frm1.txtTotBillAmt.text) < UNICDbl(frm1.txtSumBillAmt.text) Then
		IntRetCD = DisplayMsgBox("205525", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    CAll DbSave
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncCopy = False  

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		
		.vspdData.ReDraw = True
	End With
	
	If Err.number = 0 Then	
       FncCopy = True                                                            
    End If

    Set gActiveElement = document.ActiveElement  
    
End Function

'========================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    
	If frm1.vspdData.MaxRows > 1 Then 
		Call chgProtectOfBillType(frm1.vspdData.ActiveRow)
		Call BillTotalSum(C_BillAmt)
	End If
End Function

'========================================
Function FncInsertRow(pvRowCnt) 

	Dim IntRetCD
    Dim imRow,i
    On Error Resume Next                                                         
    Err.Clear                                                                    
    
    FncInsertRow = False                                                         

	If IsNumeric(Trim(pvRowCnt)) then	
		imRow = CInt(pvRowCnt)
	Else		
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If
	End If
    
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData

		ggoSpread.InsertRow, imRow
		
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
		lgBlnFlgChgValue = True
		For i = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
		
			frm1.vspdData.Col = C_XchRate
			frm1.vspdData.Row = i	
			frm1.vspdData.Text = TRim(frm1.HXchRate.value)

			frm1.vspdData.Col = C_XchCalop
			frm1.vspdData.Row = i
			frm1.vspdData.Text = TRim(frm1.HXchRateOp.value)
		Next
				
		.vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
       FncInsertRow = True                             
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncDeleteRow() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
    
	lDelRows = ggoSpread.DeleteRow
	
    lgBlnFlgChgValue = True
    
	Call BillTotalSum(C_BillAmt)
    
    End With
    
End Function

'========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()	
	
	If frm1.txtHPostFlag.value = "Y" Then
		Call SetPostYesSpreadColor(1)
	Else
		Call SetQuerySpreadColor(1)
	End If
End Sub

'========================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================
Function DbQuery() 

    Err.Clear                                                               
    
    DbQuery = False                                                         
			
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If
    
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtHBillNo.value)					
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtConBillNo.value)				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If	

	Call RunMyBizASP(MyBizASP, strVal)												
	
    DbQuery = True																	

End Function

'========================================
Function DbQueryOk()														
	
    lgIntFlgMode = parent.OPMD_UMODE												
	lgBlnFlgChgValue = False
    lgIntGrpCount = 0														
  
	If frm1.txtHPostFlag.value = "N" Then
	    Call SetToolbar("11101111001111")
    Else
		Call SetToolbar ("11100000000111")
	End If
	
	If Trim(frm1.txtSts.value) <> "" Then
		If Cint(frm1.txtSts.value) < 3 Then
			frm1.btnPostFlag.disabled = False
		Else
			frm1.btnPostFlag.disabled = True
		End If
	End If

End Function

'========================================
Function DbSave()

    Err.Clear																
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel, strMsg
	Dim	ldbBillAmt, ldbBillLocAmt
	
	strMsg = "수금액"
	
	DbSave = False                                                          
    
    On Error Resume Next

	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							'☜: 신규 
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep'☜: C=Create
		        Case ggoSpread.UpdateFlag							'☜: 수정 
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep'☜: U=Update
				Case ggoSpread.DeleteFlag							'☜: 삭제 
					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep'☜: D=Delete
					'--- 수금순번 
		            .vspdData.Col = C_BillSeq 
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1 
			End Select

			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

					'--- 수금순번 
		            .vspdData.Col = C_BillSeq 
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 수금유형 
		            .vspdData.Col = C_BillTypeCd 		            
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 수금액 
		            .vspdData.Col = C_BillAmt
		            ldbBillAmt = UNICDbl(Trim(.vspdData.Text)) 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
					'--- 수금자국금액 
		            .vspdData.Col = C_BillLocAmt 		
		            ldbBillLocAmt = UNICDbl(Trim(.vspdData.Text)) 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
					'--- 은행 
		            .vspdData.Col = C_BankCd 		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 은행계좌번호 
					.vspdData.Col = C_BankAcct		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- 어음번호 
		            .vspdData.Col = C_Note		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- 선수금번호 
		            .vspdData.Col = C_PreReceipt		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            '--- 비고 
		            .vspdData.Col = C_Remark		
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					'--- 환율 
		            .vspdData.Col = C_XchRate
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep		            
					'--- 환율연산자 
		            .vspdData.Col = C_XchCalop
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1 
		    End Select       
		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
	End With
	
    DbSave = True                                                           
    
End Function

'========================================
Function DbSaveOk()

	Call InitVariables
	frm1.txtConBillNo.value = frm1.txtHBillNo.value
	Call ggoOper.ClearField(Document, "2")
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권수금내역</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>매출채권번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConBillNo" ALT="매출채권번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSBillDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBillDtl()"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>주문처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="주문처">&nbsp;<INPUT NAME="txtSoldtoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>예외매출채권여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoExceptBill" id="rdoExceptBillYes" VALUE="Y" tag = "24" CHECKED>
										<LABEL FOR="rdoExceptBillYes">예</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoExceptBill" id="rdoExceptBillNo" VALUE="N" tag = "24">
										<LABEL FOR="rdoExceptBillNo">아니오</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>매출채권일</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5114ma1_fpDateTime1_txtBillDt.js'></script>
								</TD>							
								<TD CLASS=TD5 NOWRAP>확정여부</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoPost" id="rdoPostNo" VALUE="N" tag = "24" CHECKED>
										<LABEL FOR="rdoPostNo">미확정</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoPost" id="rdoPostYes" VALUE="Y" tag = "24">
										<LABEL FOR="rdoPostYes">확정</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>매출채권형태</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" TYPE="Text" MAXLENGTH=20 SIZE=10 tag="24XXXU">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24"></TD>
							</TR>
							<TR>							
								<TD CLASS=TD5 NOWRAP>총매출채권금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5114ma1_fpDoubleSingle2_txtTotBillAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>총매출채권자국금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5114ma1_fpDoubleSingle2_txtTotBillAmtLoc.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtLocCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24">
											</TD>											
										</TR>
									</TABLE>
								</TD>															
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>총수금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5114ma1_fpDoubleSingle4_txtSumBillAmt.js'></script>
								</TD>									
								<TD CLASS=TD5 NOWRAP>총수금자국액</TD>							
									<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5114ma1_fpDoubleSingle3_txtSumLocBillAmt.js'></script>
											</TD>

										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s5114ma1_I882900816_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>border =1>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">확정</BUTTON>&nbsp;
						<BUTTON NAME="btnGLView" CLASS="CLSMBTN">전표조회</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "vbscript:JumpChgCheck(BIZ_BillHdr_JUMP_ID)">매출채권등록</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillDtl_JUMP_ID)">예외매출채권등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HXchRate" tag="24">
<INPUT TYPE=HIDDEN NAME="HXchRateOp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSts" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHPostFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
