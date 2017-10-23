<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : B/L���� 
'*  3. Program ID           : S5115MA1
'*  4. Program Name         : B/L ���ݳ������ 
'*  5. Program Desc         : B/L ���ݳ������ 
'*  6. Comproxy List        : PS7G151.dll, PS7G158.dll, PS7G115.dll
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/05/20
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "s5115mb1.asp"												'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_BillHdr_JUMP_ID = "s5211ma1"											'��: JUMP�� �����Ͻ� ���� ASP�� 

'��: Spread Sheet�� Column�� ��� 
Dim C_BillTypeCd
Dim C_BillTypePop
Dim C_BillTypeNm
Dim C_BillAmt
Dim C_BillLocAmt
Dim C_BankCd
Dim C_BankPop
Dim C_BankNm
Dim C_BankAcct
Dim C_BankAcctPop
Dim C_Note
Dim C_NotePop
Dim C_PreReceipt
Dim C_PreReceiptPop
Dim C_Remark
Dim C_XchRate
Dim C_XchCalop
Dim C_BillSeq

<!-- #Include file="../../inc/lgvariables.inc" -->	
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim arrCollectType		'�������� �迭 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop						' Popup
Const PostFlag = "PostFlag"

'========================================================================================================
Sub initSpreadPosVariables()  

	C_BillTypeCd	= 1			'�������� 
	C_BillTypePop	= 2			'���������˾� 
	C_BillTypeNm	= 3			'���������� 
	C_BillAmt		= 4			'���ݾ� 
	C_BillLocAmt	= 5			'�����ڱ��ݾ� 
	C_BankCd		= 6			'���� 
	C_BankPop		= 7			'�����˾� 
	C_BankNm		= 8			'����� 
	C_BankAcct		= 9			'������¹�ȣ 
	C_BankAcctPop	= 10		'������¹�ȣ�˾� 
	C_Note			= 11		'������ȣ 
	C_NotePop		= 12		'������ȣ�˾� 
	C_PreReceipt	= 13		'������ 
	C_PreReceiptPop = 14		'�������˾� 
	C_Remark		= 15		'��� 
	C_XchRate		= 16		'ȯ�� 
	C_XchCalop		= 17		'ȯ�������� 
	C_BillSeq		= 18		'���ݼ��� 
End Sub

'========================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtConBLNo.focus
	frm1.btnPostFlag.disabled = True
	frm1.btnPostFlag.value = "Ȯ��"
	frm1.rdoPostNo.checked = True
	frm1.btnGLView.disabled = True

	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>	
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	With frm1.vspdData
	    ggoSpread.Source = frm1.vspdData
                           'patch version
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread
		
		.ReDraw = False
		
	    .MaxRows = 0	: .MaxCols = 0
	    	
	    .MaxCols = C_BillSeq+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols														'��: ������Ʈ�� ��� Hidden Column
	    .ColHidden = True

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("6","3","0")
		ggoSpread.SSSetEdit C_BillTypeCd, "��������", 10,,,5,2
	    ggoSpread.SSSetButton C_BillTypePop
		ggoSpread.SSSetEdit C_BillTypeNm, "����������", 20
		ggoSpread.SSSetFloat C_BillAmt,"���ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_BillLocAmt,"�����ڱ���",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_BankCd, "����", 10,,,10,2
	    ggoSpread.SSSetButton C_BankPop
		ggoSpread.SSSetEdit C_BankNm, "�����", 20
		ggoSpread.SSSetEdit C_BankAcct, "������¹�ȣ", 18,,,30,2
		ggoSpread.SSSetButton C_BankAcctPop
		ggoSpread.SSSetEdit C_Note, "������ȣ", 18,,,30,2
		ggoSpread.SSSetButton C_NotePop
		ggoSpread.SSSetEdit C_PreReceipt, "�����ݹ�ȣ", 18,,,18,2
		ggoSpread.SSSetButton C_PreReceiptPop
		ggoSpread.SSSetEdit C_Remark, "���", 50,,,200,1	
		ggoSpread.SSSetFloat C_XchRate,"ȯ��",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_XchCalop, "ȯ��������", 15
		ggoSpread.SSSetFloat C_BillSeq,"���ݼ���" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"				
		
		call ggoSpread.MakePairsColumn(C_BillTypeCd,C_BillTypePop)
		call ggoSpread.MakePairsColumn(C_BankCd,C_BankPop)
		call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPop)
		call ggoSpread.MakePairsColumn(C_Note,C_NotePop)																
		call ggoSpread.MakePairsColumn(C_PreReceipt,C_PreReceiptPop)
		
		Call ggoSpread.SSSetColHidden(C_BillSeq,C_BillSeq,True)		
	    
		.ReDraw = True
   
    End With
    
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
  
    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.Source = frm1.vspdData

	ggoSpread.SSSetProtected	C_BillSeq, pvStartRow, pvEndRow    
	ggoSpread.SSSetRequired		C_BillTypeCd, pvStartRow, pvEndRow    
	ggoSpread.SSSetProtected	C_BillTypeNm, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_BillAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_BillLocAmt, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_BankCd, pvStartRow, pvEndRow
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
    .vspdData.ReDraw = True
    
    End With

End Sub

'=========================================================================================================
Sub InitCollectType()	
	Dim i
	Dim iCodeArr, iTypeArr

	Err.Clear

	Call CommonQueryRs(" MINOR.MINOR_CD, CONFIG.REFERENCE ", " B_MINOR MINOR, B_CONFIGURATION CONFIG ", " MINOR.MINOR_CD *= CONFIG.MINOR_CD AND MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND CONFIG.SEQ_NO = " & FilterVar("4", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iTypeArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then		
		MsgBox Err.description, vbInformation,Parent.gLogoName
		Err.Clear 
		Exit Sub
	 End If

	 Redim arrCollectType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectType(i, 0) = iCodeArr(i)
		arrCollectType(i, 1) = iTypeArr(i)
	Next
End Sub

'========================================================================================================
Function OpenConBillDtl()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	iCalledAspName = AskPRAspName("s5211pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5211pa1", "x")
		IsOpenPop = False
		exit Function
	end if

	IsOpenPop = True
		
	frm1.txtConBLNo.focus
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetConBillDtl(strRet)
		frm1.txtConBLNo.focus
	End If	

End Function

'========================================================================================================
Function OpenBillTypePop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"							
	arrParam(1) = "B_CONFIGURATION Config, B_MINOR Minor"	
	arrParam(2) = Trim(strCode)								
	arrParam(3) = ""										
	arrParam(4) = "Config.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND Config.SEQ_NO = " & FilterVar("1", "''", "S") & "  " _
				& "AND Config.MINOR_CD = Minor.MINOR_CD AND Config.MAJOR_CD = Minor.MAJOR_CD " _
				& "AND Config.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("R", "''", "S") & " )"		
	arrParam(5) = "��������"							

	arrField(0) = "Config.MINOR_CD"							
	arrField(1) = "Minor.MINOR_NM"							

	arrHeader(0) = "��������"							
	arrHeader(1) = "����������"							

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBillAmtDtl(arrRet, C_BillTypePop)
	End If	
	
End Function

'========================================================================================================
Function OpenBankPop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�Ա�����"				
	arrParam(1) = "B_BANK BK, F_DPST DP"
	arrParam(2) = Trim(strCode)							
	arrParam(3) = ""									
	arrParam(4) = "BK.BANK_CD=DP.BANK_CD" 	
	arrParam(5) = "�Ա�����"						

	arrField(0) = "BK.BANK_CD"				
	arrField(1) = "BK.BANK_NM"				

	arrHeader(0) = "�Ա�����"				
	arrHeader(1) = "�Ա������"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBillAmtDtl(arrRet, C_BankPop)
	End If	
	
End Function
'========================================================================================================
Function OpenBankAcctPop(strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		frm1.vspdData.Col = C_BankCd
		If Trim(frm1.vspdData.Text) = "" Then
			Call DisplayMsgBox("205152", "X", "����", "X")
			frm1.vspdData.Action = 0
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "������¹�ȣ"				
		arrParam(1) = "B_BANK BK, F_DPST DP"
		arrParam(2) = Trim(strCode)			
		arrParam(3) = ""					
		arrParam(4) = "BK.BANK_CD=DP.BANK_CD And BK.BANK_CD = " _
			& FilterVar(Trim(frm1.vspdData.Text), "" , "S")
		arrParam(5) = "������¹�ȣ"			

		arrField(0) = "DP.BANK_ACCT_NO"	
		arrField(1) = "BK.BANK_NM"				
		arrField(2) = "BK.BANK_CD"			

		arrHeader(0) = "������¹�ȣ"			
		arrHeader(1) = "�����"					
		arrHeader(2) = "����"					

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBillAmtDtl(arrRet, C_BankAcctPop)
	End If	
	
End Function
'========================================================================================================
Function OpenNotePop(ByVal pvStrCode)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	Dim iStrBpCd, iStrChargeDt, iStrChargeLocAmt, iStrVatLocAmt, iStrTotAmt

	On Error Resume Next

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	iStrBpCd = Trim(frm1.txtApplicant.value)
	iStrChargeDt = Trim(frm1.txtBLIssueDt.text)

	iArrParam(1) = "F_NOTE"								<%' TABLE ��Ī %>
	iArrParam(2) = Trim(pvStrCode)							<%' Code Condition%>
	iArrParam(3) = ""									<%' Name Cindition%>
	iArrParam(4) = "NOTE_FG IN (" & FilterVar("D1", "''", "S") & ", " & FilterVar("CR", "''", "S") & ") AND NOTE_STS = " & FilterVar("BG", "''", "S") & " AND BP_CD =  " & FilterVar(iStrBpCd , "''", "S") & "" _
		& " AND (Convert(CHAR(10), ISSUE_DT, 112) <= '" & UniConvDateToYYYYMMDD(iStrChargeDt, Parent.gDateFormat,"") & _
		"' And Convert(CHAR(10), DUE_DT, 112) >=  " & FilterVar(UniConvDateToYYYYMMDD(iStrChargeDt, Parent.gDateFormat,""), "''", "S") & ")" <%' Where Condition%>

	iArrParam(5) = "������ȣ"						<%' TextBox ��Ī %>

	iArrField(0) = "NOTE_NO"								<%' Field��(0)%>
	iArrField(1) = "HH" & Parent.gColSep & "NOTE_AMT"			<%' Field��(1) - Hidden%>
	iArrField(2) = "F2" & Parent.gColSep & "NOTE_AMT"			<%' Field��(2)%>				
	iArrField(3) = "NOTE_FG"								<%' Field��(3)%>
	iArrField(4) = "NOTE_STS"							<%' Field��(4)%>									
		
	iArrHeader(0) = "������ȣ"						<%' Header��(0)%>
	iArrHeader(1) = "�����ݾ�"						<%' Header��(1) - Hidden%>
	iArrHeader(2) = "�����ݾ�"						<%' Header��(2)%>
	iArrHeader(3) = "��������"						<%' Header��(3)%>
	iArrHeader(4) = "��������"						<%' Header��(4)%>

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		OpenNotePop = False
		Exit Function
	Else
		Call SetBillAmtDtl(iArrRet, C_NotePop)
		OpenNotePop = True
	End If	
			
End Function

'========================================================================================================
Function OpenPreReceiptPop(ByVal prStrCode)
	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(4)
	
	OpenPreReceiptPop = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	With frm1
		iArrParam(0) = Trim(.txtBLIssueDt.Text)			'������ 
		iArrParam(1) = Trim(.txtApplicant.value)			'������ 
		iArrParam(2) = Trim(.txtApplicantNm.value)		'�����ڸ� 
		iArrParam(3) = Trim(.txtCurrency.value)			'ȭ�� 
		.vspddata.col = C_PreReceipt
		iArrParam(4) = Trim(.vspddata.text)				'�����ݹ�ȣ 
	End With	

	iCalledAspName = AskPRAspName("s5111ra7")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5111ra7", "x")
		IsOpenPop = False
		exit Function
	end if
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtFlag=CL&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent, iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetBillAmtDtl(iArrRet, C_PreReceiptPop)
		OpenPreReceiptPop = True
	End If	
			
End Function

'========================================================================================================
Function SetConBillDtl(Byval arrRet)
	frm1.txtConBLNo.value = arrRet(0)
End Function

'========================================================================================================
Function SetBillAmtDtl(Byval arrRet,ByVal iWhere)

	With frm1

		Select Case iWhere
		Case C_BillTypePop	'�������� 
			.vspdData.Col = C_BillTypeCd	:	.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BillTypeNm	:	.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_BillTypeCd, .vspdData.Row)		<% ' ������ �о�ٰ� �˷��� %>

		Case C_BankPop		'���� 
			.vspdData.Col = C_BankCd		:	.vspdData.Text = arrRet(0)
			.vspdData.Col = C_BankNm		:	.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_BankCd, .vspdData.Row)
		
		Case C_BankAcctPop	'������¹�ȣ 
			.vspdData.Col = C_BankAcct		:	.vspdData.Text = arrRet(0)
			Call vspdData_Change(C_BankAcct, .vspdData.Row)
			
		Case C_NotePop	'������ȣ 
			.vspdData.Col = C_Note		:	.vspdData.Text = arrRet(0)
			<%'�����ڱ��ݾ�(������ Local Currency�� ���ؼ��� ��ϰ���)%>
			.vspdData.Col = C_BillAmt	:	.vspdData.Text = UNIConvNumPCToCompanyByCurrency(arrRet(1), Parent.gCurrency, Parent.ggAmtOfMoneyNo, "X" , "X")
			.vspdData.Col = C_BillLocAmt:	.vspdData.Text = UNIConvNumPCToCompanyByCurrency(arrRet(1), Parent.gCurrency, Parent.ggAmtOfMoneyNo, "X" , "X")
			Call BillTotalSum(C_BillAmt)

		Case C_PreReceiptPop	'�����ݹ�ȣ 
			.vspdData.Col = C_PreReceipt		:	.vspdData.Text = arrRet(1)
			.vspdData.Col = C_XchRate			:	.vspdData.Text = UNIFormatNumber(arrRet(8), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
'			��ȭ�ݾװ�� 
			Call ClientXchRateCalcu(.vspdData.Row)
		End Select

	End With

	lgBlnFlgChgValue = True
	
End Function

<% '======================================   GetNoteInfo()  =========================================
'	Description : ���������� Fetch�Ѵ�.
'==================================================================================================== %>
Function GetNoteInfo(IRow)
	Dim strSoldToParty, strNoteNO, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strNoteInfo
	
	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_Note						'ǰ���ڵ� 
		strNoteNo = .vspdData.text

		strSoldToParty = .txtApplicant.value		'�ֹ�ó 
		strValidDt = UniConvDateToYYYYMMDD(.txtBLIssueDt.Text, Parent.gDateFormat,"")
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
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(strNoteInfo(1), Parent.gCurrency, Parent.ggAmtOfMoneyNo, "X" , "X")
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(strNoteInfo(1), Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
		Call BillTotalSum(C_BillAmt)
		Exit Function
	Else
		If Err.number <> 0 Then
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		
		If Not OpenNotePop(strNoteNo) Then
			'����� ��� �Էµ� ������ clear�Ѵ�.
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

<% '======================================   GetPreReceiptInfo()  =========================================
'	Description : �����ݹ�ȣ�� ��ȿ�� �� ȯ���� Fetch�Ѵ�.
'==================================================================================================== %>
Function GetPreReceiptInfo(byVal prIntRow)
	Dim iStrPreReceiptNo
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrPreReceiptInfo
	
	With frm1
		.vspdData.Row = prIntRow
		.vspdData.col = C_PreReceipt						'������ ��ȣ 
		iStrPreReceiptNo = .vspdData.text
	End With

	If Trim(iStrPreReceiptNo) = "" Then
		With frm1
			.vspdData.col = C_XchRate
			.vspdData.Text = .HXchRate.value
		End With
		Call ClientXchRateCalcu(prIntRow)		' ��ȭ�ݾ� ���� 
		Exit Function
	End If
		
	iStrSelectList = " FP.xch_rate "
	iStrFromList  = " f_prrcpt FP INNER JOIN a_jnl_item AJ ON (FP.prrcpt_type = AJ.jnl_cd) "
	With frm1
		iStrWhereList = " FP.bp_cd =  " & FilterVar(.txtApplicant.value , "''", "S") & " AND FP.doc_cur =  " & FilterVar(.txtCurrency.value , "''", "S") & "" & _
					   " AND FP.prrcpt_dt < '" & UniConvDateAToB(.txtBLIssueDt.Text, Parent.gDateFormat,Parent.gAPDateFormat) & "'" & _
					   " AND FP.bal_amt > 0 AND FP.conf_fg = " & FilterVar("C", "''", "S") & "  AND AJ.jnl_type = " & FilterVar("PR", "''", "S") & " AND FP.prrcpt_no =  " & FilterVar(iStrPreReceiptNo , "''", "S") & ""
	End With

    Err.Clear
    
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrPreReceiptInfo = Split(iStrRs, Chr(11))
		frm1.vspdData.Col = C_XchRate
		frm1.vspdData.text = UNIFormatNumber(iArrPreReceiptInfo(1), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		Call ClientXchRateCalcu(prIntRow)		' ��ȭ�ݾ� ���� 
		Exit Function
	Else
		If Not OpenPreReceiptPop(iStrPreReceiptNo) Then
			'����� ��� �Էµ� ������ clear�Ѵ�.
			With frm1
				.vspdData.Col = C_PreReceipt
				.vspdData.Text = ""
				.vspdData.col = C_XchRate
				.vspdData.Text = .HXchRate.value
			End With
			Call ClientXchRateCalcu(prIntRow)		' ��ȭ�ݾ� ���� 
		End if
	End if
End Function

'===================================   ClientXchRateCalcu()  ========================================
'	Description : ��ȭ�ݾ� ��� 
'==================================================================================================== 
Sub ClientXchRateCalcu(ByVal Row)

	Dim ldbBillAmt, ldbXchgRate

	frm1.vspdData.Row = Row
		
	frm1.vspdData.Col = C_BillAmt	:	ldbBillAmt = UNICDbl(Trim(frm1.vspdData.Text))
	frm1.vspdData.Col = C_XchRate	:	ldbXchgRate = UNICDbl(Trim(frm1.vspdData.Text))

	frm1.vspdData.Col = C_XchCalop
	Select Case Trim(frm1.vspdData.Text)
	Case "+"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt + ldbXchgRate, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	Case "-"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt - ldbXchgRate, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	Case "*"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt * ldbXchgRate, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	Case "/"
		frm1.vspdData.Col = C_BillLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(ldbBillAmt / ldbXchgRate, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	End Select

	Call BillTotalSum(C_BillAmt)

End Sub

'===========================================================================
' Function Desc : Protect Change of BillType Value
'============================================================================
Function chgProtectOfBillType(ByVal Row)

	<%
					'���ù�ȣ	���� 
	'-----------------------------------
	'������	 DP		Edit:O		Edit:O
	'��������NR		Edit:O		Edit:X
	'������	 PR		Edit:O		Edit:X
	'����	 CS		Edit:X		Edit:X
	%>
	Dim iCnt
	Dim strRefVal
	
	With frm1


		ggoSpread.Source = frm1.vspdData
		.vspdData.Col = C_BillTypeCd	:	.vspdData.Row = Row
		
		strRefVal = GetCollectTypeRef(UCase(Trim(.vspdData.Text)))
		Select Case strRefVal
		Case "DP"	'������			
			ggoSpread.SpreadUnLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadUnLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row

			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_BankCd, Row, Row
				ggoSpread.SSSetRequired C_BankAcct, Row, Row
			End If
		Case "NO"	'�������� 
			if frm1.txtCurrency.value <> frm1.txtLocCur.value Then
				.vspdData.Text = ""

				.vspdData.Col = C_BillTypeNm
				.vspdData.Text = ""

				Call DisplayMsgBox("206154", "X", "X", "X")
				Exit Function
			End if
			
			ggoSpread.SpreadUnLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_Note, Row, Row
			Else
				ggoSpread.SSSetProtected C_NotePop, Row, Row
			End If
			
		Case "PR"	'������ 
			ggoSpread.SpreadUnLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row

			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(Parent.gSetupMod, "A") = "Y" Then
				ggoSpread.SSSetRequired C_PreReceipt, Row, Row
			Else
				ggoSpread.SSSetProtected C_PreReceiptPop, Row, Row
			End If			

		Case "CS"	'����,��ǥ 
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadLock	C_PreReceipt, Row, C_PreReceiptPop, Row		
			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row
		Case Else
			ggoSpread.SpreadUnLock	C_BillAmt, Row, C_BillLocAmt, Row
			ggoSpread.SpreadUnLock	C_BankCd, Row, C_BankPop, Row
			ggoSpread.SpreadUnLock	C_BankAcct, Row, C_BankAcctPop, Row
			ggoSpread.SpreadUnLock	C_Note, Row, C_NotePop, Row
			ggoSpread.SpreadUnLock	C_PreReceipt, Row, C_PreReceiptPop, Row
			ggoSpread.SSSetRequired C_BillAmt, Row, Row
			ggoSpread.SSSetRequired C_BillLocAmt, Row, Row

			If GetSetupMod(Parent.gSetupMod, "A") <> "Y" Then
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
			.vspdData.Col = C_XchRate		:	.vspdData.Text = TRim(.HXchRate.value)
		End Select

	End With
	
End Function


'===========================================================================
' Function Desc : �ݾ� ����� �ڵ� ��� �� 
'============================================================================
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
				frm1.vspdData.Col = C_BillAmt		'���ݱݾ� 
				BillAmt = UNICDbl(frm1.vspdData.Text)

				SumBillAmt = SumBillAmt + BillAmt

				frm1.vspdData.Col = C_BillLocAmt	'�����ڱ��ݾ� 
				BillLocAmt = UNICDbl(frm1.vspdData.Text)

				SumBillLocAmt = SumBillLocAmt + BillLocAmt

			Case C_BillLocAmt
				frm1.vspdData.Col = C_BillLocAmt	'�����ڱ��ݾ� 
				BillLocAmt = UNICDbl(frm1.vspdData.Text)

				SumBillLocAmt = SumBillLocAmt + BillLocAmt

			End Select

		End If
	Next

	frm1.txtSumBillAmt.Text = UNIConvNumPCToCompanyByCurrency(SumBillAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	frm1.txtSumLocBillAmt.Text = UNIConvNumPCToCompanyByCurrency(SumBillLocAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	
End Sub

'================================== =====================================================
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetPostYesSpreadColor(ByVal lRow)

    With frm1

		Call SetToolbar("11100000000111")
    
		.vspdData.ReDraw = False

		'ggoSpread.SpreadLock C_BillTypeCd, -1
    
		Dim GridCol
		For GridCol = 1 To .vspdData.MaxCols
			ggoSpread.SSSetProtected GridCol, 1, .vspdData.MaxRows
		Next
    
		.vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================
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

'========================================================================================
' Function Desc : Jump�� �ش� ���ǰ� Query
'========================================================================================
Function CookiePage(Byval Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	
	Dim strTemp, arrVal

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtHBLNo.value

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtConBLNo.value =  arrVal(0)

		WriteCookie CookieSplit , ""
		
		Call DbQuery()
			
	End If
	
End Function

'===========================================================================
' Function Desc : Jump�� ����Ÿ ���濩�� üũ 
'===========================================================================
Function JumpChgCheck(DesID)

	Dim IntRetCD

	'************ ��Ƽ�� ��� **************
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
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

'======================================== BtnSpreadCheck()  ========================================
'	Description : Before Button Click, Spread Check Function
'==================================================================================================== 
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD
	ggoSpread.Source = frm1.vspdData	

	'������ ������ ���� ���� ���� üũ��, YES�̸� �۾����࿩�� üũ ���Ѵ� 
	If ggoSpread.SSCheckChange = True Then
	IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")               '����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?
	If IntRetCD = vbNo Then Exit Function
	End If

	'������ ������ �۾����࿩�� üũ 
	If ggoSpread.SSCheckChange = False Then
	IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")                '�۾��� �����Ͻðڽ��ϱ�?
	If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function

'========================================================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================
'	Description : ���������� reference���� ��ȯ�ϴ� �Լ� 
'========================================================================================================
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

'========================================================================================================
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

'=========================================================================================================
Sub Form_Load()

	Call SetDefaultVal
	Call InitVariables														'��: Initializes local global variables
    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet
	Call InitCollectType

	'����/��ȸ/�Է� 
	'/����/����/����In
	'/����Out/���/���� 
	'/����/����/���� 
	'/�μ� 
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
	Call CookiePage(0)

	frm1.txtConBLNo.focus 
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	<% '----------  Coding part  -------------------------------------------------------------%>   
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 Then				
			Select Case Col
			Case C_BillTypePop							'���������˾� 
				.Col = Col - 1
				.Row = Row
				Call OpenBillTypePop(.Text)
			Case C_BankPop								'�����˾� 
				.Col = Col - 1
				.Row = Row
				Call OpenBankPop(.Text)
			Case C_BankAcctPop 							'������¹�ȣ�˾� 
				.Col = Col - 1
				.Row = Row
				Call OpenBankAcctPop(.Text)
			Case C_NotePop 								'������ȣ 
				.Col = Col - 1
				.Row = Row
				Call OpenNotePop(.Text)
			Case C_PreReceiptPop 						'�����ݹ�ȣ 
				.Col = Col - 1
				.Row = Row
				Call OpenPreReceiptPop(.Text)	
			End Select
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
		End If		    

	End With

End Sub

'========================================================================================================
	Sub vspdData_Click(ByVal Col, ByVal Row)
		
		If frm1.txtHPostFlag.value = "N" Then
			Call SetPopupMenuItemInf("1101111111")
		Else
			Call SetPopupMenuItemInf("0000111111")
		End If
		gMouseClickStatus = "SPC"  	
	    Set gActiveSpdSheet = frm1.vspdData	
	End Sub
	
'========================================================================================================
	Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	    ggoSpread.Source = frm1.vspdData
	    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

	End Sub
		
'========================================================================================================
	Sub vspdData_MouseDown(Button , Shift , x , y)

	   If Button = 2 And gMouseClickStatus = "SPC" Then
	      gMouseClickStatus = "SPCR"
	   End If
	End Sub    

'========================================================================================================
	Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
	    ggoSpread.Source = frm1.vspdData
	    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
	    Call GetSpreadColumnPos("A")
	End Sub

'========================================================================================================
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
			' ��ȭ�ݾ� ��� 
			Call ClientXchRateCalcu(Row)
		
		Case C_BillLocAmt
			' ��ȭ�ݾ� ��� 
			Call BillTotalSum(C_BillLocAmt)

		Case C_PreReceipt
			CALL GetPreReceiptInfo(Row)			' ���� ȯ�� Fetch
			
		Case C_Note
			Call GetNoteInfo(Row)

		End Select

	End With

End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
        
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
		If CheckRunningBizProcess Then	Exit Sub
			
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End if    
End Sub

'==========================================================================================
'   Event Desc : Ȯ�� ��ư�� Ŭ���� ��� �߻� 
'==========================================================================================
Sub btnPostFlag_OnClick()

	If BtnSpreadCheck = False Then Exit Sub

	Dim strVal

	frm1.txtInsrtUserId.value = Parent.gUsrID 

	If LayerShowHide(1) = False Then
		Exit Sub
	End If

	strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag									
	strVal = strVal & "&txtHBLNo=" & Trim(frm1.txtHBLNo.value)						
	strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)

	Call RunMyBizASP(MyBizASP, strVal)												
	
End Sub

'==========================================================================================
'   Event Desc : ȸ����ǥ ��ư�� Ŭ������ �� 
'==========================================================================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	
	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True
	
	If Trim(frm1.txtGLNo.value) <> "" Then
		arrParam(0) = Trim(frm1.txtGLNo.value)	'ȸ����ǥ��ȣ 
		
		if arrParam(0) = "" THEN Exit Sub
		
		arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
		arrParam(0) = Trim(frm1.txtTempGLNo.value)	'������ǥ��ȣ 
		
		if arrParam(0) = "" THEN Exit Sub
		arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else	
		Call DisplayMsgBox("205154", "X", "X", "X")
	End If	
	IsOpenPop = False
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'�Ѹ���ä�Ǳݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtBillAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'�� ���ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtSumBillAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub


'===================================== CurFormatNumSprSheet()  ======================================
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData
		'���ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_BillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub


'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                  
    
    Err.Clear                                                         

    '-----------------------
    'Check previous data area
    '----------------------- 
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	
    Call InitVariables						

    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery											

    FncQuery = True											
        
End Function

'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                  
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")     ' condition,contents
    Call ggoOper.LockField(Document, "N")      
    Call SetToolbar("11000000000011")			
    Call SetDefaultVal
    Call InitVariables							

	frm1.txtConBLNo.focus 
	Set gActiveElement = document.ActiveElement	

    FncNew = True								

End Function

'========================================================================================================
Function FncDelete() 
    
    Exit Function
    Err.Clear                                                               '��: Protect system from crashing    
    
    FncDelete = False													
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002", "X", "X", "X")
        'Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                  '��: Clear Condition,Contents Field
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                     
    
    Err.Clear                                                           
    
    frm1.txtConBLNo.focus
    
    '-----------------------
    'Precheck area
    '-----------------------
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		'Call MsgBox("No data changed!!", vbInformation)
	    Exit Function
    End If

    
    '-----------------------
    'Check content area
    '-----------------------
    If ggoSpread.SSDefaultCheck = False Then 
       Exit Function
    End If

   	If UNICDbl(frm1.txtBillAmt.text) < UNICDbl(frm1.txtSumBillAmt.text) Then
		IntRetCD = DisplayMsgBox("205525", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    CAll DbSave				          
    
    FncSave = True        
    
End Function

'========================================================================================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    '----------  Coding part  -------------------------------------------------------------
    
	With frm1
		.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		
		.vspdData.ReDraw = True
	End With
    
End Function

'========================================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
	Call BillTotalSum(C_BillAmt)
End Function

'========================================================================================================
Function FncInsertRow(pvRowCnt) 

	Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

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

		For imRow = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1

			frm1.vspdData.Col = C_XchRate
			frm1.vspdData.Row = imRow
			frm1.vspdData.Text = TRim(frm1.HXchRate.value)

			frm1.vspdData.Col = C_XchCalop
			frm1.vspdData.Row = imRow
			frm1.vspdData.Text = TRim(frm1.HXchRateOp.value)
		Next

		.vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If
    
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
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

'========================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)		
End Function

'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI, False) 
End Function

'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
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
'========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================================
Function DbDelete() 
    On Error Resume Next                                                 
End Function

'========================================================================================================
Function DbDeleteOk()							
    On Error Resume Next                        
End Function

'========================================================================================================
Function DbQuery() 

    Err.Clear             
    
    DbQuery = False       

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtConBLNo=" & Trim(frm1.txtHBLNo.value)	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001					
		strVal = strVal & "&txtConBLNo=" & Trim(frm1.txtConBLNo.value)			
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If	

	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbQuery = True														

End Function

'========================================================================================================
Function DbQueryOk(ByVal pvStrFlag)				
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If pvStrFlag = "H" Then
	    lgIntFlgMode = Parent.OPMD_UMODE				
		lgBlnFlgChgValue = False
	    lgIntGrpCount = 0						
	  
	'    Call ggoOper.LockField(Document, "Q")
		If frm1.txtHPostFlag.value = "N" Then
		    If frm1.txtHRefFlagNo.value = "M" Then
			's_bill_hdr.ref_flag ='M' �ϰ�� -- ��Ź�������ϰ�� ���ݳ�����Ͽ��� ������� ���ϵ��ϸ��� 
			    Call SetToolbar("11100000000111")		    
			Else
				Call SetToolbar("11101111001111")
			End if
		Else
			Call SetToolbar("11100000000111")
		End If
		
		If Trim(frm1.txtSts.value) <> "" Then
			If Cint(frm1.txtSts.value) < 3 Then
				If frm1.txtHRefFlagNo.value = "M" Then
					frm1.btnPostFlag.disabled = True
					Call DisplayMsgBox("205326", "X", "X", "X")					
				Else
					frm1.btnPostFlag.disabled = False
				End if		
			Else
				frm1.btnPostFlag.disabled = True
			End If
		End If
	Else
		If frm1.txtHPostFlag.value = "Y" Then
			Call SetPostYesSpreadColor(1)
		Else
			Call SetQuerySpreadColor(1)
		End If
	End If

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus		
    Else
       frm1.txtConBLNo.focus
    End If     

End Function

'========================================================================================================
Function DbSave()

    Err.Clear													
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim CollectedAmt	

    DbSave = False                                                          '��: Processing is NG
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							'��: �ű� 
					strVal = strVal & "C" & Parent.gColSep	& lRow & Parent.gColSep'��: C=Create
		        Case ggoSpread.UpdateFlag							'��: ���� 
					strVal = strVal & "U" & Parent.gColSep	& lRow & Parent.gColSep'��: U=Update
				Case ggoSpread.DeleteFlag							'��: ���� 
					strDel = strDel & "D" & Parent.gColSep	& lRow & Parent.gColSep'��: D=Delete
					'--- ���ݼ��� 
		            .vspdData.Col = C_BillSeq 
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1 
			End Select

			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

					'--- ���ݼ��� 
		            .vspdData.Col = C_BillSeq 
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- �������� 
		            .vspdData.Col = C_BillTypeCd 		            
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- ���ݾ� 
		            .vspdData.Col = C_BillAmt
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep		             		
					'--- �����ڱ��ݾ� 
		            .vspdData.Col = C_BillLocAmt 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					'--- ���� 
		            .vspdData.Col = C_BankCd 		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- ������¹�ȣ 
					.vspdData.Col = C_BankAcct		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            '--- ������ȣ 
		            .vspdData.Col = C_Note		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            '--- �����ݹ�ȣ 
		            .vspdData.Col = C_PreReceipt		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            '--- ��� 
		            .vspdData.Col = C_Remark		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- ȯ�� 
		            .vspdData.Col = C_XchRate
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep		            
					'--- ȯ�������� 
		            .vspdData.Col = C_XchCalop
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1 
		    End Select      
		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function
'========================================================================================================
Function DbSaveOk()						

	Call InitVariables
	frm1.txtConBLNo.value = frm1.txtHBLNo.value
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L���ݳ���</font></td>
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
									<TD CLASS="TD5" NOWRAP>B/L������ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConBLNo" ALT="B/L������ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSBillDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBillDtl()"></TD>
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
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicant" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="������">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=radio CLASS="RADIO" NAME="rdoPost" id="rdoPostNo" VALUE="N" tag = "24" CHECKED>
										<LABEL FOR="rdoPostNo">��Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoPost" id="rdoPostYes" VALUE="Y" tag = "24">
										<LABEL FOR="rdoPostYes">Ȯ��</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ä������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillTypeCd" TYPE="Text" MAXLENGTH=20 SIZE=10 tag="24XXXU">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>B/L��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" ALT="B/L��ȣ" TYPE="Text" MAXLENGTH="35" SIZE=30 tag="24XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5115ma1_fpDateTime1_txtBLIssueDT.js'></script>
								</TD>							
								<TD CLASS=TD5 NOWRAP>�����׷�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��B/L�ݾ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5115ma1_fpDoubleSingle2_txtBillAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS="TD5" NOWRAP>��B/L�ڱ��ݾ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5115ma1_fpDoubleSingle3_txtLocBillAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtLocCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ѽ��ݾ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5115ma1_fpDoubleSingle4_txtSumBillAmt.js'></script>
								</TD>
								<TD CLASS=TD5 NOWRAP>�Ѽ����ڱ���</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5115ma1_fpDoubleSingle5_txtSumLocBillAmt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s5115ma1_I912761259_vspdData.js'></script>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
						<BUTTON NAME="btnGLView" CLASS="CLSMBTN">��ǥ��ȸ</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "vbscript:JumpChgCheck(BIZ_BillHdr_JUMP_ID)">B/L���</a></TD>
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
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHBLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HXchRate" tag="24">
<INPUT TYPE=HIDDEN NAME="HXchRateOp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSts" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHRefFlagNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
