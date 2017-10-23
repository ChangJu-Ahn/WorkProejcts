  ' 1.1 Constant variables For spreadsheet
Dim C_ItemCd
Dim C_ItemNm
Dim C_Qty
Dim C_Unit
Dim C_Price
Dim C_SupplyAmt
Dim C_VatType
Dim C_VatNm
Dim C_VatRate
Dim C_VATAmt
Dim C_TotalAmt
Dim C_SupplyLocAmt
Dim C_VATLocAmt
Dim C_TotalLocAmt
Dim C_Seq
Dim C_BillNo
Dim C_BillSeq
Dim C_Spec
Dim C_XchCalop
Dim C_XchRate
Dim C_VatIncflag

' 1.2.1 Global ��� ���� 
Const BIZ_PGM_ID = "s5312mb1.asp"												
Const BIZ_BillTax_JUMP_ID = "s5311ma1"											

Dim IsOpenPop						' Popup
Const PostFlag = "PostFlag"

'========================================
Sub initSpreadPosVariables()  
    C_ItemCd		= 1		'ǰ�� 
    C_ItemNm		= 2		'ǰ��� 
    C_Qty			= 3		'���� 
    C_Unit			= 4		'���� 
    C_Price			= 5		'�ܰ� 
    C_SupplyAmt		= 6		'���ް��ݾ� 
    C_VatType		= 7		'vat
    C_VatNm			= 8		'vat�� 
    C_VatRate		= 9		'vat�� 
    C_VATAmt		= 10	'VAT�ݾ� 
    C_TotalAmt		= 11	'�հ�ݾ� 
    C_SupplyLocAmt	= 12	'���ް��ڱ��ݾ� 
    C_VATLocAmt		= 13	'VAT�ڱ��ݾ� 
    C_TotalLocAmt	= 14	'�հ��ڱ��ݾ� 
    C_Seq			= 15	'���� 
    C_BillNo		= 16	'����ä�ǹ�ȣ 
    C_BillSeq		= 17	'����ä�Ǽ��� 
    C_Spec			= 18	'�԰�	
    C_XchCalop		= 19	'ȯ�������� 
    C_XchRate		= 20	'ȯ�� 
    C_VatIncflag	= 21	'�ΰ������Կ��� 
End Sub
    
'========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE       
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           

    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtTaxBillNo.focus
	Set gActiveElement = document.activeElement 
	
	frm1.btnPostFlag.disabled = True
	frm1.btnPostFlag.value = "����"
	frm1.rdoVatCalcType2.checked = True
	lgBlnFlgChgValue = False
End Sub

'========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20021120",,parent.gAllowDragDropSpread    		
		.ReDraw = False

	    .MaxRows = 0	: .MaxCols = 0
	    .MaxCols = C_VatIncFlag + 1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    
        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_ItemCd, "ǰ��", 18,,,18,2
		ggoSpread.SSSetEdit C_ItemNm, "ǰ���", 30
		ggoSpread.SSSetFloat C_Qty,"����" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_Unit, "����", 8,,,3,2
		ggoSpread.SSSetFloat C_Price,"�ܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_SupplyAmt,"���ް���",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		
		'�߰� 
		ggoSpread.SSSetEdit 	C_VatType, "VAT����", 10,,,4,2
		ggoSpread.SSSetEdit 	C_VatNm, "VAT������", 20 
		ggoSpread.SSSetFloat	C_VatRate,"VAT��",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit 	C_VatIncFlag, "VAT���Ա���", 1 
		
		ggoSpread.SSSetFloat C_VatAmt,"VAT�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_TotalAmt,"�հ�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_SupplyLocAmt,"���ް��ڱ���",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_VatLocAmt,"VAT�ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_TotalLocAmt,"�հ��ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		Call AppendNumberPlace("6","4","0")
		ggoSpread.SSSetFloat C_Seq,"����" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"		
		ggoSpread.SSSetEdit C_BillNo, "����ä�ǹ�ȣ", 18,,,18,2
		ggoSpread.SSSetFloat C_BillSeq,"����ä�Ǽ���" ,18,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit C_Spec, "ǰ��԰�", 30
		ggoSpread.SSSetEdit C_XchCalop, "ȯ��������", 15
		ggoSpread.SSSetFloat C_XchRate,"ȯ��",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"

		ggoSpread.SpreadLockWithOddEvenRowColor()
		
		Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
		Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,True)
		Call ggoSpread.SSSetColHidden(C_VatRate,C_VatRate,True)
		Call ggoSpread.SSSetColHidden(C_VatIncflag,C_VatIncflag,True)
		Call ggoSpread.SSSetColHidden(C_XchCalop,C_XchCalop,True)
		Call ggoSpread.SSSetColHidden(C_XchRate,C_XchRate,True)
		Call ggoSpread.SSSetColHidden(C_Seq,C_Seq,True)			

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    	    
		.ReDraw = True
   
    End With
    
End Sub

'========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Qty				= iCurColumnPos(3)
			C_Unit				= iCurColumnPos(4)
			C_Price				= iCurColumnPos(5)
			C_SupplyAmt			= iCurColumnPos(6)
			C_VatType			= iCurColumnPos(7)
			C_VatNm				= iCurColumnPos(8)
			C_VatRate			= iCurColumnPos(9)
			C_VATAmt			= iCurColumnPos(10)
			C_TotalAmt			= iCurColumnPos(11)
			C_SupplyLocAmt		= iCurColumnPos(12)
			C_VATLocAmt			= iCurColumnPos(13)
			C_TotalLocAmt		= iCurColumnPos(14)
			C_Seq				= iCurColumnPos(15)
			C_BillNo			= iCurColumnPos(16)
			C_BillSeq			= iCurColumnPos(17)
			C_Spec				= iCurColumnPos(18)
			C_XchCalop			= iCurColumnPos(19)
			C_XchRate			= iCurColumnPos(20)			
			C_VatIncflag		= iCurColumnPos(21)
			
	End Select

End Sub	

'=========================================
Function OpenTaxBillNo()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	IsOpenPop = True
	    
	iCalledAspName = AskPRAspName("s5311pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5311pa1", "x")
		IsOpenPop = False
		exit Function
	end if

    strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtTaxBillNo.focus

	If strRet <> "" Then
		frm1.txtTaxBillNo.value = strRet
	End If	

End Function

'=========================================
Function OpenBillDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim strParam

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	With frm1
		If Trim(.txtBilltoParty.value) = "" Then
			Call DisplayMsgBox("900002", "X", "X", "X")
			Exit Function
		End If

		If Trim(.HPostFlag.value) = "Y" Then
			Msgbox "�̹� ����� ���ݰ�꼭������ȣ�� ������ ���� �� �� �����ϴ�",vbInformation, Parent.gLogoName
			Exit Function
		End If

		strParam = ""
		strParam = strParam & Trim(.txtBilltoParty.value) & Parent.gColSep
		strParam = strParam & Trim(.txtBilltoPartyNm.value) & Parent.gColSep
		strParam = strParam & Trim(.txtCurrency.value) & Parent.gColSep
		strParam = strParam & Trim(.txtVatType.Value) & Parent.gColSep
		strParam = strParam & Trim(.txtVatTypeNm.Value) & Parent.gColSep
		strParam = strParam & Trim(.HSalesGrpCd.Value) & Parent.gColSep
		strParam = strParam & Trim(.HSalesGrpNm.Value) & Parent.gColSep
		strParam = strParam & Trim(.txtBillNo.Value) & Parent.gColSep

		if .rdoVatCalcType1.checked = True then 
			strParam = strParam & "%" & Parent.gColSep
		elseif .rdoVatCalcType2.checked = True then 
			if .rdoVatIncFlag1.checked = True then 
				strParam = strParam & "1" & Parent.gColSep
			elseif .rdoVatCalcType2.checked = True then 
				strParam = strParam & "2" & Parent.gColSep
			end if
		end if

		strParam = strParam & Trim(.txtIssueDt.Value) & Parent.gRowSep
	End With

	iCalledAspName = AskPRAspName("s5112ra3")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5112ra3", "x")
		exit Function
	end if
	
	arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,strParam), _
			"dialogWidth=780px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If Err.number <> 0 Then Err.Clear 
	
    If arrRet(0, 0) <> "" Then
		Call SetBillDtlRef(arrRet)
	End If	

End Function

'=========================================
Function SetBillDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, i
	Dim intLoopCnt
	Dim intCnt
	Dim blnEqualFlg
	Dim strBillNo,strBillSeqNo
	Dim intCntRow

	Dim strBillJungBokMsg

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False	

		TempRow = .vspdData.MaxRows											'��: ��������� MaxRows
		intLoopCnt = Ubound(arrRet, 1)										'��: Reference Popup���� ���õǾ��� Row��ŭ �߰� 
		intCntRow = 0

		strBillJungBokMsg = ""

		For intCnt = 1 to intLoopCnt	
			blnEqualFlg = False

			If TempRow <> 0 Then
			
				strBillNo = ""	:	strBillSeqNo = ""
				' ���⳻�������� ���� ����ä�ǹ�ȣ�� ����ä�Ǽ����� �ִ��� üũ�Ѵ� 
				For i = 1 To TempRow
		
					.vspdData.Row = i
					.vspdData.Col = C_BillNo					 
					strBillNo = .vspdData.text
					
					.vspdData.Col = C_BillSeq
					strBillSeqNo = .vspdData.text
					
					If strBillNo = arrRet(intCnt - 1, 11) And strBillSeqNo = arrRet(intCnt - 1, 12) Then
						blnEqualFlg = True
						strBillJungBokMsg = strBillJungBokMsg & Chr(13) & strBillNo & "-" & strBillSeqNo
						Exit For
					End If

				Next

			End If
						
			If blnEqualFlg = false then
				intCntRow = intCntRow + 1
				.vspdData.MaxRows = CLng(TempRow) + CLng(intCntRow)
				.vspdData.Row = CLng(TempRow) + CLng(intCntRow)

				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				.vspdData.Col = C_ItemCd
				.vspdData.text = arrRet(intCnt - 1, 1)
				.vspdData.Col = C_ItemNm										
				.vspdData.text = arrRet(intCnt - 1, 2)
				.vspdData.Col = C_Qty			
				.vspdData.text = arrRet(intCnt - 1, 3)
				.vspdData.Col = C_Unit			
				.vspdData.text = arrRet(intCnt - 1, 4)
				.vspdData.Col = C_Price
				.vspdData.text = arrRet(intCnt - 1, 5)
				.vspdData.Col = C_SupplyAmt	
				.vspdData.text = arrRet(intCnt - 1, 6)
				.vspdData.Col = C_SupplyLocAmt
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 7), Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
				.vspdData.Col = C_VATAmt
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 9), .txtCurrency.value, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo , "X")
				.vspdData.Col = C_VATLocAmt
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 10), Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo , "X")
				.vspdData.Col = C_BillNo
				.vspdData.text = arrRet(intCnt - 1, 11)
				.vspdData.Col = C_BillSeq
				.vspdData.text = arrRet(intCnt - 1, 12)
				.vspdData.Col = C_XchCalop	
				.vspdData.text = arrRet(intCnt - 1, 13)
				.vspdData.Col = C_XchRate		
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 14), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
				.vspdData.Col = C_VatType
				.vspdData.text = arrRet(intCnt - 1, 15)
				.vspdData.Col = C_VatRate
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 16), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 17)

				.vspdData.Col = C_Spec
				.vspdData.text = arrRet(intCnt - 1, 18)
				
				'�ش� row�� �����Ѿ��� �����Ѵ�.
				BillDtlSum CLng(TempRow) + CLng(intCntRow)
			End if
		Next

		.vspdData.ReDraw = True

		' Head�� ���ް���, �ΰ��������� ����Ѵ�.
		Call BillHdrSum()
		Call JungBokMsg(strBillJungBokMsg,"����ä�ǹ�ȣ" & "-" & "����ä�Ǽ���")

	End With
		
	lgBlnFlgChgValue = True
End Function

'====================================================
Function JungBokMsg(strJungBok,strID)

	Dim strJugBokMsg
     
	If Len(Trim(strJungBok)) Then strJungBok = strID & Chr(13) & String(35,"=") & strJungBok
	If Len(Trim(strJungBok)) Then strJugBokMsg = strJungBok & Chr(13) & Chr(13)

	If Len(Trim(strJugBokMsg)) Then
		strJugBokMsg = strJugBokMsg & "�̹� ������ ��ȣ�� ������ �����մϴ�"
		MsgBox strJugBokMsg, vbInformation, Parent.gLogoName
	End If

End Function

'====================================================
Sub BillHdrSum()

	Dim SumSupplyAmt, SupplyAmt, SumSupplyLocAmt, SupplyLocAmt
	Dim SumVatAmt, VatAmt, SumLocVatAmt, LocVatAmt
	Dim lRow

	SumSupplyAmt = 0
	SumSupplyLocAmt = 0
	SumVatAmt = 0
	SumLocVatAmt = 0

	ggoSpread.source = frm1.vspdData
	For lRow = 1 To frm1.vspdData.MaxRows 
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then

			frm1.vspdData.Col = C_SupplyAmt		:		SupplyAmt = UNICDbl(frm1.vspdData.Text)
			frm1.vspdData.Col = C_SupplyLocAmt	:		SupplyLocAmt = UNICDbl(frm1.vspdData.Text)
			frm1.vspdData.Col = C_VATAmt		:		VatAmt = UNICDbl(frm1.vspdData.Text)
			frm1.vspdData.Col = C_VATLocAmt		:		LocVatAmt = UNICDbl(frm1.vspdData.Text)

'			frm1.vspdData.Col = C_VATIncFlag
			'�ΰ��� ������� ������ ��� ���ް��׿� �ΰ������� ���ԵǾ� ����.
'			If frm1.vspdData.Text = 1 Then
				SumSupplyAmt = SumSupplyAmt + SupplyAmt
				SumSupplyLocAmt = SumSupplyLocAmt + SupplyLocAmt
'			Else
'				SumSupplyAmt = SumSupplyAmt + SupplyAmt - VatAmt
'				SumSupplyLocAmt = SumSupplyLocAmt + SupplyLocAmt - LocVatAmt
'			End If

			SumVatAmt = SumVatAmt + VatAmt
			SumLocVatAmt = SumLocVatAmt + LocVatAmt

		End If
	Next
	
	frm1.txtSupplyAmt.Text		= UNIConvNumPCToCompanyByCurrency(SumSupplyAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	frm1.txtSupplyLocAmt.Text	= UNIConvNumPCToCompanyByCurrency(SumSupplyLocAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
	frm1.txtVatAmt.Text			= UNIConvNumPCToCompanyByCurrency(SumVatAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo , "X")
	frm1.txtLocVatAmt.Text		= UNIConvNumPCToCompanyByCurrency(SumLocVatAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo , "X")
End Sub

'====================================================
Sub BillDtlSum(ByVal GRow)

	Dim SupplyAmt, SupplyLocAmt, VATAmt, VATLocAmt

	SupplyAmt = 0
	SupplyLocAmt = 0
	VATAmt = 0
	VATLocAmt = 0

	ggoSpread.source = frm1.vspdData
	frm1.vspdData.Row = GRow
	frm1.vspdData.Col = 0
	If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
		frm1.vspdData.Col = C_VATAmt		:	VATAmt = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = C_VATLocAmt		:	VATLocAmt = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = C_SupplyAmt		:	SupplyAmt = UNICDbl(frm1.vspdData.Text)
		frm1.vspdData.Col = C_SupplyLocAmt	:	SupplyLocAmt = UNICDbl(frm1.vspdData.Text)

		' �ΰ������Կ��� Check
		frm1.vspdData.Col = C_VatIncFlag
		If frm1.vspdData.Text = "1" Then
			frm1.vspdData.Col = C_TotalAmt		:	frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(VATAmt + SupplyAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
			frm1.vspdData.Col = C_TotalLocAmt	:	frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(VATLocAmt + SupplyLocAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
		Else
			'�ΰ��������� ��쿡�� ���ް��׿� �ΰ����� �����ϰ� ����.
			frm1.vspdData.Col = C_TotalAmt		:	frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(SupplyAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
			frm1.vspdData.Col = C_TotalLocAmt	:	frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(SupplyLocAmt, Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
		End if
	End If

End Sub

'====================================================
Function CookiePage(Byval Kubun)

	On Error Resume Next
	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.HTaxBillNo.value
	ElseIf Kubun = 0 Then
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		frm1.txtTaxBillNo.value =  arrVal(0)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		Call MainQuery()
			
	End If
	
End Function

'====================================================
Function JumpChgCheck(DesID)

	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)

	Call PgmJump(BIZ_BillTax_JUMP_ID)

End Function

'====================================================
Function BtnSpreadCheck()

	BtnSpreadCheck = False

	Dim IntRetCD
	ggoSpread.Source = frm1.vspdData	

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If

	BtnSpreadCheck = True

End Function

'====================================================
Sub CurFormatNumericOCX()

	With frm1
		'���ް��� 
		ggoOper.FormatFieldByObjectOfCur .txtSupplyAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'VAT�ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub

'====================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'�ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'���ް��ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_SupplyAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'VAT�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_VATAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'�հ�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_TotalAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec

	End With

End Sub

Sub LockFieldInit()
    Call FormatDoubleSingleField(frm1.txtVatRate)
    Call LockObjectField(frm1.txtVatRate,"P")

    Call FormatDoubleSingleField(frm1.txtSupplyAmt)
    Call LockObjectField(frm1.txtSupplyAmt,"P")

    Call FormatDoubleSingleField(frm1.txtSupplyLocAmt)
    Call LockObjectField(frm1.txtSupplyLocAmt ,"P")

    Call FormatDoubleSingleField(frm1.txtVatAmt)
    Call LockObjectField(frm1.txtVatAmt ,"P")

    Call FormatDoubleSingleField(frm1.txtLocVatAmt)
    Call LockObjectField(frm1.txtLocVatAmt ,"P")

End Sub

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call SetDefaultVal
	Call InitVariables														
'	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
'   Call ggoOper.LockField(Document, "N")                                   
	Call LockFieldInit
	Call InitSpreadSheet
    Call SetToolbar("11000000000011")										
	Call CookiePage(0)

	Call LockHTMLField(frm1.rdoVatIncFlag1, "P")	
	Call LockHTMLField(frm1.rdoVatIncFlag2, "P")	
	Call LockHTMLField(frm1.rdoVatCalcType1, "P")	
	Call LockHTMLField(frm1.rdoVatCalcType2, "P")	

End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	Call SetPopupMenuItemInf("1101111111")
	
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If	

	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If
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
	
	If frm1.HPostFlag.value = "N" Then
		If frm1.vspdData.MaxRows < 1 Or UNICDbl(frm1.txtSupplyAmt.Text) = 0 Then
			Msgbox "���ް��ݾ��� 0 �Դϴ�" & vbcrlf & "ǰ���� �߰��ϼ���",vbInformation, Parent.gLogoName
			Exit Sub
		End If
	End If
	
	If BtnSpreadCheck = False Then Exit Sub

	Dim strVal

	frm1.txtInsrtUserId.value = Parent.gUsrID 
			
	If LayerShowHide(1) = False Then  Exit Sub

	strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag									
	strVal = strVal & "&HTaxBillNo=" & Trim(frm1.HTaxBillNo.value)						
	strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)

	Call RunMyBizASP(MyBizASP, strVal)												
	
End Sub

'========================================
Function FncQuery() 
    Dim IntRetCD 
    
    Err.Clear                                                               

    FncQuery = False                                                        
    
    If Not chkFieldByCell(frm1.txtTaxBillNo, "A", 1) Then Exit Function 

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
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
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")
'    Call ggoOper.LockField(Document, "N")
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

    CAll DbSave
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
	Call BillHdrSum()
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
    
	Call BillHdrSum()
    
    End With
    
End Function

'========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)									
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                                         
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
End Sub

'========================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
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
			
	If LayerShowHide(1) = False Then
	    Exit Function 
    End If
    
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									
		strVal = strVal & "&txtTaxBillNo=" & Trim(frm1.HTaxBillNo.value)				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtHQuery=F"
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									
		strVal = strVal & "&txtTaxBillNo=" & Trim(frm1.txtTaxBillNo.value)				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtHQuery=T"
	End If	

	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	
	Call RunMyBizASP(MyBizASP, strVal)												
	
    DbQuery = True																	

End Function

'========================================
Function DbHdrQueryOk()														
	
    lgIntFlgMode = Parent.OPMD_UMODE												
	lgBlnFlgChgValue = False
    lgIntGrpCount = 0														

	With frm1
		If .HPostFlag.value = "Y" Then
			.btnPostFlag.value = "�������"
			Call SetToolbar("11100000000111")
		Else
			.btnPostFlag.value = "����"
		    Call SetToolbar("1110101100011")
		End If
	End With

End Function

'========================================
Function DbQueryOk()														
	If frm1.vspdData.MaxRows > 0 Then 
		frm1.btnPostFlag.disabled = False
		frm1.vspdData.Focus		
	Else
       frm1.txtTaxBillNo.focus
    End If     

End Function

'========================================
Function DbSave()

    Err.Clear																
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

	Dim iVat,iVat_Loc																'�ΰ����� 
	
    DbSave = False                                                          

    On Error Resume Next                                                   

			
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If


	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
    
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
					'--- ���� 
		            .vspdData.Col = C_Seq 
		            strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1 
			End Select

			Select Case .vspdData.Text
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

					'�ΰ����� ���� 
					.vspdData.Col = C_VatAmt 
					iVat = Trim(.vspdData.Text)
					'Local�ΰ����� ���� 
					.vspdData.Col = C_VatLocAmt
					iVat_Loc = Trim(.vspdData.Text)


					'--- ǰ�� 
		            .vspdData.Col = C_ItemCd
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- ���� 
		            .vspdData.Col = C_Unit
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- �ܰ� 
		            .vspdData.Col = C_Price
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

					'--- ���ް��ݾ� 
'		            .vspdData.Col = C_SupplyAmt
'		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

					'VAT ���Կ��ο� ���� ���ް��ݾװ�� 
'		            .vspdData.Col = C_VatIncFlag
'		            If Trim(.vspdData.Text)  = "1" Then
						'--- �ݾ� 
						.vspdData.Col = C_SupplyAmt 		
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
'		            Else 
'						.vspdData.Col = C_SupplyAmt 
'				        strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) - UNIConvNum(iVat,0) & Parent.gColSep
'					End If


					'--- ���ް��ڱ��ݾ� 
'		            .vspdData.Col = C_SupplyLocAmt
'		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep

					'VAT ���Կ��ο� ���� ���ް��ڱ��ݾװ�� 
'		            .vspdData.Col = C_VatIncFlag
'		            If Trim(.vspdData.Text)  = "1" Then
						'--- �ݾ� 
						.vspdData.Col = C_SupplyLocAmt 		
						strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
'		            Else 
'						.vspdData.Col = C_SupplyLocAmt 
'				        strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) - UNIConvNum(iVat_Loc,0) & Parent.gColSep
'					End If

					'--- ����ä�ǹ�ȣ 
		            .vspdData.Col = C_BillNo
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'--- ����ä�Ǽ��� 
		            .vspdData.Col = C_BillSeq
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
					'--- ���� 
		            .vspdData.Col = C_Seq 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
		           '--- ���� 
		            .vspdData.Col = C_Qty
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
					
					'�߰� 
					'---vatŸ�� 
					.vspdData.Col = C_VatType 		
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					'---vat�� 
					.vspdData.Col = C_VatRate 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
					
					'--- VAT�ݾ� 
		            .vspdData.Col = C_VATAmt
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
					'--- VAT�ڱ��ݾ� 
		            .vspdData.Col = C_VATLocAmt
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & Parent.gColSep
					'--- �ΰ������Կ��� 
		            .vspdData.Col = C_VATIncflag
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		            
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
	frm1.txtTaxBillNo.value = frm1.HTaxBillNo.value
	Call ggoOper.ClearField(Document, "2")
    Call MainQuery()

End Function

'========================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

