  ' External ASP File
'========================================
Const BIZ_PGM_ID = "s5112mb2.asp"            
Const BIZ_BillHdr_JUMP_ID = "s5111ma2"           
Const BIZ_BillCollect_JUMP_ID = "s5114ma1"

' Constant variables defined
'========================================
Const PostFlag = "PostFlag"

' Common variables 
'========================================= %>
Dim lgBlnFlgChgValue    ' Variable is for Dirty flag
Dim lgIntGrpCount       ' Group View Size�� ������ ���� 
Dim lgIntFlgMode        ' Variable is for Operation Status
Dim lgArrVATTypeInfo	' VAT Type ����(VAT_TYPE, VAT_TYPE_NAME, VAT_RATE)
Dim lgBlnOpenPop

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortkey


' Variables For spreadsheet
'========================================
'��: Spread Sheet�� Column
Dim C_PlantCd			'���� 
Dim C_PlantPopup		'�����˾� 
Dim C_ItemCd			'ǰ�� 
Dim C_ItemPopup			'ǰ���˾�  
Dim C_ItemNm			'ǰ��� 
Dim C_BillQty			'���� 
Dim C_BillUnit			'���� 
Dim C_UnitPopup			'�����˾�  
Dim C_VatIncFlagNm		'�ΰ������Կ���  
Dim C_VatIncFlag		'�ΰ������Կ����ڵ�  
Dim C_BillPrice			'�ܰ� 
Dim C_BillAmt			'�ݾ� 
Dim C_VatType			'VatType 
Dim C_VatPopup			'Vat Popup 
Dim C_VatTypeNm			'Vat�� 
Dim C_VatRate			'Vat�� 
Dim C_VatAmt			'VAT�ݾ� 
Dim C_BillLocAmt		'��ȭ�ݾ� 
Dim C_VatLocAmt			'VAT��ȭ�ݾ� 
Dim C_DepositPrice		'�����ݴܰ� 
Dim C_DepositAmt		'�����ݱݾ� 
Dim C_Remark			'��� 
Dim C_ItemSpec			' ǰ��԰� 
Dim C_BillSeq			'����ä�Ǽ��� 
Dim C_RetItemFlag		'��ǰ����   
Dim C_PreBillNo			'��������ä�ǹ�ȣ 
Dim C_PreBillSeq		'��������ä�Ǽ��� 
Dim C_OldBillAmt
Dim C_OldVatIncFlag
Dim C_OldVatAmt
Dim C_InitialBillAmt
Dim C_InitialVatIncFlag
Dim C_InitialVatAmt
Dim C_TrackNo		   'Tracking No (������ �߰�)
Dim C_TrackingNoPopup	   'Tracking No (������ �߰�)
Dim C_TrackingFlg

' User-defind Variables
'========================================
Dim lgStrDepositflag   ' ������ ���� ���� 
Dim lgStrVatFlag    ' ǰ�� vat���� ���� ����(Header�� vat������ ��ϵǾ� �ִ� ��� 'N')

'========================================
Sub initSpreadPosVariables()  
	
	C_PlantCd			= 1
	C_PlantPopup		= 2
	C_ItemCd			= 3
	C_ItemPopup			= 4
	C_ItemNm			= 5
	C_BillQty			= 6
	C_BillUnit			= 7
	C_UnitPopup			= 8
	C_VatIncFlagNm		= 9
	C_VatIncFlag		= 10
	C_BillPrice			= 11
	C_BillAmt			= 12
	C_VatType			= 13
	C_VatPopup			= 14
	C_VatTypeNm			= 15
	C_VatRate			= 16
	C_VatAmt			= 17
	C_BillLocAmt		= 18
	C_VatLocAmt			= 19
	C_DepositPrice		= 20
	C_DepositAmt		= 21
	C_Remark			= 22
	C_ItemSpec			= 23
	C_BillSeq			= 24
	C_RetItemFlag		= 25
	C_PreBillNo			= 26
	C_PreBillSeq		= 27
	C_OldBillAmt		= 28
	C_OldVatIncFlag		= 29
	C_OldVatAmt			= 30
	C_InitialBillAmt	= 31
	C_InitialVatIncFlag = 32
	C_InitialVatAmt		= 33	
	C_TrackNo		= 34    'Tracking No 
	C_TrackingNoPopup	= 35   'Tracking No
	
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
	With frm1
		.txtConBillNo.focus
		.btnPostFlag.disabled = True
		.btnPostFlag.value = "Ȯ��"
		.btnGLView.disabled = True
		.btnPreRcptView.disabled = True
	End With
	lgBlnFlgChgValue = False
End Sub

'=========================================
Sub SetRowDefaultVal(ByVal pvRowCnt)

	With frm1.vspdData
	
		.Row = pvRowcnt
	
		'���� Default��ó�� 
		If Parent.gPlant <> "" Then
			.Col = C_PlantCd
			.Text = Parent.gPlant
			  
			.Col = C_ItemCd
			.Action = 0
		End If
		
		.Row = pvRowcnt	' -- .Action = 0 ������ TopLeft�̺�Ʈ�� �߻���Row�� ����ȴ�.
		    
		'�ΰ��� ���Կ��� Default�� ���� 
		If frm1.rdoVatIncFlag1.Checked Then
			
			.Col = C_VatIncFlag
			.text = "1"
			
			
			.Col = C_VatIncFlagNm
			.Text  = "����"
		
		Else
			.Col = C_VatIncFlag
			.text = "2"

			.Col = C_VatIncFlagNm
			.Text  = "����"
		
		End If
		  
		' Header�� �ΰ��� ������ ��ϵǾ� �ִ� ��� �ش� ���� Default���� �����Ѵ�.
		If lgStrVatFlag = "N" Then
		
			.Col = C_VatType
			.Text = frm1.txtVatType.value
			   
			.Col = C_VatTypeNm
			.Text = frm1.txtVatTypeNm.value
			   
			.Col = C_VatRate
			.Text = frm1.txtVatRate.Text
		End If

		.Col  = C_TrackNo
		.Text = "*"
		Call ChangeTrackingSetField(pvRowCnt)

	End With

End Sub

'==========================================
Sub InitSpreadSheet()
	On Error Resume Next
	Call initSpreadPosVariables() 
	
	With ggoSpread
		.Source = frm1.vspdData
		'patch version
        .Spreadinit "V20070413",,parent.gAllowDragDropSpread    		
		frm1.vspdData.ReDraw = false
		frm1.vspdData.MaxRows = 0 : frm1.vspdData.MaxCols = 0
		frm1.vspdData.MaxCols = C_TrackingNoPopup + 1            '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
        
        Call GetSpreadColumnPos("A")
		.SSSetEdit  C_ItemCd, "ǰ��", 18,,,18,2
	    .SSSetButton C_ItemPopup
		.SSSetEdit  C_ItemNm, "ǰ���", 30
		.SSSetFloat C_BillQty,"����" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		.SSSetEdit  C_BillUnit, "����", 8,2,,3,2
	    .SSSetButton C_UnitPopup
		.SSSetFloat C_BillPrice,"�ܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetCombo C_VatIncFlagNm,"VAT���Ա���", 15,2
		.SetCombo      "����" & vbTab & "����",C_VatIncFlagNm
		.SSSetEdit  C_VatIncFlag, "VAT���Ա���", 1,2
		.SetCombo      "1" & vbTab & "2",C_VatIncFlag
		.SSSetFloat C_BillAmt,"�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetEdit  C_VatType, "VAT����",10,2,,4,2
		.SSSetButton  C_VatPopup
		.SSSetEdit  C_VatTypeNm, "VAT������", 20 
		.SSSetFloat C_VatRate,"VAT��",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_VatAmt,"VAT�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_BillLocAmt,"�ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_DepositPrice,"�����ܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_DepositAmt,"�����ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_VatLocAmt,"VAT�ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetEdit  C_PlantCd, "����", 10,,,4,2
		.SSSetButton C_PlantPopup
		.SSSetEdit  C_Remark, "���", 30,,,120
		.SSSetEdit  C_ItemSpec, "ǰ��԰�", 30

		.SSSetEdit  C_TrackNo, "Tracking No", 25,,,25,2 '������ �߰�'
	        .SSSetButton C_TrackingNoPopup

		Call AppendNumberPlace("6","4","0")
		.SSSetFloat C_BillSeq,"����ä�Ǽ���" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
		.SSSetEdit  C_PreBillNo, "��������ä�ǹ�ȣ", 10
		.SSSetFloat C_PreBillSeq,"��������ä�Ǽ���" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
		.SSSetEdit  C_RetItemFlag, "��ǰ����", 10,2,,1,2


		.SSSetFloat C_OldBillAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetEdit  C_OldVatIncFlag, "", 10,2,,1,2
		.SSSetFloat C_OldVatAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    
		ggoSpread.SSSetFloat C_InitialBillAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit  C_InitialVatIncFlag, "", 10, 2,, 1, 2
		ggoSpread.SSSetFloat C_InitialVatAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		
		Call .MakePairsColumn(C_ItemCd,C_ItemPopup)
		Call .MakePairsColumn(C_BillUnit,C_UnitPopup)
		Call .MakePairsColumn(C_PlantCd,C_PlantPopup)
		Call .MakePairsColumn(C_VatType,C_VatPopup)
		Call .MakePairsColumn(C_TrackNo,C_TrackingNoPopup)
	    
	    Call .SSSetColHidden(C_BillSeq,C_BillSeq,True)
	    Call .SSSetColHidden(C_PreBillNo,C_PreBillNo,True)
	    Call .SSSetColHidden(C_PreBillSeq,C_PreBillSeq,True)
	    Call .SSSetColHidden(C_VatIncFlag,C_VatIncFlag,True)
	    Call .SSSetColHidden(C_RetItemFlag,C_RetItemFlag,True)

	    Call .SSSetColHidden( C_OldBillAmt, C_OldBillAmt,True)
	    Call .SSSetColHidden( C_OldVatIncFlag, C_OldVatIncFlag,True)
	    Call .SSSetColHidden( C_OldVatAmt, C_OldVatAmt,True)
	    Call .SSSetColHidden( C_InitialBillAmt, C_InitialBillAmt,True)
	    Call .SSSetColHidden( C_InitialVatIncFlag , C_InitialVatIncFlag ,True)
	    Call .SSSetColHidden( C_InitialVatAmt, C_InitialVatAmt,True)
	    Call .SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)				'��: ������Ʈ�� ��� Hidden Column

    End With
	frm1.vspdData.ReDraw = true
    
End Sub

'==========================================
Sub SetSpreadLock()
End Sub

'==========================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	With ggoSpread
		If frm1.txtHRefFlag.value <> "" Then
			'���� ������ ������ ��� 
			.SSSetProtected C_ItemCd		, pvStartRow, pvEndRow
			.SSSetProtected C_BillUnit		, pvStartRow, pvEndRow
			.SSSetProtected C_PlantCd		, pvStartRow, pvEndRow
			.SSSetProtected C_VatIncflagNm	, pvStartRow, pvEndRow    
			.SSSetProtected C_TrackNo	, pvStartRow, pvEndRow
		Else
			' ������ ����� ��� 
			.SSSetRequired  C_ItemCd		, pvStartRow, pvEndRow
			.SSSetRequired  C_BillUnit		, pvStartRow, pvEndRow
			.SSSetRequired  C_PlantCd		, pvStartRow, pvEndRow
			.SSSetRequired  C_VatIncflagNm	, pvStartRow, pvEndRow    
			.SSSetRequired  C_TrackNo	, pvStartRow, pvEndRow
		End if
 
		.SSSetProtected C_ItemNm		, pvStartRow, pvEndRow    
		.SSSetRequired  C_BillQty		, pvStartRow, pvEndRow
		.SSSetRequired  C_BillPrice		, pvStartRow, pvEndRow
		.SSSetRequired  C_BillAmt		, pvStartRow, pvEndRow
		.SSSetRequired  C_VatType		, pvStartRow, pvEndRow

		'2006-03-29 ������ ���� (��ȭ�ݾ� ���� ����)
'		.SSSetProtected C_BillLocAmt	, pvStartRow, pvEndRow 
		.SSSetRequired  C_BillLocAmt	, pvStartRow, pvEndRow 

		.SSSetRequired  C_VatAmt		, pvStartRow, pvEndRow

'		.SSSetProtected C_VatLocAmt	, pvStartRow, pvEndRow
		.SSSetRequired  C_VatLocAmt	, pvStartRow, pvEndRow

		.SSSetProtected C_VatTypeNm		, pvStartRow, pvEndRow
		.SSSetProtected C_VatRate		, pvStartRow, pvEndRow
		.SSSetProtected C_DepositPrice	, pvStartRow, pvEndRow
		.SSSetProtected C_DepositAmt	, pvStartRow, pvEndRow
		.SSSetProtected C_RetItemFlag	, pvStartRow, pvEndRow
		.SSSetProtected C_ItemSpec		, pvStartRow, pvEndRow    

	End With
	
End Sub

'==========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
			  Call SetActiveCell(frm1.vspdData, iDx, iRow,"M","X","X")
              Exit For
           End If
       Next
          
    End If   
End Sub

'==========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_PlantCd			= iCurColumnPos(1)
			C_PlantPopup		= iCurColumnPos(2)
			C_ItemCd  			= iCurColumnPos(3)    
			C_ItemPopup 		= iCurColumnPos(4)
			C_ItemNm  			= iCurColumnPos(5)
			C_BillQty			= iCurColumnPos(6)
			C_BillUnit 			= iCurColumnPos(7)
			C_UnitPopup			= iCurColumnPos(8)
			C_VatIncFlagNm		= iCurColumnPos(9)
			C_VatIncFlag		= iCurColumnPos(10)
			C_BillPrice			= iCurColumnPos(11)
			C_BillAmt  			= iCurColumnPos(12)
			C_VatType			= iCurColumnPos(13)
			C_VatPopup			= iCurColumnPos(14)
			C_VatTypeNm			= iCurColumnPos(15)
			C_VatRate			= iCurColumnPos(16)
			C_VatAmt  			= iCurColumnPos(17)    
			C_BillLocAmt 		= iCurColumnPos(18)
			C_VatLocAmt 		= iCurColumnPos(19)
			C_DepositPrice		= iCurColumnPos(20)
			C_DepositAmt 		= iCurColumnPos(21)
			C_Remark			= iCurColumnPos(22)
			C_ItemSpec			= iCurColumnPos(23)
			C_BillSeq  			= iCurColumnPos(24)
			C_RetItemFlag 		= iCurColumnPos(25)
			C_PreBillNo 		= iCurColumnPos(26)
			C_PreBillSeq 		= iCurColumnPos(27)
			C_OldBillAmt 		= iCurColumnPos(28)
			C_OldVatIncFlag		= iCurColumnPos(29)
			C_OldVatAmt 		= iCurColumnPos(30)
			C_InitialBillAmt	= iCurColumnPos(31)
			C_InitialVatIncFlag = iCurColumnPos(32)
			C_InitialVatAmt		= iCurColumnPos(33)	
			C_TrackNo		= iCurColumnPos(34)	
			C_TrackingNoPopup	= iCurColumnPos(35)	
    End Select    
End Sub

'==========================================
Sub SetPostYesSpreadColor(ByVal lRow)
	Dim iIntMaxRows
	
	Call SetToolbar("11100000000111")

	iIntMaxRows = frm1.vspdData.MaxRows
	frm1.vspdData.ReDraw = False
	With ggoSpread
		.SSSetProtected C_ItemCd, lRow, iIntMaxRows    
		.SSSetProtected C_ItemNm, lRow, iIntMaxRows    
		.SSSetProtected C_ItemPopup, lRow, iIntMaxRows    
		.SSSetProtected C_BillQty, lRow, iIntMaxRows 
		.SSSetProtected C_BillUnit, lRow, iIntMaxRows
		.SSSetProtected C_UnitPopup, lRow, iIntMaxRows
		.SSSetProtected C_BillPrice, lRow, iIntMaxRows
		.SSSetProtected C_BillAmt, lRow, iIntMaxRows
		.SSSetProtected C_VatAmt, lRow, iIntMaxRows
		.SSSetProtected C_BillLocAmt, lRow, iIntMaxRows
		.SSSetProtected C_VatLocAmt, lRow, iIntMaxRows
		.SSSetProtected C_PlantCd, lRow, iIntMaxRows
		.SSSetProtected C_PlantPopup, lRow, iIntMaxRows
		.SSSetProtected C_Remark, lRow, iIntMaxRows
		.SSSetProtected C_VatType, lRow, iIntMaxRows
		.SSSetProtected C_VatTypeNm, lRow, iIntMaxRows
		.SSSetProtected C_VatRate, lRow, iIntMaxRows
		.SSSetProtected C_VatPopUp, lRow, iIntMaxRows
		.SSSetProtected C_DepositAmt, lRow, iIntMaxRows
		.SSSetProtected C_DepositPrice, lRow, iIntMaxRows
		.SSSetProtected C_VatIncflagNm, lRow, iIntMaxRows
		.SSSetProtected C_RetItemFlag, lRow, iIntMaxRows
		.SSSetProtected C_ItemSpec, lRow, iIntMaxRows
		.SSSetProtected C_TrackNo, lRow, iIntMaxRows
		.SSSetProtected C_TrackingNoPopup, lRow, iIntMaxRows

	End With

	frm1.vspdData.ReDraw = True
	
End Sub

'==========================================
Sub SetQuerySpreadColor(ByVal lRow)
	Dim iIntMaxRows
	
	iIntMaxRows = frm1.vspdData.MaxRows
	
	frm1.vspdData.ReDraw = False

	With ggoSpread 
		.SSSetProtected C_PlantCd, lRow, iIntMaxRows
		.SSSetProtected C_PlantPopup, lRow, iIntMaxRows
		.SSSetProtected C_ItemCd, lRow, iIntMaxRows    
		.SSSetProtected C_ItemNm, lRow, iIntMaxRows    
		.SSSetProtected C_ItemSpec, lRow, iIntMaxRows    
		.SSSetProtected C_ItemPopup, lRow, iIntMaxRows    
		.SSSetRequired  C_BillPrice, lRow, iIntMaxRows
		.SSSetRequired  C_BillQty, lRow, iIntMaxRows
		.SSSetRequired  C_VatType, lRow, iIntMaxRows
		.SSSetProtected C_VatRate, lRow, iIntMaxRows
		.SSSetProtected C_VatTypeNm, lRow, iIntMaxRows
		.SSSetRequired  C_VatAmt, lRow, iIntMaxRows

		'2006-03-29 ������ ���� (�ڱ��ݾ� ���� ����)	
'		.SSSetProtected C_BillLocAmt, lRow, iIntMaxRows
'		.SSSetProtected C_VatLocAmt, lRow, iIntMaxRows

		.SSSetRequired  C_BillLocAmt, lRow, iIntMaxRows
		.SSSetRequired  C_VatLocAmt, lRow, iIntMaxRows

		.SSSetProtected C_VatRate, lRow, iIntMaxRows
		.SSSetRequired  C_BillAmt, lRow, iIntMaxRows
		.SSSetProtected C_DepositPrice, lRow, iIntMaxRows
		.SSSetProtected C_DepositAmt, lRow, iIntMaxRows
		.SSSetProtected C_RetItemFlag, lRow, iIntMaxRows

		If frm1.txtHRefFlag.value = "" Then 
			.SSSetRequired  C_BillUnit, lRow, iIntMaxRows
			.SSSetRequired  C_VatIncflagNm, lRow, iIntMaxRows					   
		Else
			.SSSetProtected C_BillUnit, lRow, iIntMaxRows
			.SSSetProtected C_UnitPopup, lRow, iIntMaxRows
			.SSSetProtected C_VatIncflagNm, lRow, iIntMaxRows
		End if

		.SSSetProtected C_TrackNo, lRow, iIntMaxRows
		.SSSetProtected C_TrackingNoPopup, lRow, iIntMaxRows

	End With
	frm1.vspdData.ReDraw = True
End Sub

'==========================================
Sub SetSpreadHidden()

	With ggoSpread
		If frm1.rdoVatCalcType2.checked = True then
			Call .SSSetColHidden(C_VatType,C_VatType,True)
			Call .SSSetColHidden(C_VatTypeNm,C_VatTypeNm,True)
			Call .SSSetColHidden(C_VatRate,C_VatRate,True)
			Call .SSSetColHidden(C_VatIncFlagNm,C_VatIncFlagNm,True)
		Else
			Call .SSSetColHidden(C_VatType,C_VatType,False)
			Call .SSSetColHidden(C_VatTypeNm,C_VatTypeNm,False)
			Call .SSSetColHidden(C_VatRate,C_VatRate,False)
			Call .SSSetColHidden(C_VatIncFlagNm,C_VatIncFlagNm,False)
		End If
	
		'2006-03-29 ������ ���� ��ȭ�ݾ� ���� ���� �ϵ��� ���� s. 
'		If UCase(Parent.gCurrency) <> UCase(Trim(frm1.txtCurrency.value)) Then
'			Call .SSSetColHidden(C_BillLocAmt,C_BillLocAmt,False)
'			Call .SSSetColHidden(C_VatLocAmt,C_VatLocAmt,False)
'		Else
'			Call .SSSetColHidden(C_BillLocAmt,C_BillLocAmt,True)
'			Call .SSSetColHidden(C_VatLocAmt,C_VatLocAmt,True)
'		End If

		Call .SSSetColHidden(C_BillLocAmt,C_BillLocAmt,False)
		Call .SSSetColHidden(C_VatLocAmt,C_VatLocAmt,False)

		'2006-03-29 ������ ���� ��ȭ�ݾ� ���� ���� �ϵ��� ���� e. 
		
		If frm1.txtHRefFlag.value = "" Then 
			Call .MakePairsColumn(C_ItemCd,C_ItemPopup)
			Call .MakePairsColumn(C_BillUnit,C_UnitPopup)
			Call .MakePairsColumn(C_PlantCd,C_PlantPopup)
			
			Call .SSSetColHidden(C_ItemPopUp,C_ItemPopUp,False)
			Call .SSSetColHidden(C_UnitPopup,C_UnitPopup,False)
			Call .SSSetColHidden(C_PlantPopUp,C_PlantPopUp,False)
			
		Else
			Call .MakePairsColumn(C_ItemCd,C_ItemPopup,"1")
			Call .MakePairsColumn(C_BillUnit,C_UnitPopup,"1")
			Call .MakePairsColumn(C_PlantCd,C_PlantPopup,"1")
			
			Call .SSSetColHidden(C_ItemPopUp,C_ItemPopUp,True)
			Call .SSSetColHidden(C_UnitPopup,C_UnitPopup,True)
			Call .SSSetColHidden(C_PlantPopUp,C_PlantPopUp,True)
		End if
		
	End With
End Sub		

'==========================================
Function CookiePage(Byval Kubun)

	On Error Resume Next
	Const CookieSplit = 4877
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.txtHBillNo.value
	ElseIf Kubun = 0 Then
		strTemp = ReadCookie(CookieSplit)
		
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, Parent.gRowSep)
		 
		If arrVal(0) = "" Then Exit Function
		frm1.txtConBillNo.value =  arrVal(0)
		frm1.txtHRefFlag.value =  arrVal(1)

		WriteCookie CookieSplit , ""
		Call DBQuery()
		 
	End If
End Function

'==========================================
Function JumpChgCheck(byVal pvStrJumpPgmId)
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
	Call PgmJump(pvStrJumpPgmId)

End Function

'==========================================
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

'==========================================
Function RefCheckMessage()

	RefCheckMessage = True

	If frm1.txtHPostFlag.value = "Y" Then
		Msgbox "�̹� Ȯ���� �Ǿ ���� �� �� �����ϴ�",vbInformation, Parent.gLogoName
		Exit Function
	End If

	RefCheckMessage = False

End Function

Sub LockFieldInit()
    Call FormatDoubleSingleField(frm1.txtXchgRate)
    Call LockObjectField(frm1.txtXchgRate,"P")

    Call FormatDoubleSingleField(frm1.txtVatRate)
    Call LockObjectField(frm1.txtVatRate,"P")

    Call FormatDoubleSingleField(frm1.txtVatAmt)
    Call LockObjectField(frm1.txtVatAmt,"P")

    Call FormatDoubleSingleField(frm1.txtOriginBillAmt)
    Call LockObjectField(frm1.txtOriginBillAmt ,"P")
End Sub

'==========================================
Sub Form_Load()
	Call SetDefaultVal    
	Call InitVariables              
	Call LoadInfTB19029
'	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
'	Call ggoOper.LockField(Document, "N")                                   
	Call LockFieldInit
	Call InitSpreadSheet
	Call InitVATTypeInfo()

	Call LockHTMLField(frm1.rdoVatIncFlag1, "P")	
	Call LockHTMLField(frm1.rdoVatIncFlag2, "P")	
	Call LockHTMLField(frm1.rdoVatCalcType1, "P")	
	Call LockHTMLField(frm1.rdoVatCalcType2, "P")	

	Call SetToolbar("11000000000011")          
	Call CookiePage(0)
End Sub

'==========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================
Function FncQuery() 
    Dim IntRetCD 
    
    Err.Clear             

    FncQuery = False                                                        
                                                      
    If Not chkFieldByCell(frm1.txtConBillNo, "A", 1) Then Exit Function 

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
		If IntRetCD = vbNo Then Exit Function
    End If

    Call ggoOper.ClearField(Document, "A")
    Call SetDefaultVal
    Call InitVariables

    Call SetToolbar("11000000000011")          

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
    
    If ggoSpread.SSDefaultCheck = False Then     
       Exit Function
    End If

    CAll  DbSave                                                   
    
    FncSave = True                                                          
    
End Function

'========================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function
    
	With frm1
	 .vspdData.ReDraw = False
	 
	 ggoSpread.Source = frm1.vspdData 
	 ggoSpread.CopyRow
	 SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

	 call  ChangeTrackingSetField(.vspdData.ActiveRow ) ' 2007-06-07 ������ �߰� 

	 .vspdData.ReDraw = True
	End With
    
End Function

'========================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	Call CalcTotal("C", frm1.vspdData.ActiveRow)
	
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.EditUndo  
End Function

'========================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow,i
	On Error Resume Next                                                          
    Err.Clear
    
    FncInsertRow = False                                                         

    If IsNumeric(Trim(pvRowCnt)) Then
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
	
		For i = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1			
			SetRowDefaultVal i
		Next	
		.vspdData.ReDraw = True
				 
		lgBlnFlgChgValue = True
	End With
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          
    End If    
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================
Function FncDeleteRow() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement   
		
	ggoSpread.Source = frm1.vspdData 
			    
	Call CalcTotal("D", 0)
		
	ggoSpread.DeleteRow
End Function

'========================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
 On Error Resume Next                                                             
    Err.Clear                                                                     

    FncExcel = False                                                              

	Call parent.FncExport(Parent.C_SINGLEMULTI)	                     			  '��: ȭ�� ���� 

    If Err.number = 0 Then	 
       FncExcel = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================
Function FncFind()
	On Error Resume Next                                                          
    Err.Clear                                                                     

    FncFind = False                                                               
     
    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                              
    
    If Err.number = 0 Then	 
       FncFind = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
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
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
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
		strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtHBillNo.value)    
		strVal = strVal & "&txtHQuery=F"
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         
		strVal = strVal & "&txtConBillNo=" & Trim(frm1.txtConBillNo.value)    
		strVal = strVal & "&txtHQuery=T"
	End If 
	
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)            
	 
	Call DbQueryOk
	DbQuery = True   
                          
End Function

'========================================
Function DbQueryOk()              
	Dim intCnt
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE            
	lgBlnFlgChgValue = False
	  
	If frm1.txtHRefFlag.value <> "" Then
		Call SetToolbar("11101011000111")
	Else
		Call SetToolbar("11101111001111")
	End if

	If UNICDbl(frm1.txtSts.value) < 3 Then
		frm1.btnPostFlag.disabled = False
	Else
		frm1.btnPostFlag.disabled = True
	End If
	
	frm1.vspdData.Focus
End Function

'========================================
Function DbSave() 
    Err.Clear					                
    DbSave = False                             
    
    On Error Resume Next			              

	If LayerShowHide(1) = False Then Exit Function 

	Dim iIntRow
	Dim iArrColData
	Dim iStrDel, iStrVal
	Dim iDblAmt, iStrAmt
	Dim iColSep, iRowSep, iFormLimitByte, iChunkArrayCount
	
	Dim iLngCUTotalvalLen		'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim iLngDTotalValLen		'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]

	Dim iTmpCUBuffer			'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount		'������ ���� Position
	Dim iTmpCUBufferMaxCount	'������ ���� Chunk Size
	Dim iTmpDBuffer				'������ ���� [����] 
	Dim iTmpDBufferCount		'������ ���� Position
	Dim iTmpDBufferMaxCount		'������ ���� Chunk Size

	' �ӵ� ����� ���� Local ������ ������ 
	iColSep = parent.gColSep
	iRowSep = parent.gRowSep
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
	iChunkArrayCount = parent.C_CHUNK_ARRAY_COUNT

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	iLngCUTotalvalLen = 0
	iLngDTotalValLen  = 0
	
	Redim iArrColData(41)
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]
	
	iArrColData(12) =  ""		' ����ȣ 
	iArrColData(13) =  "0"		' ������ 
	iArrColData(14) =  ""		' S/O ��ȣ 
	iArrColData(15) =  "0"		' S/O ���� 
	iArrColData(16) =  ""		' L/C ��ȣ 
	iArrColData(17) =  "0"		' L/C ���� 

	With frm1.vspdData
	
		For iIntRow = 1 To .MaxRows
			.Row = iIntRow
			.Col = 0

			'������ ��� 
			If .Text = ggoSpread.DeleteFlag then
				iStrDel = "D" & iColSep & iIntRow & iColSep

				.Col = C_BillSeq
				iStrDel = iStrDel & UNIConvNum(.Text,0) & iRowSep

				If iLngDTotalValLen + Len(iStrDel) >  iFormLimitByte Then				'�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
					Call MakeTextArea("txtDSpread", iTmpDBuffer)
							          
				   iTmpDBufferMaxCount = iChunkArrayCount
				   ReDim iTmpDBuffer(iTmpDBufferMaxCount)
				   iTmpDBufferCount = -1
				   iLngDTotalValLen = 0 
				End If
							       
				iTmpDBufferCount = iTmpDBufferCount + 1

				If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
				   iTmpDBufferMaxCount = iTmpDBufferMaxCount + iChunkArrayCount
				   ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
				End If   
							         
				iTmpDBuffer(iTmpDBufferCount) =  iStrDel         
				iLngDTotalValLen = iLngDTotalValLen + Len(iStrDel)

			' �Է�, ������ ��� 
			Elseif .Text <> "" Then
				If .Text = ggoSpread.InsertFlag Then
					iArrColData(0) = "C"
				Else
					iArrColData(0) = "U"
				End If

				iArrColData(1) = iIntRow
				.Col = C_BillSeq	:	iArrColData(2) =  UNIConvNum(.Text,0)		' ������� 
				.Col = C_ItemCd		:	iArrColData(3) =  Trim(.Text)				' ǰ�� 
				.Col = C_BillQty	:	iArrColData(4) =  UNIConvNum(.Text,0)		' ���� 
				.Col = C_BillUnit	:	iArrColData(5) =  Trim(.Text)				' ���� 
				.Col = C_BillPrice	:	iArrColData(6) =  UNIConvNum(.Text,0)		' �ܰ� 
				.Col = C_VatType	:	iArrColData(8) =  Trim(.Text)				' VAT���� 
				.Col = C_VatRate	:	iArrColData(9) =  UNIConvNum(.Text,0)		' VAT �� 
				.Col = C_VatAmt		:   iArrColData(10) = UNIConvNum(.Text,0)		' VAT �ݾ� 
				.Col = C_Remark		:	iArrColData(11) = Trim(.Text)				' ��� 
				.Col = C_PlantCd	:	iArrColData(18) = Trim(.Text)				' ���� 
				.Col = C_VatLocAmt	:   iArrColData(20) = UNIConvNum(.Text,0)		' VAT �ڱ��ݾ� 
				.Col = C_VatIncFlag	:	iArrColData(21) = Trim(.Text)				' VAT ���Ա��� 
				' VAT ���Ա��п� ���� �ݾװ�� 
				If iArrColData(21) = "1" Then
					.Col = C_BillAmt	:   iArrColData(7) = UNIConvNum(.Text,0)
					.Col = C_BillLocAmt	:   iArrColData(19) = UNIConvNum(.Text,0)
				Else
					' Db�� Bill_amt�� �ΰ��� ���� �ݾ׸� �����Ѵ�.
					' (ȭ���� �ݾ��� '�ΰ������Կ���'�� '����'�� ��� �ΰ��� ���Աݾ��̴�.)
					.Col = C_BillAmt	:	iDblAmt = UNICDbl(.Text)
					.Col = C_VatAmt		:	iDblAmt = iDblAmt - UNICDbl(.Text)
					
					iStrAmt = UNIConvNumPCToCompanyByCurrency(iDblAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
					iArrColData(7) = UNIConvNum(iStrAmt, 0)
					
					.Col = C_BillLocAmt	:	iDblAmt = UNICDbl(.Text)
					.Col = C_VatLocAmt	:	iDblAmt = iDblAmt - UNICDbl(.Text)

					iStrAmt = UNIConvNumPCToCompanyByCurrency(iDblAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo, pvStrRndPolicyNo, "X")
					iArrColData(19) = UNIConvNum(iStrAmt, 0)
				End If
				.Col = C_DepositPrice	:   iArrColData(22) = UNIConvNum(.Text,0)
				.Col = C_DepositAmt		:   iArrColData(23) = UNIConvNum(.Text,0)
				.Col = C_RetItemFlag	:	iArrColData(24) =  Trim(.Text)
				.Col = C_TrackNo	:	iArrColData(41) =  Trim(.Text)
				
				iStrVal = Join(iArrColData, iColSep) & iRowSep
			    If iLngCUTotalvalLen + Len(iStrVal) >  iFormLimitByte Then					'�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
					Call MakeTextArea("txtCUSpread", iTmpCUBuffer)
					
			       iTmpCUBufferMaxCount = iChunkArrayCount									' �ӽ� ���� ���� �ʱ�ȭ 
			       ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			       iTmpCUBufferCount = -1
			       iLngCUTotalvalLen  = 0
			    End If
			       
			    iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			    If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                            '������ ���� ����ġ�� ������ 
			       iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + iChunkArrayCount			'���� ũ�� ���� 
			       ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			    End If   
			    iTmpCUBuffer(iTmpCUBufferCount) =  iStrVal         
			    iLngCUTotalvalLen = iLngCUTotalvalLen + Len(iStrVal)
			End If
		Next
	End With

   ' ������ ������ ó�� 
	If iTmpCUBufferCount > -1 Then Call MakeTextArea("txtCUSpread", iTmpCUBuffer)
	If iTmpDBufferCount > -1 Then Call MakeTextArea("txtDSpread", iTmpDBuffer)
		 
	frm1.txtMode.value = Parent.UID_M0002

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         
	 
	DbSave = True                                                           
    
End Function

'========================================
Function DbSaveOk()
	Call InitVariables
	frm1.txtConBillNo.value = frm1.txtHBillNo.value
	Call ggoOper.ClearField(Document, "2")
	Call RemovedivTextArea
	Call MainQuery()
End Function

'========================================
Sub MakeTextArea(ByVal pvStrName, ByRef prArrData)
	Dim iObjTEXTAREA		'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Set iObjTEXTAREA = document.createElement("TEXTAREA")            '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
	iObjTEXTAREA.name = pvStrName
	iObjTEXTAREA.value = Join(prArrData,"")
	divTextArea.appendChild(iObjTEXTAREA)
End Sub

'========================================
Function RemovedivTextArea()
	Dim iIntIndex
	
	For iIntIndex = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'========================================
Sub InitVATTypeInfo()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iArrVATType, iArrVATTypeNm, iArrVATRate
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim iIntIndex
	
	Err.Clear
	
	iStrSelectList	= " Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE "
	iStrFromList	= " B_MINOR Minor,B_CONFIGURATION Config "
	iStrWhereList	= " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 " 
	
	If CommonQueryRs(iStrSelectList, iStrFromList, iStrWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		iArrVATType		= Split(lgF0, parent.gColSep)
		iArrVATTypeNm	= Split(lgF1, parent.gColSep)
		iArrVATRate		= Split(lgF2, parent.gColSep)
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
		End If
		Exit Sub
	End If

	Redim lgArrVATTypeInfo(UBound(iArrVATType) - 1, 2)

	For iIntIndex = 0 to UBound(iArrVATType) - 1
		lgArrVATTypeInfo(iIntIndex, 0) = iArrVATType(iIntIndex)
		lgArrVATTypeInfo(iIntIndex, 1) = iArrVATTypeNm(iIntIndex)
		lgArrVATTypeInfo(iIntIndex, 2) = iArrVATRate(iIntIndex)
	Next
End Sub

'========================================
Sub SetVATTypeForSpread(ByVal pvIntRow)
	Dim iStrVATType, iStrVATTypeNm, iStrVATRate
	
	With frm1.vspdData
		.Row = pvIntRow
		.Col = C_VatType : iStrVATType = .text

		Call GetVATType(iStrVATType, iStrVATTypeNm, iStrVATRate)
		 
		.Col = C_VatTypeNm	: .Text = iStrVATTypeNm
		.Col = C_VatRate	: .Text = iStrVATRate
	End With
End Sub

'========================================
Sub GetVATType(ByVal pvStrVATType, ByRef prStrVATTypeNm, ByRef prStrVATRate)
	Dim iIntIndex

	For iIntIndex = 0 To Ubound(lgArrVATTypeInfo, 1)
		If UCase(lgArrVATTypeInfo(iIntIndex, 0)) = UCase(pvStrVATType) Then
			prStrVATTypeNm = lgArrVATTypeInfo(iIntIndex, 1)
			prStrVATRate   = lgArrVATTypeInfo(iIntIndex, 2)
			Exit Sub
		End If
	Next

	prStrVATTypeNm = ""
	prStrVATRate = "0"
End Sub

'===========================================
' Spread button popup
Function OpenSpreadPopup(ByVal pvIntCol, ByVal pvIntRow)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1.vspdData
		.Row = pvIntRow		:	.Col = pvIntCol - 1
		
		Select Case pvIntCol
			' ���� 
			Case C_PlantPopup
				.Col = C_ItemCd
				If Trim(.Text) = "" Then
					iArrParam(1) = "B_PLANT PT INNER JOIN B_ITEM_BY_PLANT IP ON (IP.PLANT_CD = PT.PLANT_CD) "' FROM Clause
				Else
					iArrParam(1) = "B_PLANT PT INNER JOIN B_ITEM_BY_PLANT IP ON (IP.PLANT_CD = PT.PLANT_CD) "' FROM Clause
					iArrParam(4) = "IP.ITEM_CD = "	& FilterVar(.Text, "''", "S") & ""		' Where Condition
				End If

				.Col = C_PlantCd				
				iArrParam(2) = .Text													' Code Condition
				iArrParam(3) = ""														' Name Cindition
				
				iArrField(0) = "ED15" & Parent.gColSep & "PT.PLANT_CD"		' ���� 
				iArrField(1) = "ED30" & Parent.gColSep & "PT.PLANT_NM"		' ����� 

				.Row = 0
				iArrHeader(0) = .Text					' Header��(0)
				iArrHeader(1) = "�����"			' Header��(1)

			' ǰ�� 
			Case C_ItemPopup
				lgBlnOpenPop = False
				OpenSpreadPopup = OpenItem(.Text)
				
				Exit Function

			' ���� 
			Case C_UnitPopup
				iArrParam(1) = " B_UNIT_OF_MEASURE "
				iArrParam(2) = .Text
				iArrParam(3) = ""
				iArrParam(4) = "DIMENSION = " & FilterVar("CT", "''", "S") & ""
					
				iArrField(0) = "ED15" & Parent.gColSep & "UNIT"
				iArrField(1) = "ED30" & Parent.gColSep & "UNIT_NM"
				    
				.Row = 0
				iArrHeader(0) = .Text
				iArrHeader(1) = "������"

			' �ΰ������� 
			Case C_VatPopup
				iArrParam(1) = "B_MINOR MI INNER JOIN B_CONFIGURATION CF ON (CF.MAJOR_CD = MI.MAJOR_CD AND CF.MINOR_CD = MI.MINOR_CD) "
				iArrParam(2) = .Text
				iArrParam(3) = ""
				iArrParam(4) = "MI.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND CF.SEQ_NO = 1 "
					
				iArrField(0) = "ED15" & Parent.gColSep & "MI.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MI.MINOR_NM"
				iArrField(2) = "ED8" & Parent.gColSep & "CF.REFERENCE"
				    
				.Row = 0
				iArrHeader(0) = .Text
				.Col = C_VatTypeNm	:	iArrHeader(1) = .Text
				.Col = C_VatRate	:	iArrHeader(2) = .Text

			' Tracking �˾� ( 2007-04-16 ������ ���� ) 
			Case C_TrackingNoPopup  
				iArrParam(1) = "s_so_tracking a, b_item_by_plant b, b_item c"    
				iArrParam(2) = .Text
				iArrParam(3) = ""
				iArrParam(4) = "a.item_cd = b.item_cd and a.sl_cd = b.plant_cd and b.item_cd = c.item_cd"      		 
				iArrParam(5) = "Tracking No"   
		    
				iArrField(0) = "a.tracking_no"       
				iArrField(1) = "a.item_cd"   
				iArrField(2) = "c.item_nm"   
				iArrField(3) = "c.spec"   	
		  
				iArrHeader(0) = "Tracking No"		
				iArrHeader(1) = "ǰ��"
				iArrHeader(2) = "ǰ���"
				iArrHeader(3) = "Spec"


		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' �˾� ��Ī 
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� TextBox ��Ī 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	If iArrRet(0) <> "" Then
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvIntCol, pvIntRow)
	End If	

End Function

'===========================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvIntCol, ByVal pvIntRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvIntRow
		.Col = pvIntCol	- 1	:	.Text = pvArrRet(0)
		
		Select Case pvIntCol
			Case C_VatPopup
				.Col = C_VatTypeNm		: .Text = pvArrRet(1)
				.Col = C_VatRate		: .Text = pvArrRet(2)

			Case C_TrackingNoPopup  ' ������ �߰� 
				.Col = C_TrackNo		: .Text = pvArrRet(0)
		End Select
	End With

	Call SetRowStatus(pvIntRow)
	SetSpreadPopup = True
End Function

'===========================================
Function OpenConBillDtl()
	Dim iStrRet
	 
	Dim iCalledAspName
	Dim IntRetCD
	 
	If lgBlnOpenPop Then Exit Function
	   
	lgBlnOpenPop = True
	
	iCalledAspName = AskPRAspName("s5111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If
		 
	iStrRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=Y", Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	lgBlnOpenPop = False

	frm1.txtConBillNo.focus
	If iStrRet <> "" Then
		 frm1.txtConBillNo.value = iStrRet 
	End If 
 
End Function

'===========================================
Function OpenBillDtlRef()
	Dim iArrRet
	Dim iArrParam(8)
	
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
	 
	If RefCheckMessage Then Exit Function
	   
	If frm1.txtHRefflag.value <> "B"  Then    
		Call DisplayMsgBox("205237", "x", "x", "x")
		Exit Function
	End If
	  
	If lgIntFlgMode = Parent.OPMD_CMODE Then    
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("s5112ra2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5112ra2", "X")
		lblnWinEvent = False
		Exit Function
	End If

	With frm1
		iArrParam(0) = Trim(.txtSoldToParty.value)  
		iArrParam(1) = Trim(.txtSoldToPartyNm.value) 
		iArrParam(2) = Trim(.txtSalesGrpCd.value)    
		iArrParam(3) = Trim(.txtSalesGrpNm.value) 
		iArrParam(4) = Trim(.txtCurrency.value)

		'�ΰ��� ������ؿ� ���� ��ȸ���� ���� 
		If .rdoVatCalcType1.checked Then
			iArrParam(5) = "%"      '�ΰ��� ���� 
			iArrParam(6) = "%"      '�ΰ��� ���Ա��� 
		Else
			'������ ��� 
			iArrParam(5) = .txtVatType.value   '�ΰ��� ���� 
			'�ΰ��� ���Ա��� 
			If .rdoVatIncflag1.checked Then
				iArrParam(6) = .rdoVatIncFlag1.value
			Else
				iArrParam(6) = .rdoVatIncFlag2.value
			End if
		End if
		'���������ȣ 
		iArrParam(7) = .txtRefBillNo.value   
		iArrParam(8) = .txtHBillDt.value
		  
		iArrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & .txtCurrency.value, Array(window.parent,iArrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		lblnWinEvent = False
		
	End With
	  
	If iArrRet(0,0) <> "" Then
		Call SetBillDtlRef(iArrRet)
	End If 
End Function

'===========================================
Sub SetRowStatus(intRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow intRow
	lgBlnFlgChgValue = True
End Sub

'===========================================
Function SetBillDtlRef(iArrRet)
	Dim iIntRow, iIntStartRow, iIntIndex, iIntLastIndex
	Dim iDblBillAmtLoc, iDblVatRate, iDblTotBillAmt, iDblTotVatAmt
	Dim iBlnExists
	Dim iStrVatIncFlag, iStrVatAmt, iStrPreBillJungBokMsg

	iDblTotBillAmt = 0
	iDblTotVatAmt = 0
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False 

		iIntStartRow = .MaxRows						'��: ��������� MaxRows
		iIntLastIndex = Ubound(iArrRet, 1)          '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� 

		iStrPreBillJungBokMsg = ""
		   
		For iIntIndex = 0 To iIntLastIndex
			iBlnExists = False

			' �ش� ǰ���� �� ���뿩�� Check
			If iIntStartRow <> 0 Then
				For iIntRow = 1 To iIntStartRow
					.Row = iIntRow
					.Col = C_PreBillNo
					If Trim(.text) = iArrRet(iIntIndex, 0) Then
						.Col = C_PreBillSeq
						If Trim(.text) = iArrRet(iIntIndex, 1) Then
							iBlnExists = True
							iStrPreBillJungBokMsg = iStrPreBillJungBokMsg & Chr(13) & iArrRet(iIntIndex, 0) & "-" & iArrRet(iIntIndex, 1)
							Exit For
						End If
					End If
				Next
			End If

			If iBlnExists = False Then
				.MaxRows = .MaxRows + 1
				.Row = .MaxRows
				.Col = 0			:	.Text = ggoSpread.InsertFlag
				.Col = C_PreBillNo	:	.text = iArrRet(iIntIndex, 0)
				.Col = C_PreBillSeq :	.text = iArrRet(iIntIndex, 1)
				.Col = C_ItemCd     :	.text = iArrRet(iIntIndex, 2)
				.Col = C_ItemNm     :	.text = iArrRet(iIntIndex, 3)
				.Col = C_BillUnit   :	.text = iArrRet(iIntIndex, 4)
				.Col = C_BillQty    :	.text = iArrRet(iIntIndex, 5)
				.Col = C_BillPrice  :	.text = iArrRet(iIntIndex, 6)
				.Col = C_BillAmt    :	.text = iArrRet(iIntIndex, 7)
				.Col = C_OldBillAmt	:	.text = iArrRet(iIntIndex, 7)
				
				'��ǰ���� ���� 
				.Col = C_RetItemFlag
				If UNICDbl(iArrRet(iIntIndex, 7)) >= 0 Then
					.text = "N"
				Else
					.text = "Y"
				End If
				    
				.Col = C_PlantCd	:	.text = iArrRet(iIntIndex, 8)
				.Col = C_VatType    :	.text = iArrRet(iIntIndex, 9)
				.Col = C_VatTypeNm  :	.text = iArrRet(iIntIndex, 10)
				.Col = C_VatRate	:	.text = UNIConvNumPCToCompanyByCurrency(iArrRet(iIntIndex, 11), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
				.Col = C_VatAmt		:	.text = UNIConvNumPCToCompanyByCurrency(iArrRet(iIntIndex, 12), frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo  , "X")
				iStrVatAmt = .Text
				.Col = C_OldVatAmt	:	.text = iStrVatAmt
				iDblTotVatAmt = iDblTotVatAmt + iArrRet(iIntIndex, 12)
				
				.Col = C_VatIncflag	:	.text = iArrRet(iIntIndex, 13)
				.Col = C_OldVatIncFlag	:	.text = iArrRet(iIntIndex, 13)
				
				.Col = C_VatIncflagNm
				If iArrRet(iIntIndex, 13) = "1" Then
					.text = "����"
					iDblTotBillAmt = iDblTotBillAmt + UNICDbl(iArrRet(iIntIndex, 7))				
				Else
					.text = "����"
					iDblTotBillAmt = iDblTotBillAmt + UNICDbl(iArrRet(iIntIndex, 7)) - iArrRet(iIntIndex, 12)
				End if

				'--- �ڱ��ݾ�, VAT�ڱ��ݾ� 
				' 2006-03-29 ������ ���� (�ڱ��ݾ� ���� �����ϵ��� )
'				if UCase(frm1.txtCurrency.value) = UCase(Parent.gCurrency) Then
'					.Col = C_BillLocAmt
'					.text = UNIConvNumPCToCompanyByCurrency(iArrRet(iIntIndex, 16),Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo  , "X")
'					.Col = C_VatLocAmt
'					.Text = UNIConvNumPCToCompanyByCurrency(iArrRet(iIntIndex, 17),Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo  , "X")
'				Else
					.Col = C_BillLocAmt	: .Text = FncCalcAmtLoc(UNICDbl(iArrRet(iIntIndex, 7)), UNICDbl(frm1.txtXchgRate.Text), frm1.txtXchgOp.value, Parent.gLocRndPolicyNo)
					iDblBillAmtLoc = UNICDbl(.Text)

					.Col = C_VatIncFlag	: iStrVatIncFlag = .Text
					.Col = C_VatRate	: iDblVatRate	 = UNICDbl(.Text)
					.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)
'				End if
			
				' ǰ��԰�	
				.Col = C_ItemSpec		: .Text = iArrRet(iIntIndex, 18)

				'Setting Initial Amount 
				.Col = C_InitialBillAmt : .Text = 0
				.Col = C_InitialVatAmt  : .Text = 0
				.Col = C_InitialVatIncFlag  : .Text = iArrRet(iIntIndex, 13)

			End If
		Next

		' �߰��� Row�� ���� Color���� 
		If iIntStartRow <> .MaxRows Then
			SetSpreadColor iIntStartRow, .MaxRows
			' Head�� �ѱݾ� ��� 
			Call SetTotal(iDblTotBillAmt, iDblTotVatAmt)
		End If
		
		.ReDraw = True

		If Trim(iStrPreBillJungBokMsg) <> "" Then
			iStrPreBillJungBokMsg = "��������ä�ǹ�ȣ" & "-" & "��������ä�Ǽ���" & vbCrLf & _
									String(40,"=") & vbCrLf & _
									iStrPreBillJungBokMsg & vbCrLf & vbCrLf & _
									"�̹� ������ ��ȣ�� ������ �����մϴ�"
			MsgBox iStrPreBillJungBokMsg, vbInformation, Parent.gLogoName
		End If
	End With
End Function

' Description : ǰ���� �������� Fetch�Ѵ�.
'===========================================
Function GetItemInfo(ByVal pvIntRow)

	Dim strSoldToParty, strPlantCd, strItemCd, strPayMeth, strCurrency, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strItemInfo
	Dim strItemCd2
	
	GetItemInfo = False

	Call ChangeTrackingSetField(pvIntRow)
	
	With frm1.vspdData
		.Row = pvIntRow
		.Col = C_PlantCd       '���� 
		strPlantCd = .text
		.col = C_ItemCd        'ǰ���ڵ� 
		strItemCd = .text
		strItemCd2 = .text
		    
		strSoldToParty = frm1.txtSoldToParty.value  '�ֹ�ó 
		strPayMeth = frm1.txtPayTermsCd.value    '������� 
		strCurrency = frm1.txtCurrency.value    'ȭ����� 
		strValidDt = UniConvDateToYYYYMMDD(frm1.txtHBillDt.value, Parent.gDateFormat,"")
	 
		If Trim(strItemCd) = "" Then Exit Function
		 
		' dbo.ufn_s_GetItemInfo (@plant_cd, @item_cd, @bp_cd, @deal_type, @pay_meth, @currency, @valid_dt, @price_flag, @vat_flag, @deposit_flag
		strSelectList = " plant_cd, item_cd, item_nm, spec, unit, hs_cd, vat_type, vat_nm, vat_rate, item_price, deposit_price "
		strFromList = " dbo.ufn_s_GetItemInfo ( " & FilterVar(strPlantCd, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & ",  " & FilterVar(strSoldToParty, "''", "S") & ", " & _
						" " & FilterVar("*", "''", "S") & " ,  " & FilterVar(strPayMeth, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ", " & FilterVar("S", "''", "S") & " ,  " & FilterVar(lgStrVatFlag, "''", "S") & ",  " & FilterVar(lgStrDepositFlag, "''", "S") & ")"
		strWhereList = ""
		 
		Err.Clear
		    
		'ǰ������ Fetch
		If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
			strItemInfo = Split(strRs, Chr(11))
			.Col = C_PlantCd	:	.text = Trim(strItemInfo(1))
			.Col = C_ItemNm		:	.text = Trim(strItemInfo(3))
			.Col = C_ItemSpec	:	.Text = Trim(strItemInfo(4))
			.Col = C_BillUnit	:	.text = Trim(strItemInfo(5))
			.Col = C_BillPrice	:	.text = UNIConvNumPCToCompanyByCurrency(strItemInfo(10),frm1.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")
			.Col = C_DepositPrice	:	.text = UNIConvNumPCToCompanyByCurrency(strItemInfo(11),frm1.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")

			' ���ܸ���ä�� ������ VAT Type�� ��ϵ��� ���� ��� ǰ�� �Ҵ�� VAT ������ �����Ѵ�.
			If lgStrVatFlag = "Y" Then
				.Col = C_VatType	:	.text = Trim(strItemInfo(7))
				.Col = C_VatTypeNm	:	.text = Trim(strItemInfo(8))
				.Col = C_VatRate	:	.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(strItemInfo(9)), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
			End If

			'�ݾ� ���� 
			Call CalcAmt(pvIntRow)
			
			GetItemInfo = True
		Else
			If Err.number = 0 Then
				'Editing�Ѱ�� �ش� ǰ�������� �������� ������ ǰ�� Popup�� Display�Ѵ�.
				GetItemInfo = OpenItem(strItemCd2)
			Else
				MsgBox Err.description, vbObjectError, Parent.gLogoName 
				Err.Clear
			End If
		End if
	End With

End Function


'===========================================
Function ChangeTrackingSetField(ByVal IRow) ' ������ �߰� 
	Dim strTrackingFlag, strPlantCd	, strItemCd

	With frm1.vspdData
		.Row = IRow
		.Col = C_PlantCd       '���� 
		strPlantCd = .text
		.col = C_ItemCd        'ǰ���ڵ� 
		strItemCd =  .text
	
	        If Trim(strItemCd) = "" Then 
         	   .Col  = C_TrackNo
	           .Text = "*"

                   ggoSpread.SSSetProtected C_TrackNo, IRow, IRow
  	           ggoSpread.SSSetProtected C_TrackingNoPopup, IRow, IRow 
        	End If	

	End With
	
	Call CommonQueryRs(" tracking_flg  ", " b_item_by_plant (nolock) ",  " PLANT_CD = " & FilterVar(strPlantCd , "''", "S")  &  "  AND  ITEM_CD = " & FilterVar(strItemCd , "''", "S")   , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	

	strTrackingFlag = Replace(lgF0, chr(11), "") 

	If  strTrackingFlag = "Y" Then
 	    ggoSpread.SpreadUnLock  C_TrackNo, IRow, C_TrackNo, IRow
	    ggoSpread.SSSetRequired C_TrackNo, IRow, IRow
	    ggoSpread.SpreadUnLock  C_TrackingNoPopup, IRow, C_TrackingNoPopup, IRow	    
	Else	 
            frm1.vspdData.Col  = C_TrackNo
	    frm1.vspdData.Text = "*"
	    ggoSpread.SSSetProtected C_TrackNo, IRow, IRow
	    ggoSpread.SSSetProtected C_TrackingNoPopup, IRow, IRow
	End If

End Function

'===========================================
Function OpenItem(ByVal pvStrCode)
	Dim iArrParam(1)
	Dim iStrRet, intCurRow
	Dim iCalledAspName

	OpenItem = False
	  
	If lgBlnOpenPop = True Then Exit Function
	   
	lgBlnOpenPop = True

	iCalledAspName = AskPRAspName("S3112PA2")
		
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "S3112PA2", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	With frm1.vspdData
		iArrParam(0) = pvStrCode						' ǰ���ڵ� 
		.Col = C_PlantCd	:	iArrParam(1) = .text	' ���� 

		iStrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
										"dialogWidth=820px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		lgBlnOpenPop = False

		If iStrRet(0) <> "" Then
			.Col = C_ItemCd		:	.Text = iStrRet(0)
			.Col = C_PlantCd	:	.Text = iStrRet(2)
			OpenItem = GetItemInfo(.Row)
		End If 
	End With
End Function

' Document�ݾ� ��� 
Sub CalcAmt(ByVal pvIntRow)
	Dim iDblBillAmt, iDblBillQty, iDblBillPrice
	Dim iStrVatIncFlag

	With frm1.vspdData
		.Row = pvIntRow

		ggoSpread.source = frm1.vspdData
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		.Col = C_BillQty	: iDblBillQty = UNICDbl(.Text)
		.Col = C_BillPrice	: iDblBillPrice = UNICDbl(.Text)
		
		iDblBillAmt = iDblBillQty * iDblBillPrice
		.Col = C_BillAmt
		If iDblBillAmt <> 0 Then
			.Text = UNIConvNumPCToCompanyByCurrency(iDblBillAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
			iDblBillAmt = UNICDbl(.Text)
		Else
			.Text = 0
		End If
	End With
	
	'0811 SMJ
'	Call SetRetItemFlag(pvIntRow, iDblBillAmt)
	
	' �ڱ��ݾ� ��� 
	Call CalcAmtLoc(pvIntRow)
End Sub

' �ڱ��ݾ� / Vat�ݾ� / VAT �ڱ��ݾ� ��� 
Sub CalcAmtLoc(ByVal pvIntRow)
	Dim iDblBillAmt, iDblBillAmtLoc, iDblVatRate
	Dim iStrVatIncFlag

	With frm1.vspdData
		.Row = pvIntRow
		
		.Col = C_BillAmt : iDblBillAmt = UNICDbl(.Text)
		If iDblBillAmt <> 0 Then
			' �ڱ��ݾװ�� 
			.Col = C_BillLocAmt	: .Text = FncCalcAmtLoc(iDblBillAmt, UNICDbl(frm1.txtXchgRate.Text), frm1.txtXchgOp.value, Parent.gLocRndPolicyNo)
			iDblBillAmtLoc = UNICDbl(.Text)
						
			.Col = C_VatIncFlag	: iStrVatIncFlag = .Text
			.Col = C_VatRate	: iDblVatRate	 = UNICDbl(.Text)

			.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
			.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)
		Else
			.Col = C_BillLocAmt	: .Text = "0"
			.Col = C_VatAmt		: .Text = "0"
			.Col = C_VatLocAmt	: .Text = "0"
		End If
	End With
	
	'0811 SMJ calcamt �� �ִ� setretitemfalg�� calcamtloc()�� �̵���.
	Call SetRetItemFlag(pvIntRow, iDblBillAmt)
	' �ѱݾ� ��� 
	Call CalcTotal("U", pvIntRow)
End Sub

' VAT �ݾ� ��� 
Sub CalcVatAmt(ByVal pvIntRow)
	Dim iDblBillAmt, iDblBillAmtLoc, iDblVatRate
	Dim iStrVatIncFlag

	With frm1.vspdData
		.Row = pvIntRow
		.Col = C_BillAmt	: iDblBillAmt = UNICDbl(.Text)
		.Col = C_BillLocAmt	: iDblBillAmtLoc = UNICDbl(.Text)
							
		.Col = C_VatIncFlag	: iStrVatIncFlag = .Text
		.Col = C_VatRate	: iDblVatRate	 = UNICDbl(.Text)
							
		.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
		.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)
	End With

	' �ѱݾ� ��� 
	Call CalcTotal("U", pvIntRow)
End Sub

' ���հ�ݾ��� �����Ѵ�.
Sub CalcTotal(ByVal pvStrFlag, ByVal pvIntRow)
	On Error Resume Next
	
	Dim iLngRow, iLngFirstRow, iLngLastRow
	Dim iDblBillAmt, iDblVatAmt, iDblOldBillAmt, iDblOldVatAmt, iDblDiffNetAmt, iDblDiffVatAmt
	Dim iStrBillAmt, iStrVatAmt, iStrVatIncFlag
	
	With frm1.vspdData
		Select Case pvStrFlag
			' �߰�/���� 
			Case "U"
				.Row = pvIntRow
				.Col = C_OldBillAmt	: iDblOldBillAmt = UNICDbl(.Text)
				.Col = C_OldVatAmt	: iDblOldVatAmt = UNICDbl(.Text)
				
				.Col = C_VatAmt		: iStrVatAmt = .Text	:	iDblDiffVatAmt = UNICDbl(.Text) - iDblOldVatAmt
				
				.Col = C_VatIncFlag	: iStrVatIncFlag = .Text
				If iStrVatIncFlag = "1" Then
					.Col = C_BillAmt	: iStrBillAmt = .Text
					.Col = C_OldVatIncFlag
					If .Text = "1" Then
						' �ݾ��� ����� ��� 
						iDblDiffNetAmt = UNICDbl(iStrBillAmt) - iDblOldBillAmt
					Else
						' VAT ���Կ��ΰ� ����� ��� 
						iDblDiffNetAmt = iDblOldVatAmt
					End If
				Else
					.Col = C_BillAmt	: iStrBillAmt = .Text
					.Col = C_OldVatIncFlag
					If .Text = "1" Then
						' VAT���Կ��ΰ� ����� ��� 
						.Col = C_VatAmt	: iDblDiffNetAmt = -UNICDbl(.Text)
					Else
						' �ݾ��� ����� ��� 
						iDblDiffNetAmt = UNICDbl(iStrBillAmt) - iDblOldBillAmt - iDblDiffVatAmt
					End If
				End If

				' ������ �� ���� 
				.Col = C_OldBillAmt		:	.Text = iStrBillAmt
				.Col = C_OldVatIncFlag	:	.Text = iStrVatIncFlag
				.Col = C_OldVatAmt		:	.Text = iStrVatAmt

			' ��� 
			Case "C"
				ggoSpread.Source = frm1.vspdData 
	
				.Row = pvIntRow
				.Col = C_OldBillAmt	: iDblOldBillAmt = UNICDbl(.Text)
				.Col = C_OldVatAmt	: iDblOldVatAmt = UNICDbl(.Text)
				.Col = 0
				Select Case	.Text
					Case ggoSpread.InsertFlag
						.Col = C_VatIncFlag
						If .Text = "1" Then
							iDblDiffNetAmt = - iDblOldBillAmt
						Else
							iDblDiffNetAmt = -(iDblOldBillAmt - iDblOldVatAmt)
						End If
						
						iDblDiffVatAmt = - iDblOldVatAmt
'					    ggoSpread.EditUndo
						    
					Case ggoSpread.UpdateFlag
'					    ggoSpread.EditUndo
					    .Col = C_InitialVatAmt		: iDblDiffVatAmt = UNICDbl(.Text) - iDblOldVatAmt
						.Col = C_VatIncFlag
						If .Text = "1" Then
							.Col = C_InitialBillAmt	: iDblDiffNetAmt = UNICDbl(.Text) - iDblOldBillAmt
						Else
							.Col = C_InitialBillAmt	: iDblDiffNetAmt = UNICDbl(.Text) - iDblOldBillAmt - iDblDiffVatAmt
						End If

					Case ggoSpread.DeleteFlag
'					    ggoSpread.EditUndo
					    .Col = C_InitialVatAmt		: iDblDiffVatAmt = UNICDbl(.Text)
					    
						.Col = C_VatIncFlag
						If .Text = "1" Then
							.Col = C_InitialBillAmt	: iDblDiffNetAmt = UNICDbl(.Text)
						Else
							.Col = C_InitialBillAmt	: iDblDiffNetAmt = UNICDbl(.Text) - iDblDiffVatAmt
						End If
						
				End Select

			' ���� 
			Case "D"
				ggoSpread.Source = frm1.vspdData 
				iLngFirstRow = .SelBlockRow
				If iLngFirstRow = -1 Then
					iLngFirstRow = 1
					iLngLastRow = .MaxRows
				Else
					iLngLastRow = .SelBlockRow2
				End If
						
				For	iLngRow = iLngFirstRow To iLngLastRow
					.Row = iLngRow
					.Col = 0
					If .Text <> ggoSpread.DeleteFlag And .Text <> ggoSpread.InsertFlag Then
						.Col = C_BillAmt	: iDblBillAmt = UNICDbl(.Text)
						.Col = C_VatAmt		: iDblVatAmt = UNICDbl(.Text)
						
						.Col = C_VatIncFlag
						If .Text = "1" Then 
							iDblDiffNetAmt = iDblDiffNetAmt - iDblBillAmt
						Else
							iDblDiffNetAmt = iDblDiffNetAmt - iDblBillAmt + iDblVatAmt
						End If
						
						iDblDiffVatAmt = iDblDiffVatAmt - iDblVatAmt
					End If
				Next
				
		End Select
	End With
	
	Call SetTotal(iDblDiffNetAmt, iDblDiffVatAmt)
End Sub

Sub SetTotal(ByVal pvDblNetAmt, ByVal pvDblVatAmt)
	Dim iDblTotNetAmt, iDblTotVatAmt
	With frm1
		iDblTotNetAmt = UNICDbl(.txtOriginBillAmt.Text) + pvDblNetAmt
		iDblTotVatAmt = UNICDbl(.txtVatAmt.Text) + pvDblVatAmt

		.txtOriginBillAmt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotNetAmt,.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
		.txtVatAmt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotVatAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo  , "X")
	End With
End Sub

'--------------------------- The begin of the test scripts
' �ڱ��ݾ��� ����Ѵ�.
' pvDblAmt : Document�ݾ� - Double�� 
' pvDblXchgRate : ȯ�� - Double�� 
' pvStrXchgRateOp : ȯ�������� 
' ���ǻ��� : ȯ�������ڰ� �Էµ��� ������ ���������� ó���Ѵ�.
' �Լ��� Return ���� Formató���� �����̴�.
Function FncCalcAmtLoc( ByVal pvDblAmt, _
						ByVal pvDblXchgRate, _
						ByVal pvStrXchgRateOp, _
						ByVal pvStrRndPolicyNo)
    Dim iDblAmtLoc
    
    If pvStrXchgRateOp = "*" Then
        iDblAmtLoc = pvDblAmt * pvDblXchgRate
    Else
        iDblAmtLoc = pvDblAmt / pvDblXchgRate
    End If
        
    ' �ڱ��ݾ� ���� ó�� 
    FncCalcAmtLoc = UNIConvNumPCToCompanyByCurrency(iDblAmtLoc,Parent.gCurrency,Parent.ggAmtOfMoneyNo, pvStrRndPolicyNo, "X")
End Function

' �ΰ��� �ݾ��� ����Ѵ�.
Function FncCalcVatAmt(ByVal pvDblAmt, _
					ByVal pvStrVatIncFlag, _
					ByVal pvDblVatRate, _
					ByVal pvStrCurrency)
	Dim iDblVatAmt
	
    ' �ΰ����� ������ ��� 
    If pvStrVatIncFlag = "1" Then
        iDblVatAmt = pvDblAmt * pvDblVatRate / 100
    Else
        iDblVatAmt = pvDblAmt * pvDblVatRate / (100 + pvDblVatRate)
    End If
    
	FncCalcVatAmt = UNIConvNumPCToCompanyByCurrency(iDblVatAmt, pvStrCurrency, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo  , "X")

End Function

' Function Desc : �ݾ�/�ܰ� ����� ��ǰ flag ���� 
'===========================================
Sub SetRetItemFlag(ByVal pvIntRow, ByVal pvDblBillAmt)
	With frm1
		.vspdData.Row = pvIntRow
		.vspdData.Col = C_RetItemFlag
		
		If pvDblBillAmt >= 0 Then
			.vspdData.Text = "N"
		Else
			.vspdData.Text = "Y"
		End If
	End With
End Sub

' Description : ǰ��/���������� �ܰ� �ڵ� Pad
'===========================================
Function GetItemPrice(IRow)
	Dim strSoldToParty, strItemCd, strBillUnit, strPayMeth, strCurrency, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, strItemInfo

	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_ItemCd      'ǰ���ڵ� 
		strItemCd = .vspdData.text
		.vspdData.Col = C_BillUnit      '���� 
		strBillUnit = .vspdData.text

		strSoldToParty = .txtSoldToParty.value  '�ֹ�ó 
		strPayMeth = .txtPayTermsCd.value    '������� 
		strCurrency = .txtCurrency.value    'ȭ����� 
		strValidDt = UniConvDateToYYYYMMDD(.txtHBillDt.value, Parent.gDateFormat,"")
	End With

	If Len(Trim(strItemCd)) = 0 Or Len(Trim(strBillUnit)) = 0 Then Exit Function

	'dbo.ufn_s_GetItemSalesPrice( @bp_cd, @item_cd, @deal_type, @pay_meth, @sales_unit, @currency, @valid_dt)
	strSelectList = " dbo.ufn_s_GetItemSalesPrice( " & FilterVar(strSoldToParty, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & ", Default,  " & FilterVar(strPayMeth, "''", "S") & "," & _
													" " & FilterVar(strBillUnit, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ")"
	strFromList  = ""
	strWhereList = ""
	 
	Err.Clear
	    
	'ǰ������ �ܰ� Fetch
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		strItemInfo = Split(strRs, Chr(11))
		frm1.vspdData.Col = C_BillPrice
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(strItemInfo(1),frm1.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")
		Exit Function
	Else
		If Err.number <> 0 Then
		MsgBox Err.description 
		Err.Clear 
		Exit Function
		End If
	End if
End Function

' Description : ǰ��/������ ������ �ܰ� Fetch
'===========================================
Function GetDepositPrice(IRow)

	If lgstrDepositFlag = "2" Then Exit Function
	 
	Dim strSoldToParty, strItemCd, strBillUnit, strCurrency, strValidDt
	Dim strSelectList, strFromList, strWhereList
	Dim strRs, ldbBillQty
	Dim arrDepositPrice

	With frm1
		.vspdData.Row = IRow
		.vspdData.col = C_ItemCd      'ǰ���ڵ� 
		strItemCd = .vspdData.text
		.vspdData.Col = C_BillUnit      '���� 
		strBillUnit = .vspdData.text

		strSoldToParty = .txtSoldToParty.value  '�ֹ�ó 
		strCurrency = .txtCurrency.value    'ȭ����� 
		strValidDt = UniConvDateToYYYYMMDD(.txtHBillDt.value, Parent.gDateFormat,"")
	End With

	If Len(Trim(strItemCd)) = 0 Or Len(Trim(strBillUnit)) = 0 Then Exit Function

	strSelectList = " dbo.ufn_s_GetDepositPrice( " & FilterVar(strSoldToParty, "''", "S") & ",  " & FilterVar(strItemCd, "''", "S") & "," & _
					" " & FilterVar(strBillUnit, "''", "S") & ",  " & FilterVar(strCurrency, "''", "S") & ",  " & FilterVar(strValidDt, "''", "S") & ",  " & FilterVar(lgstrDepositFlag, "''", "S") & ")"
	strFromList  = ""
	strWhereList = ""
	 
	Err.Clear
	    
	'ǰ������ �ܰ� Fetch
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrDepositPrice = Split(strRs, Chr(11))
		frm1.vspdData.Col = C_DepositPrice
		frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrDepositPrice(1),frm1.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")

		'�����ܰ� ����� �����ݾ� ���� 
		frm1.vspdData.Col = C_BillQty
		ldbBillQty = UNICDbl(frm1.vspdData.Text)
		Call CalcDepositAmt(IRow, ldbBillQty)
		Exit Function
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear 
			Exit Function
		End If
	End if
End Function

' Description : �����ݾ� ��� 
'===========================================
Function CalcDepositAmt(IRow, BillQty)
	Dim DepositPrice
	 
	'��������� �����ݾ� ���� 
	frm1.vspdData.Col = C_DepositPrice : DepositPrice = UNICDbl(frm1.vspdData.Text)
	frm1.vspdData.Col = C_DepositAmt : frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(BillQty * DepositPrice,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
End Function

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
    
    If frm1.txtHPostFlag.Value <> "Y" Then
		Call SetQuerySpreadColor(1) 
    Else
		Call SetPostYesSpreadColor(1)
    End If	
    Call SetSpreadHidden()
End Sub

'========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Row <= 0 Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		If OpenSpreadPopup(Col, Row) Then
			Select Case Col
				' ���� 
				Case C_UnitPopup
					Call GetItemPrice(Row)				' �ܰ� Fetch
					Call GetDepositPrice(Row)			' �����ܰ� Fetch
					Call CalcAmt(Row)
				
				' Vat
				Case C_VatPopup
					CalcVatAmt(Row)
					
			End Select
		End If
		Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
	End With

End Sub

'========================================
Sub vspdData_Click(ByVal Col , ByVal Row )
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
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If
End Sub

'========================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iStrVatAmt
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.Row = Row
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		Call SetRowStatus(Row)

		Select Case Col

			Case C_ItemCd
				Call GetItemInfo(Row) 

			Case C_BillUnit
				Call GetItemPrice(Row)				' �ܰ� Fetch
				Call GetDepositPrice(Row)			' �����ܰ� Fetch
				Call CalcAmt(Row)

			Case C_BillQty

				Call CalcAmt(Row)
				
			Case C_BillPrice
				Call CalcAmt(Row)
			
			Case C_BillAmt
				Call CalcAmtLoc(Row)

			Case C_VatType
				Call SetVATTypeForSpread(Row)
				Call CalcVatAmt(Row)

			Case C_VatAmt
				.Row = Row
				.Col = C_VatAmt	: iStrVatAmt = .Text

				'Document Currency�� Local Currency�� ������ ��� Vat Amount, Vat Amount Local�� �����ϰ� ���� 
				If UCase(Parent.gCurrency) = UCase(Trim(frm1.txtCurrency.value)) Then
					.Col = C_VatLocAmt	:	.Text = iStrVatAmt
				'Document Currency�� Local Currency�� �ٸ� ��� Vat Amount Local �ٽ� ��� 
				Else
					.Col = C_BillAmt
					If UNICDbl(.Text) = 0 Then
						.Col = C_VatLocAmt	:	.Text = FncCalcAmtLoc(UNICDbl(iStrVatAmt), UNICDbl(frm1.txtXchgRate.Text), frm1.txtXchgOp.value, Parent.gTaxRndPolicyNo)
					End If
				End if
				
				Call CalcTotal("U", Row)
		End Select

	End With

End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End if    

End Sub

'========================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex
	With frm1.vspdData
		.Row = Row
	    .Col = Col
		 intIndex = .Value
	 
		 .Col = C_VatIncFlag
		 .Value = intIndex+1
	End With
	
	' �ΰ��� �ݾ� ����.
	Call SetRowStatus(Row)
	Call CalcVatAmt(Row)
End Sub 

'========================================
Sub btnPostFlag_OnClick()

	If BtnSpreadCheck = False Then Exit Sub
	 
	Dim strVal

	If LayerShowHide(1) = False Then Exit Sub

	strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag         
	strVal = strVal & "&txtHBillNo=" & Trim(frm1.txtHBillNo.value)      
	strVal = strVal & "&txtChangeOrgId=" & Parent.gChangeOrgId
	  
	Call RunMyBizASP(MyBizASP, strVal)            

End Sub

'==========================================
Sub btnGLView_OnClick()
	Dim iArrRet
	Dim iArrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
 	
	If Trim(frm1.txtGLNo.value) <> "" Then
		 iArrParam(0) = Trim(frm1.txtGLNo.value) 'ȸ����ǥ��ȣ 
		 iArrParam(1) = Trim(frm1.txtHBillNo.value) 'Reference��ȣ 
		 
		 if iArrParam(0) = "" THEN Exit Sub
		 
		 iCalledAspName = AskPRAspName("a5120ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If

		 iArrRet = window.showModalDialog(iCalledAspName , Array(window.parent,iArrParam), _
		      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		      
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
	     iArrParam(0) = Trim(frm1.txtTempGLNo.value) '������ǥ��ȣ 
	     iArrParam(1) = Trim(frm1.txtHBillNo.value) 'Reference��ȣ 
	 
	     if iArrParam(0) = "" THEN Exit Sub
	     
	     iCalledAspName = AskPRAspName("a5130ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If
		 
	     iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else 
	     Call DisplayMsgBox("205154", "X", "X", "X")
	End If 
	     lblnWinEvent = False
End Sub

'==========================================
Sub btnPreRcptView_OnClick()
     Dim iArrRet
     Dim iArrParam(4)
	 Dim iCalledAspName
	 Dim IntRetCD
	 Dim lblnWinEvent
 
     iArrParam(0) = Trim(frm1.txtHBillDt.Value)    '����ä���� 
     iArrParam(1) = Trim(frm1.txtSoldToParty.value)  '�ֹ�ó 
     iArrParam(2) = Trim(frm1.txtSoldToPartyNm.value)  '�ֹ�ó 
     iArrParam(3) = Trim(frm1.txtCurrency.value)   'ȭ�� 
     iArrParam(4) = ""         '�����ݹ�ȣ 
 
     iCalledAspName = AskPRAspName("s5111ra7")
		 
	 If Trim(iCalledAspName) = "" Then
	      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111ra7", "X")
		  lblnWinEvent = False
		       Exit Sub
		  End If
     iArrRet = window.showModalDialog(iCalledAspName & "?txtFlag=BH&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,iArrParam), _
     "dialogWidth=860px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
     lblnWinEvent = False	
End Sub

'========================================
Sub CurFormatNumericOCX()
	With frm1
		'����ä�Ǳݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtOriginBillAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		'VAT�ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With
End Sub

'========================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		'����ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_BillPrice,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'����ä�Ǳݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_BillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'VAT�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec 
		'�����ܰ� 
		ggoSpread.SSSetFloatByCellOfCur C_DepositPrice,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'�����ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_DepositAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec

		' ������ ����ä�Ǳݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_OldBillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		' ������ VAT�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_OldVatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec 

		' ����ä�Ǳݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_InitialBillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		' VAT�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_InitialVatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec 

	End With

End Sub

