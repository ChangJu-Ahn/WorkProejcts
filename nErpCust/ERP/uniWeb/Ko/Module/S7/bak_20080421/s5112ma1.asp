<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5112MA1
'*  4. Program Name         : ����ä�ǳ������ 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr
'*  7. Modified date(First) : 2002/11/12
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : AHN TAE HEE
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd ȭ�� Layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� Layout
'*                            -2001/12/18 : Date ǥ������ 
'*							  -2002/06/26 : VB conversion
'*							  -2002/11/12 : UI���� ����		
'**********************************************************************************************%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/JpQuery.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID = "s5112mb1.asp"            
Const BIZ_BillHdr_JUMP_ID = "s5111ma1"           
Const BIZ_BillCollect_JUMP_ID = "s5114ma1"

' Constant variables defined
'========================================
Const PostFlag = "PostFlag"

' Common variables 
'========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows

Dim lgSortKey

' Variables For spreadsheet
'========================================
'��: Spread Sheet�� Column
Dim C_ItemCd           'ǰ�� 
Dim C_ItemNm           'ǰ��� 
Dim C_TrackNo		   'Tracking No 
Dim C_BillQty          '���� 
Dim C_BillUnit         '���� 
Dim C_VatIncFlag       '�ΰ������Կ��� 
Dim C_VatIncFlagNm	   '�ΰ������Կ��θ� 
Dim C_BillPrice        '�ܰ� 
Dim C_BillAmt          '�ݾ� 

'�߰� 
Dim C_VatType          'vatŸ�� 
Dim C_VatPopup         'vat�˾� 
Dim C_VatNm            'vat�� 
Dim C_VatRate          'vat�� 

Dim C_VatAmt           'VAT�ݾ� 
Dim C_BillLocAmt       '��ȭ�ݾ� 
Dim C_VatLocAmt        'VAT��ȭ�ݾ� 
Dim C_DepositPrice     '�����ܰ� 
Dim C_DepositAmt       '�����ݾ� 
Dim C_FOBAmt           'FOB�ݾ� 
Dim C_Remark           '��� 
Dim C_DnNo             '���Ϲ�ȣ 
Dim C_DnSeqNo          '���ϼ��� 
Dim C_SoNo             '���ֹ�ȣ 
Dim C_SoSeqNo          '���ּ��� 
Dim C_LlcNo            'Local L/C��ȣ 
Dim C_LlcSeq           'Local L/C���� 
Dim C_BillSeq          '������� 
Dim C_PlantCd          '�����ڵ� 
Dim C_ItemSpec         'ǰ��԰� 
Dim C_RetItemFlag      '��ǰ���� 
Dim C_OldBillAmt
Dim C_OldVatIncFlag
Dim C_OldVatAmt
Dim C_InitialBillAmt
Dim C_InitialVatAmt

' User-defind Variables
'========================================
Dim IsOpenPop      ' Popup

Dim arrCollectVatType

'========================================
Sub initSpreadPosVariables()  
	
	C_ItemCd		= 1    'ǰ�� 
	C_ItemNm		= 2    'ǰ��� 
	C_TrackNo		= 3    'Tracking No 
	C_BillQty		= 4    '���� 
	C_BillUnit		= 5    '���� 
	C_VatIncFlag	= 6    '�ΰ������Կ��� 
	C_VatIncFlagNm	= 7    '�ΰ������Կ��θ� 
	C_BillPrice		= 8    '�ܰ� 
	C_BillAmt		= 9    '�ݾ� 
	'�߰� 
	C_VatType		= 10    'vatŸ�� 
	C_VatPopup		= 11    'vat�˾� 
	C_VatNm			= 12   'vat�� 
	C_VatRate		= 13   'vat�� 
	C_VatAmt		= 14    'VAT�ݾ� 
	C_BillLocAmt	= 15    '��ȭ�ݾ� 
	C_VatLocAmt		= 16    'VAT��ȭ�ݾ� 
	C_DepositPrice	= 17    '�����ݴܰ� 
	C_DepositAmt	= 18    '�����ݾ� 
	C_FOBAmt		= 19    'FOB�ݾ� 
	C_Remark		= 20    '��� 
	C_DnNo			= 21    '���Ϲ�ȣ 
	C_DnSeqNo		= 22    '���ϼ��� 
	C_SoNo			= 23    '���ֹ�ȣ 
	C_SoSeqNo		= 24    '���ּ��� 
	C_LlcNo			= 25    'Local L/C��ȣ 
	C_LlcSeq		= 26    'Local L/C���� 
	C_BillSeq		= 27    '������� 
	C_PlantCd		= 28    '�����ڵ� 
	C_ItemSpec		= 29    'ǰ��԰� 
	C_RetItemFlag   = 30    '��ǰ���� 

	'Total �ݾ� ������ ������ ���� �߰� 
	C_OldBillAmt	= 31
	C_OldVatIncFlag = 32
	C_OldVatAmt		= 33
	C_InitialBillAmt= 34
	C_InitialVatAmt = 35

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
		Set gActiveElement = document.activeElement 

		.btnPostFlag.disabled = True
		.btnPostFlag.value = "Ȯ��"
		.btnGLView.disabled = True
		.btnPreRcptView.disabled = True
	End With

	lgBlnFlgChgValue = False
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With ggoSpread
		
		.Source = frm1.vspdData
		'patch version
		.Spreadinit "V20031001",,parent.gAllowDragDropSpread    
		frm1.vspdData.ReDraw = False
		
	    frm1.vspdData.MaxRows = 0 : frm1.vspdData.MaxCols = 0	     
	    frm1.vspdData.MaxCols = C_InitialVatAmt + 1           '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    
        Call GetSpreadColumnPos("A")
	 
	    .SSSetEdit C_ItemCd, "ǰ��", 18,,,18,2
	    .SSSetEdit C_ItemNm, "ǰ���", 30
	    .SSSetEdit C_TrackNo, "Tracking No", 18,,,25,2
	    .SSSetFloat C_BillQty,"����" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    .SSSetEdit C_BillUnit, "����", 8,,,3,2
	    .SSSetFloat C_BillPrice,"�ܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetFloat C_BillAmt,"�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		 
	    '�߰� 
	    .SSSetEdit  C_VatType, "VAT����", 10,2,,4,2
	    .SSSetButton  C_VatPopup
	    .SSSetEdit  C_VatNm, "VAT������", 20 
	    .SSSetFloat C_VatRate,"VAT��",15,Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		         
	    .SSSetFloat C_VatAmt,"VAT�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetFloat C_BillLocAmt,"�ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetFloat C_VatLocAmt,"VAT�ڱ��ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		 
	    .SSSetFloat C_DepositPrice,"�����ܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetFloat C_DepositAmt,"�����ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetFloat C_FOBAmt,"FOB�ݾ�",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    .SSSetEdit C_Remark, "���", 30,,,120
	    .SSSetEdit C_DnNo, "���Ϲ�ȣ", 18,,,18,2
	    
	    Call AppendNumberPlace("6","5","0")
	    
	    .SSSetFloat C_DnSeqNo,"���ϼ���" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	    .SSSetEdit C_SoNo, "���ֹ�ȣ", 18,,,18,2
	    .SSSetFloat C_SoSeqNo,"���ּ���" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	    .SSSetEdit C_LlcNo, "LOCAL L/C��ȣ", 18,,,18,2
	    .SSSetFloat C_LlcSeq,"LOCAL L/C����" ,18,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	    .SSSetFloat C_BillSeq,"�������" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
	    .SSSetEdit C_PlantCd, "����", 10,,,4,2
	    .SSSetEdit C_VatIncFlag, "VAT���Ա���", 1,,,1,2
		.SSSetEdit C_VatIncFlagNm, "VAT���Ա���", 15,2
	    .SSSetEdit C_ItemSpec, "ǰ��԰�", 30,,,50,2
	    .SSSetEdit C_RetItemFlag, "��ǰ����", 10,2,,1,2

		.SSSetFloat C_OldBillAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetEdit  C_OldVatIncFlag, "", 10, 2,, 1, 2
		.SSSetFloat C_OldVatAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    
		.SSSetFloat C_InitialBillAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		.SSSetFloat C_InitialVatAmt,"",15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	    Call .MakePairsColumn(C_VatType,C_VatPopup)
	    
	    Call .SSSetColHidden(C_BillSeq,C_BillSeq,True)
	    Call .SSSetColHidden(C_PlantCd,C_PlantCd,True)
	    Call .SSSetColHidden(C_FOBAmt,C_FOBAmt,True)
	    Call .SSSetColHidden(C_VatIncFlag,C_VatIncFlag,True)
	    Call .SSSetColHidden(C_OldBillAmt, frm1.vspdData.MaxCols, True)				'��: ������Ʈ�� ��� Hidden Column

	    Call .SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)				'��: ������Ʈ�� ��� Hidden Column
	 
		frm1.vspdData.ReDraw = True
   
    End With
    
End Sub

'==========================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	With frm1
		ggoSpread.SSSetProtected C_ItemCd		, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_ItemNm		, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_TrackNo		, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_BillQty		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BillUnit		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_VatIncFlagNm	, pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired  C_BillPrice	, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_BillAmt		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DepositPrice	, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_DepositAmt	, pvStartRow, pvEndRow    
		 
		if .rdoVatCalcType2.checked = True then
			ggoSpread.SpreadUnLock  C_VatType	, pvStartRow, pvEndRow
		elseif .rdoVatCalcType1.checked = True then
			ggoSpread.SSSetRequired  C_VatType	, pvStartRow, pvEndRow
		end if
		ggoSpread.SpreadUnLock    C_VatPopup	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_VatNm		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected  C_VatRate		, pvStartRow, pvEndRow

		ggoSpread.SSSetRequired  C_VatAmt		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_FOBAmt		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DnNo			, pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected C_DnSeqNo		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoNo			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_SoSeqNo		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LlcNo		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LlcSeq		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RetItemFlag	, pvStartRow, pvEndRow

		If UCase(Parent.gCurrency) <> UCase(Trim(frm1.txtCurrency.value)) Then
			ggoSpread.SSSetRequired  C_BillLocAmt, pvStartRow, pvEndRow
			'goSpread.SSSetProtected C_VatLocAmt	, pvStartRow, pvEndRow
		End If
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
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
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
			C_ItemCd          = iCurColumnPos(1)
			C_ItemNm          = iCurColumnPos(2)
			C_TrackNo		  = iCurColumnPos(3)    
			C_BillQty		  = iCurColumnPos(4)
			C_BillUnit		  = iCurColumnPos(5)
			C_VatIncFlag      = iCurColumnPos(6)
			C_VatIncFlagNm    = iCurColumnPos(7)
			C_BillPrice       = iCurColumnPos(8)
			C_BillAmt		  = iCurColumnPos(9)
			C_VatType         = iCurColumnPos(10)
			C_VatPopup        = iCurColumnPos(11)
			C_VatNm			  = iCurColumnPos(12)
			C_VatRate         = iCurColumnPos(13)
			C_VatAmt		  = iCurColumnPos(14)
			C_BillLocAmt      = iCurColumnPos(15)
			C_VatLocAmt       = iCurColumnPos(16)
			C_DepositPrice    = iCurColumnPos(17)
			C_DepositAmt      = iCurColumnPos(18)
			C_FOBAmt		  = iCurColumnPos(19)    
			C_Remark		  = iCurColumnPos(20)
			C_DnNo			  = iCurColumnPos(21)
			C_DnSeqNo         = iCurColumnPos(22)
			C_SoNo			  = iCurColumnPos(23)
			C_SoSeqNo         = iCurColumnPos(24)
			C_LlcNo			  = iCurColumnPos(25)
			C_LlcSeq		  = iCurColumnPos(26)
			C_BillSeq		  = iCurColumnPos(27)
			C_PlantCd		  = iCurColumnPos(28)
			C_ItemSpec		  = iCurColumnPos(29)
			C_RetItemFlag     = iCurColumnPos(30)
			C_OldBillAmt 	  = iCurColumnPos(31)
			C_OldVatIncFlag	  = iCurColumnPos(32)
			C_OldVatAmt 	  = iCurColumnPos(33)
			C_InitialBillAmt  = iCurColumnPos(34)
			C_InitialVatAmt   = iCurColumnPos(35)
    End Select    
End Sub

'==========================================
Sub SetPostYesSpreadColor(ByVal lRow)
	With frm1

		Call SetToolbar("11100000000111")
		    
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected C_ItemCd, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_ItemNm, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_TrackNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillQty, lRow, .vspdData.MaxRows 
		ggoSpread.SSSetProtected C_BillUnit, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatIncFlagNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillPrice, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillAmt, lRow, .vspdData.MaxRows

		'�߰� 
		ggoSpread.SSSetProtected C_VatType, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatPopup, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatRate, lRow, .vspdData.MaxRows

		ggoSpread.SSSetProtected C_VatAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillLocAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatLocAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DepositPrice, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DepositAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_FOBAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DnNo, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_DnSeqNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoSeqNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LlcNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LlcSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_Remark, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemSpec, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_RetItemFlag, lRow, .vspdData.MaxRows
		
		.vspdData.ReDraw = True
	    
	End With
End Sub

'==========================================
Sub SetQuerySpreadColor(ByVal lRow)
	With frm1
	    
		.vspdData.ReDraw = False
	
		ggoSpread.SSSetProtected C_ItemCd, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_ItemNm, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_TrackNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetRequired  C_BillQty, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_BillUnit, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatIncFlagNm, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetRequired  C_BillPrice, lRow, .vspdData.MaxRows
		ggoSpread.SSSetRequired  C_BillAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DepositPrice, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_DepositAmt, lRow, .vspdData.MaxRows    
		
		If .rdoVatCalcType2.checked = True then
			ggoSpread.SpreadUnLock  C_VatType, lRow, .vspdData.MaxRows
		ElseIf .rdoVatCalcType1.checked = True Then
			ggoSpread.SSSetRequired  C_VatType, lRow, .vspdData.MaxRows
		End If
		  
		ggoSpread.SpreadUnLock   C_VatPopup, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected  C_VatNm, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_VatRate, lRow, .vspdData.MaxRows
		  
		ggoSpread.SSSetRequired  C_VatAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_FOBAmt, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_DnNo, lRow, .vspdData.MaxRows    
		ggoSpread.SSSetProtected C_DnSeqNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_SoSeqNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LlcNo, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_LlcSeq, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_ItemSpec, lRow, .vspdData.MaxRows
		ggoSpread.SSSetProtected C_RetItemFlag, lRow, .vspdData.MaxRows

		If UCase(Parent.gCurrency) <> UCase(Trim(frm1.txtCurrency.value)) Then
			ggoSpread.SSSetRequired  C_BillLocAmt, lRow, .vspdData.MaxRows
			' ggoSpread.SSSetProtected C_VatLocAmt, lRow, .vspdData.MaxRows
		Else
		End If
		
		.vspdData.ReDraw = True
	    
	End With
End Sub

'==========================================
Sub SetSpreadHidden()
	With frm1
		
		' ������ ��� VAT ������ Hiddenó�� 
		If .rdoVatCalcType2.checked = True then
			Call ggoSpread.SSSetColHidden(C_VatIncFlagNm,C_VatIncFlagNm,True)
			Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
			Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,True)
			Call ggoSpread.SSSetColHidden(C_VatRate,C_VatRate,True)
		ElseIf .rdoVatCalcType1.checked = True Then
			Call ggoSpread.SSSetColHidden(C_VatIncFlagNm,C_VatIncFlagNm,False)
			Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,False)
			Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,False)
			Call ggoSpread.SSSetColHidden(C_VatRate,C_VatRate,False)
		End If
		
		If UCase(Parent.gCurrency) <> UCase(Trim(frm1.txtCurrency.value)) Then
			Call ggoSpread.SSSetColHidden(C_BillLocAmt,C_BillLocAmt,False)
			Call ggoSpread.SSSetColHidden(C_VatLocAmt,C_VatLocAmt,False)
		Else
			Call ggoSpread.SSSetColHidden(C_BillLocAmt,C_BillLocAmt,True)
			Call ggoSpread.SSSetColHidden(C_VatLocAmt,C_VatLocAmt,True)
		End If
		
	End With
End Sub		

'==========================================
Sub InitComboBox()
End Sub

'==========================================
Sub InitData()
End Sub

'==========================================
Function CookiePage(Byval Kubun)
	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim strTemp, arrVal

	If Kubun = 1 Then
		WriteCookie CookieSplit , frm1.txtHBillNo.value
	ElseIf Kubun = 0 Then
		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
		arrVal = Split(strTemp, Parent.gRowSep)
		If arrVal(0) = "" Then Exit Function
		frm1.txtConBillNo.value =  arrVal(0)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()
		WriteCookie CookieSplit , ""
	End If
End Function

'==========================================
Function JumpChgCheck(ByVal pvStrJumpPgmId)

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
Function RefCheckMessage(strRefFlag)

 RefCheckMessage = False

 If frm1.HPostFlag.value = "Y" Then
  Msgbox "�̹� Ȯ���� �Ǿ ���� �� �� �����ϴ�",vbInformation, Parent.gLogoName
  Exit Function
 End If

 If strRefFlag <> Trim(frm1.txtRefFlag.value) Then
  Select Case Trim(frm1.txtRefFlag.value)
  Case "L"
   Call DisplayMsgBox("209002", "X", "L/C", "L/C��������")
   Exit Function
  Case "S"
   Call DisplayMsgBox("209002", "X", "����", "���ֳ�������")
   Exit Function
  Case "D"
   Call DisplayMsgBox("209002", "X", "����", "���ϳ�������")
   Exit Function
  End Select
 End If

 RefCheckMessage = True

End Function

Function JungBokMsg(strJungBok,strID)

 Dim strJugBokMsg

 If Len(Trim(strJungBok)) Then strJungBok = strID & Chr(13) & String(30,"=") & strJungBok
 If Len(Trim(strJungBok)) Then strJugBokMsg = strJungBok & Chr(13) & Chr(13)
 If Len(Trim(strJugBokMsg)) Then
  strJugBokMsg = strJugBokMsg & "�̹� ������ ��ȣ�� ������ �����մϴ�"
  MsgBox strJugBokMsg, vbInformation, Parent.gLogoName
 End If

End Function

Sub LockFieldInit()
    Call FormatDoubleSingleField(frm1.txtXchgRate)
    Call LockObjectField(frm1.txtXchgRate,"P")

    Call FormatDoubleSingleField(frm1.txtVatAmt)
    Call LockObjectField(frm1.txtVatAmt,"P")

    Call FormatDoubleSingleField(frm1.txtOriginBillAmt)
    Call LockObjectField(frm1.txtOriginBillAmt ,"P")

    Call FormatDoubleSingleField(frm1.txtTotBillAmt)
    Call LockObjectField(frm1.txtTotBillAmt ,"P")
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
Function FncQuery() 
    Dim IntRetCD 
    
    Err.Clear                                                               

    FncQuery = False                                                        
    
'    If Not chkField(Document, "1") Then Exit Function
    If Not chkFieldByCell(frm1.txtConBillNo, "A", 1) Then Exit Function 

	 ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
	End If
    
    Call ggoOper.ClearField(Document, "2")          
    Call InitVariables

    Call DbQuery                

    FncQuery = True                
    
    Set gActiveElement = document.activeElement 
        
End Function

'========================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then  Exit Function
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

    CAll DbSave														                <%'��: Save db data%>
    
    FncSave = True													            
    
End Function

'========================================
Function FncCancel() 
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function

	Call CalcTotal("C", frm1.vspdData.ActiveRow)

	ggoSpread.Source = frm1.vspdData 
	ggoSpread.EditUndo  
'	Call RefBillHdrSum("DL")
End Function

'========================================
Function FncDeleteRow() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

		.vspdData.focus
		ggoSpread.Source = .vspdData 

	 	Call CalcTotal("D", 0)
    
		lDelRows = ggoSpread.DeleteRow
 
		lgBlnFlgChgValue = True
    
'		Call RefBillHdrSum("DL")
    End With
    
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
    
    DbQuery = True                 

    End Function

'========================================
Function DbQueryOk()              
 
    lgIntFlgMode = Parent.OPMD_UMODE            
	lgBlnFlgChgValue = False
    lgIntGrpCount = 0              
  
    Call SetToolbar("11101011000111")        

	If UNICDbl(frm1.txtSts.value) < 3 Then
	 frm1.btnPostFlag.disabled = False
	Else
	 frm1.btnPostFlag.disabled = True
	End If
	
	frm1.vspdData.focus

End Function

'========================================
Function DbSave() 

    Err.Clear                
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim iVat,iVat_Loc                '�ΰ����� 
 
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
    
  '-----------------------
  'Data manipulate area
  '-----------------------
  For lRow = 1 To .vspdData.MaxRows
    
      .vspdData.Row = lRow
      .vspdData.Col = 0

      Select Case .vspdData.Text
          Case ggoSpread.InsertFlag       '��: �ű� 
     strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep'��: C=Create
          Case ggoSpread.UpdateFlag       '��: ���� 
     strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep'��: U=Update
    Case ggoSpread.DeleteFlag       '��: ���� 
     strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep'��: D=Delete
     '--- ������� 
              .vspdData.Col = C_BillSeq 
              strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gRowSep

              lGrpCnt = lGrpCnt + 1 
   End Select

   Select Case .vspdData.Text
    case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
     
     '�ΰ����� ���� 
     .vspdData.Col = C_VatAmt 
     iVat = Trim(.vspdData.Text)
     'Local�ΰ����� ���� 
     .vspdData.Col = C_VatLocAmt
     iVat_Loc = Trim(.vspdData.Text)
          
     '--- ������� 
              .vspdData.Col = C_BillSeq 
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- ǰ�� 
              .vspdData.Col = C_ItemCd               
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- ���� 
              .vspdData.Col = C_BillQty   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- ���� 
              .vspdData.Col = C_BillUnit   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- �ܰ� 
              .vspdData.Col = C_BillPrice   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     
     'VAT ���Կ��ο� ���� �ݾװ�� 
              .vspdData.Col = C_VatIncFlag
              If Trim(.vspdData.Text)  = "1" Then
      '--- �ݾ� 
      .vspdData.Col = C_BillAmt   
      strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
              Else 
      .vspdData.Col = C_BillAmt 
            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) - UNIConvNum(iVat,0) & Parent.gColSep
     End If
              
     '�߰� 
     '---vatŸ�� 
     .vspdData.Col = C_VatType   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '---vat�� 
     .vspdData.Col = C_VatRate   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- VAT�ݾ� 
              .vspdData.Col = C_VatAmt   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- ��� 
              .vspdData.Col = C_Remark   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- ���Ϲ�ȣ 
              .vspdData.Col = C_DnNo   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- ���ϼ��� 
              .vspdData.Col = C_DnSeqNo   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep              
     '--- ���ֹ�ȣ 
              .vspdData.Col = C_SoNo   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- ���ּ��� 
              .vspdData.Col = C_SoSeqNo   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
              '--- L/C��ȣ 
              .vspdData.Col = C_LlcNo   
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
     '--- L/C���� 
              .vspdData.Col = C_LlcSeq  
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- ���� 
              .vspdData.Col = C_PlantCd  
     strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep         
     
     '--- VAT ���Կ��ο� ���� ��ȭ�ݾװ�� 
              .vspdData.Col = C_VatIncFlag
              If Trim(.vspdData.Text)  = "1" Then
      '--- ��ȭ�ݾ� 
      .vspdData.Col = C_BillLocAmt   
      strVal = strVal &UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     Else
      .vspdData.Col = C_BillLocAmt 
            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) - UNIConvNum(iVat_Loc,0) & Parent.gColSep
     End If
         
     '--- VAT��ȭ�ݾ� 
              .vspdData.Col = C_VatLocAmt   
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep

     '--- VAT ���Կ��� 
              .vspdData.Col = C_VatIncFlag
              strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

     '--- �����ܰ� 
              .vspdData.Col = C_DepositPrice
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- �����ݾ� 
              .vspdData.Col = C_DepositAmt
              strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & Parent.gColSep
     '--- ��ǰ���� 
              .vspdData.Col = C_RetItemFlag
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
 frm1.txtConBillNo.value = frm1.txtHBillNo.value
 Call ggoOper.ClearField(Document, "2")
    Call MainQuery()

End Function

'========================================
Sub InitCollectType()
 Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
 Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

 Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", _
                    " B_MINOR Minor,B_CONFIGURATION Config ", _
                    " Minor.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", _
                    lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

 If Err.number <> 0 Then
  MsgBox Err.description 
  Err.Clear 
  Exit Sub
 End If

 Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

 For i = 0 to UBound(iCodeArr) - 1
  arrCollectVatType(i, 0) = iCodeArr(i)
  arrCollectVatType(i, 1) = iNameArr(i)
  arrCollectVatType(i, 2) = iRateArr(i)
 Next
End Sub

'=========================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

 Dim iCnt

 For iCnt = 0 To Ubound(arrCollectVatType)  
  If arrCollectVatType(iCnt, 0) = UCase(VatType) Then
   VatTypeNm = arrCollectVatType(iCnt, 1)
   VatRate   = arrCollectVatType(iCnt, 2)
   Exit Sub
  End If
 Next
 VatTypeNm = ""
 VatRate = ""
End Sub

'=========================================
Sub SetVatType(ByVal Row)
 Dim VatType, VatTypeNm, VatRate
    
' frm1.vspdData.Row = frm1.vspdData.ActiveRow
 frm1.vspdData.Row = Row
    frm1.vspdData.Col = C_VatType

 VatType = Trim(frm1.vspdData.text)
 Call InitCollectType
 Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

    frm1.vspdData.Col = C_VatNm              'vat�� 
 frm1.vspdData.text = VatTypeNm
    
 frm1.vspdData.Col = C_VatRate            'vat�� 
 frm1.vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(VatRate), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")

 lgBlnFlgChgValue = True
End Sub

'===========================================
Function OpenVat()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

' If IsOpenPop = True Or UCase(frm1.txtVattype.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
    If IsOpenPop = True Then Exit Function 

 IsOpenPop = True
 
    frm1.vspdData.Col=C_VatType
 frm1.vspdData.Row=frm1.vspdData.ActiveRow 

 arrParam(0) = "VAT����"    
 arrParam(1) = "B_MINOR,b_configuration" 
 
 arrParam(2) = Trim(frm1.vspdData.Text)  
  
 arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd " 
 arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
 arrParam(5) = "VAT����"     
 
    arrField(0) = "b_minor.MINOR_CD"   
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE" 
    
    arrHeader(0) = "VAT����"     
    arrHeader(1) = "VAT���¸�"    
    arrHeader(2) = "VAT��"
 arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetVat(arrRet)
 End If 
End Function

'===========================================
Function OpenConBillDtl()

	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
	 
	iCalledAspName = AskPRAspName("s5111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	  
	strRet = window.showModalDialog(iCalledAspName & "?txtExceptFlag=N", Array(window.parent), _
	"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	frm1.txtConBillNo.focus
	
	If strRet <> "" Then  frm1.txtConBillNo.value = strRet 

End Function

'===========================================
Function OpenSODtlRef()
	Dim arrRet
	Dim strParam
	
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		'Call MsgBox("��ȸ�� �����Ͻʽÿ�.", Parent.VB_INFORMATION)
		Exit Function
	End IF

	If Not RefCheckMessage("S") Then Exit Function

	'If IsOpenPop Then Exit Function

	'IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112ba1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3112ba1", "X")
		lblnWinEvent = False
		Exit Function
	End If

	With frm1
		strParam = ""
		strParam = strParam & .txtSoNo.value & Parent.gColSep & .txtSoldtoParty.value & Parent.gColSep & .txtSoldtoPartyNm.value & Parent.gColSep
		strParam = strParam & .txtHSalesGrpCd.value & Parent.gColSep & .txtHSalesGrpNm.value & Parent.gColSep & .txtPayTermsCd.value & Parent.gColSep & .txtPayTermsNm.value & Parent.gColSep
		strParam = strParam & .txtCurrency.value & Parent.gColSep & .txtHBillDt.value & Parent.gColSep

		'vat��,����,�������, �ΰ��� ���Ա��� 
		if frm1.rdoVatCalcType1.checked then 
			strParam = strParam & .HVatRate.value & Parent.gColSep & "%" & Parent.gColSep & "1" & Parent.gColSep & "%" & Parent.gColSep
		else
			'�ΰ������Կ��� 
			if frm1.rdoVatIncFlag1.checked then 
				strParam = strParam & .HVatRate.value & Parent.gColSep & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "1" & Parent.gColSep
			else
				strParam = strParam & .HVatRate.value & Parent.gColSep & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "2" & Parent.gColSep
			end if
		end if
		'ȯ�� 
		strParam = strParam & .txtXchgRate.Text & Parent.gColSep & .txtXchgOp.value & Parent.gColSep & .txtBillTypeCd.value & Parent.gRowSep
		 
		arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & .txtCurrency.value,Array(window.parent,strParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End With

	lblnWinEvent = False

	If arrRet(0,0) = "" Then
		If Err.Number <> 0 Then Err.Clear 
	Else
		Call SetSODtlRef(arrRet)
	End If 

End Function

'===========================================
Function OpenLCDtlRef()
	Dim arrRet
	Dim strParam
	
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		'Call MsgBox("��ȸ�� �����Ͻʽÿ�.", Parent.VB_INFORMATION)
		Exit Function
	End IF

	If Not RefCheckMessage("L") Then Exit Function
	 
	'If IsOpenPop Then Exit Function

	'IsOpenPop = True
	
	iCalledAspName = AskPRAspName("s3112ba2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3112ba2", "X")
		lblnWinEvent = False
		Exit Function
	End If

	With frm1
		strParam = ""
		strParam = strParam & .txtSoNo.value & Parent.gColSep & .txtSoldtoParty.value & Parent.gColSep & .txtSoldtoPartyNm.value & Parent.gColSep
		strParam = strParam & .txtHSalesGrpCd.value & Parent.gColSep & .txtHSalesGrpNm.value & Parent.gColSep & .txtPayTermsCd.value & Parent.gColSep & .txtPayTermsNm.value & Parent.gColSep
		strParam = strParam & .txtCurrency.value & Parent.gColSep & .txtHBillDt.value & Parent.gColSep
		'vat��,����,�������, �ΰ��� ���Ա��� 
		if frm1.rdoVatCalcType1.checked then 
			strParam = strParam& .HVatRate.value & Parent.gColSep & "%" & Parent.gColSep & "1" & Parent.gColSep & "%" & Parent.gColSep
		else
			'�ΰ������Կ��� 
			if frm1.rdoVatIncFlag1.checked then 
				strParam = strParam & .HVatRate.value & Parent.gColSep & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "1" & Parent.gColSep
			else
				strParam = strParam & .HVatRate.value & Parent.gColSep & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "2" & Parent.gColSep
			end if
		end if
		'ȯ�� 
		strParam = strParam & .txtXchgRate.Text & Parent.gColSep & .txtXchgOp.value & Parent.gColSep & .txtBillTypeCd.value & Parent.gRowSep
		arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & frm1.txtCurrency.value,Array(window.parent,strParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End With

	lblnWinEvent = False

	If arrRet(0,0) = "" Then
		If Err.Number <> 0 Then Err.Clear 
	Else
		Call SetLCDtlRef(arrRet)
	End If 

End Function

'===========================================
Function OpenDNDtlRef()
	Dim arrRet
	Dim strParam

	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent

	On Error Resume Next

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		'Call MsgBox("��ȸ�� �����Ͻʽÿ�.", Parent.VB_INFORMATION)
		Exit Function
	End IF

	If Not RefCheckMessage("D") Then Exit Function

	iCalledAspName = AskPRAspName("s3112ba3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3112ba3", "X")
		lblnWinEvent = False
		Exit Function
	End If
	 
	With frm1
		strParam = ""
		strParam = strParam & .txtSoNo.value & Parent.gColSep & .txtSoldtoParty.value & Parent.gColSep & .txtSoldtoPartyNm.value & Parent.gColSep
		strParam = strParam & .txtHSalesGrpCd.value & Parent.gColSep & .txtHSalesGrpNm.value & Parent.gColSep & .txtPayTermsCd.value & Parent.gColSep
		strParam = strParam & .txtPayTermsNm.value & Parent.gColSep & .txtCurrency.value & Parent.gColSep & .txtHBillDt.value & Parent.gColSep
		'vat��,����,�������,ȯ�� 
		strParam = strParam & .HVatRate.value & Parent.gColSep
		'�ΰ��� ������� - ���� 
		if .rdoVatCalcType1.checked then 
			strParam = strParam & "%" & Parent.gColSep & "1" & Parent.gColSep & "%" & Parent.gColSep
		else
			'�ΰ������Կ��� 
			if .rdoVatIncFlag1.checked then 
				strParam = strParam & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "1" & Parent.gColSep
			else
				strParam = strParam & .HVatType.value & Parent.gColSep & "2" & Parent.gColSep & "2" & Parent.gColSep
			end if
		end if
		 
		strParam = strParam & .txtXchgRate.Text & Parent.gColSep & .txtXchgOp.value & Parent.gColSep & .txtBillTypeCd.value & Parent.gRowSep

		arrRet = window.showModalDialog(iCalledAspName & "?txtCurrency=" & .txtCurrency.value,Array(window.parent,strParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End With
	 
	lblnWinEvent = False

	' Popup���� Cancel�� ��� 
	If UBound(arrRet, 2) = 0 Then
		If Err.Number <> 0 Then Err.Clear 
	Else
		Call SetDNDtlRef(arrRet)
	End If 

End Function

'===========================================
Function SetVat(byval arrRet)
	With frm1.vspdData
		.Col = C_VatType
		.Text = arrRet(0)
		.Col = C_VatNm
		.Text = arrRet(1)
		.Col = C_VatRate
		.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(arrRet(2)), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
	End With
    
	'Call BillTotalSum(C_VatAmt,frm1.vspdData.ActiveRow)
	Call vspdData_Change(C_VatType , frm1.vspdData.ActiveRow )
	  
	lgBlnFlgChgValue = True
End Function

'===========================================
Function SetSODtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim iIntStartRow, I, j
	Dim intLoopCnt
	Dim intCnt
	Dim blnEqualFlg
	Dim strSoNo,strSoSeqNo
	Dim intCntRow
	Dim strSOJungBokMsg
	Dim iDblAccuBillAmt, iDblAccuVatAmt

	iDblAccuBillAmt = 0
	iDblAccuVatAmt = 0

	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False 

		iIntStartRow = .MaxRows           <% '��: ��������� MaxRows %>
		intLoopCnt = Ubound(arrRet, 1)          <% '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� %>
		intCntRow = 0

		strSOJungBokMsg = ""
		For intCnt = 1 to intLoopCnt 
			blnEqualFlg = False

			If iIntStartRow <> 0 Then
				strSoNo=""
				strSoSeqNo=""

				<% '---> ���������� ���� ���ֹ�ȣ�� ���ּ����� �ִ��� üũ�Ѵ� %>
				For j = 1 To iIntStartRow
					.Row = j
					<% '���ֹ�ȣ %>
					.Col = C_SoNo		:	strSoNo = .text
					<% '���ּ��� %>
					.Col = C_SoSeqNo    :	strSoSeqNo = .text

					If strSoNo = arrRet(intCnt - 1, 0) And strSoSeqNo = arrRet(intCnt - 1, 9) Then
						blnEqualFlg = True
						strSOJungBokMsg = strSOJungBokMsg & Chr(13) & strSoNo & "-" & strSoSeqNo
						Exit For
					End If

				Next
			End If
		   
			If blnEqualFlg = False then
				intCntRow = intCntRow + 1
				.MaxRows = CLng(iIntStartRow) + CLng(intCntRow)
				.Row = CLng(iIntStartRow) + CLng(intCntRow)

				.Col = 0		:		.Text = ggoSpread.InsertFlag

				<% '���ֹ�ȣ %>
				.Col = C_SoNo		:	.text = arrRet(intCnt - 1, 0)
				<% 'ǰ�� %>
				.Col = C_ItemCd		:	.text = arrRet(intCnt - 1, 1)
				<% 'ǰ��� %>
				.Col = C_ItemNm     :	.text = arrRet(intCnt - 1, 2)
				<% 'Tracking No %>
				.Col = C_TrackNo    :	.text = arrRet(intCnt - 1, 3)
				<% '�̸������ %>	
				.Col = C_BillQty    :	.text = arrRet(intCnt - 1, 4)
				<% '���� %>
				.Col = C_BillUnit   :	.text = arrRet(intCnt - 1, 5)
				<% '�̸���ܰ� %>
				.Col = C_BillPrice  :	.text = arrRet(intCnt - 1, 6)
				<% '�̸���ݾ� %>
				.Col = C_BillAmt    :	.text = arrRet(intCnt - 1, 7)
				<% '���ּ��� %>
				.Col = C_SoSeqNo    :	.text = arrRet(intCnt - 1, 9)
				<% '���� %>
				.Col = C_PlantCd    :	.text = arrRet(intCnt - 1, 10)
				<% 'ǰ��԰� %>
				.Col = C_ItemSpec	:	.text = arrRet(intCnt - 1, 12)
				<% 'VAT ���� %>
				.Col = C_VatType    :	.text = arrRet(intCnt - 1, 13)
				<% 'VAT ������ %>
				.Col = C_VatNm      :	.text = arrRet(intCnt - 1, 14)
				<% 'VAT �� %>
				.Col = C_VatRate    :	.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 15), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
				<% 'VAT ���Ա��� %>
				.Col = C_VatIncFlag	:	.text = arrRet(intCnt - 1, 16)
				
				' VAT ���Ա��и� 
				.Col = C_VatIncFlagNm
				If arrRet(intCnt - 1, 16) = "1" Then
					.Text = "����"
				Else
					.Text = "����"
				End If
				
				<% '�����ܰ� %>
				.Col = C_DepositPrice	:	.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 17), frm1.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")
				<% '��ǰ���� %>
				.Col = C_RetItemFlag	:	.text = arrRet(intCnt - 1, 18)
				<% 'VAT ���Ա��� %>
				.Col = C_OldVatIncFlag	:	.text = arrRet(intCnt - 1, 16)
				<% 'Initial ����ݾ� %>
				.Col = C_InitialBillAmt :	.text = 0
				<% 'Initial VAT�ݾ� %>
				.Col = C_InitialVatAmt  :	.text = 0

				Call CalcRefAmt(CLng(iIntStartRow) + CLng(intCntRow), iDblAccuBillAmt, iDblAccuVatAmt)

			End if
		Next

		If iIntStartRow <> .MaxRows Then
			SetSpreadColor iIntStartRow, .MaxRows
			Call SetTotal(iDblAccuBillAmt, iDblAccuVatAmt)
		End If
		.ReDraw = True

	End With

	' Call RefBillHdrSum("DL")
	Call JungBokMsg(strSOJungBokMsg,"���ֹ�ȣ" & "-" & "���ּ���")

	lgBlnFlgChgValue = True

End Function

'===========================================
Function SetDNDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim iIntStartRow, I, j
	Dim intLoopCnt
	Dim intCnt
	Dim blnEqualFlg
	Dim strSoNo,strSoSeqNo
	Dim intCntRow
	Dim iDblAccuBillAmt, iDblAccuVatAmt

	Dim strSOJungBokMsg
	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		iIntStartRow = .vspdData.MaxRows           <% '��: ��������� MaxRows %>
		intLoopCnt = Ubound(arrRet, 1)          <% '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� %>
		intCntRow = 0

		strSOJungBokMsg = ""
		For intCnt = 1 to intLoopCnt 
			blnEqualFlg = False

			If iIntStartRow <> 0 Then

				strSoNo=""
				strSoSeqNo=""

				<% '---> ���������� ���� ���Ϲ�ȣ�� ���ϼ����� �ִ��� üũ�Ѵ� %>
				For j = 1 To iIntStartRow
					.vspdData.Row = j
					<% '���Ϲ�ȣ %>
					.vspdData.Col = C_DnNo
					strSoNo = .vspdData.text
					<% '���ϼ��� %>
					.vspdData.Col = C_DnSeqNo          
					strSoSeqNo = .vspdData.text

					If strSoNo = arrRet(intCnt - 1, 9) And strSoSeqNo = arrRet(intCnt - 1, 10) Then
						blnEqualFlg = True
						strSOJungBokMsg = strSOJungBokMsg & Chr(13) & strSoNo & "-" & strSoSeqNo
						Exit For
					End If
				Next

			End If
      
			If blnEqualFlg = false then
				intCntRow = intCntRow + 1
				.vspdData.MaxRows = CLng(iIntStartRow) + CLng(intCntRow)
				.vspdData.Row = CLng(iIntStartRow) + CLng(intCntRow)

				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				<% '���ֹ�ȣ %>
				.vspdData.Col = C_SoNo
				.vspdData.text = arrRet(intCnt - 1, 0)
				<% 'ǰ�� %>
				.vspdData.Col = C_ItemCd          
				.vspdData.text = arrRet(intCnt - 1, 1)
				<% 'ǰ��� %>
				.vspdData.Col = C_ItemNm          
				.vspdData.text = arrRet(intCnt - 1, 2)
				<% 'Tracking No %>
				.vspdData.Col = C_TrackNo          
				.vspdData.text = arrRet(intCnt - 1, 3)
				<% '�̸������ %>
				.vspdData.Col = C_BillQty          
				.vspdData.text = arrRet(intCnt - 1, 4)
				<% '���� %>
				.vspdData.Col = C_BillUnit          
				.vspdData.text = arrRet(intCnt - 1, 5)
				<% '�̸���ܰ� %>
				.vspdData.Col = C_BillPrice          
				.vspdData.text = arrRet(intCnt - 1, 6)
				<% '�̸���ݾ� %>
				.vspdData.Col = C_BillAmt          
				.vspdData.text = arrRet(intCnt - 1, 7)
				<% '���Ϲ�ȣ %>
				.vspdData.Col = C_DnNo         
				.vspdData.text = arrRet(intCnt - 1, 9)
				<% '���ϼ��� %>
				.vspdData.Col = C_DnSeqNo        
				.vspdData.text = arrRet(intCnt - 1, 10)
				<% 'LC��ȣ %>
				.vspdData.Col = C_LlcNo
				.vspdData.text = arrRet(intCnt - 1, 11)
				<% 'LC���� %>
				.vspdData.Col = C_LlcSeq          
				.vspdData.text = arrRet(intCnt - 1, 12)
				<% '���� %>
				.vspdData.Col = C_PlantCd         
				.vspdData.text = arrRet(intCnt - 1, 13)
				<% 'ǰ��԰� %>
				.vspdData.Col = C_ItemSpec
				.vspdData.text = arrRet(intCnt - 1, 15)

				<% '���ּ��� %>
				.vspdData.Col = C_SoSeqNo          
				.vspdData.text = arrRet(intCnt - 1, 16)

				.vspdData.Col = C_VatType         
				.vspdData.text = Trim(arrRet(intCnt - 1, 17))
				        
				.vspdData.Col = C_VatNm         
				.vspdData.text = arrRet(intCnt - 1, 18)
				                    
				.vspdData.Col = C_VatRate         
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 19), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")

				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 20)
				
				.vspdData.Col = C_VatIncFlagNm
				If arrRet(intCnt - 1, 20) = "1" Then
					.vspdData.text = "����"
				Else
					.vspdData.text = "����"
				End If
				<% '�����ܰ� %>
				.vspdData.Col = C_DepositPrice
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 21),.txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")
				<% '��ǰ���� %>
				.vspdData.Col = C_RetItemFlag
				.vspdData.text = arrRet(intCnt - 1, 22)
				<% 'VAT ���Ա��� %>
				.vspdData.Col = C_OldVatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 20)
				<% 'Initial ����ݾ� %>
				.vspdData.Col = C_InitialBillAmt          
				.vspdData.text = 0
				<% 'Initial VAT�ݾ� %>
				.vspdData.Col = C_InitialVatAmt          
				.vspdData.text = 0

				Call CalcRefAmt(CLng(iIntStartRow) + CLng(intCntRow), iDblAccuBillAmt, iDblAccuVatAmt)

'				SetSpreadColor CLng(TempRow) + CLng(intCntRow),CLng(TempRow) + CLng(intCntRow)
			End if
		Next

		If iIntStartRow <> .vspdData.MaxRows Then
			SetSpreadColor iIntStartRow, .vspdData.MaxRows
			Call SetTotal(iDblAccuBillAmt, iDblAccuVatAmt)
		End If
		.vspdData.ReDraw = True
	End With

	' Call RefBillHdrSum("DL")
	Call JungBokMsg(strSOJungBokMsg,"���Ϲ�ȣ" & "-" & "���ϼ���")

	lgBlnFlgChgValue = True

End Function

'===========================================
Function SetLCDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim iIntStartRow, I, j
	Dim intLoopCnt
	Dim intCnt
	Dim blnEqualFlg
	Dim strLcNo,strLcSeqNo
	Dim intCntRow
	Dim iDblAccuBillAmt, iDblAccuVatAmt

	Dim strLCJungBokMsg

	With frm1
		.vspdData.focus
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False 

		iIntStartRow = .vspdData.MaxRows           <% '��: ��������� MaxRows %>
		intLoopCnt = Ubound(arrRet, 1)          <% '��: Reference Popup���� ���õǾ��� Row��ŭ �߰� %>
		intCntRow = 0

		strLCJungBokMsg = ""

		For intCnt = 1 to intLoopCnt 
			blnEqualFlg = False

			If iIntStartRow <> 0 Then
				strLcNo=""
				strLcSeqNo=""

				<% '---> L/C������ ���� L/C��ȣ�� L/C������ �ִ��� üũ�Ѵ� %>
				For j = 1 To iIntStartRow
					.vspdData.Row = j
					<% 'LC��ȣ %>
					.vspdData.Col = C_LlcNo
					strLcNo = .vspdData.text
					<% 'LC���� %>
					.vspdData.Col = C_LlcSeq          
					strLcSeqNo = .vspdData.text

					If strLcNo = arrRet(intCnt - 1, 0) And strLcSeqNo = arrRet(intCnt - 1, 11) Then
						blnEqualFlg = True
						strLCJungBokMsg = strLCJungBokMsg & Chr(13) & strLCNo & "-" & strLCSeqNo
						Exit For
					End If
				Next
			End If
	      
			If blnEqualFlg = false then
				intCntRow = intCntRow + 1
				.vspdData.MaxRows = CLng(iIntStartRow) + CLng(intCntRow)
				.vspdData.Row = CLng(iIntStartRow) + CLng(intCntRow)

				.vspdData.Col = 0
				.vspdData.Text = ggoSpread.InsertFlag

				<% 'L/C��ȣ %>
				.vspdData.Col = C_LlcNo
				.vspdData.text = arrRet(intCnt - 1, 0)
				<% '���ֹ�ȣ %>
				.vspdData.Col = C_SoNo
				.vspdData.text = arrRet(intCnt - 1, 1)
				<% 'ǰ�� %>
				.vspdData.Col = C_ItemCd          
				.vspdData.text = arrRet(intCnt - 1, 2)
				<% 'ǰ��� %>
				.vspdData.Col = C_ItemNm          
				.vspdData.text = arrRet(intCnt - 1, 3)
				<% 'Tracking No %>
				.vspdData.Col = C_TrackNo          
				.vspdData.text = arrRet(intCnt - 1, 4)
				<% '�̸������ %>
				.vspdData.Col = C_BillQty          
				.vspdData.text = arrRet(intCnt - 1, 5)
				<% '���� %>
				.vspdData.Col = C_BillUnit          
				.vspdData.text = arrRet(intCnt - 1, 6)
				<% '�̸���ܰ� %>
				.vspdData.Col = C_BillPrice          
				.vspdData.text = arrRet(intCnt - 1, 7)
				<% '�̸���ݾ� %>
				.vspdData.Col = C_BillAmt          
				.vspdData.text = arrRet(intCnt - 1, 8)
				<% '���ּ��� %>
				.vspdData.Col = C_SoSeqNo          
				.vspdData.text = arrRet(intCnt - 1, 10)
				<% 'L/C���� %>
				.vspdData.Col = C_LlcSeq          
				.vspdData.text = arrRet(intCnt - 1, 11)
				<% 'ǰ��԰� %>
				.vspdData.Col = C_ItemSpec
				.vspdData.text = arrRet(intCnt - 1, 12)
				<% 'VAT type %>
				.vspdData.Col = C_VatType         
				.vspdData.text = arrRet(intCnt - 1, 13)
				<% 'VAT type�� %>
				.vspdData.Col = C_VatNm         
				.vspdData.text = arrRet(intCnt - 1, 14)
				<% 'VAT Rate %>
				.vspdData.Col = C_VatRate         
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 15), Parent.gCurrency, Parent.ggExchRateNo, "X" , "X")
				<% 'VAT ���Ա��� %>
				.vspdData.Col = C_VatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 16)

				.vspdData.Col = C_VatIncFlagNm
				If arrRet(intCnt - 1, 16) = "1" Then
					.vspdData.text = "����"
				Else
					.vspdData.text = "����"
				End If
				
				<% '�����ܰ� %>
				.vspdData.Col = C_DepositPrice
				.vspdData.text = UNIConvNumPCToCompanyByCurrency(arrRet(intCnt - 1, 17), .txtCurrency.value,Parent.ggUnitCostNo, "X" , "X")
				<% '���� %>
				.vspdData.Col = C_PlantCd
				.vspdData.text = arrRet(intCnt - 1, 18)
				<% '��ǰ���� %>
				.vspdData.Col = C_RetItemFlag
				.vspdData.text = arrRet(intCnt - 1, 19)
				<% 'Original ����ݾ� %>
				.vspdData.Col = C_OldBillAmt          
				.vspdData.text = 0
				<% 'VAT ���Ա��� %>
				.vspdData.Col = C_OldVatIncFlag
				.vspdData.text = arrRet(intCnt - 1, 16)
				<% 'Original VAT�ݾ� %>
				.vspdData.Col = C_OldVatAmt          
				.vspdData.text = 0
				<% 'Initial ����ݾ� %>
				.vspdData.Col = C_InitialBillAmt          
				.vspdData.text = 0
				<% 'Initial VAT�ݾ� %>
				.vspdData.Col = C_InitialVatAmt          
				.vspdData.text = 0

				Call CalcRefAmt(CLng(iIntStartRow) + CLng(intCntRow), iDblAccuBillAmt, iDblAccuVatAmt)
'				SetSpreadColor CLng(TempRow) + CLng(intCntRow),CLng(TempRow) + CLng(intCntRow)
			End if
		Next

		If iIntStartRow <> .vspdData.MaxRows Then
			SetSpreadColor iIntStartRow, .vspdData.MaxRows
			Call SetTotal(iDblAccuBillAmt, iDblAccuVatAmt)
		End If
		.vspdData.ReDraw = True
	End With

	' Call RefBillHdrSum("DL")
	 Call JungBokMsg(strLCJungBokMsg,"L/C��ȣ" & "-" & "L/C����")

	 lgBlnFlgChgValue = True
End Function

'========================================
'������ �������� �ݾ� ��� 
Sub CalcRefAmt(ByVal pvIntRow, ByRef iDblAccuBillAmt, ByRef iDblAccuVatAmt)
	Dim iDblBillAmt, iDblBillQty, iDblBillPrice, iDblBillAmtLoc, iDblVatRate, iDblVatAmt, iDblDepositAmt, iDblDepositPrice 
	Dim iStrVatIncFlag

	With frm1.vspdData
		.Row = pvIntRow

		ggoSpread.source = frm1.vspdData
		
		.Col = C_BillQty		: iDblBillQty = UNICDbl(.Text)
		.Col = C_BillAmt		: iDblBillAmt = UNICDbl(.Text)
		.Col = C_DepositPrice	: iDblDepositPrice = UNICDbl(.Text)
		
'		iDblBillAmt = iDblBillQty * iDblBillPrice
'		.Col = C_BillAmt
		If iDblBillAmt <> 0 Then
'			.Text = UNIConvNumPCToCompanyByCurrency(iDblBillAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")

			.Col = C_OldBillAmt	: .Text = UNIConvNumPCToCompanyByCurrency(iDblBillAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")	 
			iDblBillAmt = UNICDbl(.Text)

			iDblAccuBillAmt = iDblAccuBillAmt + iDblBillAmt

			.Col = C_BillLocAmt	: .Text = FncCalcAmtLoc(iDblBillAmt, UNICDbl(frm1.txtXchgRate.Text), frm1.txtXchgOp.value, Parent.gLocRndPolicyNo)

			iDblBillAmtLoc = UNICDbl(.Text)
						
			.Col = C_VatIncFlag	: iStrVatIncFlag = .Text
			.Col = C_VatRate	: iDblVatRate	 = UNICDbl(.Text)
						
'			.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
			.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency.Value)

'			.Col = C_OldVatAmt  : .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
			.Col = C_OldVatAmt  : .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency.Value)

			iDblAccuVatAmt = iDblAccuVatAmt + UNICDbl(.Text)

			.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)			
			
			If iStrVatIncFlag = "2" Then						
				iDblAccuBillAmt = UNICDbl(iDblAccuBillAmt) - iDblAccuVatAmt
			End If
			
		Else
'			.Col = C_BillAmt	: .Text = "0"
			.Col = C_BillLocAmt	: .Text = "0"
			.Col = C_VatAmt		: .Text = "0"
			.Col = C_VatLocAmt	: .Text = "0"
		End If

		iDblDepositAmt = iDblBillQty * iDblDepositPrice

		.Col = C_DepositAmt
		If iDblDepositAmt <> 0 Then
			.Text = UNIConvNumPCToCompanyByCurrency(iDblDepositAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
		Else
			.Col = C_DepositAmt  : .Text = "0"
		End If

	End With
End Sub

'========================================
' Document�ݾ� ��� 
Sub CalcAmt(ByVal pvIntRow)
	Dim iDblBillAmt, iDblBillQty, iDblBillPrice, iDblDepositAmt, iDblDepositPrice 
	Dim iStrVatIncFlag

	With frm1.vspdData
		.Row = pvIntRow

		ggoSpread.source = frm1.vspdData
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		.Col = C_BillQty		: iDblBillQty = UNICDbl(.Text)
		.Col = C_BillPrice		: iDblBillPrice = UNICDbl(.Text)
		.Col = C_DepositPrice	: iDblDepositPrice = UNICDbl(.Text)
		
		iDblBillAmt = iDblBillQty * iDblBillPrice
		.Col = C_BillAmt
		If iDblBillAmt <> 0 Then
			.Text = UNIConvNumPCToCompanyByCurrency(iDblBillAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
			iDblBillAmt = UNICDbl(.Text)
		Else
			.Text = 0
		End If

		iDblDepositAmt = iDblBillQty * iDblDepositPrice
		.Col = C_DepositAmt
		If iDblDepositAmt <> 0 Then
			.Text = UNIConvNumPCToCompanyByCurrency(iDblDepositAmt,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
			iDblDepositAmt = UNICDbl(.Text)
		Else
			.Text = 0
		End If


	End With
	
	' �ڱ��ݾ� ��� 
	Call CalcAmtLoc(pvIntRow)
End Sub

'========================================
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
						
'			.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
			.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency.Value)

			.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)
		Else
			.Col = C_BillLocAmt	: .Text = "0"
			.Col = C_VatAmt		: .Text = "0"
			.Col = C_VatLocAmt	: .Text = "0"
		End If
	End With
	
	' �ѱݾ� ��� 
	Call CalcTotal("U", pvIntRow)
End Sub

'========================================
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
							
'		.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency)
		.Col = C_VatAmt		: .Text = FncCalcVatAmt(iDblBillAmt, iStrVatIncFlag, iDblVatRate, frm1.txtCurrency.Value)

		.Col = C_VatLocAmt	: .Text = FncCalcVatAmt(iDblBillAmtLoc, iStrVatIncFlag, iDblVatRate, parent.gCurrency)
	End With

	' �ѱݾ� ��� 
	Call CalcTotal("U", pvIntRow)
End Sub

'========================================
' ���հ�ݾ��� �����Ѵ�.
Sub CalcTotal(ByVal pvStrFlag, ByVal pvIntRow)
'	On Error Resume Next
	
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
					iDblDiffNetAmt = UNICDbl(iStrBillAmt) - iDblOldBillAmt
				Else
					.Col = C_BillAmt	: iStrBillAmt = .Text
					.Col = C_OldVatIncFlag
					iDblDiffNetAmt = UNICDbl(iStrBillAmt) - iDblOldBillAmt - iDblDiffVatAmt
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
					    .Col = C_InitialVatAmt	: iDblDiffVatAmt = UNICDbl(.Text) - iDblOldVatAmt
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

'========================================
Sub SetTotal(ByVal pvDblNetAmt, ByVal pvDblVatAmt)
	Dim iDblTotNetAmt, iDblTotVatAmt
	
	With frm1	
			iDblTotNetAmt = UNICDbl(.txtOriginBillAmt.Text) + pvDblNetAmt
			iDblTotVatAmt = UNICDbl(.txtVatAmt.Text) + pvDblVatAmt

			.txtOriginBillAmt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotNetAmt,.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")
			.txtVatAmt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotVatAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, Parent.gTaxRndPolicyNo  , "X")
			.txtTotBillAmt.Text = UNIConvNumPCToCompanyByCurrency(iDblTotNetAmt+iDblTotVatAmt,.txtCurrency.value,Parent.ggAmtOfMoneyNo, "X" , "X")	
	End With
		
End Sub

'========================================
' �ڱ��ݾ��� ����Ѵ�.
' pvDblAmt : Document�ݾ� - Double�� 
' pvDblXchgRate : ȯ�� - Double�� 
' pvStrXchgRateOp : ȯ�������� 
' ���ǻ��� : ȯ�������ڰ� �Էµ��� ������ ���������� ó���Ѵ�.
' �Լ��� Return ���� Formató���� �����̴�.
'========================================
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

'========================================
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
    Call SetSpreadHidden()
    If frm1.HPostFlag.Value <>"Y" Then
		Call SetQuerySpreadColor(1) 
    Else
		Call SetPostYesSpreadColor(1)
    End If	
End Sub

'========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Row <= 0 Then Exit Sub

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		.Col = Col - 1
		.Row = Row
					       
		Select Case Col
					                
			Case C_VatPopup             '�߰� 
				Call OpenVat()
		End Select
					      
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
			ggoSpread.SSSort Col			'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		Exit Sub
	End If    		
	
	If frm1.HPostFlag.Value <>"Y" Then   
		Call SetPopupMenuItemInf("0101111111")   
	Else
		Call SetPopupMenuItemInf("0000111111")   
	End IF	
End Sub

'========================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim iStrVatAmt

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True
	 
	With frm1

		.vspdData.Row = Row

		Select Case Col 
		Case C_BillQty
		 	Call CalcAmt(Row)
		'   Call BillTotalSum(C_BillQty,Row)
		Case C_BillPrice
			Call CalcAmt(Row)
		'   Call BillTotalSum(C_BillPrice,Row)
		Case C_BillAmt
			Call CalcAmtLoc(Row)
		'   Call BillTotalSum(C_BillAmt,Row)
		Case C_VatAmt
			.vspdData.Row = Row
			.vspdData.Col = C_VatAmt	: iStrVatAmt = .vspdData.text
			'Document Currency�� Local Currency�� ������ ��� Vat Amount, Vat Amount Local�� �����ϰ� ���� 
			If UCase(Parent.gCurrency) = UCase(Trim(frm1.txtCurrency.value)) Then
				.vspdData.Col = C_VatLocAmt	:	.vspdData.Text = iStrVatAmt
			'Document Currency�� Local Currency�� �ٸ� ��� Vat Amount Local �ٽ� ��� 
			Else
				.vspdData.Col = C_BillAmt
				If UNICDbl(.vspdData.Text) = 0 Then
					.vspdData.Col = C_VatLocAmt	:	.vspdData.Text = FncCalcAmtLoc(UNICDbl(iStrVatAmt), UNICDbl(frm1.txtXchgRate.Text), frm1.txtXchgOp.value, Parent.gTaxRndPolicyNo)
				End If
			End if
				
			Call CalcTotal("U", Row)
		'   Call BillTotalSum(C_VatAmt,Row)
		Case C_BillLocAmt
			Call CalcVatAmt(Row)
		'   Call LocBillTotalSum(Col,Row)
		Case C_VatType
		'	Call SetVatType()
			Call SetVatType(Row)
			Call CalcVatAmt(Row)
		'   Call BillTotalSum(C_BillAmt,Row)
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
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess Then Exit Sub
		    
		Call DisableToolBar(Parent.TBC_QUERY)
		Call DBQuery
	End If
End Sub



'=============================================
' 2005.11.10 SMJ
'=============================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)	
	ggoSpread.Source = frm1.vspdData
'	Call JumpPgm()
End Sub


Function JumpPgm()	
	Dim pvSelmvid, pvFB_fg,pvKeyVal,StrNVar,StrNPgm,pvSingle
	
	if frm1.vspddata.Maxrows  < 1 then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	ggoSpread.Source = frm1.vspdData
	
	frm1.vspddata.row = 0
    frm1.vspddata.col = frm1.vspddata.Activecol


    Select case frm1.vspddata.value
    
   	
	Case "���Ϲ�ȣ"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow		

		if 	frm1.vspddata.value <> "" then
	
				
					pvKeyVal =   frm1.vspddata.value
					
									
					pvSingle =   ""
				
					pvFB_fg = "B"
					pvSelmvid = "DN_NO"
	
						Call Jump_Pgm (	pvSelmvid, _
										pvFB_fg, _
										pvSingle,  _
										pvKeyVal)
										
										
										
	End if 											
		 
	End select
End Function


'========================================
Sub btnPostFlag_OnClick()

 If BtnSpreadCheck = False Then Exit Sub

 Dim strVal

 frm1.txtInsrtUserId.value = Parent.gUsrID 
   
  If   LayerShowHide(1) = False Then
             Exit Sub
        End If

 strVal = BIZ_PGM_ID & "?txtMode=" & PostFlag         
 strVal = strVal & "&txtHBillNo=" & Trim(frm1.txtHBillNo.value)      
 strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
 strVal = strVal & "&txtChangeOrgId=" & Parent.gChangeOrgId

 Call RunMyBizASP(MyBizASP, strVal)            
 
End Sub

'==========================================
Sub btnGLView_OnClick()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lblnWinEvent
 	
	If Trim(frm1.txtGLNo.value) <> "" Then
		 arrParam(0) = Trim(frm1.txtGLNo.value) 'ȸ����ǥ��ȣ 
		 arrParam(1) = Trim(frm1.txtHBillNo.value) 'Reference��ȣ 
		 
		 if arrParam(0) = "" THEN Exit Sub
		 
		 iCalledAspName = AskPRAspName("a5120ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If

		 arrRet = window.showModalDialog(iCalledAspName , Array(window.parent,arrParam), _
		      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		      
	ElseIf Trim(frm1.txtTempGLNo.value) <> "" Then
	     arrParam(0) = Trim(frm1.txtTempGLNo.value) '������ǥ��ȣ 
	     arrParam(1) = Trim(frm1.txtHBillNo.value) 'Reference��ȣ 
	 
	     if arrParam(0) = "" THEN Exit Sub
	     
	     iCalledAspName = AskPRAspName("a5130ra1")
		 
		 If Trim(iCalledAspName) = "" Then
		      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		      lblnWinEvent = False
		      Exit Sub
		 End If
		 
	     arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else 
	     Call DisplayMsgBox("205154", "X", "X", "X")
	End If 
	     lblnWinEvent = False
End Sub

'==========================================
Sub btnPreRcptView_OnClick()
     Dim arrRet
     Dim arrParam(4)
	 Dim iCalledAspName
	 Dim IntRetCD
	 Dim lblnWinEvent
 
     arrParam(0) = Trim(frm1.txtHBillDt.value)    '����ä���� 
     arrParam(1) = Trim(frm1.txtSoldToParty.value)  '�ֹ�ó 
     arrParam(2) = Trim(frm1.txtSoldToPartyNm.value)  '�ֹ�ó 
     arrParam(3) = Trim(frm1.txtCurrency.value)   'ȭ�� 
     arrParam(4) = ""         '�����ݹ�ȣ 
 
     iCalledAspName = AskPRAspName("s5111ra7")
		 
	 If Trim(iCalledAspName) = "" Then
	      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s5111ra7", "X")
		  lblnWinEvent = False
		       Exit Sub
		  End If
     arrRet = window.showModalDialog(iCalledAspName & "?txtFlag=BH&txtCurrency=" & frm1.txtCurrency.value, Array(window.parent,arrParam), _
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
  '�Ѹ���ä�Ǳݾ� 
  ggoOper.FormatFieldByObjectOfCur .txtTotBillAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
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
  ggoSpread.SSSetFloatByCellOfCur C_OldBillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  ggoSpread.SSSetFloatByCellOfCur C_InitialBillAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec  
  
  'VAT�ݾ� 
  ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  ggoSpread.SSSetFloatByCellOfCur C_OldVatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  ggoSpread.SSSetFloatByCellOfCur C_InitialVatAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec 
  
  '�����ܰ� 
  ggoSpread.SSSetFloatByCellOfCur C_DepositPrice,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  '�����ݾ� 
  ggoSpread.SSSetFloatByCellOfCur C_DepositAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  'FOB�ݾ� 
  ggoSpread.SSSetFloatByCellOfCur C_FOBAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
  
 End With

End Sub
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ä�ǳ������</font></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
          </TR>
      </TABLE>
     </TD>
     <TD WIDTH=* align=right><A href="vbscript:OpenSODtlRef">���ֳ�������</A>&nbsp;|&nbsp;<A href="vbscript:OpenLCDtlRef">L/C��������</A>&nbsp;|&nbsp;<A href="vbscript:OpenDNDtlRef">���ϳ�������</A></TD>
     <TD WIDTH=10>&nbsp;</TD>
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
         <TD CLASS="TD5" NOWRAP>����ä�ǹ�ȣ</TD>
         <TD CLASS="TD6"><INPUT NAME="txtConBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSBillDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBillDtl()"></TD>
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
        <TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldtoParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="�ֹ�ó" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtSoldtoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>����ó</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillToPartyCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtBillToPartyNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>�������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>����ä�Ǳݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <TABLE CELLSPACING=0 CELLPADDING=0>
          <TR>
           <TD>
            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtOriginBillAmt" CLASS=FPDS140 Alt="����ä�Ǳݾ�" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
           </TD>
           <TD>
            &nbsp;<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1">
           </TD>
          </TR>
         </TABLE>
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>ȯ��</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchgRate" CLASS=FPDS100 tag="24X5" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
        <TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtVatAmt" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
        <TD CLASS=TD6><INPUT NAME="txtSoNo" TYPE="Text" MAXLENGTH=18 SiZE=30 tag="24XXXU" class = protected readonly = true TABINDEX="-1"></TD>
        <TD CLASS=TD5 NOWRAP>�Ѹ���ä�Ǳݾ�</TD>
        <TD CLASS=TD6 NOWRAP>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtTotBillAmt" CLASS=FPDS140 Alt="�Ѹ���ä�Ǳݾ�" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
        </TD>
       </TR>
       <TR>
        <TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag1" value="1" tag = "24">
          <label ID="lblVatIncFlag1" for="rdoVatIncFlag1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag2" value="2" tag = "24" checked>
          <label ID="lblVatIncFlag2" for="rdoVatIncFlag2">����</label>
        </TD>
        <TD CLASS=TD5 NOWRAP>VAT�������</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoVatCalcType" id="rdoVatCalcType1" value="1" tag = "24">
          <label ID="lblVatIncFlag1" for="rdoVatCalcType1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoVatCalcType" id="rdoVatCalcType2" value="2" tag = "24" checked>
          <label ID="lblVatIncFlag2" for="rdoVatCalcType2">����</label>
        </TD>
       </TR>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" Title="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
      <BUTTON NAME="btnGLView" CLASS="CLSMBTN">��ǥ��ȸ</BUTTON>&nbsp;
      <BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">��������Ȳ</BUTTON></TD>
     <TD WIDTH=* Align=Right><a href = "vbscript:JumpChgCheck(BIZ_BillHdr_JUMP_ID)">����ä�ǵ��</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillCollect_JUMP_ID)">����ä�Ǽ��ݳ������</a></TD>
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

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtRefFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBillTypeCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtReverseFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HVatRate" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HVatType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSts" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtXchgOp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPostFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtLocBillAmt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtLocVatAmt" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHBillDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrpCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrpNm" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
  </DIV>
</BODY>
</HTML>