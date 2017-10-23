<%@ LANGUAGE="VBSCRIPT" %>
<!--
*****************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M4212ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관란 등록 ASP	     												*
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Sun-jung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'*							  2. 2000/04/11 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'============================================  1.1.2 공통 Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="VBS">
Option Explicit	
	

Const BIZ_PGM_QRY_ID = "m4213mb1.asp"	
Const BIZ_PGM_SAVE_ID = "m4213mb2.asp"	
Const CC_DETAIL_ENTRY_ID = "m4212ma1"
	
<!-- #Include file="../../inc/lgvariables.inc" -->


Dim gblnWinEvent
	
Dim C_LanNo 								'란번호 
Dim C_HsCd 									'HS번호 
Dim C_HsNm 									'H/S명 
Dim C_Unit 									'단위 
Dim C_CIFDocAmt 							'CIF금액(US)
Dim C_CIFLocAmt 							'CIF원화금액 
Dim C_TraiffRate 							'세율 
Dim C_ReduRate 								'감면율 
Dim C_TaxLocAmt								'세액 
Dim C_NetWeight								'순중량 
Dim C_TotQty 								'총수량 
Dim C_TotDocAmt								'총금액		

'=============================  initSpreadPosVariables()  ==============================================
Sub InitSpreadPosVariables()
	C_LanNo			= 1							'란번호 
	C_HsCd			= 2							'HS번호 
	C_HsNm			= 3							'H/S명 
	C_Unit			= 4							'단위 
	C_CIFDocAmt		= 5							'CIF금액(US)
	C_CIFLocAmt		= 6							'CIF원화금액 
	C_TraiffRate	= 7							'세율 
	C_ReduRate		= 8							'감면율 
	C_TaxLocAmt		= 9							'세액 
	C_NetWeight		= 10						'순중량 
	C_TotQty		= 11						'총수량 
	C_TotDocAmt		= 12						'총금액 

End Sub
<!--
'==========================================  2.1.1 InitVariables()  =====================================
-->
Function InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE			
	lgBlnFlgChgValue = False			
	lgIntGrpCount = 0					
	lgStrPrevKey = ""					
	lgLngCurRows = 0 					
		
	gblnWinEvent = False
	frm1.vspdData.MaxRows = 0
End Function

<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
Sub SetDefaultVal()
	Call SetToolBar("1110000000001111")
	frm1.txtCCNo.focus
	Set gActiveElement = document.activeElement
End Sub

<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

<!--
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
    
    With frm1
		
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread  
		
		.vspdData.ReDraw = False
		
		.vspdData.MaxCols = C_TotDocAmt + 1
		.vspdData.Col = .vspdData.MaxCols :		.vspdData.ColHidden = True
		
    	Call GetSpreadColumnPos("A")
    		
		ggoSpread.SSSetEdit		C_LanNo,		"란번호", 10
		ggoSpread.SSSetEdit		C_HsCd,			"H/S부호", 20, 0
		ggoSpread.SSSetEdit		C_HsNm,			"H/S명", 25, 0
		ggoSpread.SSSetEdit		C_Unit,			"단위", 10, 0
		SetSpreadFloatLocal 	C_CIFDocAmt,	"CIF금액(US)",15,1 ,2
		SetSpreadFloatLocal 	C_CIFLocAmt,	"CIF자국금액",15,1 ,2
		SetSpreadFloatLocal 	C_TraiffRate,	"세율(%)",15,1,5
		SetSpreadFloatLocal 	C_ReduRate,		"감면율(%)",15,1,5
		SetSpreadFloatLocal 	C_TaxLocAmt,	"세액",15,1,2
		SetSpreadFloatLocal 	C_NetWeight,	"순중량",15,1,3
		SetSpreadFloatLocal 	C_TotQty,		"총수량",15,1,3
		SetSpreadFloatLocal 	C_TotDocAmt,	"총금액",15,1,2
		
		Call SetSpreadLock()
		.vspdData.ReDraw = True

	End With
End Sub

<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
Sub SetSpreadLock()
    
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
		ggoSpread.SpreadLock C_LanNo, -1, C_LanNo, -1
		ggoSpread.SpreadLock C_HsCd, -1, C_HsCd, -1
		ggoSpread.SpreadLock C_HsNm, -1, C_HsNm, -1
		
		ggoSpread.SpreadLock C_TotQty, -1, C_TotQty, -1
		ggoSpread.SpreadLock C_TotDocAmt, -1, C_TotDocAmt, -1
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, -1
	
	Call SetSpreadColor(-1,-1)
	
	frm1.vspdData.ReDraw = True	
    
End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
	ggoSpread.Source = frm1.vspdData
	
	ggoSpread.SSSetProtected C_LanNo,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_HsCd,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_HsNm,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Unit,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_NetWeight,	pvStartRow, pvEndRow
		
	ggoSpread.SSSetProtected C_TotQty,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TotDocAmt,	pvStartRow, pvEndRow
		
	ggoSpread.SSSetRequired C_CIFDocAmt,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_CIFLocAmt,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_TraiffRate,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_ReduRate,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_TaxLocAmt,	pvStartRow, pvEndRow
		
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStartRow, pvEndRow
		
End Sub

'===================================  GetSpreadColumnPos()  =======================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_LanNo			= iCurColumnPos(1)
			C_HsCd			= iCurColumnPos(2)
			C_HsNm			= iCurColumnPos(3)
			C_Unit			= iCurColumnPos(4)
			C_CIFDocAmt		= iCurColumnPos(5)
			C_CIFLocAmt		= iCurColumnPos(6)
			C_TraiffRate	= iCurColumnPos(7)
			C_ReduRate		= iCurColumnPos(8)
			C_TaxLocAmt		= iCurColumnPos(9)
			C_NetWeight		= iCurColumnPos(10)
			C_TotQty		= iCurColumnPos(11)
			C_TotDocAmt		= iCurColumnPos(12)
				
	End Select

End Sub	
	
<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCcNoPop()  ++++++++++++++++++++++++++++++++++++++
-->
Function OpenCcNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtCCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
   	iCalledAspName = AskPRAspName("M4211PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4211PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
		frm1.txtCCNo.focus
		Exit Function
	Else
		frm1.txtCCNo.value = strRet
		frm1.txtCCNo.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

	With frm1
		ggoSpread.Source = frm1.vspdData
		'CIF금액 
		ggoSpread.SSSetFloatByCellOfCur C_CIFDocAmt,-1, "USD",  Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'CIF자국금액 
		ggoSpread.SSSetFloatByCellOfCur C_CIFLocAmt,-1, parent.gCurrency,  Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'총통관금액 
		ggoSpread.SSSetFloatByCellOfCur C_TotDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec ,,,"Z"
    End With
End Sub
<!--
'=====================================  SetSpreadFloatLocal()  ===================================
-->
Sub SetSpreadFloatLocal(ByVal iCol,ByVal Header,ByVal dColWidth,ByVal HAlign,ByVal iFlag)
	        
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
    End Select
         
End Sub
<!--
'============================================  2.5.1 TotalSum()  ======================================
-->
Sub TotalSum()
	Dim SumTotal, lRow
		
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
		
	For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0
		If UNICDbl(frm1.vspdData.Text) <> UNICDbl(ggoSpread.DeleteFlag) then
			frm1.vspdData.Col = C_DocAmt
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
			
		end if
	Next
		
	frm1.txtTotItemAmt.Text = UNIConvNumPCToCompanyByCurrency(Cstr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,parent.gLocRndPolicyNo,"X")
End Sub

<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
	frm1.txtCCNo.Value = ReadCookie("CCNo")
	WriteCookie "txtCCNo", ""

	If UCase(Trim(frm1.txtCCNo.value)) <> "" Then
		Call MainQuery
	End If
		
End Function

<!--
'=============================================  2.5.3 LoadCcDtl()  ======================================
-->
Function LoadCcDtl()

	If Trim(frm1.txtCCNo.value)="" Then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	WriteCookie "CCNo", UCase(Trim(frm1.txtCCNo.value))
	PgmJump(CC_DETAIL_ENTRY_ID)

End Function

<!--
'==========================================  2.5.4 SetQuerySpreadColor()  ====================================
-->
Sub SetQuerySpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
	    
		.Redraw = False
		
		ggoSpread.SSSetProtected C_LanNo,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_HsCd,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_HsNm,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_Unit,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_NetWeight,	lRow, .MaxRows
		
		ggoSpread.SSSetProtected C_TotQty,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_TotDocAmt,	lRow, .MaxRows
		
		ggoSpread.SSSetRequired C_CIFDocAmt,	lRow, .MaxRows
		ggoSpread.SSSetRequired C_CIFLocAmt,	lRow, .MaxRows
		ggoSpread.SSSetRequired C_TraiffRate,	lRow, .MaxRows
		ggoSpread.SSSetRequired C_ReduRate,		lRow, .MaxRows
		ggoSpread.SSSetRequired C_TaxLocAmt,	lRow, .MaxRows
		
		.ReDraw = True
	End With
End Sub

<!--
'==========================================  2.5.5 CookiePage()  ======================================
-->
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4875	
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 1 Then
		
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
		
		WriteCookie CookieSplit , frm1.txtCCNo.value 
		Call PgmJump(CC_DETAIL_ENTRY_ID)
		
	ElseIf Kubun = 0 Then
		
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtCCNo.value =  arrVal(0) 

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()
					
		WriteCookie CookieSplit , ""

	End IF

End Function

<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
								
	Call LoadInfTB19029		
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet			

	Call SetDefaultVal
	Call InitVariables
		
	Call CookiePage(0)
		
End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
	    
End Sub
	

'========================================  vspdData_Click()  ===================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	IF frm1.vspdData.MaxRows = 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0001111111")
	End If
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
		   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
		   
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
'========================================  vspdData_DblClick()  ===================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
   
End Sub
'========================================  vspdData_ColWidthChange()  ===================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================  vspdData_MouseDown()  ===================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================  FncSplitColumn()  ===================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================  PopSaveSpreadColumnInf()  ===================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================  PopRestoreSpreadColumnInf()  ===================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    'Call SetQuerySpreadColor(1)
End Sub
<!--
'======================================  3.2.1 btnCcNo_OnClick()  ====================================
-->
Sub btnCcNo_Click()
	Call OpenCcNoPop()
End Sub

<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
Sub vspdData_Change(ByVal Col, ByVal Row )
	Dim CIFLocAmt, CIFDocAmt, XchRt, TraiffRate, ReduRate, TaxLocAmt
	Dim temp
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
        
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
        
	Select Case Col
	Case C_CIFDocAmt
		
		frm1.vspdData.Col = Col
			
		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
			CIFDocAmt = 0
		Else
			CIFDocAmt = UNICDbl(frm1.vspdData.Text)
		End If
			
		if Trim(frm1.hdnXchRt.Value) = "" OR IsNull(frm1.hdnXchRt.Value) then
			XchRt = 0
		else
			XchRt = UNICDbl(frm1.hdnXchRt.Value)
		end if
			
		frm1.vspdData.Col  = C_CIFLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(CIFDocAmt * XchRt),Parent.gCurrency,Parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo ,"X")
			
		frm1.vspdData.Col = C_CIFLocAmt
		CIFLocAmt = UNICDbl(frm1.vspdData.Text)
			
		frm1.vspdData.Col = C_TraiffRate
		TraiffRate = UNICDbl(frm1.vspdData.Text)
			
		frm1.vspdData.Col = C_ReduRate
		ReduRate = UNICDbl(frm1.vspdData.Text)
		TaxLocAmt = CIFLocAmt * (TraiffRate - ReduRate) / 100
			
		frm1.vspdData.Col = C_TaxLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(TaxLocAmt),Parent.gCurrency,Parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo ,"X")
			
	Case C_CIFLocAmt, C_TraiffRate, C_ReduRate
			
		frm1.vspdData.Col = C_CIFLocAmt
		CIFLocAmt = UNICDbl(frm1.vspdData.Text)
			
		frm1.vspdData.Col = C_TraiffRate
		TraiffRate = UNICDbl(frm1.vspdData.Text)
			
		frm1.vspdData.Col = C_ReduRate
		ReduRate = UNICDbl(frm1.vspdData.Text)
		TaxLocAmt = CIFLocAmt * (TraiffRate - ReduRate) / 100
			
		frm1.vspdData.Col = C_TaxLocAmt
		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(TaxLocAmt),Parent.gCurrency,Parent.ggAmtOfMoneyNo, parent.gTaxRndPolicyNo ,"X")
	End select
End Sub
	
<!--
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
-->
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub
	
<!--
'========================================  3.3.3 vspdData_TopLeftChange()  ==================================
-->
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

<!--
'=========================================  5.1.1 FncQuery()  ===========================================
-->
Function FncQuery()
	Dim IntRetCD

	FncQuery = False	

	Err.Clear			

	If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	Call InitVariables						

	If Not chkField(Document, "1") Then	
		Exit Function
	End If

	If DbQuery = False Then Exit Function

	FncQuery = True	
	Set gActiveElement = document.activeElement
End Function
	
<!--
'===========================================  5.1.2 FncNew()  ===========================================
-->
Function FncNew()
	Dim IntRetCD 

	FncNew = False  

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "A")						
	Call ggoOper.LockField(Document, "N")						
	Call SetDefaultVal
	Call InitVariables											

	FncNew = True												
	Set gActiveElement = document.activeElement
End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
Function FncDelete()
	If lgIntFlgMode <> Parent.OPMD_UMODE Then		
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	If DbDelete = False Then Exit Function

	FncDelete = True
	Set gActiveElement = document.activeElement	
End Function
	
<!--
'===========================================  5.1.4 FncSave()  ==========================================
-->
Function FncSave()
	Dim IntRetCD
		
	FncSave = False									
		
	Err.Clear										
		
	ggoSpread.Source = frm1.vspdData                
	If ggoSpread.SSCheckChange = False Then         
	    IntRetCD = DisplayMsgBox("900001","X","X","X") 
	    Exit Function
	End If
    
	ggoSpread.Source = frm1.vspdData                   
	
	If Not ggoSpread.SSDefaultCheck  Then   Exit Function
		
	If DbSave = False Then Exit Function
		
	FncSave = True
	Set gActiveElement = document.activeElement	
End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
Function FncCopy()
	Dim IntRetCD

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")	
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	lgIntFlgMode = Parent.OPMD_CMODE									
	
	frm1.vspdData.ReDraw = False
		
	if frm1.vspdData.Maxrows < 1	then exit function

	ggoSpread.Source = frm1.vspdData	
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo	
	Set gActiveElement = document.activeElement			
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow

	On Error Resume Next
	Err.Clear

	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End IF
		
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow , imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
    End With
	
	If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
	
	if frm1.vspdData.Maxrows < 1	then exit function
		
	With frm1.vspdData 
	
		.focus
		ggoSpread.Source = frm1.vspdData

		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
	End With

	Call TotalSum()
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
Function FncPrint()
   Call parent.FncPrint()
   Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
Function FncPrev() 
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then					
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgPrevNo = "" Then							
		Call DisplayMsgBox("900011","X","X","X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
Function FncNext()
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then	
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	ElseIf lgNextNo = "" Then			
		Call DisplayMsgBox("900012","X","X","X")
	End If
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
	Set gActiveElement = document.activeElement
End Function

<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
		
	Dim IntRetCD
		
	FncExit = False
		
    ggoSpread.Source = frm1.vspdData
	    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")  
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	    
    FncExit = True
	Set gActiveElement = document.activeElement  
End Function

<!--
'=============================================  5.2.1 DbQuery()  ========================================
-->
Function DbQuery()
	Dim strVal

	Err.Clear			

		
	DbQuery = False		
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtHCCNo.value)	
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)	
	End If
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	'수정(2003.06.10)
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
	
	Call RunMyBizASP(MyBizASP, strVal)								
	
	DbQuery = True													
End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim ColSep, RowSep
	
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size	
			
    Err.Clear														
    		
    DbSave = False
    
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '초기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '초기 버퍼의 설정[삭제]
   
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep  
	                                               
	If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		.txtMode.value = Parent.UID_M0002
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
	    
		For lRow = 1 To .vspdData.MaxRows
	    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							
					strVal = "C" & ColSep	& lRow & ColSep
		        Case ggoSpread.UpdateFlag							
					strVal = "U" & ColSep	& lRow & ColSep
				Case ggoSpread.DeleteFlag							
					strDel = "D" & ColSep	& lRow & ColSep

		            .vspdData.Col = C_LanNo 
		            strDel = strDel & Trim(.vspdData.Text) & RowSep

		            lGrpCnt = lGrpCnt + 1 

			End Select

			Select Case .vspdData.Text
				case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

		            .vspdData.Col = C_LanNo 		
		            strVal = strVal & Trim(.vspdData.Text) & ColSep
		            
		            .vspdData.Col = C_HsCd 		
		            strVal = strVal & Trim(.vspdData.Text) & ColSep
		            
		            .vspdData.Col = C_CIFDocAmt 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_CIFLocAmt 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_TraiffRate		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_ReduRate 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_TaxLocAmt		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_NetWeight 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_TotQty 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & ColSep
		            
		            .vspdData.Col = C_TotDocAmt 		
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & RowSep

		            lGrpCnt = lGrpCnt + 1 
		    End Select  
		    
		    '=====================
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
				                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
				 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
				       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
				      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
				          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
				       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
				         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select  

			'=====================     
		Next

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If
			
		'.txtMaxRows.value = lGrpCnt
		'.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
		
	End With
		
    DbSave = True                               

End Function

'======================================  RemovedivTextArea()  =================================
Function RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function	
<!--
'=============================================  5.2.3 DbDelete()  =======================================
-->
Function DbDelete()
	On Error Resume Next                        
End Function
	
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
Function DbQueryOk()					
	lgIntFlgMode = Parent.OPMD_UMODE			

	lgBlnFlgChgValue = False

	Call ggoOper.LockField(Document, "Q")
	Call SetToolBar("11101001000111")	
	
	Call RemovedivTextArea		
	'Call SetQuerySpreadColor(1)
End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()								
	Call InitVariables
	Call MainQuery()
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()							
	On Error Resume Next                        
End Function


'=========================================  vspdData_ScriptDragDropBlock()  =====================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관 란정보등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>통관 관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU"  ALT="통관 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" ALIGN=top TYPE="BUTTON" ONCLICK="VBSCRIPT:btnCcNo_Click()"></TD>
									<TD CLASS=TD6>&nbsp;</TD>
									<TD CLASS=TD6>&nbsp;</TD>
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
								<TD CLASS=TD5 NOWRAP>신고번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIDNo" ALT="신고번호" TYPE=TEXT MAXLENGTH=35 SIZE=34  TAG="24XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>신고일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m4213ma1_fpDateTime1_txtIDDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>면허번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIPNo" ALT="면허번호" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="24XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>면허일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m4213ma1_fpDateTime2_txtIPDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>총통관금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="환율">&nbsp;
											</TD>
											<TD>
												<script language =javascript src='./js/m4213ma1_fpDoubleSingle1_txtDocAmt.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>수출자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
								<script language =javascript src='./js/m4213ma1_I786608707_vspdData.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TD WIDTH=* ALIGN=RIGHT><A href="VBSCRIPT:CookiePage(1)">수입통관내역등록</A></TD>
				<TD WIDTH=50>&nbsp;</TD>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>

<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtHCCNo" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="hdnXchRt" TAG="24" TABINDEX=-1>
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
