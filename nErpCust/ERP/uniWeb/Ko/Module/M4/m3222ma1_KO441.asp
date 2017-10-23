
<%@ LANGUAGE="VBSCRIPT" %>
<!--
*****************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3222ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C Amend Detail 등록 ASP											*
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/04/03																*
'*  8. Modified date(Last)  : 2003/05/22																*
'*  9. Modifier (First)     : Sun-jung Lee
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : 화면 design												*
'*							  2. 2000/04/03 : Coding Start												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
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

	Const BIZ_PGM_QRY_ID = "m3222mb1_KO441.asp"	 
	Const BIZ_PGM_SAVE_ID = "m3222mb2_KO441.asp"	 
	Const LCAMEND_HEADER_ENTRY_ID = "m3221ma1"

	Dim C_AmdFlg
	Dim C_AmdFlgstr
	Dim C_ItemCd
	Dim C_ItemNm
	Dim C_Spec
	Dim C_Unit
	Dim C_BeQty
	Dim C_AtQty
	Dim C_Price
	Dim C_DocAmt
	Dim C_PORemainQty
	Dim C_HsCd
	Dim C_HsNm
	Dim C_AmendSeq
	Dim C_LcSeq
	Dim C_PoNo
	Dim C_PoSeq
	Dim C_OverTolerance
	Dim C_UnderTolerance
	Dim C_LcAmdSeq
	Dim C_ChgFlg
	'총품목금액계산을 위해 추가(2003.05)
	Dim C_OrgDocAmt		'변화값 저장 
	Dim C_OrgDocAmt1	'조회후 초기값 저장 
	'참조시 사용(2003.04.08)-Lee Eun Hee
	Dim C_AtQty_Ref
	
<!-- #Include file="../../inc/lgvariables.inc" -->
	
 Dim gblnWinEvent

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
 End Function

<!--
'==========================================  2.2.1 SetDefaultVal()  =====================================
-->
 Sub SetDefaultVal()
 	frm1.txtAmendDt.Text = UniConvDateAToB("<%= GetSvrDate %>", Parent.gServerDateFormat, Parent.gDateFormat)
 	frm1.txtDocAmt.text = UNIFormatNumber(CStr(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 	Call SetToolbar("1110000000001111")
 	frm1.txtLCAmdNo.focus
 	Set gActiveElement = document.activeElement
 End Sub

<!--
'==========================================  2.2.2 LoadInfTB19029()  ====================================
-->
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
 	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
 End Sub

'=========================================  2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
		C_AmdFlg		= 1					 
		C_AmdFlgstr		= 2 
		C_ItemCd		= 3
		C_ItemNm		= 4
		C_Spec			= 5
		C_Unit			= 6
		C_BeQty			= 7
		C_AtQty			= 8
		C_Price			= 9
		C_DocAmt		= 10
		C_PORemainQty	= 11
		C_HsCd			= 12
		C_HsNm			= 13
		C_AmendSeq		= 14
		C_LcSeq			= 15
		C_PoNo			= 16
		C_PoSeq			= 17
		C_OverTolerance	= 18
		C_UnderTolerance= 19
		C_LcAmdSeq		= 20
		C_ChgFlg		= 21
		C_OrgDocAmt		= 22
		C_OrgDocAmt1	= 23
End Sub
<!--
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
 Sub InitSpreadSheet()
 	Call InitSpreadPosVariables()

     With frm1
    
 		ggoSpread.Source = .vspdData
 		ggoSpread.Spreadinit "V20030530",,Parent.gAllowDragDropSpread  

 		.vspdData.ReDraw = False
			
 		.vspdData.MaxCols = C_OrgDocAmt1 + 1
 		.vspdData.MaxRows = 0
			
 		Call GetSpreadColumnPos("A")
			
 		ggoSpread.SSSetCombo	C_AmdFlg, "변경구분", 10, 0, False
 		ggoSpread.SSSetEdit		C_AmdFlgStr, "변경내용", 10, 0
 		ggoSpread.SSSetEdit		C_ItemCd, "품목", 10, 0
 		ggoSpread.SSSetEdit		C_ItemNm, "품목명", 20, 0
 		ggoSpread.SSSetEdit		C_Spec, "품목규격", 20, 0
 		ggoSpread.SSSetEdit		C_Unit, "단위", 10, 2
 		SetSpreadFloatLocal		C_BeQty,  "변경전수량", 15, 1, 3
 		SetSpreadFloatLocal		C_AtQty,  "변경후수량", 15, 1, 3
 		SetSpreadFloatLocal		C_Price, "단가", 15, 1, 4
 		SetSpreadFloatLocal		C_DocAmt, "L/C금액", 15, 1, 2
 		SetSpreadFloatLocal		C_PORemainQty,  "발주잔량", 15, 1, 3
 		ggoSpread.SSSetEdit		C_HsCd, "HS부호", 10, 0
 		ggoSpread.SSSetEdit		C_HsNm, "HS명", 20, 0
 		ggoSpread.SSSetEdit		C_AmendSeq, "AMEND순번", 10, 0
 		ggoSpread.SSSetEdit		C_LcSeq, "L/C순번", 10, 2
 		ggoSpread.SSSetEdit		C_PoNo, "발주번호", 18, 0
 		ggoSpread.SSSetEdit		C_PoSeq, "발주순번", 10, 2
 		SetSpreadFloatLocal				C_OverTolerance, "과부족허용율(+)", 15, 1, 5
 		SetSpreadFloatLocal				C_UnderTolerance, "과부족허용율(-)", 15, 1, 5
 		ggoSpread.SSSetEdit		C_LcAmdSeq, "AMEND 순번", 10, 2
 		ggoSpread.SSSetEdit		C_ChgFlg, "Chgfg", 1, 2
 		SetSpreadFloatLocal		C_OrgDocAmt, "C_OrgDocAmt", 15, 1, 2
 		SetSpreadFloatLocal		C_OrgDocAmt1, "C_OrgDocAmt1", 15, 1, 2

 		Call ggoSpread.SSSetColHidden(C_LcAmdSeq,C_OrgDocAmt1,True)
 		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)		
			
 		ggoSpread.SetCombo "U" & vbTab & "D", C_AmdFlg
			
 		SetSpreadLock "", 0, -1, ""

 		.vspdData.ReDraw = True
			
 	End With
 End Sub

<!--
'=====================================  2.2.4 SetSpreadLock()  =====================================
-->
 Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
     With frm1
 		ggoSpread.Source = .vspdData
			
 		.vspdData.ReDraw = False
		
 	    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
 	    ggoSpread.SpreadLock C_AmdFlgStr, lRow, 1
 		ggoSpread.SpreadLock C_ItemCd, lRow, 1
 		ggoSpread.SpreadLock C_ItemNm, lRow, 1
 		ggoSpread.SpreadLock C_Spec, lRow, 1
 		ggoSpread.SpreadLock C_Unit, lRow, 1
 		ggoSpread.SpreadLock C_BeQty, lRow, 1
 		ggoSpread.SpreadUnLock C_AtQty, lRow, 1
 		ggoSpread.SSSetRequired C_AtQty, lRow, lRow
 		ggoSpread.SpreadLock C_PORemainQty , lRow, -1  
 		ggoSpread.SpreadUnLock C_Price, lRow, 1
 		ggoSpread.SSSetRequired C_Price, lRow, lRow
 		ggoSpread.SpreadLock C_DocAmt, lRow, 1
 		ggoSpread.SpreadLock C_PORemainQty, lRow, 1
 		ggoSpread.SpreadLock C_HsCd, lRow, 1
 		ggoSpread.SpreadLock C_HsNm, lRow, 1
 		ggoSpread.SpreadLock C_AmendSeq, lRow, 1
 		ggoSpread.SpreadLock C_LcSeq, lRow, 1
 		ggoSpread.SpreadLock C_PoNo, lRow, 1
 		ggoSpread.SpreadLock C_PoSeq, lRow, 1
 		ggoSpread.SpreadLock C_OverTolerance, lRow, 1
 		ggoSpread.SpreadLock C_UnderTolerance, lRow, 1
 		ggoSpread.SpreadLock C_LcAmdSeq, lRow, 1
 		ggoSpread.SpreadLock C_ChgFlg, lRow, 1
			
 		.vspdData.ReDraw = True
 	End With
 End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 Dim iCurColumnPos
	
 Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData
			
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

 		C_AmdFlg			= iCurColumnPos(1)
 		C_AmdFlgstr			= iCurColumnPos(2)
 		C_ItemCd			= iCurColumnPos(3)
 		C_ItemNm			= iCurColumnPos(4)
 		C_Spec				= iCurColumnPos(5)
 		C_Unit				= iCurColumnPos(6)
 		C_BeQty				= iCurColumnPos(7)
 		C_AtQty				= iCurColumnPos(8)
 		C_Price				= iCurColumnPos(9)
 		C_DocAmt			= iCurColumnPos(10)
 		C_PORemainQty		= iCurColumnPos(11)
 		C_HsCd				= iCurColumnPos(12)
 		C_HsNm				= iCurColumnPos(13)
 		C_AmendSeq			= iCurColumnPos(14)
 		C_LcSeq				= iCurColumnPos(15)
 		C_PoNo				= iCurColumnPos(16)
 		C_PoSeq				= iCurColumnPos(17)
 		C_OverTolerance		= iCurColumnPos(18)
 		C_UnderTolerance	= iCurColumnPos(19)
 		C_LcAmdSeq			= iCurColumnPos(20)
 		C_ChgFlg			= iCurColumnPos(21)
 		C_OrgDocAmt			= iCurColumnPos(22)
 		C_OrgDocAmt1		= iCurColumnPos(23)
 End Select

End Sub	
	
<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
 Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 	ggoSpread.Source = frm1.vspdData

     With frm1.vspdData
	    
 		.Redraw = False

 	    ggoSpread.SSSetProtected frm1.vspddata.maxcols, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_AmdFlg, pvStartRow, pvEndRow
			
 		ggoSpread.SSSetProtected C_AmdFlgStr, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_BeQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_AtQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_Price, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_DocAmt, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_HsCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_HsNm, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LcSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_PoNo, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_PoSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_OverTolerance, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_UnderTolerance, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LcAmdSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ChgFlg, pvStartRow, pvEndRow

 		.Col = 1
 		.Row = .ActiveRow
 		.Action = 0
 		.EditMode = True

 		.ReDraw = True
 	End With
 End Sub

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenLCAmdNoPop()  ++++++++++++++++++++++++++++++++++++++
'+	Name : OpenLCAmdNoPop()																				+
'+	Description : Master L/C Amend No PopUp Call														+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenLCAmdNoPop()
 	Dim strRet,IntRetCD
 	Dim iCalledAspName
		
 	If gblnWinEvent = True Or UCase(frm1.txtLCAmdNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
 	gblnWinEvent = True
		
 	iCalledAspName = AskPRAspName("M3221PA1_KO441")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3221PA1_KO441", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	strRet = window.showModalDialog(iCalledAspName, Array(window.parent), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False
		
 	If strRet = "" Then
 		frm1.txtLCAmdNo.focus
 		Set gActiveElement = document.activeElement
 		Exit Function
 	Else
 		frm1.txtLCAmdNo.value = strRet
 		frm1.txtLCAmdNo.focus
 		Set gActiveElement = document.activeElement
 	End If	
 End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenLCDtlRef()  +++++++++++++++++++++++++++++++++++++++
'+	Name : OpenLCDtlRef()																				+
'+	Description : Open L/C Reference Window Call														+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenLCDtlRef()
 	Dim arrRet
 	Dim arrParam(13)
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	if lgIntFlgMode = Parent.OPMD_CMODE then
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	End If
		
 	arrParam(0) = UCase(Trim(frm1.hdnPoNo.value))
 	arrParam(1) = UCase(Trim(frm1.hdnPayMethCd.Value))
 	arrParam(2) = Trim(frm1.hdnPayMethNm.Value)
 	arrParam(3) = UCase(Trim(frm1.hdnIncotermsCd.Value))
 	arrParam(4) = Trim(frm1.hdnIncotermsNm.Value)
 	arrParam(5) = UCase(Trim(frm1.txtCurrency.Value))
 	arrParam(6) = UCase(Trim(frm1.txtBeneficiary.Value))
 	arrParam(7) = Trim(frm1.txtBeneficiaryNm.Value)
 	arrParam(8) = UCase(Trim(frm1.hdnGrpCd.Value))
 	arrParam(9) = Trim(frm1.hdnGrpNm.Value)
 	arrParam(10)= "M"
 	arrParam(11) = Trim(frm1.txtLCDocNo.value)
 	arrParam(12) = Trim(frm1.txtLCAmendSeq.value)
 	arrParam(13) = Trim(frm1.txtLCNo.Value)
		
 	iCalledAspName = AskPRAspName("M3212RA1")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212RA1", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If

 	If gblnWinEvent = True Then Exit Function
		
 	gblnWinEvent = True
		
 	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
 	gblnWinEvent = False

 	If arrRet(0, 0) = "" Then
 		frm1.txtLCAmdNo.focus
 		Set gActiveElement = document.activeElement
 		Exit Function
 	Else
 		Call SetLCDtlRef(arrRet)
 	End If	
 End Function
<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenPODtlRef()  +++++++++++++++++++++++++++++++++++++++
'+	Name : OpenPODtlRef()																				+
'+	Description : S/O Reference Window Call																+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function OpenPODtlRef()
 	Dim arrRet
 	Dim arrParam(10)
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	if lgIntFlgMode = Parent.OPMD_CMODE then
 		Call DisplayMsgBox("900002", "X", "X", "X")	
 		Exit Function
 	End If
        
 	arrParam(0) = UCase(Trim(frm1.hdnPoNo.value))
 	arrParam(1) = UCase(Trim(frm1.hdnPayMethCd.Value))
 	arrParam(2) = Trim(frm1.hdnPayMethNm.Value)
 	arrParam(3) = UCase(Trim(frm1.hdnIncotermsCd.Value))
 	arrParam(4) = Trim(frm1.hdnIncotermsNm.Value)
 	arrParam(5) = UCase(Trim(frm1.txtCurrency.Value))
 	arrParam(6) = UCase(Trim(frm1.txtBeneficiary.Value))
 	arrParam(7) = Trim(frm1.txtBeneficiaryNm.Value)
 	arrParam(8) = UCase(Trim(frm1.hdnGrpCd.Value))
 	arrParam(9) = Trim(frm1.hdnGrpNm.Value)
 	arrParam(10)= "M"

 	If gblnWinEvent = True Then Exit Function
		
 	gblnWinEvent = True

 	iCalledAspName = AskPRAspName("M3112RA1")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA1", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
 	gblnWinEvent = False
				
 	If arrRet(0, 0) = "" Then
 		frm1.txtLCAmdNo.focus
 		Set gActiveElement = document.activeElement
 		Exit Function
 	Else
 		Call SetPODtlRef(arrRet)
 	End If	
		
 End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetLCDtlRef()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : SetLCDtlRef()																				+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function SetLCDtlRef(arrRet)
 	Dim TempRow, I, j, Row1
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim intMinusCnt
 	Dim strMessage

 	Const C_Ref_ItemCd			= 0
 	Const C_Ref_ItemNm			= 1
 	Const C_Ref_LcQty			= 2
 	Const C_Ref_Spec			= 3
 	Const C_Ref_Unit			= 4
 	Const C_Ref_Price			= 5
 	Const C_Ref_DocAmt			= 6
 	Const C_Ref_LcSeq			= 7
 	Const C_Ref_PoNo			= 8
 	Const C_Ref_PoSeq			= 9
 	Const C_Ref_HsCd			= 10
 	Const C_Ref_HsNm			= 11
 	Const C_Ref_OverTolerance	= 12
 	Const C_Ref_UnderTolerance	= 13
 	Const C_Ref_PORemainQty		= 14
		
 	With frm1 
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows					
 		intLoopCnt = Ubound(arrRet, 1)				
			
 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_LcSeq

 					If .vspdData.Text = arrRet(intCnt - 1, C_Ref_LcSeq) Then
 						intMinusCnt = intMinusCnt + 1
 						strMessage = strMessage & arrRet(intCnt - 1, C_Ref_LcSeq) & " "
 						blnEqualFlg = True
 						Exit For
 					End If
 				Next
 			End If

 			If blnEqualFlg = False Then
					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				Row1 = .vspdData.Row
				
				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_AmdFlg,	Row1, "U")
				Call .vspdData.SetText(C_AmdFlgStr,	Row1, "내용변경")
 				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
 				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
 				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
 				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
 				Call .vspdData.SetText(C_BeQty,	Row1, arrRet(intCnt - 1, C_Ref_LcQty))
 				Call .vspdData.SetText(C_AtQty,	Row1, 0)	'arrRet(intCnt - 1, 2)
 				Call .vspdData.SetText(C_Price,	Row1, arrRet(intCnt - 1, C_Ref_Price))
 				Call .vspdData.SetText(C_DocAmt,	Row1, 0)	'arrRet(intCnt - 1, 9)
 				Call .vspdData.SetText(C_PORemainQty,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_PORemainQty),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
 				Call .vspdData.SetText(C_HsCd,	Row1, arrRet(intCnt - 1, C_Ref_HsCd))
 				Call .vspdData.SetText(C_HsNm,	Row1, arrRet(intCnt - 1, C_Ref_HsNm))
 				Call .vspdData.SetText(C_LcSeq,	Row1, arrRet(intCnt - 1, C_Ref_LcSeq))
 				Call .vspdData.SetText(C_PoNo,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
 				Call .vspdData.SetText(C_PoSeq,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
 				Call .vspdData.SetText(C_OverTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_OverTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_UnderTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_Ref_UnderTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
 				Call .vspdData.SetText(C_ChgFlg,	Row1, .vspdData.Row)

					
 				'SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)

 				'ggoSpread.spreadUnlock C_AmdFlg,CLng(TempRow) + CLng(intCnt), C_AmdFlg,CLng(TempRow) + CLng(intCnt)
					
 				'ggoSpread.SSSetRequired  C_AmdFlg, CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
					
 			End If
 		Next
		
		Call SetSpreadColor(CLng(TempRow)+1,.vspdData.MaxRows)	
		
		'200310 실제로 추가된 부분이 있어야만 그리드에 Coloring 작업을 수행한다.		
		.vspdData.Col = 0
		.vspdData.Row = .vspdData.MaxRows
		
		If Trim(.vspdData.text) = ggoSpread.InsertFlag then 		
			ggoSpread.spreadUnlock C_AmdFlg,CLng(TempRow) + 1, C_AmdFlg, .vspdData.MaxRows
			ggoSpread.SSSetRequired  C_AmdFlg, CLng(TempRow) + 1, .vspdData.MaxRows
		End if
						
 		if strMessage<>"" then
 			Call DisplayMsgBox("17a005", "X",strmessage,"L/C순번")
 			.vspdData.ReDraw = True
 			Exit Function
 		End if
			
 		.vspdData.ReDraw = True

 	End With
 End Function

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetPODtlRef()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : SetPODtlRef()																				+
'+	Description : Set Return array from S/O Reference Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
 Function SetPODtlRef(arrRet)
 	Dim TempRow, I, j, Row1
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim strMessage
 	Dim tempstr1,tempstr2
		
 	Const C_REF_ItemCd			= 0
 	Const C_REF_ItemNm			= 1
 	Const C_REF_LcQty			= 2
 	Const C_REF_ItemSpec		= 3
 	Const C_REF_Unit			= 4
 	Const C_REF_Price			= 5
 	Const C_REF_DocAmt			= 6
 	Const C_REF_PoNo			= 7
 	Const C_REF_PoSeq			= 8
 	Const C_REF_HsCd			= 9
 	Const C_REF_HsNm			= 10
 	Const C_REF_OverTolerance	= 11
 	Const C_REF_UnderTolerance	= 12

 	With frm1 
 		.vspdData.focus
 		ggoSpread.Source = .vspdData
 		.vspdData.ReDraw = False	

 		TempRow = .vspdData.MaxRows							
 		intLoopCnt = Ubound(arrRet, 1)						
			
 		For intCnt = 1 to intLoopCnt + 1
 			blnEqualFlg = False

 			If TempRow <> 0 Then
 				For j = 1 To TempRow
 					.vspdData.Row = j
 					.vspdData.Col = C_PoNo
 					tempstr1 = .vspdData.text
 					.vspdData.Col = C_PoSeq
 					tempstr2 = .vspdData.text
 					If tempstr1 = arrRet(intCnt - 1, C_REF_PoNo) And tempstr2 = arrRet(intCnt - 1, C_REF_PoSeq) Then
 						blnEqualFlg = True
 						strMessage = strMessage & arrRet(intCnt - 1, C_REF_PoNo) & "-" & arrRet(intCnt - 1, C_REF_PoSeq) & " "
 						Exit For
 					End If
 				Next
 			End If

 			If blnEqualFlg = False Then
					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = CLng(TempRow) + CLng(intCnt)
				Row1 = .vspdData.Row
				
				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_AmdFlg,	Row1, "C")
				Call .vspdData.SetText(C_AmdFlgStr,	Row1, "품목추가")
 				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_REF_ItemCd))
 				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_REF_ItemNm))
 				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_REF_ItemSpec))
 				Call .vspdData.SetText(C_PORemainQty,	Row1, arrRet(intCnt - 1, C_REF_LcQty))
 				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_REF_Unit))
 				Call .vspdData.SetText(C_AtQty,	Row1, arrRet(intCnt - 1, C_REF_LcQty))
 				Call .vspdData.SetText(C_Price,	Row1, arrRet(intCnt - 1, C_REF_Price))
 				Call .vspdData.SetText(C_DocAmt,	Row1, 0)	'arrRet(intCnt - 1, 6)
 				Call .vspdData.SetText(C_HsCd,	Row1, arrRet(intCnt - 1, C_REF_HsCd))
 				Call .vspdData.SetText(C_HsNm,	Row1, arrRet(intCnt - 1, C_REF_HsNm))
 				Call .vspdData.SetText(C_PoNo,	Row1, arrRet(intCnt - 1, C_REF_PoNo))
 				Call .vspdData.SetText(C_PoSeq,	Row1, arrRet(intCnt - 1, C_REF_PoSeq))
 				Call .vspdData.SetText(C_OverTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_REF_OverTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
				Call .vspdData.SetText(C_UnderTolerance,	Row1, UNIFormatNumber(arrRet(intCnt - 1, C_REF_UnderTolerance),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit))
 				Call .vspdData.SetText(C_ChgFlg,	Row1, .vspdData.Row)
					
 				'수정(이은희)-2003.04.08
 				Call vspdData_Change(C_AtQty_Ref, .vspdData.Row)
					
 			End If
 		Next
			
 		Call SetSpreadColor(CLng(TempRow)+1,.vspdData.MaxRows)
 		Call TotalSum()
			
 		if strMessage<>"" then
 			Call DisplayMsgBox("17a005", "X",strmessage,"발주번호" & "," & "발주순번")
 			.vspdData.ReDraw = True
 			Exit Function
 		End if
			
 		.vspdData.ReDraw = True

 	End With
 End Function

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()
 ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()
 With frm1

 	ggoSpread.Source = frm1.vspdData
 	'단가 
 	ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	'금액 
 	ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt1,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

 End With

End Sub
<!--
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
-->
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                 ByVal dColWidth , ByVal HAlign , _
                 ByVal iFlag )
	        
Select Case iFlag
     Case 2                                                              '금액 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
     Case 3                                                              '수량 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
     Case 4                                                              '단가 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
     Case 5                                                              '환율 
         ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
 End Select
         
End Sub
<!--
'==========================================  2.5.1 LoadLCAmendHdr()  ====================================
'=	Event Name : LoadLCAmendHdr																			=
'========================================================================================================
-->
 Function LoadLCAmendHdr()
 	Dim strDtlOpenParam
 	Dim IntRetCD

     If lgIntFlgMode <> Parent.OPMD_UMODE Then                  
         Call DisplayMsgBox("900002", "X", "X", "X")
         Exit Function
     End if
	    	
     If ggoSpread.SSCheckChange = true Then
 		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
     End If

 	WriteCookie "LCAmdNo", UCase(Trim(frm1.txtLCAmdNo.value))

 	PgmJump(LCAMEND_HEADER_ENTRY_ID)

 End Function

<!--
'============================================  2.5.2 OpenCookie()  ======================================
-->
 Function OpenCookie()
 Dim strTemp	
 	strTemp = ReadCookie("LCAmdNo")
 	frm1.txtLCAmdNo.Value = strTemp
 	WriteCookie "LCAmdNo", ""
 	If Trim(frm1.txtLCAmdNo.value) <> "" Then
 		Call dbQuery()
 	End If
		
 End Function

<!--
'============================================  2.5.1 TotalSum()  ======================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
-->
 Sub TotalSum()
 	Dim SumTotal, lRow
		
 	SumTotal = UNICDbl(frm1.txtDocAmt.Text)
 	ggoSpread.source = frm1.vspdData
		
 	For lRow = 1 To frm1.vspdData.MaxRows 		
 		frm1.vspdData.Row = lRow
 		frm1.vspdData.Col = 0
 		If frm1.vspdData.Text = ggoSpread.InsertFlag then
 			frm1.vspdData.Col = C_DocAmt
 			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
 		end if
 	Next
		
 	frm1.txtDocAmt.text = UNIConvNumPCToCompanyByCurrency(CStr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

 End Sub

'########################################################################################
'============================================  2.5.1 TotalSumNew()  ======================================
'=	Name : TotalSumNew()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
Sub TotalSumNew(ByVal row)
		
    Dim SumTotal, lRow, tmpGrossAmt

	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_DocAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = C_OrgDocAmt							
	SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))
        
    frm1.txtDocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
End Sub
'######################################################################################
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
 	Call OpenCookie()
 	'Call SetToolbar("1110110100101111")

 End Sub
	
<!--
'=========================================  3.1.2 Form_QueryUnload()  ===================================
-->
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub
	
'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 gMouseClickStatus = "SPC"   
 Set gActiveSpdSheet = frm1.vspdData
	   
 'Call SetPopupMenuItemInf("0101111111")
 If lgIntFlgMode = Parent.OPMD_UMODE Then
 	If frm1.vspddata.maxRows > 0 Then
 		Call SetPopupMenuItemInf("0101111111")
 	Else	
 		Call SetPopupMenuItemInf("0001111111")
 	End If
 Else	
 	Call SetPopupMenuItemInf("0000111111")
 End If
	
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

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

If Button = 2 And gMouseClickStatus = "SPC" Then
   gMouseClickStatus = "SPCR"
End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
 ggoSpread.Source = frm1.vspdData
 Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
 ggoSpread.Source = frm1.vspdData
 Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
 Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
 If Row <= 0 Then
 	Exit Sub
 End If
 If frm1.vspddata.MaxRows=0 Then
 	Exit Sub
 End If	
End Sub

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Function FncSplitColumn()
    
If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
    Exit Function
 End If

 ggoSpread.Source = gActiveSpdSheet
 ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
 ggoSpread.Source = gActiveSpdSheet
 Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
 ggoSpread.Source = gActiveSpdSheet
    
 Call ggoSpread.RestoreSpreadInf()
 Call InitSpreadSheet()      
 Call ggoSpread.ReOrderingSpreadData()
 Call SetSpreadColor(1, frm1.vspdData.MaxRows)
End Sub

<!--
'======================================  3.2.1 btnLCAmdNoOnClick()  ====================================
-->
 Sub btnLCAmdNoOnClick()
 	Call OpenLCAmdNoPop()
 End Sub
	
<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
 Sub vspdData_Change(ByVal Col, ByVal Row )
 	Dim Qty, Price, DocAmt
	
 	ggoSpread.Source = frm1.vspdData
 	ggoSpread.UpdateRow Row

 	Frm1.vspdData.Row = Row
 	Frm1.vspdData.Col = Col
		
 	'공통함수로 변경 - 2002/10/14 KSH			
 	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
		
 	Select Case col
 	Case C_AtQty, C_AtQty_Ref
 		frm1.vspdData.Col = C_AtQty
 		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
 			Qty = 0
 		Else
 			Qty = UNICDbl(frm1.vspdData.Text)
 		End If
			
 		frm1.vspdData.Col = C_Price
 		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
 			Price = 0
 		Else
 			Price = UNICDbl(frm1.vspdData.Text)
 		End If
			
 		DocAmt = Qty * Price
 		frm1.vspdData.Col = C_DocAmt
 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(DocAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
 		If col <> C_AtQty_Ref Then
 			Call TotalSumNew(Row)					'총품목금액합계 
 		End If
 		'총금액계산을 위해 필요(2003.05)
 		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt
			
 	Case C_Price
 		frm1.vspdData.Col = C_Price
 		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
 			Price = 0
 		Else
 			Price = UNICDbl(frm1.vspdData.Text)
 		End If
			
 		frm1.vspdData.Col = C_AtQty
 		If Trim(frm1.vspdData.Text) = "" OR IsNull(frm1.vspdData.Text) then
 			Qty = 0
 		Else
 			Qty = UNICDbl(frm1.vspdData.Text)
 		End If
			
 		DocAmt = Qty * Price
 		frm1.vspdData.Col = C_DocAmt
 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(DocAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

 		Call TotalSumNew(Row)					'총품목금액합계 
 		'총금액계산을 위해 필요(2003.05)
 		frm1.vspdData.Col = C_DocAmt
		DocAmt = frm1.vspdData.Text
		frm1.vspdData.Col = C_OrgDocAmt		
		frm1.vspdData.Text = DocAmt
 	Case C_AmdFlg
			
 		frm1.vspdData.Row = Row
 		frm1.vspdData.Col = C_AmdFlg
					
 		Select Case Trim(frm1.vspdData.text)			
 			Case "U"	
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AmdFlgStr
							
 				frm1.vspdData.text = "내용변경"
					
 				ggoSpread.spreadUnlock C_AtQty,frm1.vspdData.ActiveRow,C_Price,frm1.vspdData.ActiveRow
 				ggoSpread.SSSetRequired  C_AtQty, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
 				ggoSpread.SSSetRequired  C_Price, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
						
 			Case "D"
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AmdFlgStr
 				frm1.vspdData.text = "품목삭제"
					
 				frm1.vspdData.Col = C_AtQty
 				frm1.vspdData.text = 0
					
 				frm1.vspdData.Col = C_Price
 				frm1.vspdData.text = 0
					
 				frm1.vspdData.Col = C_DocAmt
 				frm1.vspdData.text = 0
					
 				Call ggoSpread.DeleteRow
					
 				ggoSpread.SSSetProtected C_AtQty, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
 				ggoSpread.SSSetProtected C_Price, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
					
 			Case "C"
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AmdFlgStr
							
 				frm1.vspdData.text = "품목추가"
				
 				ggoSpread.spreadUnlock C_AtQty,frm1.vspdData.ActiveRow,C_Price,frm1.vspdData.ActiveRow
 				ggoSpread.SSSetRequired  C_AtQty, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
 				ggoSpread.SSSetRequired  C_Price, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
					
 			Case Else
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AmdFlg
 				frm1.vspdData.text = ""
	
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AmdFlgStr
 				frm1.vspdData.text = ""
					
 		End Select
			
 	end select
 End Sub
	
<!--
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
-->
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
 	With frm1.vspdData
 		If Row >= NewRow Then
 			Exit Sub
 		End If

 		If NewRow = .MaxRows Then
 			If lgStrPrevKey <> "" Then							
 				DbQuery
 			End If
 		End If
 	End With
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
		
 	ggoSpread.Source = frm1.vspdData
		
 	If ggoSpread.SSCheckChange = true Then
 		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
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
		
 	frm1.hdnQueryType.Value = "Query"
				
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

 	ggoSpread.Source = frm1.vspdData
		
 	If ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")	
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
		
 	ggoSpread.Source = frm1.vspdData
		
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then							
 		Call DisplayMsgBox("900002", "X", "X")
 		Exit Function
 	End If

		
 	Call DbDelete											

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
    
 	If ggoSpread.SSCheckChange = False  Then  
 	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")  
 	    Exit Function
 	End If
    
 	'If Not chkField(Document, "2") Then           
 	'   Exit Function
 	'End If

 	ggoSpread.Source = frm1.vspdData              
 	If Not ggoSpread.SSDefaultCheck  Then  
 	   Exit Function
 	End If
		
 	If DbSave = False Then Exit Function
		
 	If frm1.txtHLCAmdNo.value <> frm1.txtLCAmdNo.value then			'---?????---
 		frm1.txtLCAmdNo.value =	frm1.txtHLCAmdNo.value		
 	End If
		
 	FncSave = True											
	Set gActiveElement = document.activeElement		
 End Function

<!--
'===========================================  5.1.5 FncCopy()  ==========================================
-->
 Function FncCopy()
		
 	Dim IntRetCD

 	ggoSpread.Source = frm1.vspdData
		
 	lgIntFlgMode = Parent.OPMD_CMODE								

 	frm1.vspdData.ReDraw = False
 	if frm1.vspdData.Maxrows < 1	then exit function

 	ggoSpread.Source = frm1.vspdData	
 	ggoSpread.CopyRow
 	SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

 	frm1.vspdData.ReDraw = True
	Set gActiveElement = document.activeElement	
 End Function

<!--
'===========================================  5.1.6 FncCancel()  ========================================
-->
 Function FncCancel() 
	Dim SumTotal,tmpGrossAmt,orgtmpGrossAmt, Row, CUDflag
 	if frm1.vspdData.Maxrows < 1	then exit function
 	'총금액계산수정(2003.05.28)
	'---------------------------------------------
    SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	Row = frm1.vspdData.SelBlockRow
		
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_DocAmt
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = C_OrgDocAmt1
	orgtmpGrossAmt = UNICDbl(frm1.vspdData.Text)
	    
	frm1.vspdData.Col = 0
	CUDflag = frm1.vspdData.Text
				
    If CUDflag = ggoSpread.UpdateFlag Then
        SumTotal = SumTotal + (orgtmpGrossAmt - tmpGrossAmt )
    ElseIf CUDflag = ggoSpread.InsertFlag  Then
        SumTotal = SumTotal - tmpGrossAmt
    End If

	frm1.txtDocAmt.Text = SumTotal
	'--------------------------------------------
 	ggoSpread.Source = frm1.vspdData
	
     ggoSpread.EditUndo   
     Set gActiveElement = document.activeElement 
 End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
 Function FncInsertRow()
				
 	With frm1
 		.vspdData.focus
 		ggoSpread.Source = .vspdData

 		'.vspdData.EditMode = True

 		.vspdData.ReDraw = False
 		ggoSpread.InsertRow
 		.vspdData.ReDraw = True

 		SetSpreadColor .vspdData.ActiveRow,.vspdData.ActiveRow
     End With
	 Set gActiveElement = document.activeElement   
 End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
 Function FncDeleteRow()
		
 	Dim lDelRows
 	Dim iDelRowCnt, i
 	Dim index, count
		
 	if frm1.vspdData.Maxrows < 1	then exit function
 	With frm1.vspdData 
	
 		.focus
 		ggoSpread.Source = frm1.vspdData

 		lDelRows = ggoSpread.DeleteRow
			
 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
 		count = 1
			
 		'for index = .SelBlockRow to .SelBlockRow2
 		'	frm1.vspdData.Row = index
 		'	frm1.vspdData.Col = C_AmdFlg
 		'	frm1.vspdData.text = "D"
 		'	frm1.vspdData.Col = C_AmdFlgStr
 		'	frm1.vspdData.text = "내용삭제"
 		'Next
			
 	End With
	Set gActiveElement = document.activeElement	
 End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
 Function FncPrint()
 	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
 End Function

<!--
'============================================  5.1.10 FncPrev()  ========================================
-->
 Function FncPrev() 
		
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then					
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	ElseIf lgPrevNo = "" Then							
 		Call DisplayMsgBox("900011", "X", "X", "X")
 	End If
 	Set gActiveElement = document.activeElement
 End Function

<!--
'============================================  5.1.11 FncNext()  ========================================
-->
 Function FncNext()
 	ggoSpread.Source = frm1.vspdData
		
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then			
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	ElseIf lgNextNo = "" Then					
 		Call DisplayMsgBox("900012", "X", "X", "X")
 	End If
 	Set gActiveElement = document.activeElement
 End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
 Function FncExcel() 
 	ggoSpread.Source = frm1.vspdData
 	Call parent.FncExport(Parent.C_MULTI)
 	Set gActiveElement = document.activeElement
 End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
 Function FncFind() 
 	ggoSpread.Source = frm1.vspdData
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
 		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X") 
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

 	if LayerShowHide(1) =false then
 	    exit Function
 	end if

 	If lgIntFlgMode = Parent.OPMD_UMODE Then
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001				
 		strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtHLCAmdNo.value)	 		
 	Else
 		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001				
 		strVal = strVal & "&txtLCAmdNo=" & Trim(frm1.txtLCAmdNo.value)	
 	End If
 	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
 	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	'수정(2003.06.10)
	strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)	
 	strVal = strVal & "&txtQueryType=" & Trim(frm1.hdnQueryType.value)

 	frm1.hdnmaxrow.value = frm1.vspdData.MaxRows

 	Call RunMyBizASP(MyBizASP, strVal)									
	
 	DbQuery = True														
 End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
 Function DbSave() 
 	Dim lRow
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
				
 	DbSave = False														
    
    ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '초기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '초기 버퍼의 설정[삭제]
   
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
 	if LayerShowHide(1) =false then
 	    exit Function
 	end if

 	With frm1
 		.txtMode.value = Parent.UID_M0002

 		strVal = ""
 		strDel = ""

 		For lRow = 1 To .vspdData.MaxRows
 			.vspdData.Row = lRow
 			.vspdData.Col = 0

 			Select Case .vspdData.Text
 				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag	
 				
 				If .vspdData.Text=ggoSpread.InsertFlag Then
					strVal = "C" & ColSep	'0
					.vspdData.Col = C_AmdFlg		
 					If Trim(.vspdData.Text)="" Then
 						strVal = strVal & "C" & ColSep
 					Else
 						strVal = strVal & Trim(.vspdData.Text) & ColSep
 					End If
				Else
					strVal = "U" & ColSep
					
					.vspdData.Col = C_AmdFlg								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep
				End If 							
						
 					.vspdData.Col = C_AmdFlg
 					if Trim(.vspdData.Text) <> "D" then
 						.vspdData.Col = C_AtQty
 						if Trim(UNICDbl(.vspdData.Text)) = "0" or Trim(UNICDbl(.vspdData.Text)) = "" then
 							Call DisplayMsgBox("970021", "X","변경후수량", "X")
 							Call SetActiveCell(frm1.vspdData,C_AtQty,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End if
						
 						.vspdData.Col = C_Price
 						if Trim(UNICDbl(.vspdData.Text)) = "0" or Trim(UNICDbl(.vspdData.Text)) = "" then
 							Call DisplayMsgBox("970021", "X","단가", "X")
 							Call SetActiveCell(frm1.vspdData,C_Price,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End if
 					End if

 					.vspdData.Col = C_ItemCd								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_Unit								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep
						
 					.vspdData.Col = C_BeQty								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_AtQty								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_Price								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_DocAmt								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					'local amount
 					strVal = strVal & "0" & ColSep

 					.vspdData.Col = C_HsCd								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_AmendSeq								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_LcSeq								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_PoNo								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_PoSeq								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_OverTolerance								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_UnderTolerance								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_ChgFlg								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep
						
 	                strVal = strVal & lRow & RowSep

 				Case ggoSpread.DeleteFlag	
 					strDel = "D" & ColSep	
					
					strDel = strDel & ColSep & ColSep & ColSep
					
					.vspdData.Col = C_BeQty								
 					strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_AtQty								
 					strDel = strDel & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

					strDel = strDel & ColSep & ColSep & ColSep & ColSep 
					
					.vspdData.Col = C_AmendSeq								
 					strDel = strDel & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_LcSeq								
 					strDel = strDel & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_PoNo								
 					strDel = strDel & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_PoSeq								
 					strDel = strDel & Trim(.vspdData.Text) & ColSep
 					
					strDel = strDel & ColSep & ColSep & ColSep 
					
					strDel = strDel & lRow & RowSep
								
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
 End Function
	
<!--
'=============================================  5.2.4 DbQueryOk()  ======================================
-->
 Function DbQueryOk()								
 	Dim index

 	lgIntFlgMode = Parent.OPMD_UMODE						

 	lgBlnFlgChgValue = False

 	'Call TotalSum									
	Call RemovedivTextArea
	
 	Call ggoOper.LockField(Document, "Q")			
		
 	if frm1.vspdData.MaxRows > 0 then
 		Call SetToolbar("11101011000111")
 		frm1.vspdData.focus
 	else
 		Call SetToolbar("11101001000111")
 		frm1.txtLCAmdNo.focus
 	end if
	
		
 	frm1.vspdData.ReDraw = False
 	For index = cint(frm1.hdnmaxrow.value)+1 to frm1.vspdData.MaxRows
 		frm1.vspdData.Row = index
 		frm1.vspdData.Col = C_AmdFlg
 		If Trim(frm1.vspdData.Text) = "D" then
 			ggoSpread.SpreadLock C_AmdFlg , index,frm1.vspdData.MaxCols, index
 		End If
 	Next
		
 	ggoSpread.SpreadLock C_AmdFlg , -1,C_AmdFlg
		
 	frm1.vspdData.ReDraw = True
		
 End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
 Function DbSaveOk()									
 	Call InitVariables
 	frm1.vspdData.MaxRows = 0
 	Call MainQuery()
 End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
 Function DbDeleteOk()								
'		Call FncNew()
 End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C AMEND 내역</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenLCDtlRef">L/C내역참조</A>&nbsp;|&nbsp;<A><A href="vbscript:OpenPODtlRef">발주내역참조</A></TD>
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
										<TD CLASS=TD5 NOWRAP>L/C AMEND 관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo"  SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="L/C AMEND 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCAmdNoOnClick()"></TD>
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
									<TD CLASS=TD5 NOWRAP>L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C번호" TYPE=TEXT MAXLENGTH=35 SIZE=29 TAG="24XXXU">&nbsp;-
														 <INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수출자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총AMEND금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="통화">&nbsp;
												</TD>
												<TD NOWRAP>
													<script language =javascript src='./js/m3222ma1_fpDoubleSingle5_txtDocAmt.js'></script>
												</TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3222ma1_fpDateTime1_txtAmendDt.js'></script></TD>
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									    <script language =javascript src='./js/m3222ma1_I667966985_vspdData.js'></script>
									</TD>
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
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE<%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadLCAmendHdr">L/C AMEND등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<P ID="divTextArea"></P>
	
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMethNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnGrpNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
