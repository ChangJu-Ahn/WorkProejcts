<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3222ma2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Import Local L/C Amend Detail 등록 ASP									*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/05/02																*
'*  8. Modified date(Last)  : 2003/05/22																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
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
	

	Const BIZ_PGM_QRY_ID = "m3222mb5_KO441.asp"	
	Const BIZ_PGM_SAVE_ID = "m3222mb6_KO441.asp"	
	Const LCAMEND_HEADER_ENTRY_ID = "m3221ma2"
	Const BIZ_PGM_CAL_AMT_ID = "m3211mb10.asp"
	
	Dim C_LCAmdFlg
	Dim C_LCAmdFlgDtl
	Dim C_ItemCd
	Dim C_ItemNm
	Dim C_SPEC
	Dim C_Unit
	Dim C_BeQty
	Dim C_AtQty
	Dim C_AtPrice
	Dim C_AtDocAmt
	Dim C_AtLocAmt
	Dim C_PORemainQty
	Dim C_HsCd
	Dim C_HsNm
	Dim C_LCAmendSeq
	Dim C_LCSeq
	Dim C_PONo
	Dim C_POSeq
	Dim C_OverTolerance
	Dim C_UnderTolerance
	Dim C_ChgFlg
	'총품목금액계산을 위해 추가(2003.05)
	Dim C_OrgDocAmt		'변화값 저장 
	Dim C_OrgDocAmt1	'조회후 초기값 저장 
	
	'참조시 사용(2003.04.08)-Lee Eun Hee
	Dim C_AtQty_Ref

<!-- #Include file="../../inc/lgvariables.inc" -->
	
	Dim gblnWinEvent
	Dim dblAmt
	
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
 	frm1.txtTotDocAmt.text = UNIFormatNumber(UNICDbl(0),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
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

'=========================================  2.2.3	InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()

	C_LCAmdFlg		= 1
	C_LCAmdFlgDtl	= 2
	C_ItemCd		= 3				
	C_ItemNm		= 4
	C_SPEC			= 5
	C_Unit			= 6
	C_BeQty			= 7
	C_AtQty			= 8
	C_AtPrice		= 9
	C_AtDocAmt		= 10
	C_AtLocAmt		= 11
	C_PORemainQty	= 12
	C_HsCd			= 13
	C_HsNm			= 14
	C_LCAmendSeq	= 15
	C_LCSeq			= 16
	C_PONo			= 17
	C_POSeq			= 18
	C_OverTolerance	= 19
	C_UnderTolerance= 20
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

 		ggoSpread.SSSetCombo		C_LCAmdFlg, "변경구분", 10, 0, False
 		ggoSpread.SSSetEdit			C_LCAmdFlgDtl, "변경내용", 10, 0
 		ggoSpread.SSSetEdit			C_ItemCd, "품목", 18, 0
 		ggoSpread.SSSetEdit			C_ItemNm, "품목명", 20, 0
 		ggoSpread.SSSetEdit			C_SPEC, "품목규격", 20, 0
 		ggoSpread.SSSetEdit			C_Unit, "단위", 10, 2
 		SetSpreadFloatLocal			C_BeQty,  "변경전수량", 15, 1, 3
 		SetSpreadFloatLocal			C_AtQty,  "변경후수량", 15, 1, 3
 		SetSpreadFloatLocal			C_AtPrice, "단가", 15, 1, 4
 		SetSpreadFloatLocal			C_AtDocAmt, "금액", 15, 1, 2
 		SetSpreadFloatLocal			C_AtLocAmt, "원화금액", 15, 1, 2
 		SetSpreadFloatLocal			C_PORemainQty,  "발주잔량", 15, 1, 3
 		ggoSpread.SSSetEdit			C_HsCd, "HS부호", 20, 0
 		ggoSpread.SSSetEdit			C_HsNm, "HS명", 20, 0
 		ggoSpread.SSSetEdit			C_LCAmendSeq, "AMEND순번", 10, 2
 		ggoSpread.SSSetEdit			C_LCSeq, "L/C순번", 10, 2
 		ggoSpread.SSSetEdit			C_PONo, "발주번호", 18, 0
 		ggoSpread.SSSetEdit			C_POSeq, "발주순번", 10, 2
 		SetSpreadFloatLocal			C_OverTolerance, "과부족허용율(+)", 15, 1, 5
 		SetSpreadFloatLocal			C_UnderTolerance, "과부족허용율(-)", 15, 1, 5
 		ggoSpread.SSSetEdit			C_ChgFlg, "Chgfg", 1, 2
		SetSpreadFloatLocal			C_OrgDocAmt, "C_OrgDocAmt", 15, 1, 2
		SetSpreadFloatLocal			C_OrgDocAmt1, "C_OrgDocAmt1", 15, 1, 2
			
 		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_OrgDocAmt1,True)	
 		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)	

 		ggoSpread.SetCombo "U" & vbTab & "D", C_LCAmdFlg
 		SetSpreadLock "", 0, -1, ""

 		.vspdData.ReDraw = True
 	End With
 End Sub

<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
 Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
     With frm1
 		ggoSpread.Source = .vspdData
			
 		.vspdData.ReDraw = False

 	    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
 		ggoSpread.SpreadLock C_LCAmdFlgDtl, lRow, -1
 		ggoSpread.SpreadLock C_ItemCd, lRow, -1
 		ggoSpread.SpreadLock C_ItemNm , lRow, -1
 		ggoSpread.SpreadLock C_SPEC , lRow, -1
 		ggoSpread.SpreadLock C_Unit , lRow, -1
 		ggoSpread.SpreadLock C_BeQty , lRow, -1
 		ggoSpread.SpreadUnLock C_AtQty, lRow, -1 
 		ggoSpread.SSSetRequired C_AtQty, lRow, lRow
 		ggoSpread.SSSetRequired C_AtPrice, lRow, lRow
 		ggoSpread.SSSetRequired C_AtDocAmt, lRow, lRow
 		ggoSpread.SSSetRequired C_AtLocAmt, lRow, -1
 		ggoSpread.SSSetRequired C_AtLocAmt, lRow, lRow
 		ggoSpread.SpreadLock C_PORemainQty , lRow, -1  
 		ggoSpread.SpreadLock C_HsCd, lRow, -1
 		ggoSpread.SpreadLock C_HsNm, lRow, -1
 		ggoSpread.SpreadLock C_LCAmendSeq, lRow, -1
 		ggoSpread.SpreadLock C_LCSeq, lRow, -1
 		ggoSpread.SpreadLock C_PONo, lRow, -1
 		ggoSpread.SpreadLock C_PoSeq, lRow, -1
 		ggoSpread.SpreadLock C_OverTolerance, lRow, -1
 		ggoSpread.SpreadLock C_UnderTolerance, lRow, -1
			
 		.vspdData.ReDraw = True
 	End With
 End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
 Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 	ggoSpread.Source = frm1.vspdData

     With frm1.vspdData
 		.Redraw = False
 	    ggoSpread.SSSetProtected frm1.vspddata.maxcols, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LCAmdFlg, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LCAmdFlgDtl, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_SPEC, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_Unit, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_BeQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_AtQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_AtPrice, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_AtDocAmt, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired  C_AtLocAmt, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_PORemainQty, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_HsCd, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_HsNm, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LCAmendSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_LCSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_PoNo, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_PoSeq, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_OverTolerance, pvStartRow, pvEndRow
 		ggoSpread.SSSetProtected C_UnderTolerance, pvStartRow, pvEndRow
 		.Col = 1
 		.Row = .ActiveRow
 		.Action = 0
 		.EditMode = True
 		.ReDraw = True
 	End With
 End Sub
<!--
'==========================================  GetSpreadColumnPos()  ====================================
-->
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 Dim iCurColumnPos
	
 Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		C_LCAmdFlg		= iCurColumnPos(1)
 		C_LCAmdFlgDtl	= iCurColumnPos(2)
 		C_ItemCd		= iCurColumnPos(3)
 		C_ItemNm		= iCurColumnPos(4)
 		C_SPEC			= iCurColumnPos(5)
 		C_Unit			= iCurColumnPos(6)
 		C_BeQty			= iCurColumnPos(7)
 		C_AtQty			= iCurColumnPos(8)
 		C_AtPrice		= iCurColumnPos(9)
 		C_AtDocAmt		= iCurColumnPos(10)
 		C_AtLocAmt		= iCurColumnPos(11)
 		C_PORemainQty	= iCurColumnPos(12)
 		C_HsCd			= iCurColumnPos(13)
 		C_HsNm			= iCurColumnPos(14)
 		C_LCAmendSeq	= iCurColumnPos(15)
 		C_LCSeq			= iCurColumnPos(16)
 		C_PONo			= iCurColumnPos(17)
 		C_POSeq			= iCurColumnPos(18)
 		C_OverTolerance	= iCurColumnPos(19)
 		C_UnderTolerance= iCurColumnPos(20)
 		C_ChgFlg		= iCurColumnPos(21)
 		C_OrgDocAmt		= iCurColumnPos(22)
 End Select

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
		
 	iCalledAspName = AskPRAspName("M3221PA2_KO441")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3221PA2_KO441", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
		
 	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
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
 	Dim arrParam(10)
 	Dim iCalledAspName
 	Dim IntRetCD
		
 	If Trim(frm1.txtLCNo.value) = "" Then
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	End If
		
 	arrParam(0) = Trim(frm1.txtLCDocNo.value)
 	arrParam(1) = Trim(frm1.txtLCAmendSeq.value)	
 	arrParam(2) = Trim(frm1.txtHPurGrp.value)					
 	arrParam(3) = Trim(frm1.txtHPurGrpNm.value)	
 	arrParam(4) = Trim(frm1.txtBeneficiary.value)			
 	arrParam(5) = Trim(frm1.txtBeneficiaryNm.value)	
 	arrParam(6) = Trim(frm1.txtCurrency.value) 									
 	arrParam(7) = Trim(frm1.txtHPayTerms.value)	
 	arrParam(8) = Trim(frm1.txtHPayTermsNm.value)
 	arrParam(9) = Trim(frm1.txtLCNo.value)
		
 	iCalledAspName = AskPRAspName("M3212RA2")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3212RA2", "X")
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
		
 	If Trim(frm1.txtLCNo.value) = "" Then
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	End If
		
 	arrParam(0) = Trim(frm1.txtHPurGrp.value)					
 	arrParam(1) = Trim(frm1.txtHPurGrpNm.value)	
 	arrParam(2) = Trim(frm1.txtBeneficiary.value)			
 	arrParam(3) = Trim(frm1.txtBeneficiaryNm.value)	
 	arrParam(4) = Trim(frm1.txtCurrency.value) 									
 	arrParam(5) = Trim(frm1.txtHPayTerms.value)	
 	arrParam(6) = Trim(frm1.txtHPayTermsNm.value)
 	arrParam(7) = Trim(frm1.txtPONo.value)
		
 	If gblnWinEvent = True Then Exit Function
		
 	gblnWinEvent = True
				
 	iCalledAspName = AskPRAspName("M3112RA3")

 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA3", "X")
 		gblnWinEvent = False
 		Exit Function
 	End If
		
 	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
 			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
 	gblnWinEvent = False

 	If arrRet(0, 0) = ""  Then
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
 	Dim temp
 	Dim TempRow, I, j, intEndRow, Row1
 	Dim blnEqualFlg
 	Dim intLoopCnt
 	Dim intCnt
 	Dim strMessage

 	Const C_Ref_ItemCd	= 0
 	Const C_Ref_ItemNm	= 1
 	Const C_Ref_LcQty	= 2	
 	Const C_Ref_Spec	= 3
 	Const C_Ref_Unit	= 4
 	Const C_Ref_Price	= 5
 	Const C_Ref_DocAmt	= 6
 	Const C_Ref_LcSeq	= 7
 	Const C_Ref_PoNo	= 8
 	Const C_Ref_PoSeq	= 9
 	Const C_Ref_PoQty	= 10
 	Const C_Ref_HsCd	= 11
 	Const C_Ref_LcNo	= 12

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
 					.vspdData.Col = C_LCSeq
						
 					If .vspdData.Text = arrRet(intCnt - 1, C_Ref_LcSeq) Then
 						strMessage = arrRet(intCnt - 1, C_Ref_LcSeq) 
 						blnEqualFlg = True
 						Exit For
 					End If
					
 				Next
 			End If

 			If blnEqualFlg = False Then
					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = .vspdData.MaxRows	
 				Row1 = .vspdData.Row
					
				'C_Ref_PoQty는 히든필드(2003.06.13)	
 				temp = UNIFormatNumber(CDbl(arrRet(intCnt - 1, C_Ref_PoQty)) - UNICDbl(arrRet(intCnt - 1, C_Ref_LcQty)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
					
 				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
 				Call .vspdData.SetText(C_LCAmdFlg,	Row1, "U")
 				Call .vspdData.SetText(C_LCAmdFlgDtl,	Row1, "내용변경")
 				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
 				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
 				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
 				Call .vspdData.SetText(C_BeQty,	Row1, arrRet(intCnt - 1, C_Ref_LcQty))
 				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
 				Call .vspdData.SetText(C_AtPrice,	Row1, arrRet(intCnt - 1, C_Ref_Price))
 				Call .vspdData.SetText(C_LCSeq,	Row1, arrRet(intCnt - 1, C_Ref_LcSeq))
 				Call .vspdData.SetText(C_PONo,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
 				Call .vspdData.SetText(C_POSeq,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
 				Call .vspdData.SetText(C_PORemainQty,	Row1, temp)
 				Call .vspdData.SetText(C_HSCd,	Row1, arrRet(intCnt - 1, C_Ref_HsCd))
 				'요기!!!!
 				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
										
 				SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)
					
 				ggoSpread.spreadUnlock C_LCAmdFlg,CLng(TempRow) + CLng(intCnt),C_LCAmdFlg,CLng(TempRow) + CLng(intCnt)
					
 				ggoSpread.SSSetRequired  C_LCAmdFlg, CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
										
 			End If
 		Next
 		Call SetToolbar("11101011000000")			
			
 		if strMessage<>"" then
 			Call DisplayMsgBox("17a005", "X",strMessage,"L/C순번")
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

 	Const C_Ref_ItemCd			= 0
 	Const C_Ref_ItemNm			= 1
 	Const C_Ref_PORemainQty		= 2
 	Const C_Ref_Spec			= 3 
 	Const C_Ref_Unit			= 4 
 	Const C_Ref_Price			= 5
 	Const C_Ref_DocAmt			= 6
 	Const C_Ref_PoNo			= 7
 	Const C_Ref_PoSeq			= 8
 	Const C_Ref_HsCd			= 9
 	Const C_Ref_OverTolerance	= 10
 	Const C_Ref_UnderTolerance	= 11

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
 					.vspdData.Col = C_PoSeq

 					If .vspdData.Text = arrRet(intCnt - 1, C_Ref_PoSeq) Then
 						.vspdData.Row = j
 						.vspdData.Col = C_PoNo
 						If .vspdData.Text = arrRet(intCnt - 1, C_Ref_PoNo) Then
 							blnEqualFlg = True
 							strMessage = arrRet(intCnt - 1, C_Ref_PoNo) & "-" & arrRet(intCnt - 1, C_Ref_PoSeq)
 							Exit For
 						End if
 					End If						
 				Next
 			End If

 			If blnEqualFlg = False Then
					
 				.vspdData.MaxRows = .vspdData.MaxRows + 1
 				.vspdData.Row = .vspdData.MaxRows
 				Row1 = .vspdData.Row
					
 				Call .vspdData.SetText(0       ,	Row1, ggoSpread.InsertFlag)
 				Call .vspdData.SetText(C_LCAmdFlg,	Row1, "C")
 				Call .vspdData.SetText(C_LCAmdFlgDtl,	Row1, "품목추가")
 				Call .vspdData.SetText(C_ItemCd,	Row1, arrRet(intCnt - 1, C_Ref_ItemCd))
 				Call .vspdData.SetText(C_ItemNm,	Row1, arrRet(intCnt - 1, C_Ref_ItemNm))
 				Call .vspdData.SetText(C_Spec,	Row1, arrRet(intCnt - 1, C_Ref_Spec))
 				Call .vspdData.SetText(C_PORemainQty,	Row1, arrRet(intCnt - 1, C_Ref_PORemainQty))
 				Call .vspdData.SetText(C_AtQty,	Row1, arrRet(intCnt - 1, C_Ref_PORemainQty))
 				Call .vspdData.SetText(C_Unit,	Row1, arrRet(intCnt - 1, C_Ref_Unit))
 				Call .vspdData.SetText(C_AtPrice,	Row1, arrRet(intCnt - 1, C_Ref_Price))
 				Call .vspdData.SetText(C_AtDocAmt,	Row1, arrRet(intCnt - 1, C_Ref_DocAmt))
 				Call .vspdData.SetText(C_PONo,	Row1, arrRet(intCnt - 1, C_Ref_PoNo))
 				Call .vspdData.SetText(C_POSeq,	Row1, arrRet(intCnt - 1, C_Ref_PoSeq))
 				'Call .vspdData.SetText(C_ChgFlg,	Row1, temp)
					
 				.vspdData.Col = C_ChgFlg					
 				.vspdData.text = .vspdData.Row

 				Call vspdData_Change(C_AtQty_Ref, .vspdData.Row)				

 				'SetSpreadColor CLng(TempRow) + CLng(intCnt),CLng(TempRow) + CLng(intCnt)
 			End If
 		Next
			
 		Call SetSpreadColor(CLng(TempRow)+1,.vspdData.MaxRows)
 		Call TotalSum()
 		Call SetToolbar("11101011000000")			
			
 		if strMessage<>"" then
 			Call DisplayMsgBox("17a005", "X",strMessage,"발주번호" & "," & "발주순번")
 			.vspdData.ReDraw = True
 			Exit Function
 		End if
 		.vspdData.ReDraw = True

 	End With
 End Function

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

 With frm1
 	ggoOper.FormatFieldByObjectOfCur .txtTotDocAmt, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
 End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

 With frm1

 	ggoSpread.Source = frm1.vspdData
 	'단가 
 	ggoSpread.SSSetFloatByCellOfCur C_AtPrice,-1, .txtCurrency.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	'금액 
 	ggoSpread.SSSetFloatByCellOfCur C_AtDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
 	ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt1,-1, .txtCurrency.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

 End With

End Sub
<!--
'===================================  3.2.33 SetSpreadDeleteRow() =======================================
-->
 Sub SetSpreadDeleteRow(ByVal lRow)
    With frm1
 		ggoSpread.Source = .vspdData
			
 		ggoSpread.SSSetProtected C_AtQty, lRow, lRow
 		ggoSpread.SSSetProtected C_AtPrice, lRow, lRow
 		ggoSpread.SSSetProtected C_AtDocAmt, lRow, lRow
 		ggoSpread.SSSetProtected C_AtLocAmt, lRow, lRow
		
 	End With
 End Sub
<!--
'===================================  3.2.33 SetReleaseDeleteRow() =======================================
-->
 Sub SetReleaseDeleteRow(ByVal lRow)
    With frm1
 		ggoSpread.Source = .vspdData
			
 		.vspdData.ReDraw = False

 		ggoSpread.SpreadUnLock C_AtQty, lRow, C_AtQty, lRow 
 		ggoSpread.SpreadUnLock C_AtPrice, lRow, C_AtPrice, lRow  
 		ggoSpread.SpreadUnLock C_AtDocAmt, lRow, C_AtDocAmt, lRow  
 		ggoSpread.SpreadUnLock C_AtLocAmt, lRow, C_AtLocAmt, lRow 
 		ggoSpread.SSSetRequired C_AtQty, lRow, lRow
 		ggoSpread.SSSetRequired C_AtPrice, lRow, lRow
 		ggoSpread.SSSetRequired C_AtDocAmt, lRow, lRow
 		ggoSpread.SSSetRequired C_AtLocAmt, lRow, lRow
	
 		.vspdData.ReDraw = True
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
'==========================================  2.5.1 LoadLCAmendHdr()  ====================================
-->
 Function LoadLCAmendHdr()
 	Dim strDtlOpenParam
 	Dim IntRetCD

     If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
         Call DisplayMsgBox("900002", "X", "X", "X")
         Exit Function
     End if
	    	
     If lgBlnFlgChgValue = True Then
 		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
     End If

 	WriteCookie "txtLCAmdNo", UCase(Trim(frm1.txtLCAmdNo.value))

 	PgmJump(LCAMEND_HEADER_ENTRY_ID)

 End Function

<!--
'============================================  2.5.2 OpenCookie()  ======================================
-->
 Function OpenCookie()		
 	frm1.txtLCAmdNo.Value = ReadCookie("txtLCAmdNo")
 	frm1.hdnQueryType.Value = "autoQuery"
		
 	WriteCookie "txtLCAmdNo", ""
 End Function

<!--
'============================================  2.5.1 TotalSum()  ========================================
'=	Name : TotalSum()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
-->
 Sub TotalSum()
 	Dim SumTotal, lRow
		
 	SumTotal = UNICDbl(frm1.txtTotDocAmt.Text)
 	ggoSpread.source = frm1.vspdData
 	For lRow = 1 To frm1.vspdData.MaxRows 		
 		frm1.vspdData.Row = lRow
 		frm1.vspdData.Col = 0
 		If frm1.vspdData.Text = ggoSpread.InsertFlag then
 			frm1.vspdData.Col = C_AtDocAmt
 			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
 		End If
 	Next
		
 	frm1.txtTotDocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal,frm1.txtCurrency.value,Parent.ggAmtOfMoneyNo,"X","X")'UNIFormatNumber(CStr(SumTotal),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)


 End Sub
 '########################################################################################
'============================================  2.5.1 TotalSumNew()  ======================================
'=	Name : TotalSumNew()																					=
'=	Description : Master L/C Header 화면으로부터 넘겨받은 parameter setting(Cookie 사용)				=
'========================================================================================================
Sub TotalSumNew(ByVal row)
		
    Dim SumTotal, lRow, tmpGrossAmt

	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtTotDocAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_AtDocAmt				
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = C_OrgDocAmt							
	SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))

        
    frm1.txtTotDocAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo, "X" , "X")
	
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
		
 	Call InitVariables
 	Call SetDefaultVal	
 	Call OpenCookie()
		
 	If UCase(Trim(frm1.txtLCAmdNo.value)) <> "" Then
 		Call dbQuery()
 	End If
		
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
'=	Event Name : btnLCAmdNoOnClick																		=
'========================================================================================================
-->
 Sub btnLCAmdNoOnClick()
 	Call OpenLCAmdNoPop()
 End Sub
<!--
'*********************************************  환율계산  **********************************************
'* Change Event 처리																		*
'********************************************************************************************************
 Sub TxtdblAmt(ByVal Row)
		
 	DIM RateAmt
		
 	frm1.vspdData.Row = Row
 	frm1.vspdData.Col = C_AtDocAmt
		
 	IF Trim(frm1.hdnDiv.value) = "*" THEN
			
 		RateAmt = UNICDbl(frm1.vspdData.text) * UNICDbl(frm1.txtXchRate.value)
 		frm1.vspdData.Col = C_AtLocAmt
 		'frm1.vspdData.Text = UNIFormatNumber(cstr(RateAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(RateAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
 	ELSEIF Trim(frm1.hdnDiv.value) = "/" THEN
			
 		RateAmt = UNICDbl(frm1.vspdData.text) / UNICDbl(frm1.txtXchRate.value)
 		frm1.vspdData.Col = C_AtLocAmt
 		'frm1.vspdData.Text = UNIFormatNumber(cstr(RateAmt),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
 		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(RateAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
 	END IF

 End Sub	
-->
	
<!--
'==========================================  3.3.1 vspdData_Change()  ===================================
-->
 Sub vspdData_Change(ByVal Col, ByVal Row )
 	Dim dblQty
 	Dim dblPrice, DocAmt
 	Dim iwhere
 	Dim strVal
 	Dim Todate

 	lgBlnFlgChgValue = True

 	ggoSpread.Source = frm1.vspdData
 	ggoSpread.UpdateRow Row

 	Select Case Col
 		Case C_AtQty, C_AtQty_Ref				
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_AtQty

 			dblQty = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_AtPrice

 			dblPrice = frm1.vspdData.Text

 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_AtDocAmt
				
 			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(dblAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
				
 			If frm1.txtCurrency.value = Parent.gCurrency Then
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AtLocAmt
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(dblAmt,Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 			Else
 				'Todate = "<%=EndDate%>"
					
 				'strVal = BIZ_PGM_CAL_AMT_ID & "?txtCurrency=" & Trim(frm1.txtCurrency.value)
 				'strVal = strVal & "&txtApplDt=" & Todate									
 				'strVal = strVal & "&txtXchRate=" & Trim(frm1.txtXchRate.Value)
 				'strVal = strVal & "&txtDocAmt=" & Trim(dblAmt)
 				'strVal = strVal & "&txtLocCurrency=" & Parent.gCurrency
 				'strVal = strVal & "&Row=" & Row			
 				'strVal = strVal & "&txtAmendFlg=AMEND"
 				'Call RunMyBizASP(MyBizASP, strVal)											

 				'frm1.vspdData.Row = Row
 				'frm1.vspdData.Col = C_AtLocAmt

 				Call TxtdblAmt(Row)
 				'Call TotalSum
 			End If
 			If col <> C_AtQty_Ref Then
 				Call TotalSumNew(Row)
 			End If
 			'총금액계산을 위해 필요(2003.05)
			frm1.vspdData.Col = C_AtDocAmt
			DocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgDocAmt		
			frm1.vspdData.Text = DocAmt
 			
 		Case C_AtPrice 
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col

 			dblPrice = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_AtQty

 			dblQty = frm1.vspdData.Text

 			dblAmt = UNICDbl(dblQty) * UNICDbl(dblPrice)

 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = C_AtDocAmt
 			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(dblAmt, frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
				
 			dblAmt = frm1.vspdData.Text				
				
 			If frm1.txtCurrency.value = Parent.gCurrency Then
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AtLocAmt
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(dblAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 			Else
 				'Todate = "<%=EndDate%>"
					
 				'strVal = BIZ_PGM_CAL_AMT_ID & "?txtCurrency=" & Trim(frm1.txtCurrency.value)		
 				'strVal = strVal & "&txtApplDt=" & Todate											
 				'strVal = strVal & "&txtXchRate=" & Trim(frm1.txtXchRate.value)
 				'strVal = strVal & "&txtDocAmt=" & Trim(dblAmt)
 				'strVal = strVal & "&txtLocCurrency=" & Parent.gCurrency
 				'strVal = strVal & "&Row=" & Row			
 				'strVal = strVal & "&txtAmendFlg=AMEND"
 				'Call RunMyBizASP(MyBizASP, strVal)													

 				'frm1.vspdData.Row = Row
 				'frm1.vspdData.Col = C_AtLocAmt

 				Call TxtdblAmt(Row)
 				
 			End If
 			Call TotalSumNew(Row)
 			'총금액계산을 위해 필요(2003.05)
			frm1.vspdData.Col = C_AtDocAmt
			DocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgDocAmt		
			frm1.vspdData.Text = DocAmt
			
 		Case C_AtDocAmt 
 			'frm1.vspdData.Row = Row
 			'frm1.vspdData.Col = Col

 			'dblPrice = frm1.vspdData.Text

 			'frm1.vspdData.Row = Row
 			'frm1.vspddata.Col = C_AtQty

 			'dblQty = frm1.vspdData.Text

 			frm1.vspdData.Row = Row
 			frm1.vspddata.Col = C_AtDocAmt
				
 			dblAmt = frm1.vspdData.Text		

 			If frm1.txtCurrency.value = Parent.gCurrency Then
 				frm1.vspdData.Row = Row
 				frm1.vspdData.Col = C_AtLocAmt
 				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(dblAmt),Parent.gCurrency, Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo , "X")
 			Else

 				'Todate = "<%=EndDate%>"
					
 				'strVal = BIZ_PGM_CAL_AMT_ID & "?txtCurrency=" & Trim(frm1.txtCurrency.value)		
 				'strVal = strVal & "&txtApplDt=" & Todate											
 				'strVal = strVal & "&txtXchRate=" & Trim(frm1.txtXchRate.value)
 				'strVal = strVal & "&txtDocAmt=" & Trim(dblAmt)
 				'strVal = strVal & "&txtLocCurrency=" & Parent.gCurrency
 				'strVal = strVal & "&Row=" & Row			
 				'strVal = strVal & "&txtAmendFlg=AMEND"
 				'Call RunMyBizASP(MyBizASP, strVal)													

 				'frm1.vspdData.Row = Row
 				'frm1.vspdData.Col = C_AtLocAmt

 				Call TxtdblAmt(Row)
					
 				'Call TotalSum
 			End If
 			Call TotalSumNew(Row)
 			'총금액계산을 위해 필요(2003.05)
			frm1.vspdData.Col = C_AtDocAmt
			DocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgDocAmt		
			frm1.vspdData.Text = DocAmt
 		'Case C_AtLocAmt
				
 		'	frm1.vspdData.Row = Row
 		'	frm1.vspdData.Col = C_AtLocAmt
				
 		'	frm1.txtTotDocAmt.Text = frm1.vspdData.Text

 		Case C_LCAmdFlg
 			frm1.vspdData.Row = Row
 			frm1.vspdData.Col = Col
				
 			iwhere = frm1.vspdData.text 
			
 			Select Case iwhere									
 				Case "U"	
 					frm1.vspdData.Row = Row
 					frm1.vspdData.Col = C_LCAmdFlgDtl
								
 					frm1.vspdData.text = "내용변경"
 					Call SetReleaseDeleteRow(Row)	
							
 				Case "D"
 					frm1.vspdData.Row = Row
 					frm1.vspdData.Col = C_LCAmdFlgDtl
								
 					frm1.vspdData.text = "품목삭제"
 					frm1.vspdData.Col = C_AtQty
 					frm1.vspdData.text = 0
						
 					frm1.vspdData.Col = C_AtPrice
 					frm1.vspdData.text = 0
						
 					frm1.vspdData.Col = C_AtDocAmt
 					frm1.vspdData.text = 0
						
 					frm1.vspdData.Col = C_AtLocAmt
 					frm1.vspdData.text = 0
						
 					ggoSpread.SSSetProtected C_AtQty, Row, Row
 					ggoSpread.SSSetProtected C_AtPrice, Row, Row
 					ggoSpread.SSSetProtected C_AtDocAmt, Row, Row
 					ggoSpread.SSSetProtected C_AtLocAmt, Row, Row
 					'Call SetSpreadDeleteRow(Row)
					
 				Case "C"
	
 				Case Else
 					frm1.vspdData.Row = Row
 					frm1.vspdData.Col = C_LCAmdFlg
 					frm1.vspdData.text = ""
		
 					frm1.vspdData.Row = Row
 					frm1.vspdData.Col = C_LCAmdFlgDtl
 					frm1.vspdData.text = ""
 			End Select
 	End Select
 	'Call TotalSum
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

 	If lgBlnFlgChgValue = True Then
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

 	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
 		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")	<% '⊙: "Will you destory previous data" %>
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

 	Call ggoOper.ClearField(Document, "A")				
 	Call ggoOper.LockField(Document, "N")				
 	Call InitVariables									
		
 	Call SetDefaultVal

 	FncNew = True										
 	Set gActiveElement = document.activeElement
 End Function
	
<!--
'===========================================  5.1.3 FncDelete()  ========================================
-->
 Function FncDelete()
		
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then					
 		Call DisplayMsgBox("900002", "X", "X", "X")
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
    
 	If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  
 	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")            
 	    Exit Function
 	End If

    
 	'If Not chkField(Document, "2") Then               
 	'   Exit Function
 	'End If

 	ggoSpread.Source = frm1.vspdData                  
 	If Not ggoSpread.SSDefaultCheck         Then      
 	   Exit Function
 	End If
		
 	If DbSave = False Then Exit Function
				
 	If frm1.txtHLCAmdNo.value <> frm1.txtLCAmdNo.value then
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

 	If lgBlnFlgChgValue = True Then
 		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
 		If IntRetCD = vbNo Then
 			Exit Function
 		End If
 	End If

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
	SumTotal = UNICDbl(frm1.txtTotDocAmt.Text)
	Row = frm1.vspdData.SelBlockRow
			
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = C_AtDocAmt
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

	frm1.txtTotDocAmt.Text = SumTotal
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
	
 	if frm1.vspdData.Maxrows < 1	then exit function
 	With frm1.vspdData 
	
 		.focus
 		ggoSpread.Source = frm1.vspdData

 		lDelRows = ggoSpread.DeleteRow

 		lgBlnFlgChgValue = True
 	End With
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
	
 	If lgIntFlgMode <> Parent.OPMD_UMODE Then			
 		Call DisplayMsgBox("900002", "X", "X", "X")
 		Exit Function
 	ElseIf lgNextNo = "" Then					
 		Call DisplayMsgBox("900012", "X", "X", "X")
 	End If
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

 	If lgBlnFlgChgValue = True Then
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
	    
     lgIntFlgMode = Parent.OPMD_UMODE
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
	
 	On Error Resume Next											

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
					.vspdData.Col = C_LCAmdFlg		
 					If Trim(.vspdData.Text)="" Then
 						strVal = strVal & "C" & ColSep
 					Else
 						strVal = strVal & Trim(.vspdData.Text) & ColSep
 					End If
				Else
					strVal = "U" & ColSep
					
					.vspdData.Col = C_LCAmdFlg								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep
				End If 	
				'---------------------------		
						
 					.vspdData.Col = C_ItemCd							
 					strVal = strVal & Trim(.vspdData.Text) & ColSep
						
 					.vspdData.Col = C_Unit								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_BeQty								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_AtQty								
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
						
 					.vspdData.Col = C_LCAmdFlg
 					If Trim(.vspdData.Text) <> "D" Then
 						.vspdData.Col = C_AtQty	
 						If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
 							Call DisplayMsgBox("970021", "X","변경후수량", "X")
 							Call SetActiveCell(frm1.vspdData,C_AtQty,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End If
 						.vspdData.Col = C_AtPrice
 						If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" Then
 							Call DisplayMsgBox("970021", "X","단가", "X")
 							Call SetActiveCell(frm1.vspdData,C_AtPrice,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End If
 					End If
						
 					.vspdData.Col = C_AtPrice							
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep

 					.vspdData.Col = C_AtDocAmt							
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
						
 					.vspdData.Col = C_LCAmdFlg
 					if Trim(.vspdData.Text) <> "D" Then
 						.vspdData.Col = C_AtDocAmt					
 						if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 							Call DisplayMsgBox("970021", "X","금액", "X")
 							Call SetActiveCell(frm1.vspdData,C_AtDocAmt,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End if
 					End if								

 					.vspdData.Col = C_AtLocAmt							
 					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & ColSep
						
 					.vspdData.Col = C_LCAmdFlg
 					if Trim(.vspdData.Text) <> "D" Then
 						.vspdData.Col =  C_AtLocAmt		
 						if Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) = "0" then
 							Call DisplayMsgBox("970021", "X","원화금액", "X")
 							Call SetActiveCell(frm1.vspdData,C_AtLocAmt,lRow,"M","X","X")
 							Call LayerShowHide(0)
 							Exit Function
 						End if
 					End if		

 					.vspdData.Col = C_HsCd								
 					strVal = strVal & Trim(.vspdData.Text) & ColSep

 					.vspdData.Col = C_LCAmendSeq						
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
					
					.vspdData.Col = C_LCAmendSeq								
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
		
 	lgIntFlgMode = Parent.OPMD_UMODE										

 	lgBlnFlgChgValue = False

 	'Call TotalSum
	Call RemovedivTextArea
	
 	With frm1
 		ggoSpread.Source = .vspdData
		
 		Dim TempRow, i
		
 		TempRow = .vspdData.MaxRows									

 		.vspdData.ReDraw = False
 		For i = cInt(frm1.hdnmaxrow.value)+1 to TempRow 
 			ggoSpread.SSSetProtected C_LCAmdFlg, i, i	
				
 			.vspdData.Row = i
 			.vspdData.Col = C_LCAmdFlg
				
 			If .vspdData.text = "D" Then
 				Call SetSpreadDeleteRow(i) 				
 			End If
 		Next
 		.vspdData.ReDraw = True
 	End With
		
 	Call ggoOper.LockField(Document, "Q")							
 	Call SetToolbar("11101011000111")								
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtLCAmdNo.focus
	End If
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
						<TD CLASS="CLSLTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>LOCAL L/C AMEND 내역</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenLCDtlRef">LOCAL L/C 내역참조</A>&nbsp;|&nbsp;<A><A href="vbscript:OpenPODtlRef">발주내역참조</A></TD>
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
										<TD CLASS=TD5 NOWRAP>LOCAL L/C AMEND 관리번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCAmdNo"  SIZE=20 MAXLENGTH=18 TAG="12XXXU" ALT="LOCAL L/C AMEND 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCAmdNo" ALIGN=top TYPE="BUTTON" onclick="vbscript:btnLCAmdNoOnClick()"></TD>
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
									<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL L/C번호" TYPE=TEXT MAXLENGTH=35  SIZE=28  TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
									<TD CLASS=TD5 NOWRAP>수혜자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총AMEND금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="통화"></TD>
												<TD>&nbsp;<script language =javascript src='./js/m3222ma2_fpDoubleSingle1_txtTotDocAmt.js'></script></TD>	
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>AMEND일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3222ma2_fpDateTime_txtAmendDt.js'></script></TD>
								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<script language =javascript src='./js/m3222ma2_vaSpread_vspdData.js'></script>
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
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadLCAmendHdr()">L/C AMEND등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
<P ID="divTextArea"></P>
	
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPONo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCAmdNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCDocNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHBeneficiary" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHTotDocAmt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHAmendDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPODtlRefFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurGrp" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurGrpNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTermsNm" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
