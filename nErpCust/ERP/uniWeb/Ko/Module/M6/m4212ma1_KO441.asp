<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M4212ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 수입통관 Detail 등록 ASP													*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2003/05/29																*
'*  9. Modifier (First)     : Sun-joung Lee	 Kim JH														*
'* 10. Modifier (Last)      : Jin-hyun Shin	 Ma Jin Ha(2002/2/8) Kim Jin Ha								*
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
	
	Const BIZ_PGM_QRY_ID = "m4212mb1_KO441.asp"		
	Const BIZ_PGM_SAVE_ID = "m4212mb2_KO441.asp"		
	Const CC_HEADER_ENTRY_ID = "m4211ma1"
	Const CC_LAN_ENTRY_ID = "m4213ma1"			
	Const CHARGE_HDR_ENTRY_ID = "m6111ma2"		
	
	<!-- #Include file="../../inc/lgvariables.inc" -->
	
	
	Dim gblnWinEvent
	
	Dim C_ItemCd 								'품목코드 
	Dim C_ItemNm 								'품목명 
	Dim C_Spec	 								'품목규격 
	Dim C_TrackingNo 							'Tracking No	
	Dim C_Unit 									'단위 
	Dim C_CcQty 								'통관수량 
	Dim C_Price 								'단가 
	Dim C_DocAmt 								'금액 
	Dim C_NetWeight 							'순중량 
	Dim C_CIFDocAmt 							'CIF금액(US)
	Dim C_CIFLocAmt 							'CIF원화금액 
	Dim C_HsCd 								    'H/S번호 
	Dim C_HsNm 									'H/S명 
	Dim C_BlQty 								'B/L수량 
	Dim C_InputQty								'입고수량 
	Dim C_CcSeq 							    '통관순번 
	Dim C_BlNo 									'B/L관리번호 
	Dim C_BlSeq 								'B/L순번 
	Dim C_BlDocNo 								'B/L문서번호 
	Dim C_PoNo 									'P/O번호 
	Dim C_PoSeq 								'P/O순번 
	Dim C_LcNo 									'L/C번호 
	Dim C_LcSeq 								'L/C순번 
	Dim C_BlAmt 								'B/L금액 
	Dim C_BlCcQty 								'B/L내역 통관수량 
	'총품목금액계산을 위해 추가(2003.06.02)
	Dim C_OrgDocAmt								'변화값 저장 
	Dim C_OrgDocAmt1							'조회후 초기값 저장 
	
	Dim C_CcQty_Ref1				

'==================================  initSpreadPosVariables()  =====================================================
Sub InitSpreadPosVariables()
	 C_ItemCd		= 1							'품목코드 
	 C_ItemNm		= 2							'품목명 
	 C_Spec			= 3							'품목규격 
	 C_TrackingNo	= 4							'Tracking No	
	 C_Unit			= 5							'단위 
	 C_CcQty		= 6							'통관수량 
	 C_Price		= 7							'단가 
	 C_DocAmt		= 8							'금액 
	 C_NetWeight	= 9							'순중량 
	 C_CIFDocAmt	= 10						'CIF금액(US)
	 C_CIFLocAmt	= 11						'CIF원화금액 
	 C_HsCd			= 12						'H/S번호 
	 C_HsNm			= 13						'H/S명 
	 C_BlQty		= 14						'B/L수량 
	 C_InputQty		= 15						'입고수량 
	 C_CcSeq		= 16						'통관순번 
	 C_BlNo			= 17						'B/L관리번호 
	 C_BlSeq		= 18						'B/L순번 
	 C_BlDocNo		= 19						'B/L문서번호 
	 C_PoNo			= 20						'P/O번호 
	 C_PoSeq		= 21						'P/O순번 
	 C_LcNo			= 22						'L/C번호 
	 C_LcSeq		= 23						'L/C순번 
	 C_BlAmt		= 24						'B/L금액 
	 C_BlCcQty		= 25						'B/L내역 통관수량 
	 C_OrgDocAmt	= 26				
	 C_OrgDocAmt1	= 27

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
		ggoSpread.Spreadinit "V20030602",,Parent.gAllowDragDropSpread  
		.vspdData.ReDraw = False
		
		.vspdData.MaxCols = C_OrgDocAmt1 + 1
		.vspdData.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
    	
		ggoSpread.SSSetEdit		C_ItemCd,		"품목", 18, 0
		ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 20, 0
		ggoSpread.SSSetEdit		C_Spec,			"품목규격", 25, 0
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 15, 0
		ggoSpread.SSSetEdit		C_Unit,			"단위", 10, 2
		SetSpreadFloatLocal 	C_CcQty,		"통관수량", 15, 1, 3
		SetSpreadFloatLocal 	C_Price,		"단가",15, 1, 4
		SetSpreadFloatLocal 	C_DocAmt,		"통관금액",15, 1, 2
		SetSpreadFloatLocal 	C_NetWeight,	"순중량",15, 1, 3
		SetSpreadFloatLocal 	C_CIFDocAmt,	"CIF금액(US)",15, 1, 2
		SetSpreadFloatLocal 	C_CIFLocAmt,	"CIF자국금액",15, 1, 2
		ggoSpread.SSSetEdit		C_HsCd,			"H/S부호", 20, 0
		ggoSpread.SSSetEdit		C_HsNm,			"H/S명", 20, 0
		SetSpreadFloatLocal		C_BlQty,		"B/L수량",15, 1, 3
		SetSpreadFloatLocal		C_InputQty,		"입고수량",15,1, 3
		ggoSpread.SSSetEdit		C_CcSeq,		"통관순번", 10, 2
		ggoSpread.SSSetEdit		C_BlNo,			"B/L관리번호", 18, 0
		ggoSpread.SSSetEdit		C_BlSeq,		"B/L순번", 10, 2
		ggoSpread.SSSetEdit		C_BlDocNo,		"B/L번호", 20, 0
		ggoSpread.SSSetEdit		C_PoNo,			"발주번호", 18, 0
		ggoSpread.SSSetEdit		C_PoSeq,		"발주순번", 10, 2
		ggoSpread.SSSetEdit		C_LcNo,			"L/C관리번호", 20, 0
		ggoSpread.SSSetEdit		C_LcSeq,		"L/C순번", 10, 2
		SetSpreadFloatLocal 	C_BlAmt,		"B/L금액",15, 1, 2
		SetSpreadFloatLocal		C_BlCcQty,		"B/L통관수량",15, 1, 3
		'ggoSpread.SSSetEdit	C_BlCcQty + 1,	"", 10, 0
		SetSpreadFloatLocal		C_OrgDocAmt,	"C_OrgDocAmt",15,1,2
		SetSpreadFloatLocal		C_OrgDocAmt1,	"C_OrgDocAmt1",15,1,2
		
		Call ggoSpread.SSSetColHidden(C_CIFDocAmt,C_CIFLocAmt,True)	
		Call ggoSpread.SSSetColHidden(C_BlAmt,C_OrgDocAmt1,True)
		
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,.vspdData.MaxCols,True)	
		
		'Call ggoSpread.SSSetSplit2(2)
		
		.vspdData.ReDraw = True
		Call SetSpreadLock()
		
	End With
End Sub

<!--
'==========================================  2.2.4 SetSpreadLock()  =====================================
-->
Sub SetSpreadLock()
    
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
		 
    With ggoSpread
		.SpreadLock C_ItemCd,-1,		C_ItemCd,-1
		.SpreadLock C_ItemNm,-1,		C_ItemNm,-1
		.SpreadLock C_Spec,-1,			C_Spec,-1
		.SpreadLock C_TrackingNo,-1,	C_TrackingNo,-1
		.SpreadLock C_Unit,	-1,			C_Unit,	-1
		
		.SpreadLock C_HsCd,	-1,			C_HsCd,	-1
		.SpreadLock C_HsNm,	-1,			C_HsNm,	-1
		.SpreadLock C_BlQty,-1,			C_BlQty,-1
		.SpreadLock C_InputQty,	-1,		C_InputQty,	-1
		.SpreadLock C_CcSeq,-1,			C_CcSeq,-1
		.SpreadLock C_BlNo,	-1,			C_BlNo,	-1
		.SpreadLock C_BlSeq,-1,			C_BlSeq,-1
		.SpreadLock C_BlDocNo, -1,		C_BlDocNo,-1
		.SpreadLock C_PoNo,	-1,			C_PoNo,	-1
		.SpreadLock C_PoSeq,-1,			C_PoSeq,-1
		.SpreadLock C_LcNo,	-1,			C_LcNo,	-1
		.SpreadLock C_LcSeq,-1,			C_LcSeq,-1
		
		.SSSetProtected frm1.vspdData.MaxCols, -1
    End With
    '수정(화면성능개선관련)-2003.04.03-Lee Eun Hee
    Call SetSpreadColor(-1,-1)
    frm1.vspdData.ReDraw = True
    
End Sub

<!--
'==========================================  2.2.5 SetSpreadColor()  ====================================
-->
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.Source = frm1.vspdData
	 
    With frm1.vspdData
	    '수정(화면성능개선관련)-2003.04.03-Lee Eun Hee
		'.Redraw = False
		
		ggoSpread.SSSetProtected C_ItemCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Spec,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrackingNo,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Unit,			pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected C_HsCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_HsNm,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BlQty,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InputQty,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_CcSeq,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BlNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BlSeq,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BlDocNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PoSeq,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LcNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_LcSeq,			pvStartRow, pvEndRow
		
		ggoSpread.SSSetRequired C_CcQty,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_Price,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_DocAmt,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_NetWeight,		pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected C_Price,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_DocAmt,			pvStartRow, pvEndRow
		
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStartRow, pvEndRow
		'수정(화면성능개선관련)-2003.04.03-Lee Eun Hee
		'.ReDraw = True
	End With
End Sub

'===========================  GetSpreadColumnPos()  ================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Spec				= iCurColumnPos(3)
			C_TrackingNo		= iCurColumnPos(4)
			C_Unit				= iCurColumnPos(5)
			C_CcQty				= iCurColumnPos(6)
			C_Price				= iCurColumnPos(7)
			C_DocAmt			= iCurColumnPos(8)
			C_NetWeight			= iCurColumnPos(9)
			C_CIFDocAmt			= iCurColumnPos(10)
			C_CIFLocAmt			= iCurColumnPos(11)
			C_HsCd				= iCurColumnPos(12)
			C_HsNm				= iCurColumnPos(13)
			C_BlQty				= iCurColumnPos(14)
			C_InputQty			= iCurColumnPos(15)
			C_CcSeq				= iCurColumnPos(16)
			C_BlNo				= iCurColumnPos(17)
			C_BlSeq				= iCurColumnPos(18)
			C_BlDocNo			= iCurColumnPos(19)
			C_PoNo				= iCurColumnPos(20)
			C_PoSeq				= iCurColumnPos(21)
			C_LcNo				= iCurColumnPos(22)
			C_LcSeq				= iCurColumnPos(23)
			C_BlAmt				= iCurColumnPos(24)
			C_BlCcQty			= iCurColumnPos(25)
			C_OrgDocAmt			= iCurColumnPos(26)
			C_OrgDocAmt1		= iCurColumnPos(27)
		
	End Select

End Sub	

<!--
'++++++++++++++++++++++++++++++++++++++++++++++  OpenCcNoPop()  ++++++++++++++++++++++++++++++++++++++
-->
Function OpenCcNoPop()
	
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtCCNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
   	iCalledAspName = AskPRAspName("M4211PA1_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4211PA1_KO441", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	gblnWinEvent = False
		
	If arrRet = "" Then
		frm1.txtCCNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtCCNo.value = arrRet
		frm1.txtCCNo.focus
		Set gActiveElement = document.activeElement
	End If
			
End Function

<!--
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenBlDtlRef()  +++++++++++++++++++++++++++++++++++++++
-->
Function OpenBlDtlRef()
	Dim arrRet
	Dim strCCNo
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(1)
	
	If Trim(frm1.txtIDDt.Text) = "" Then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	gblnWinEvent = True
	
	arrParam(0) = UCase(Trim(frm1.txtCCNo.value))
	arrParam(1) = frm1.txtCurrency.value
	
	iCalledAspName = AskPRAspName("M5212RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M5212RA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If arrRet(0,0) = "" Then
		Exit Function
	Else
		Call SetBlDtlRef(arrRet)
	End If	
		
End Function


<!--
'++++++++++++++++++++++++++++++++++++++++++++++  SetBlDtlRef()  +++++++++++++++++++++++++++++++++++++++++
-->
Function SetBlDtlRef(arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I, j,Row1
	Dim blnEqualFlg
	Dim intLoopCnt
	Dim intCnt, intCnt2
	Dim strBlNo
	Dim strBlSeq
	Dim strMessage
	Dim temp, temp_CcQty, temp_DocAmt, temp_BlAmt
	
	Const C_ItemCd_ref = 0								'품목코드 
	Const C_ItemNm_ref = 1								'품목명	
	Const C_TrackingNo_ref = 2							'Tracking No	
	Const C_BlQty_ref = 3								'B/L수량 
	Const C_CcQty_ref = 4								'통관수량 
	Const C_Spec_ref = 5								'규격 
	Const C_Unit_ref = 6								'단위 
	Const C_BlNo_ref = 7								'B/L관리번호 
	Const C_BlSeq_ref = 8								'B/L순번 
	Const C_BlDocNo_ref = 9								'B/L문서번호 
	Const C_PoNo_ref = 10							    'P/O번호 
	Const C_PoSeq_ref = 11								'P/O순번 
	Const C_LcNo_ref = 12								'L/C번호 
	Const C_LcSeq_ref = 13								'L/C순번 
	Const C_HsCd_ref = 14							    'H/S번호 
	Const C_HsNm_ref = 15								'H/S명 
	Const C_PlantCd_ref = 16							
	Const C_PlantNm_ref = 17							
	Const C_DocAmt_ref = 18								'금액(hidden)
	Const C_NetWeight_ref = 19							'순중량(hidden)
	Const C_Weight_Unit_ref = 20						'순중량단위(hidden)
	Const C_Price_ref = 21								'단가 
	
	With frm1 
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'수정(화면성능개선관련)-2003.04.03-Lee Eun Hee
		.vspdData.ReDraw = False	

		TempRow = .vspdData.MaxRows								
		intLoopCnt = Ubound(arrRet, 1)							
		intCnt2 = 0
		
		Redim preserve arrRet(Ubound(arrRet,1),24)
		For intCnt = 1 to intLoopCnt
			blnEqualFlg = False

			If TempRow <> 0 Then

				strBlNo=""
				strBlSeq=""

				For j = 1 To TempRow

					.vspdData.Row = j
					.vspdData.Col = C_BlNo
					strBlNo = .vspdData.Text

					.vspdData.Row = j
					.vspdData.Col = C_BlSeq
					strBlSeq = .vspdData.Text

					If strBlNo = arrRet(intCnt - 1, C_BlNo_ref) And strBlSeq = arrRet(intCnt - 1, C_BlSeq_ref) Then
						blnEqualFlg = True
						strMessage = strMessage & strBlNo & "-" & strBlSeq & ";"
						Exit For
					Else
						blnEqualFlg = False
					End If

				Next

			End If

			If blnEqualFlg = False Then
				'참조시 같은 번호가 있는 것이 포함 되었을때 같지 않은 것은 추가되어야 한다.
				intCnt2 = intCnt2 + 1	
				.vspdData.MaxRows = CLng(TempRow) + CLng(intCnt2)
				.vspdData.Row = CLng(TempRow) + CLng(intCnt2)
				Row1 = .vspdData.Row
				
				temp_CcQty = UNIFormatNumber(CStr(UNICDbl(arrRet(intCnt - 1, C_BlQty_ref)) - UNICDbl(arrRet(intCnt - 1, C_CcQty_ref))),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
				
				temp_DocAmt = UNIConvNumPCToCompanyByCurrency(CDBL(arrRet(intCnt - 1, C_DocAmt_ref)), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
				
				temp_BlAmt = UNIConvNumPCToCompanyByCurrency(CDBL(arrRet(intCnt - 1, C_DocAmt_ref)), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
				
				Call .vspdData.SetText(0			,	Row1, ggoSpread.InsertFlag)
				Call .vspdData.SetText(C_ItemCd		,	Row1, arrRet(intCnt - 1, C_ItemCd_ref))
				Call .vspdData.SetText(C_ItemNm		,	Row1, arrRet(intCnt - 1, C_ItemNm_ref))
				Call .vspdData.SetText(C_Spec		,	Row1, arrRet(intCnt - 1, C_Spec_ref))
				Call .vspdData.SetText(C_TrackingNo	,	Row1, arrRet(intCnt - 1, C_TrackingNo_ref))
				Call .vspdData.SetText(C_BlQty		,	Row1, arrRet(intCnt - 1, C_BlQty_ref))
				Call .vspdData.SetText(C_BlCcQty	,	Row1, arrRet(intCnt - 1, C_CcQty_ref))
				Call .vspdData.SetText(C_CcQty		,	Row1, temp_CcQty)
				Call .vspdData.SetText(C_Unit		,	Row1, arrRet(intCnt - 1, C_Unit_ref))
				Call .vspdData.SetText(C_BlNo		,	Row1, arrRet(intCnt - 1, C_BlNo_ref))
				Call .vspdData.SetText(C_BlSeq		,	Row1, arrRet(intCnt - 1, C_BlSeq_ref))
				Call .vspdData.SetText(C_BlDocNo	,	Row1, arrRet(intCnt - 1, C_BlDocNo_ref))
				Call .vspdData.SetText(C_PoNo		,	Row1, arrRet(intCnt - 1, C_PoNo_ref))
				Call .vspdData.SetText(C_PoSeq		,	Row1, arrRet(intCnt - 1, C_PoSeq_ref))
				Call .vspdData.SetText(C_LcNo		,	Row1, arrRet(intCnt - 1, C_LcNo_ref))
				Call .vspdData.SetText(C_LcSeq		,	Row1, arrRet(intCnt - 1, C_LcSeq_ref))
				Call .vspdData.SetText(C_HsCd		,	Row1, arrRet(intCnt - 1, C_HsCd_ref))
				Call .vspdData.SetText(C_HsNm		,	Row1, arrRet(intCnt - 1, C_HsNm_ref))
				Call .vspdData.SetText(C_DocAmt		,	Row1, temp_DocAmt)
				Call .vspdData.SetText(C_BlAmt		,	Row1, temp_BlAmt)
				
				If Trim(arrRet(intCnt - 1, C_NetWeight_ref)) <> "" Then
					'C_NetWeight_ref는 히든필드(2003.06.13)
					Call .vspdData.SetText(C_NetWeight,	Row1, UNIFormatNumber(CDbl(arrRet(intCnt - 1, C_NetWeight_ref)),ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))
				Else
					Call .vspdData.SetText(C_NetWeight,	Row1, 0)
				End If
				
				Call .vspdData.SetText(C_Price		,	Row1, arrRet(intCnt - 1, C_Price_ref))
				Call .vspdData.SetText(C_CIFDocAmt	,	Row1, 0)
				Call .vspdData.SetText(C_CIFLocAmt	,	Row1, 0)
				Call .vspdData.SetText(C_InputQty	,	Row1, 0)
				Call .vspdData.SetText(C_CcSeq		,	Row1, 0)
				

				Call vspdData_Change(C_CcQty_Ref1, Row1)	
				'수정(화면성능개선관련)-2003.04.03-Lee Eun Hee	
				'SetSpreadColor CLng(TempRow) + CLng(intCnt), CLng(TempRow) + CLng(intCnt)
			End If
		Next
		'수정(화면성능개선관련)-2003.04.03-Lee Eun Hee
		Call TotalSum
		Call SetSpreadColor(CLng(TempRow)+1,.vspdData.MaxRows)
			
		if strMessage<>"" then
			Call DisplayMsgBox("17a005","X",strmessage,"B/L번호" & "," & "B/L순번")
			.vspdData.ReDraw = True
			Exit Function
		End if
			
		.vspdData.ReDraw = True

	End With
		
End Function

'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Price,-1, .txtCurrency.value, parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_DocAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		'B/L금액 
		ggoSpread.SSSetFloatByCellOfCur C_BlAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		
		ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloatByCellOfCur C_OrgDocAmt1,-1, .txtCurrency.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec,,,"Z"

	End With

End Sub
<!--
'=====================================  SetSpreadFloatLocal()  ========================================
-->
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )

   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
     
End Sub
<!--
'============================================  2.5.1 OpenCookie()  ======================================
-->
Function OpenCookie()
	frm1.txtCCNo.Value = ReadCookie("CCNo")
	
	frm1.hdnQueryType.Value = "autoQuery"
	
	WriteCookie "CCNo", ""
	If UCase(Trim(frm1.txtCCNo.value)) <> "" Then
		Call dbQuery()
	End If
		
End Function

<!--
'============================================  2.5.2 TotalSum()  ======================================
-->
Sub TotalSum()
	Dim SumTotal, lRow

	SumTotal = UNICDbl(frm1.txtDocAmt.Text)
	ggoSpread.source = frm1.vspdData
	For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = 0

		If frm1.vspdData.Text = ggoSpread.InsertFlag Then
			frm1.vspdData.Col = C_DocAmt
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	Next

	frm1.txtDocAmt.text = UNIConvNumPCToCompanyByCurrency(Cstr(SumTotal), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")

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
'=============================================  2.5.5 LoadChargeHdr()  ======================================
-->
Function LoadChargeHdr()

	Dim IntRetCD

    If Trim(lgIntFlgMode) <> Trim(Parent.OPMD_UMODE) Then                          
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End if
	    	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'Process Step 
	WriteCookie "Process_Step", "VD"
	'통관관리번호 
	WriteCookie "Po_No", UCase(Trim(frm1.txtCCNo.value))
	'면허번호 
	'WriteCookie "TmpNo", UCase(Trim(frm1.txtIPNo.value))
	'구매그룹(수입담당)
	WriteCookie "Pur_Grp", UCase(Trim(frm1.txtPurGrp.value))
	'화폐 
	'WriteCookie "Currency", UCase(Trim(frm1.txtCurrency.value))
	'환율 
	'WriteCookie "XchRate", UCase(Trim(frm1.txtXchRate.value))

	PgmJump(CHARGE_HDR_ENTRY_ID)

End Function

<!--
'==========================================  2.5.6 SetQureySpreadColor()  ====================================
-->
Sub SetQureySpreadColor(ByVal lRow)
	ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
	    
		.Redraw = False
		
		ggoSpread.SSSetProtected C_ItemCd,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_ItemNm,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_Spec,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_TrackingNo,	lRow, .MaxRows
		ggoSpread.SSSetProtected C_Unit,		lRow, .MaxRows
		
		ggoSpread.SSSetProtected C_HsCd,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_HsNm,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_BlQty,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_InputQty,	lRow, .MaxRows
		ggoSpread.SSSetProtected C_CcSeq,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_BlNo,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_BlSeq,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_BlDocNo,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_PoNo,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_PoSeq,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_LcNo,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_LcSeq,		lRow, .MaxRows
		
		ggoSpread.SSSetRequired C_CcQty,		lRow, .MaxRows
		ggoSpread.SSSetRequired C_Price,		lRow, .MaxRows
		ggoSpread.SSSetRequired C_DocAmt,		lRow, .MaxRows
		ggoSpread.SSSetRequired C_NetWeight,	lRow, .MaxRows
		
		ggoSpread.SSSetProtected C_Price,		lRow, .MaxRows
		ggoSpread.SSSetProtected C_DocAmt,		lRow, .MaxRows
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, lRow, .MaxRows
		
		.ReDraw = True
	End With
End Sub

<!--
'==========================================  2.5.7 CookiePage()  ======================================
-->
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4875				
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 1 Then

	    If Trim(lgIntFlgMode) <> Trim(Parent.OPMD_UMODE) Then  
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie CookieSplit , frm1.txtCCNo.value
		
		Call PgmJump(CC_LAN_ENTRY_ID)
		
	elseIf Kubun = 2 Then

	    If Trim(lgIntFlgMode) <> Trim(Parent.OPMD_UMODE) Then          
	        Call DisplayMsgBox("900002","X","X","X")
	        Exit Function
	    End if
	    	
	    If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie CookieSplit , frm1.txtCCNo.value
		
		Call PgmJump(CC_HEADER_ENTRY_ID)
		
	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtCCNo.value =  arrVal(0) 
		'2003.06.03 추가 
		frm1.hdnQueryType.Value = "autoQuery"

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call DbQuery()
					
		WriteCookie CookieSplit , ""

	End IF

End Function

<!--
'=========================================  3.1.1 Form_Load()  ==========================================
-->
Sub Form_Load()
		
	Call LoadInfTB19029							
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart ,ggStrMaxPart)
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

'=========================================  vspdData_Click()  ===================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
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
	frm1.vspdData.Row = Row
End Sub

'==================================  vspdData_DblClick()  =====================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
   
End Sub

'==================================  vspdData_ColWidthChange()  ===================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'==================================  vspdData_MouseDown()  ====================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
   
End Sub    
'==================================  FncSplitColumn()  ======================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=================================  PopSaveSpreadColumnInf()  ================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'================================   PopRestoreSpreadColumnInf()  =============================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call CurFormatNumSprSheet() 
    Call ggoSpread.ReOrderingSpreadData()
    'Call SetQureySpreadColor(1) 
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
	Dim Qty, Price, DocAmt, LocAmt, XchRate, BlAmt
	Dim CcQty, BlCcQty, BlQty, InputQty
	
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Select Case Col
		Case C_CcQty, C_Price, C_CcQty_Ref1

			frm1.vspdData.Col = C_CcQty
			CcQty = UNICDbl(frm1.vspdData.text)
			frm1.vspdData.Col = C_BlCcQty
			BlCcQty = UNICDbl(frm1.vspdData.text)
			frm1.vspdData.Col = C_BlQty
			BlQty = UNICDbl(frm1.vspdData.text)
			frm1.vspdData.Col = C_BlAmt			
			BlAmt = UNICDbl(frm1.vspdData.text)
			frm1.vspdData.Col = C_InputQty
			InputQty = UNICDbl(frm1.vspdData.text)

			frm1.vspdData.Col = C_CcQty
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
				
			If BlQty = 0 then 
				DocAmt = Qty * BlAmt
			Else
				DocAmt = Qty/BlQty * BlAmt
			End if
				
			frm1.vspdData.Col = C_DocAmt
			if CcQty = BlQty then

				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(Cstr(BlAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")
			else
				frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt), frm1.txtCurrency.value, Parent.ggAmtOfMoneyNo,"X","X")				
			end if
			'수정(화면성능개선관련)-2003.04.03-Lee Eun Hee	
			If col <> C_CcQty_Ref1 Then
				Call TotalSumNew(Row)			
			End If
			'총금액계산을 위해 필요(2003.05)
			frm1.vspdData.Col = C_DocAmt
			DocAmt = frm1.vspdData.Text
			frm1.vspdData.Col = C_OrgDocAmt		
			frm1.vspdData.Text = DocAmt
			
		Case C_CIFDocAmt
				
			frm1.vspdData.Col = Col
			DocAmt = UNICDbl(Trim(frm1.vspdData.text)) 
				
			frm1.vspdData.Col = C_CIFLocAmt
			frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(Cstr(DocAmt) * UNICDbl(Trim(frm1.txtXchRate.Value)),Parent.gCurrency, Parent.ggAmtOfMoneyNo,"X","X")
			
	End select
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    	
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

	If Not chkField(Document, "1") Then	Exit Function
	
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
	Set gActiveElement = document.ActiveElement 	
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
	Set gActiveElement = document.ActiveElement 	
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
	If Not ggoSpread.SSDefaultCheck         Then             
	   Exit Function
	End If
		
	If DbSave = False Then Exit Function

	If Trim(frm1.txtHCCNo.value) <> Trim(frm1.txtCCNo.value) then
		Trim(frm1.txtCCNo.value) =	Trim(frm1.txtHCCNo.value)		
	End If
			
	FncSave = True
	Set gActiveElement = document.ActiveElement 
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
	Set gActiveElement = document.ActiveElement 
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
		'Call TotalSum()	
	Set gActiveElement = document.ActiveElement 	
End Function

<!--
'==========================================  5.1.7 FncInsertRow()  ======================================
-->
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    
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
        ggoSpread.InsertRow ,imRow
        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK
    
    Set gActiveElement = document.ActiveElement   
End Function
<!--
'==========================================  5.1.8 FncDeleteRow()  ======================================
-->
Function FncDeleteRow()
	Dim lDelRows
	Dim iDelRowCnt, i
		
	if frm1.vspdData.Maxrows < 1 then Exit Function
		
	With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow
		lgBlnFlgChgValue = True
	End With

	'Call TotalSum() -->2003.06.03
	Set gActiveElement = document.ActiveElement 
End Function

<!--
'============================================  5.1.9 FncPrint()  ========================================
-->
Function FncPrint()
   Call parent.FncPrint()
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
End Function

<!--
'===========================================  5.1.12 FncExcel()  ========================================
-->
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

<!--
'===========================================  5.1.13 FncFind()  =========================================
-->
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

<!--
'===========================================  5.1.14 FncExit()  =========================================
-->
Function FncExit()
	Dim IntRetCD

	FncExit = False

	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

<!--
'=============================================  5.2.1 DbQuery()  ========================================
-->
Function DbQuery()
	Dim strVal

	Err.Clear														

	DbQuery = False													

	If LayerShowHide(1) = False Then Exit Function
	
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtHCCNo.value)	
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001			
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)	
	End If
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
		strVal = strVal & "&txtQueryType=" & Trim(frm1.hdnQueryType.value)
		'수정(2003.06.10)
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		
	Call RunMyBizASP(MyBizASP, strVal)								
	
	DbQuery = True	
	Set gActiveElement = document.ActiveElement 												
End Function
	
<!--
'=============================================  5.2.2 DbSave()  =========================================
-->
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim strUnit, strCcQty, strPrice, strDocAmt, strNetWeight, strCIFDocAmt, strCIFLocAmt, strHsCd, strBlQty
	Dim strInputQty, strCcSeq, strBlNo, strBlSeq, strBlDocNo, strPoNo, strPoSeq, strLcNo, strLcSeq, strTrackingNo
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
				Case ggoSpread.DeleteFlag
					strDel = "D" & ColSep	& lRow & ColSep

		            .vspdData.Col = C_CcSeq		'2
		            strDel = strDel & Trim(.vspdData.Text) & RowSep

		            lGrpCnt = lGrpCnt + 1 
		            
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				
				If .vspdData.Text=ggoSpread.InsertFlag Then
					strVal = "C" & ColSep	& lRow & ColSep		'0	'1
				Else
					strVal = "U" & ColSep	& lRow & ColSep		'0	'1
				End If   

		            .vspdData.Col = C_Unit		'2
		            strUnit = UCase(Trim(.vspdData.Text))
		            
		            .vspdData.Col = C_CcQty		'3
		            If Trim(UNICDbl(.vspdData.Text)) = "0" or Trim(UNICDbl(.vspdData.Text)) = "" then
						Call DisplayMsgBox("970021","X","통관수량","X")
						Call SetActiveCell(frm1.vspdData,C_CcQty,lRow,"M","X","X")
						Call LayerShowHide(0)
						Exit Function
					End if
					strCcQty = UNIConvNum(UCase(Trim(.vspdData.Text)),0)
		            
		            '2007.2 패치 금액 입력필수 삭제- KSJ
		            '.vspdData.Col = C_Price		'4
		            'If Trim(UNICDbl(.vspdData.Text)) = "0" or Trim(UNICDbl(.vspdData.Text)) = "" then
					'	Call DisplayMsgBox("970021","X","단가","X")
					'	Call SetActiveCell(frm1.vspdData,C_Price,lRow,"M","X","X")
					'	Call LayerShowHide(0)
					'	Exit Function
					'End if 	
					'2007.2 패치End 금액 입력필수 삭제- KSJ
					strPrice = UNIConvNum(UCase(Trim(.vspdData.Text)),0)
		            
		            .vspdData.Col = C_DocAmt 	'5
		            strDocAmt = UNIConvNum(UCase(Trim(.vspdData.Text)),0)	
		            
		            .vspdData.Col = C_NetWeight	'6
		            If Trim(UNICDbl(.vspdData.Text)) = "0" or Trim(UNICDbl(.vspdData.Text)) = "" then
						Call DisplayMsgBox("970021","X","순중량","X")
						Call SetActiveCell(frm1.vspdData,C_NetWeight,lRow,"M","X","X")
						Call LayerShowHide(0)
						Exit Function
					End if 
					strNetWeight = UNIConvNum(UCase(Trim(.vspdData.Text)),0)
		            
		            .vspdData.Col = C_CIFDocAmt '7
		            strCIFDocAmt = UNIConvNum(UCase(Trim(.vspdData.Text)),0)		
		            
		            .vspdData.Col = C_CIFLocAmt '8
		            strCIFLocAmt = UNIConvNum(UCase(Trim(.vspdData.Text)),0)		
		            
		            .vspdData.Col = C_HsCd 		'9
		            strHsCd = UCase(Trim(.vspdData.Text))	
		            
		            .vspdData.Col = C_BlQty		'10	
		            strBlQty = UNIConvNum(UCase(Trim(.vspdData.Text)),0)	
		            
		            .vspdData.Col = C_InputQty	'11
		            strInputQty = UNIConvNum(UCase(Trim(.vspdData.Text)),0)		
		            
		            .vspdData.Col = C_CcSeq 	'12
		            strCcSeq = Trim(.vspdData.Text)
		            
		            .vspdData.Col = C_BlNo 		'13
		            strBlNo = UCase(Trim(.vspdData.Text))	
		            
		            .vspdData.Col = C_BlSeq		'14
		            strBlSeq = UCase(Trim(.vspdData.Text))		
		            
		            .vspdData.Col = C_BlDocNo	'15
		            strBlDocNo = UCase(Trim(.vspdData.Text))		
		            
		            .vspdData.Col = C_PoNo 		'16
		            strPoNo = UCase(Trim(.vspdData.Text))	
		            
		            .vspdData.Col = C_PoSeq		'17
		            strPoSeq = UCase(Trim(.vspdData.Text))	
		            
		            .vspdData.Col = C_LcNo 		'18
		            strLcNo = UCase(Trim(.vspdData.Text))	
		            
		            .vspdData.Col = C_LcSeq		'19
		            strLcSeq = UCase(Trim(.vspdData.Text))		
		            
		            .vspdData.Col = C_TrackingNo '20
		            strTrackingNo = UCase(Trim(.vspdData.Text))		
					
					strVal = strVal & strUnit & ColSep & strCcQty & ColSep & strPrice & ColSep &strDocAmt & ColSep &strNetWeight & ColSep &strCIFDocAmt & ColSep & strCIFLocAmt & ColSep & _   
							strHsCd & ColSep & strBlQty & ColSep & strInputQty & ColSep & strCcSeq & ColSep & strBlNo & ColSep &strBlSeq & ColSep &strBlDocNo & ColSep &strPoNo & ColSep & _
							strPoSeq & ColSep & strLcNo & ColSep & strLcSeq & ColSep & strTrackingNo & ColSep & lRow & RowSep

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
	        
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					
		
	End With
	
		
    DbSave = True                                                   
	Set gActiveElement = document.ActiveElement 
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
	Call RemovedivTextArea
		
	If frm1.vspdData.MaxRows > 0 Then
		Call SetToolBar("1110101100011111")
		frm1.vspdData.focus
	Else
		Call SetToolBar("1110100100011111")
		frm1.txtCCNo.focus
	End If
	    
	'Call SetQureySpreadColor(1)
End Function
	
<!--
'=============================================  5.2.5 DbSaveOk()  =======================================
-->
Function DbSaveOk()		
	Call InitVariables
	Call MainQuery()
	Set gActiveElement = document.ActiveElement 
End Function
	
<!--
'=============================================  5.2.6 DbDeleteOk()  =====================================
-->
Function DbDeleteOk()												
	On Error Resume Next                                            
End Function

'==========================================  vspdData_ScriptDragDropBlock()  =============================
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관 내역등록</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenBlDtlRef">B/L내역참조</A></TD>
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
									<TD CLASS=TD5 NOWRAP>통관 관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=32 MAXLENGTH=18 TAG="12XXXU" ALT="통관 관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" ALIGN=top TYPE="BUTTON" ONCLICK="VBSCRIPT:btnCcNo_Click()"></TD>
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
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtIDDt" style="HEIGHT: 20px; WIDTH: 81px" tag="24X1" ALT="신고일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>면허번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIPNo" ALT="면허번호" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="24XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>면허일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtIPDt" style="HEIGHT: 20px; WIDTH: 81px" tag="24X1" ALT="면허일" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>총통관금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU">&nbsp;
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 163px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>수출자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" Title="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH=* ALIGN=RIGHT><A href="VBSCRIPT:CookiePage(1)">수입통관란정보등록</A>&nbsp;|&nbsp;<A href="VBSCRIPT:CookiePage(2)">수입통관등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadChargeHdr()">경비등록</A></TD>
				<TD WIDTH=10>&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="txtPurGrp" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="txtXchRate" TAG="24" TABINDEX=-1>
<INPUT TYPE=HIDDEN NAME="hdnQueryType" tag="14">
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
