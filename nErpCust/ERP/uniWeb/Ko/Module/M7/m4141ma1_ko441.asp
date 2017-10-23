<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : 
'*  3. Program ID           : m4141ma1
'*  4. Program Name         : 구매반품등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/06/02
'*  9. Modifier (First)     : Shin Jin-hyun
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<!--meta http-equiv="Content-type" content="text/html; charset=euc-kr"-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit														

Const BIZ_PGM_ID = "m4141mb1_ko441.asp"			

<!-- #Include file="../../inc/lgvariables.inc" -->
'==============================================================================================================================
Dim IsOpenPop          
Dim lblnWinEvent
Dim interface_Account

Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec	
Dim C_TrackingNo
Dim C_GrQty	
Dim C_Unit
Dim C_Prc		' 2005.10.19추가 
Dim C_DocAmt	' 2005.10.19추가 
Dim	C_Cur		' 2005.10.19추가 
Dim C_SlCd	
Dim C_SlPop	
Dim C_SlNm	
Dim C_LotNo	
Dim C_LotNoPop	
Dim C_LotNoSeq	
Dim C_RetType	
Dim C_RetTypeNm
Dim C_RemarkDtl
Dim C_PoNo	
Dim C_PoSeq	
Dim C_GmNo	
Dim C_GmSeq	
Dim C_MvmtNo
Dim C_RetOrdQty

'==============================================================================================================================
Function ChangeTag(Byval Changeflg)
	
	Dim index

	If Changeflg = true then
		ggoOper.SetReqAttr	frm1.txtGrNo1, "Q"
		ggoOper.SetReqAttr	frm1.txtRemark, "Q"
		frm1.vspdData.ReDraw = false
		For index = 1 to frm1.vspdData.MaxCols
			ggoSpread.SpreadLock index , -1
		Next
		frm1.vspdData.ReDraw = true
	Else
		Call ggoOper.LockField(Document, "N")
		ggoOper.SetReqAttr	frm1.txtGrNo1, "D"
	End if 
	
End Function 

'==============================================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd		= 1
	C_PlantNm		= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_TrackingNo	= 6
	C_GrQty			= 7
	C_Unit			= 8
	C_Prc			= 9
	C_DocAmt		= 10
	C_Cur			= 11
	C_SlCd			= 12
	C_SlPop			= 13
	C_SlNm			= 14
	C_LotNo			= 15
	C_LotNoPop		= 16
	C_LotNoSeq		= 17
	C_RetType	    = 18
	C_RetTypeNm		= 19
	C_RemarkDtl		= 20
	C_PoNo			= 21
	C_PoSeq			= 22
	C_GmNo			= 23
	C_GmSeq			= 24
	C_MvmtNo		= 25
	C_RetOrdQty		= 26

End Sub
'==============================================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
    lgStrPrevKey = ""        
    lgLngCurRows = 0         
    frm1.vspdData.MaxRows = 0
    
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	frm1.txtGmDt.Text = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
	
	frm1.txtGroupCd.Value = Parent.gPurGrp
    Call SetToolBar("1110000000001111")					 				
    frm1.txtGrNo.focus 
    Set gActiveElement = document.activeElement
    interface_Account = GetSetupMod(Parent.gSetupMod, "a")
	frm1.btnGlSel.disabled = true 
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20051201",,Parent.gAllowDragDropSpread  
		
		.ReDraw = false
		
		.MaxCols = C_RetOrdQty + 1
		.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 		C_PlantCd,		"공장", 10
		ggoSpread.SSSetEdit 		C_PlantNm,		"공장명", 20 
		ggoSpread.SSSetEdit 		C_ItemCd,		"품목", 10
		ggoSpread.SSSetEdit 		C_ItemNm,		"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,			"품목규격", 20 
		    
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15    
		SetSpreadFloatLocal 		C_GrQty,		"입출고수량",15,1, 3
		ggoSpread.SSSetEdit 		C_Unit,			"단위", 10
		SetSpreadFloatLocal 		C_Prc,			"단가",15,1, 3
		SetSpreadFloatLocal 		C_DocAmt,		"금액",15,1, 3
		ggoSpread.SSSetEdit 		C_Cur,			"화폐", 10
		ggoSpread.SSSetEdit 		C_SlCd,			"창고", 10,,,7,2
		ggoSpread.SSSetButton 		C_SlPop
		ggoSpread.SSSetEdit 		C_SLNm,			"창고명", 20
		ggoSpread.SSSetEdit 		C_LotNo,		"LOT NO", 20,,,25,2
		ggoSpread.SSSetButton 		C_LotNoPop
		ggoSpread.SSSetEdit 		C_LotNoSeq,		"LOT NO 순번", 15
		ggoSpread.SSSetEdit  		C_RetType,		"반품유형", 10
		ggoSpread.SSSetEdit 		C_RetTypeNm,	"반품유형명", 15
		
		ggoSpread.SSSetEdit 		C_RemarkDtl,	"비고", 20    
    
		ggoSpread.SSSetEdit 		C_PoNo,			"발주번호", 20
		ggoSpread.SSSetFloat 		C_PoSeq,		"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_GmNo,			"재고처리번호", 20
		ggoSpread.SSSetFloat 		C_GmSeq,		"재고처리순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_MvmtNo,		"", 10
		SetSpreadFloatLocal 		C_RetOrdQty, "반품수량",15,1, 3

		Call ggoSpread.MakePairsColumn(C_SlCd,C_SlPop)
		Call ggoSpread.MakePairsColumn(C_LotNo,C_LotNoPop)
		Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo ,True)	
		Call ggoSpread.SSSetColHidden(C_RetOrdQty,C_RetOrdQty ,True)	
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols ,True)	
			
		.ReDraw = true
		
		Call SetSpreadLock()
    End With
   
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	
    With ggoSpread
    
		.SpreadLock 	C_PlantCd , -1,C_PlantCd , -1
		.SpreadLock 	C_PlantNm , -1,C_PlantNm , -1
		.SpreadLock 	C_ItemCd , -1,C_ItemCd , -1
		.SpreadLock 	C_ItemNm , -1,C_ItemNm , -1
		.SpreadLock 	C_Spec , -1,C_Spec , -1
		.SpreadLock 	C_TrackingNo , -1,C_TrackingNo, -1   
    	.spreadlock 	C_Unit , -1,C_Unit , -1
    	.spreadlock 	C_Prc , -1,C_Prc , -1
    	.spreadlock 	C_DocAmt , -1,C_DocAmt , -1
    	.spreadlock 	C_Cur , -1,C_Cur , -1
		.spreadlock 	C_SLNm, -1,C_SLNm , -1
		.spreadlock 	C_RetType, -1,C_RetTypeNm , -1    
		.spreadlock 	C_PoNo, -1
		.SSSetProtected frm1.vspdData.MaxCols, -1
    
    End With
    frm1.vspdData.ReDraw = True
End Sub
'==============================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With ggoSpread
        frm1.vspdData.ReDraw = False
		
		.SSSetProtected	C_PlantCd,		pvStartRow, pvEndRow
		.SSSetProtected	C_PlantNm,		pvStartRow, pvEndRow
		.SSSetProtected	C_ItemCd,		pvStartRow, pvEndRow
		.SSSetProtected	C_ItemNm,		pvStartRow, pvEndRow
		.SSSetProtected	C_Spec,			pvStartRow, pvEndRow    
		.SSSetProtected	C_TrackingNo,	pvStartRow, pvEndRow
		.SSSetRequired	C_GrQty,		pvStartRow, pvEndRow
		.SSSetProtected	C_Unit,			pvStartRow, pvEndRow
		.SSSetProtected	C_Prc,			pvStartRow, pvEndRow
		.SSSetProtected	C_DocAmt,		pvStartRow, pvEndRow
		.SSSetProtected	C_Cur,			pvStartRow, pvEndRow
		.SSSetRequired	C_SlCd,			pvStartRow, pvEndRow
		.SSSetProtected	C_SLNm,			pvStartRow, pvEndRow
		.SSSetProtected	C_RetType,		pvStartRow, pvEndRow    
		.SSSetProtected	C_RetTypeNm,	pvStartRow, pvEndRow
		.SSSetProtected	C_PoNo,			pvStartRow, pvEndRow
		.SSSetProtected	C_PoSeq,		pvStartRow, pvEndRow
		.SSSetProtected	C_GmNo,			pvStartRow, pvEndRow
		.SSSetProtected	C_GmSeq,		pvStartRow, pvEndRow
		.SSSetProtected frm1.vspdData.MaxCols, pvStartRow,			pvEndRow
		
		frm1.vspdData.ReDraw = True
    End With
End Sub
'==============================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_TrackingNo	= iCurColumnPos(6)
			C_GrQty			= iCurColumnPos(7)
			C_Unit			= iCurColumnPos(8)
			C_Prc			= iCurColumnPos(9)
			C_DocAmt		= iCurColumnPos(10)
			C_Cur			= iCurColumnPos(11)
			C_SlCd			= iCurColumnPos(12)
			C_SlPop			= iCurColumnPos(13)
			C_SlNm			= iCurColumnPos(14)
			C_LotNo			= iCurColumnPos(15)
			C_LotNoPop		= iCurColumnPos(16)
			C_LotNoSeq		= iCurColumnPos(17)
			C_RetType	    = iCurColumnPos(18)
			C_RetTypeNm		= iCurColumnPos(19)
			C_RemarkDtl		= iCurColumnPos(20)
			C_PoNo			= iCurColumnPos(21)
			C_PoSeq			= iCurColumnPos(22)
			C_GmNo			= iCurColumnPos(23)
			C_GmSeq			= iCurColumnPos(24)
			C_MvmtNo		= iCurColumnPos(25)
			C_RetOrdQty		= iCurColumnPos(26)
	End Select

End Sub	
'==============================================================================================================================
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnGlNo.value)
	arrParam(1) = ""
	
   If frm1.hdnGlType.Value = "A" Then               '회계전표팝업 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			lblnWinEvent = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif frm1.hdnGlType.Value = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다. 
    End if

	lblnWinEvent = False
	
End Function
'==============================================================================================================================
Function OpenLotNo()
	
	Dim strRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	Dim lgCurRow
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	lgCurRow = frm1.vspdData.ActiveRow
	
	arrParam(0) = UCase(Trim(GetSpreadText(frm1.vspdData,C_SlCd,lgCurRow,"X","X")))
	arrParam(1) = UCase(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lgCurRow,"X","X")))
	arrParam(2) = ""						'tracking No
	arrParam(3) = UCase(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lgCurRow,"X","X")))
	arrParam(4) = "J"						'Userflag
	arrParam(5) = Trim(GetSpreadText(frm1.vspdData,C_LotNo,lgCurRow,"X","X"))
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = Trim(GetSpreadText(frm1.vspdData,C_Unit,lgCurRow,"X","X"))
	
	iCalledAspName = AskPRAspName("I2212RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "I2212RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, _
		Array(window.parent,arrParam(0),arrParam(1),arrParam(2),arrParam(3),arrParam(4),arrParam(5),arrParam(6),arrParam(7),arrParam(8)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	
	If strRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_LotNo,		lgCurRow, strRet(3))
		Call frm1.vspdData.SetText(C_LotNoSeq,	lgCurRow, strRet(4))
	End If	
				
End Function
'==============================================================================================================================
Function OpenPoRef()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","신규등록이 아닌 경우","발주내역참조" )
		Exit Function
	End if 

	if Trim(frm1.txtMvmtType.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "입출고유형","X")
		frm1.txtMvmtType.focus
		Set gActiveElement = document.activeElement
		Exit Function
	elseif Trim(frm1.txtSupplierCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공급처","X")
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	if (UCase(Trim(frm1.hdnImportflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="Y") or _
	   (UCase(Trim(frm1.hdnSubcontraflg.Value))="Y" and UCase(Trim(frm1.hdnRcptflg.Value))="N") then
		Call DisplayMsgBox("17A012", "X","입출고유형" & frm1.txtMvmtType.Value & "(" & frm1.txtMvmtTypeNm.value & ")","발주내역참조" )
		'입출고유형 는(은) 발주참조를 할수 없습니다."
		Exit Function
	End if
		
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	arrParam(1) = Trim(frm1.txtSupplierNm.value)
	arrParam(2) = Trim(frm1.txtGroupCd.value)
	arrParam(3) = ""		'plant
	arrParam(4) = "N"		'Clsflg
	arrParam(5) = "Y"		'Releaseflg
	arrParam(8) = "GR"		'Rcptflg
	arrParam(9) = Trim(frm1.txtMvmtType.Value)
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""
	
	
	iCalledAspName = AskPRAspName("M3112RA7_KO441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3112RA7_KO441", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetPoRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetPoRef(strRet)

	Dim Index1, Index3, Count1
	Dim temp, temp1, temp2
	Dim LotCheck
	Dim iCurRow, TempRow
	Dim strMessage
	Dim IntIflg
		
	Const C_PoNo_Ref		= 0
	Const C_PoSeq_Ref		= 1
	Const C_PlantCd_Ref		= 2
	Const C_SlCd_Ref		= 3
	Const C_ItemCd_Ref		= 4
	Const C_ItemNm_Ref		= 5
	Const C_Spec_Ref		= 6
	Const C_TrackingNo_Ref  = 7
	Const C_Qty_Ref			= 8
	Const C_Unit_Ref		= 9
	Const C_Prc_Ref			= 10
	Const C_Amt_Ref			= 11
	Const C_Cur_Ref			= 12
	Const C_DlvyDt_Ref		= 13
	Const C_PlantNm_Ref		= 14
	Const C_SLNm_Ref		= 15
	Const C_RcptQty_Ref		= 16
	Const C_LcQty_Ref		= 17
	Const C_Lotflg_ref		= 18
	Const C_Insflg_ref		= 19
	Const C_RetType_ref	    = 20
	Const C_RetTypeNm_ref	= 21
	Const C_Insmethod_ref	= 22
	Const C_Pur_Grp_Ref	    = 23
	Const C_LotNo_ref		= 24
	Const C_LotSeq_Ref		= 25

	Count1 = Ubound(strRet,1)
	strMessage = ""
	IntIflg = true
		
	With frm1
		
		.vspdData.Redraw = False
		
		TempRow = .vspdData.MaxRows					'리스트 max값 
		
		For index1 = 0 to Count1
		
			'If TempRow <> 0 Then

				'For Index3 = 1 To TempRow				'같은 No가 있으면 Row를 추가하지 않는다.
				'	.vspdData.Row = Index3
				'	.vspdData.Col = C_PoNo
				'	temp1 = .vspdData.Text
				'	.vspdData.Col = C_PoSeq
				'	temp2 = .vspdData.Text
				'	If temp1 = strRet(index1,C_PoNo_Ref) And temp2 = strRet(index1,C_PoSeq_Ref) Then
				'		strMessage = strMessage & strRet(Index1,C_PoNo_Ref) & "-" & strRet(index1,C_PoSeq_Ref) & ";"
				'		intIflg=False
				'		Exit for
				'	End if 
				'Next
			
			'End If

			.vspdData.Row = Index1 + 1
		
			If IntIflg <> False then
				'Call fncinsertrow(Count1 + 1)
				Call fncinsertrow(1)
				
				iCurRow = .vspdData.ActiveRow + index1
					
				Call .vspdData.SetText(C_PlantCd,	iCurRow, strRet(index1,C_PlantCd_Ref))
				Call .vspdData.SetText(C_PlantNm,	iCurRow, strRet(index1,C_PlantNm_Ref))
				Call .vspdData.SetText(C_itemCd,	iCurRow, strRet(index1,C_ItemCd_Ref))
				Call .vspdData.SetText(C_itemNm,	iCurRow, strRet(index1,C_ItemNm_Ref))
				Call .vspdData.SetText(C_Spec,		iCurRow, strRet(index1,C_Spec_Ref))
				Call .vspdData.SetText(C_TrackingNo,iCurRow, strRet(index1,C_TrackingNo_Ref))
				temp = UNICDbl(strRet(index1,C_Qty_Ref)) - UNICDbl(strRet(index1,C_RcptQty_Ref))
				Call .vspdData.SetText(C_GrQty,		iCurRow, UNIFormatNumber(temp,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit))							
				Call .vspdData.SetText(C_Unit,		iCurRow, strRet(index1,C_Unit_Ref))
				Call .vspdData.SetText(C_SlCd,		iCurRow, strRet(index1,C_SlCd_Ref))
				Call .vspdData.SetText(C_SLNm,		iCurRow, strRet(index1,C_SLNm_Ref))
				Call .vspdData.SetText(C_PoNo,		iCurRow, strRet(index1,C_PoNo_Ref))
				Call .vspdData.SetText(C_PoSeq,		iCurRow, strRet(index1,C_PoSeq_Ref))
						
				LotCheck = Trim(strRet(index1,C_Lotflg_ref))
				
				if LotCheck = "Y" then
					Call .vspdData.SetText(C_LotNo,	iCurRow, strRet(index1,C_LotNo_ref))
					Call .vspdData.SetText(C_LotNoSeq,	iCurRow, strRet(index1,C_LotSeq_Ref))	
					
					ggoSpread.spreadUnlock 	C_LotNo, iCurRow, C_LotNoSeq , iCurRow
					if frm1.hdnRcptflg.value <> "Y" then
						if Trim(strRet(index1,C_LotNo_ref)) = "" then
							ggoSpread.SSSetRequired C_LotNo, iCurRow, iCurRow
							ggoSpread.SSSetRequired C_LotNoSeq, iCurRow, iCurRow
						else
							ggoSpread.SSSetProtected	C_LotNo, iCurRow, iCurRow
							ggoSpread.SSSetProtected	C_LotNoPop, iCurRow, iCurRow
							ggoSpread.SSSetProtected	C_LotNoSeq, iCurRow, iCurRow
						end if
					end if
'					if frm1.hdnRcptflg.value = "Y" then
'						ggoSpread.SSSetRequired C_LotNo, frm1.vspdData.Row, frm1.vspdData.Row
'						ggoSpread.SSSetRequired C_LotNoSeq, frm1.vspdData.Row, frm1.vspdData.Row
'					ELSE
'						ggoSpread.SSSetProtected	C_LotNo, frm1.vspdData.Row, frm1.vspdData.Row
'						ggoSpread.SSSetProtected	C_LotNoPop, frm1.vspdData.Row, frm1.vspdData.Row
'						ggoSpread.SSSetProtected	C_LotNoSeq, frm1.vspdData.Row, frm1.vspdData.Row
'					end if
				Else
					ggoSpread.spreadlock 	C_LotNo, iCurRow,C_LotNoSeq , iCurRow
					ggoSpread.SSSetProtected	C_LotNo, iCurRow, iCurRow
					ggoSpread.SSSetProtected	C_LotNoPop, iCurRow, iCurRow
					ggoSpread.SSSetProtected	C_LotNoSeq, iCurRow, iCurRow
					Call .vspdData.SetText(C_LotNo,		iCurRow, "*")
				End if
							
				Call .vspdData.SetText(C_RetType,	iCurRow, strRet(index1,C_RetType_ref))
				Call .vspdData.SetText(C_RetTypeNm,	iCurRow, strRet(index1,C_RetTypeNm_ref))
						
				'C_Pur_Grp_Ref
				IF Index1 = 0 Then
				.txtGroupCd.value = strRet(Index1,C_Pur_Grp_Ref)		
				End IF
			Else
				IntIFlg = True
			End if 
		Next

		If strMessage <> "" Then
			Call DisplayMsgBox("17a005","X",strmessage,"발주번호" & "," & "발주순번")
			.vspdData.ReDraw = True
			Exit Function
		End if
		
		.vspdData.ReDraw = True
		Call SetToolBar("11101001000111")
	End with
	
End Function
'==============================================================================================================================
Function OpenGrType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "입출고유형"	
	arrParam(1) = "M_Mvmt_type"
	
	arrParam(2) = Trim(frm1.txtMvmtType.Value)
	
	arrParam(4) = "Ret_Flg=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
	arrParam(5) = "입출고유형"			
	
    arrField(0) = "IO_Type_Cd"
    arrField(1) = "IO_Type_NM"
    
    arrHeader(0) = "입출고유형"		
    arrHeader(1) = "입출고유형명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else 
		frm1.txtMvmtType.Value	= arrRet(0)		
		frm1.txtMvmtTypeNm.Value= arrRet(1)
		Call changeMvmtType()
		lgBlnFlgChgValue = True
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCurRow
	
	iCurRow = frm1.vspdData.ActiveRow		
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "창고"
	arrParam(1) = "B_STORAGE_LOCATION"
	arrParam(2) = UCase(Trim(GetSpreadText(frm1.vspdData,C_SLCd,iCurRow,"X","X")))
	arrParam(4) = "PLANT_CD= " & FilterVar(UCase(GetSpreadText(frm1.vspdData,C_PlantCd,iCurRow,"X","X")), "''", "S") & " "
	arrParam(5) = "창고"
	
    arrField(0) = "SL_CD"
    arrField(1) = "SL_NM"
    
    arrHeader(0) = "창고"
    arrHeader(1) = "창고명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_SLCd,	iCurRow, arrRet(0))
		Call frm1.vspdData.SetText(C_SLNm,	iCurRow, arrRet(1))
		Call vspdData_Change(C_SLCd, frm1.vspdData.ActiveRow)
	End If	
End Function
'==============================================================================================================================
Function OpenMvmtNo()
	
		Dim strRet
		Dim arrParam(3)
		Dim iCalledAspName
		Dim IntRetCD
	
		If lblnWinEvent = True Or UCase(frm1.txtGrNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		lblnWinEvent = True

		arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
		arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
		arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
		arrParam(3) = ""'Rcpt flg , which must be "Y" or "N" or ""
				
		iCalledAspName = AskPRAspName("M4141PA1_KO441")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M4141PA1_KO441", "X")
			lblnWinEvent = False
			Exit Function
		End If
	
		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		lblnWinEvent = False
		If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
		If strRet(0) = "" Then
			frm1.txtGrNo.focus	
			Set gActiveElement = document.activeElement
			Exit Function
		Else
			frm1.txtGrNo.value = strRet(0)
			frm1.txtGrNo.focus	
			Set gActiveElement = document.activeElement
		End If	
		
End Function
'==============================================================================================================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
	
	arrParam(4) = "B_Pur_Grp.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    arrHeader(2) = "구매조직"		
    arrHeader(3) = "구매조직명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"			
	arrParam(1) = "B_Biz_Partner"
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	arrParam(3) = ""								
	
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "		'사외거래처만"	
	arrParam(5) = "공급처"			
	
    arrField(0) = "BP_CD"				
    arrField(1) = "BP_NM"				

	arrHeader(0) = "공급처"		
	arrHeader(1) = "공급처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
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
        Case 6                                                              'Lot 순번 Maker Lot 순번 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub
'==============================================================================================================================
Sub CookiePage()

	Dim strTemp

	strTemp = ReadCookie("MvmtNo")
	
	If strTemp = "" then Exit sub
	
	frm1.txtGrNo.value = ReadCookie("MvmtNo")
	
	Call WriteCookie("MvmtNo" , "")
	
	FncQuery()
End Sub
'==============================================================================================================================
Sub Form_Load()
	    
    Call LoadInfTB19029                                                  
    Call ggoOper.LockField(Document, "N")                                
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                 
    Call SetDefaultVal
    Call InitVariables
    Call CookiePage()
End Sub
'==============================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    IF lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows <= 0 Then
		Call SetPopupMenuItemInf("0000111111")
	ElseIf lgIntFlgMode <> Parent.OPMD_UMODE And frm1.vspdData.MaxRows > 0 Then	'참조시 
		Call SetPopupMenuItemInf("0001111111")
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
	
End Sub
'==============================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then Exit Sub
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub
End Sub
'==============================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'==============================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'==============================================================================================================================
Sub PopRestoreSpreadColumnInf()
    
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    Call ChangeTag(True)
End Sub
'==============================================================================================================================
Sub txtGmDt_DblClick(Button)
	if Button = 1 then
		frm1.txtGmDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtGmDt.focus
	End if
End Sub
'==============================================================================================================================
Sub txtGmDt_Change()
	lgBlnFlgChgValue = true	
End Sub
'==============================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
    
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
End Sub
'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
    
		if Col = C_LotNoPop then
			Call OpenLotNo()
		elseif Col = C_SlPop then
			Call OpenSl()
		End if
    End With
End Sub
'==============================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'==============================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	On Error Resume Next                                   
	
	ggoSpread.Source = frm1.vspdData
	
    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    
    If Not chkField(Document, "1") Then	Exit Function
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.ActiveElement   
    
End Function
'==============================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    On Error Resume Next                                   
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ChangeTag(False)
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call SetDefaultVal
    Call InitVariables
    
    FncNew = True                                           
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncDelete() 
    
	Dim IntRetCD
	
	On Error Resume Next 
    Err.Clear                                               
    
    FncDelete = False
    
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function
    														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then            
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    FncDelete = True                            
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim intIndex
    
    FncSave = False                             
    
    On Error Resume Next                       
    Err.Clear                                   

	ggoSpread.Source = frm1.vspdData            
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  
        IntRetCD = DisplayMsgBox("900001","X","X","X")    
        Exit Function
    End If

    If Not chkField(Document, "2") Then  Exit Function
    
    ggoSpread.Source = frm1.vspdData                 
    If Not ggoSpread.SSDefaultCheck Then  Exit Function
    If frm1.vspdData.Maxrows < 1 then Exit Function
    
    For intIndex = 1 to frm1.vspdData.MaxCols 
		frm1.vspdData.SetColItemData intindex,0
	Next
	
    If DbSave = False Then Exit Function
    
    FncSave = True                                                       
    Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function FncCopy() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
	With frm1
		if .vspdData.Maxrows < 1 then exit function

		.vspdData.Col=C_LotFlg
		.vspdData.Row = .vspdData.ActiveRow
		if .vspdData.Text <> "Y" then exit function
    
		ggoSpread.Source = .vspdData	
		ggoSpread.CopyRow

		.vspdData.ReDraw = False
	
		if .vspdData.Text <> "Y" then
			ggoSpread.spreadUnlock C_LotNo, .vspdData.Row, C_LotNoPop, .vspdData.Row
			ggoSpread.SSSetRequired		C_LotNo, .vspdData.Row, .vspdData.Row
		else
			ggoSpread.spreadlock C_LotNo, .vspdData.Row, C_LotNoSeq, .vspdData.Row
		end if
	
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

		.vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncCancel()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo 
    
    Set gActiveElement = document.ActiveElement                                                  
End Function
'==============================================================================================================================
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
		ggoSpread.InsertRow, imRow
		'Call SetSpreadLock()
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
	End With
    
    If Err.number = 0 Then FncInsertRow = True
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    	.focus
		ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncPrint()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncExcel()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear
	
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_SINGLEMULTI)	
    
    Set gActiveElement = document.ActiveElement							
End Function
'==============================================================================================================================
Function FncFind() 
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False) 
    Set gActiveElement = document.ActiveElement                                 
End Function
'==============================================================================================================================
Function FncExit()
	Dim IntRetCD
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
	
	FncExit = False
	ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		If IntRetCD = vbNo Then Exit Function
	End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function DbDelete() 
    Dim strVal
    
    On Error Resume Next       
    Err.Clear                                                           
	
	DbDelete = False													
    frm1.txtMode.value = Parent.UID_M0003
    
    If LayerShowHide(1) = False Then Exit Function
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)							
    DbDelete = True                                             
	Set gActiveElement = document.ActiveElement 
End Function
'==============================================================================================================================
Function DbDeleteOk()											
	lgBlnFlgChgValue = False
	Call MainNew()
End Function
'==============================================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey     
    Dim strVal
        
    DbQuery = False
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear    
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & .hdnGrNo.value
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtGrNo=" & Trim(.txtGrNo.value)
		End if
    
		Call RunMyBizASP(MyBizASP, strVal)							
    End With
    
    DbQuery = True
	Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function DbQueryOk()											
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear
    
    lgIntFlgMode = Parent.OPMD_UMODE									
    
    Call ggoOper.LockField(Document, "Q")						
	Call ChangeTag(True)
	lgBlnFlgChgValue = False	
	Call SetToolBar("11101011000111")
	
	Call RemovedivTextArea
	
	if interface_Account = "N" then		
		frm1.btnGlSel.disabled = true
	Else 
		frm1.btnGlSel.disabled = False		
	End if
	frm1.vspdData.focus
End Function
'==============================================================================================================================
Function DbSave() 

msgbox "DbSave"
	'On Error Resume Next                                                          '☜: If process fails
    Err.Clear  
    Dim lRow        
    Dim strVal, strDel
	Dim iColSep, iRowSep
	
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
    
	Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar

    If LayerShowHide(1) = False Then
		Exit Function
	End If 

    iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	
	strVal = ""
	strDel = ""
    
	With frm1
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"))) = "0" then
						Call DisplayMsgBox("970021","X","입출고수량","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Exit Function
					End if
					
					strVal = "C" & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantNm,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Spec,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow, "X","X"))				& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlPop,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SlNm,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNoPop,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_LotNoSeq,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RetType,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RetTypeNm,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PoSeq,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GmNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_GmSeq,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))					& iColSep 
					strVal = strVal & lRow & iColSep

					' 2005.12.16 Remark_Dtl 추가 
					strVal = strVal & iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & iColSep & iColSep & iColSep & iColSep & iColSep
					strVal = strVal & iColSep & iColSep
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_RemarkDtl,lRow, "X","X"))				& iRowSep
				Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PlantNm,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_ItemNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_Spec,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow, "X","X"))				& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_GrQty,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_SlCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_SlPop,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_SlNm,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_LotNo,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_LotNoPop,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_LotNoSeq,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_RetType,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_RetTypeNm,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PoNo,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_PoSeq,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GmNo,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_GmSeq,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_MvmtNo,lRow, "X","X"))					& iColSep 
					If Trim(UNICDbl(GetSpreadText(frm1.vspdData,C_RetOrdQty,lRow, "X","X"))) >  "0" then 
						Call DisplayMsgBox("172126","X",lRow & "행","X")
						Call RemovedivTextArea
						Call LayerShowHide(0)
						Exit Function
					End if
					strDel = strDel & lRow & iRowSep
		   	End Select 

			.vspdData.Row = lRow
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
		Next
	End With
	
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
	
	'------ Developer Coding part (End ) -------------------------------------------------------------- 
msgbox "DbSave 100"
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
msgbox "DbSave 200"

	If Err.number = 0 Then	 
	   DbSave = True                                                             '☜: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement         
End Function
'==============================================================================================================================
Function DbSaveOk()												

	Call InitVariables
	Call ChangeTag(true)
	Call MainQuery()
	
End Function
'==============================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================
Function changeMvmtType()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                            
  	If CheckRunningBizProcess = True Then
		Exit Function
	End If      
    changeMvmtType = False               
    
    If LayerShowHide(1) = False Then
         Exit Function
    End If  
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeMvmtType"
    strVal = strVal & "&txtMvmtType=" & Trim(frm1.txtMvmtType.Value)
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true
    
    changeMvmtType = True                

End Function
'==============================================================================================================================
Function changeSpplCd()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    If CheckRunningBizProcess = True Then
		Exit Function
	End If 
    changeSpplCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeSpplCd"
    strVal = strVal & "&txtSupplierCd=" & FilterVar(Trim(frm1.txtSupplierCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true
    
    changeSpplCd = True            

End Function
'==============================================================================================================================
Function changeGroupCd()
	
	If gLookUpEnable = False Then
		Exit Function
	End If
	
    Err.Clear                        
    If CheckRunningBizProcess = True Then
		Exit Function
	End If 
    changeGroupCd = False           
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
    
    Dim strVal    
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "changeGroupCd"
    strVal = strVal & "&txtGroupCd=" & FilterVar(Trim(frm1.txtGroupCd.Value),"","SNM")
    
    Call RunMyBizASP(MyBizASP, strVal)
	
	lgBlnFlgChgValue = true
    
    changeGroupCd = True            

End Function
'==============================================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구매반품</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPORef()">반품발주내역참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>입출고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입출고번호" NAME="txtGrNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGmNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
								<TD CLASS="TD5" NOWRAP>입출고유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="23NXXU" OnChange="VBScript:changeMvmtType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>입출고일</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입출고일 NAME="txtGmDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="23N1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" OnChange="VBScript:changeSpplCd()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>	
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="25XXXU" OnChange="VBScript:changeGroupCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupNm" SIZE=20 tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입출고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고번호" NAME="txtGrNo1" SIZE=34 MAXLENGTH=18 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="비고" NAME="txtRemark" MAXLENGTH=120 SIZE=34 tag="21XXXU"></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td>						
		         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
					</td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGrNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
