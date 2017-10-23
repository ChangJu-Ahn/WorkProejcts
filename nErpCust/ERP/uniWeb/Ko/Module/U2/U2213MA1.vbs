'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2004/08/12
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************

Const BIZ_PGM_ID	= "U2213MB1.asp"			'☆: List & Manage SCM Orders

Dim C_Dt
Dim C_ItemCd
Dim C_ItemPop
Dim C_ItemNm
Dim C_Spec
Dim	C_Qty
Dim C_Unit
Dim C_UnitPop
Dim C_RetType
Dim C_RetTypePop
Dim C_RetTypeNm
Dim C_STATUS
Dim C_STATUS2
Dim C_PlantCd
Dim C_PlantPop
Dim C_PlantNm
Dim C_BpCd
Dim C_BpPop
Dim C_BpNm
Dim C_PoNo
Dim C_PoSeqNo

Dim IsOpenPop
'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgStrPrevKey1 = ""
    lgLngCurRows = 0
    lgSortKey1 = 1
	lgSortKey2 = 1
	lgOldRow = 0
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
End Sub

'================================================================================================================================
Sub InitSpreadComboBox()

End Sub

'================================================================================================================================
Sub InitData()

	Dim intRow
    Dim intIndex
    Dim j
    
    With frm1.vspdData1
	
		j = .row
		For j = 1 To .MaxRows
		.Row = j
		.Col = C_STATUS2

		If .value <> "P" Then
			ggoSpread.SpreadLock    C_Dt, j , C_PoSeqNo     ,j 
		End If

		Next

    End With

End Sub

'================================================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	call SetBPCD()
	
	Call ExtractDateFrom(LocSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
	frm1.txtRetFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtRetToDt.text = UniConvDateAToB(UNIDateAdd ("M", 0, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	
	If parent.gPlant <> "" and frm1.txtPlantCd.Value = "" Then
		frm1.txtPlantCd.Value = parent.gPlant
		frm1.txtPlantNm.Value = parent.gPlantNm
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement 
	End If
	
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call DisplayMsgBox("210033","X","X","X")
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtbpcd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtbpcd2,"Q")
		Call ggoOper.SetReqAttr(frm1.txtRetFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtRetToDt,"Q")
		
		Call ggoOper.SetReqAttr(frm1.rdoflg1,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoflg2,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoflg3,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoflg4,"Q")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
	    Call SetToolBar("11001111000111")								'⊙: 버튼 툴바 제어 
	End If

	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpCd.value = parent.gUsrId
	frm1.txtBpNm.value = lgF0(1)
	
End Sub

'================================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20030805", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False
	
			.vspdData1.MaxCols = C_PoSeqNo + 1
			.vspdData1.MaxRows = 0

			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetDate 	C_Dt,		"반품일자",12, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ItemCd,	"품목"    ,15
			ggoSpread.SSSetButton 	C_itemPop
			ggoSpread.SSSetEdit		C_ItemNm,	"품목명"  ,20
			ggoSpread.SSSetEdit		C_Spec,		"규격"    ,15,,,18,2
			ggoSpread.SSSetFloat	C_Qty,		"수량"    ,12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit,		"단위"    ,7
			ggoSpread.SSSetButton 	C_UnitPop
			ggoSpread.SSSetEdit		C_RetType,	"반품유형"    ,15
			ggoSpread.SSSetButton 	C_RetTypePop
			ggoSpread.SSSetEdit		C_RetTypeNm,"반품유형명"  ,20
			ggoSpread.SSSetEdit		C_Status,	"상태"    ,7
			ggoSpread.SSSetEdit		C_Status2,	"상태코드"    ,7
			ggoSpread.SSSetEdit		C_PlantCd,	"납품업체",12
			ggoSpread.SSSetButton 	C_PlantPop
			ggoSpread.SSSetEdit		C_PlantNm,	"납품업체명", 15
			ggoSpread.SSSetEdit		C_BpCd,		"업체"    ,12
			ggoSpread.SSSetButton 	C_BpPop
			ggoSpread.SSSetEdit		C_BpNm,		"업체명"  ,15
			ggoSpread.SSSetEdit		C_PoNo,		"발주번호",15
			ggoSpread.SSSetEdit		C_PoSeqNo,	"순번"    , 7,,,3,2
			
			Call ggoSpread.SSSetColHidden( C_BPCD, C_BPNM , True)
			'Call ggoSpread.SSSetColHidden( C_STATUS2, C_STATUS2 , True)
			Call ggoSpread.SSSetColHidden( .vspdData1.MaxCols, .vspdData1.MaxCols , True)
			
			.vspdData1.ReDraw = True
    
			Call SetSpreadLock("A")
			
			.vspdData1.ReDraw = true    
    
		End With
	
End Sub

'================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		With frm1
			ggoSpread.Source = .vspdData1
	
			.vspdData1.ReDraw = False
			
   			ggoSpread.SpreadLock	 C_Dt	     ,-1, C_Dt
   			ggoSpread.SpreadLock	 C_ItemCd    ,-1, C_ItemCd
   			ggoSpread.SpreadLock	 C_ITEMPop   ,-1, C_ITEMPop
			ggoSpread.SpreadLock	 C_ItemNm    ,-1, C_ItemNm
			ggoSpread.SpreadLock	 C_Spec      ,-1, C_Spec
			ggoSpread.SSSetRequired  C_QTY       ,-1
			ggoSpread.SSSetRequired  C_Unit		 ,-1
			ggoSpread.SpreadLock	 C_RetTypeNm ,-1, C_RetTypeNm
			ggoSpread.SpreadLock	 C_Status	 ,-1, C_Status
			ggoSpread.SpreadLock	 C_Status2	 ,-1, C_Status2
			ggoSpread.SpreadLock	 C_PlantCd   ,-1, C_PlantCd
			ggoSpread.SpreadLock	 C_PlantPop  ,-1, C_PlantPop
			ggoSpread.SpreadLock	 C_PlantNm   ,-1, C_PlantNm
			ggoSpread.SpreadLock	 C_BPCd      ,-1, C_BPCd
			ggoSpread.SpreadLock	 C_BPPop     ,-1, C_BPPop
			ggoSpread.SpreadLock	 C_BPNm      ,-1, C_BPNm
			ggoSpread.SpreadLock	 C_PONO      ,-1, C_PONO
			ggoSpread.SpreadLock	 C_POSEQNO   ,-1, C_POSEQNO
			
			.vspdData1.ReDraw = True
	
		End With
	End If

End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData1 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData1
    
    ggoSpread.SSSetRequired  C_DT,          pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_ITEMCD,      pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_QTY    ,     pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_unit   ,     pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_PLANTCD,     pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_BPCD,	    pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderNo,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderSeq,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCode,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Status,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Status2,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RetTypeNm,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BpNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PONO,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_POSEQNO,		pvStartRow, pvEndRow
    
    
    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()	
	
	' Grid 1(vspdData1)
		
	C_Dt		= 1
	C_ItemCd	= 2
	C_ItemPop	= 3
	C_ItemNm	= 4
	C_Spec		= 5
	C_Qty		= 6
	C_Unit		= 7
	C_UnitPop	= 8
	C_RetType	= 9
	C_RetTypePop= 10	
	C_RetTypeNm	= 11
	C_Status	= 12
	C_Status2	= 13
	C_PlantCd	= 14
	C_PlantPop	= 15
	C_PlantNm	= 16
	C_BpCd		= 17
	C_BpPop		= 18
	C_BpNm		= 19
	C_PoNo		= 20
	C_PoSeqno	= 21
	
End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Dt		= iCurColumnPos(1)
			C_ItemCd	= iCurColumnPos(2)
			C_ItemPop	= iCurColumnPos(3)
			C_ItemNm	= iCurColumnPos(4)
			C_Spec		= iCurColumnPos(5)
			C_Qty		= iCurColumnPos(6)
			C_Unit		= iCurColumnPos(7)
			C_UnitPop	= iCurColumnPos(8)
			C_RetType	= iCurColumnPos(9)
			C_RetTypePop= iCurColumnPos(10)
			C_RetTypeNm	= iCurColumnPos(11)
			C_Status	= iCurColumnPos(12)
			C_Status2	= iCurColumnPos(13)
			C_PlantCd	= iCurColumnPos(14)
			C_PlantPop	= iCurColumnPos(15)
			C_PlantNm	= iCurColumnPos(16)
			C_BpCd		= iCurColumnPos(17)
			C_BpPop		= iCurColumnPos(18)
			C_BpNm		= iCurColumnPos(19)
			C_PoNo		= iCurColumnPos(20)
			C_PoSeqNo	= iCurColumnPos(21)
			
 	End Select
 
End Sub


'------------------------------------------  OpenConItemCd()  --------------------------------------------
'	Name : OpenConItemCd()
'	DeScription : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(iWhere) 						
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    arrField(2) = 3 							' Field명(1) : "SPEC"
    arrField(3) = 4 							' Field명(1) : "BASIC_UNIT"

	iCalledAspName = AskPRAspName("B1B11PA4")
	    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		If arrRet(0) <> "" Then
				.vspdData1.Col = C_ItemCd
				.vspdData1.Text = arrRet(0)
				.vspdData1.Col = C_ItemNM
				.vspdData1.Text = arrRet(1)
				.vspdData1.Col = C_SPEC
				.vspdData1.Text = arrRet(2)
				.vspdData1.Col = C_UNIT
				.vspdData1.Text = arrRet(3)
		End If	
	End With	
	
	Call SetFocusToDocument("M")
	frm1.vspddata1.focus

End Function

Function OpenPLANTCD(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업체 팝업"					
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.vspdData1.Text)	
	arrParam(4) = ""	
	arrParam(5) = "업체"					
	
	arrParam(3) = ""						
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "업체"					
    arrHeader(1) = "업체명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData1.Col = C_PLANTCD
			.vspdData1.Text = arrRet(0)
			.vspdData1.Col = C_PLANTNM
			.vspdData1.Text = arrRet(1)
			
		end with
	End If	
	
End Function

Function OpenBPCD(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업체 팝업"					
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.vspdData1.Text)	
	arrParam(4) = ""	
	arrParam(5) = "업체"					
	
	arrParam(3) = ""						
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "업체"					
    arrHeader(1) = "업체명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData1.Col = C_BPCD
			.vspdData1.Text = arrRet(0)
			.vspdData1.Col = C_BPNM
			.vspdData1.Text = arrRet(1)
			ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
		end with
	End If	
	
End Function

Function OpenRetType(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "반품사유"					
	arrParam(1) = "B_MINOR"						
	arrParam(2) = Trim(frm1.vspdData1.Text)	
	arrParam(4) = " MAJOR_CD = 'b9017' "	
	arrParam(5) = "반품사유"					
	
	arrParam(3) = ""						
	
    arrField(0) = "MINOR_CD"					
    arrField(1) = "MINOR_NM"					
    
    arrHeader(0) = "반품사유"					
    arrHeader(1) = "반품사유명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData1.Col = C_RetType
			.vspdData1.Text = arrRet(0)
			.vspdData1.Col = C_RetTypeNm
			.vspdData1.Text = arrRet(1)
			ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
		end with
	End If	
	
End Function

Function OpenPLANT(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장 팝업"					
	arrParam(1) = "B_PLANT "						
	arrParam(2) = Trim(iWhere)	
	arrParam(4) = ""	
	arrParam(5) = "공장"					
	
	arrParam(3) = ""						
	
    arrField(0) = "PLANT_CD"					
    arrField(1) = "PLANT_NM"					
    
    arrHeader(0) = "공장"					
    arrHeader(1) = "공장명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.txtPlantcd.value = arrRet(0)
			.txtPlantnm.value = arrRet(1)
		end with
	End If	
	
End Function

Function OpenBPCD2(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업체 팝업"					
	arrParam(1) = "B_BIZ_PARTNER "
	arrParam(2) = Trim(iWhere)	
	arrParam(4) = " BP_TYPE IN ('S' ,'CS') "	
	arrParam(5) = "업체"					
	
	arrParam(3) = ""						
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "업체"					
    arrHeader(1) = "업체명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.txtBPcd2.value = arrRet(0)
			.txtBPnm2.value = arrRet(1)
		end with
	End If	
	
End Function

Function OpenUnit(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위 팝업"					
	arrParam(1) = " B_UNIT_OF_MEASURE"						
	arrParam(2) = Trim(frm1.vspdData1.Text)	
	arrParam(4) = " DIMENSION <> 'TM' "	
	arrParam(5) = "단위"
	
	arrParam(3) = ""						
	
    arrField(0) = "UNIT"					
    arrField(1) = "UNIT_NM"					
    
    arrHeader(0) = "단위"					
    arrHeader(1) = "단위명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		with frm1
			.vspdData1.Col = C_UNIT
			.vspdData1.Text = arrRet(0)
			ggoSpread.UpdateRow frm1.vspdData1.ActiveRow
		end with
	End If	
	
End Function

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
 
'================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("1101111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    	
End Sub

'================================================================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
 
'================================================================================================================================
Sub vspdData1_KeyPress(index , KeyAscii )
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub

'================================================================================================================================
Sub vspdData1_ButtonClicked(Col, Row, ButtonDown)
	With frm1.vspdData1
		 ggoSpread.Source = frm1.vspdData1
		 .Row = Row
         .Col = Col
		If Row > 0 Then
			Select Case Col
				
				Case C_ItemPop
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenITEMCD(.text)
				
				Case C_PlantPop
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenPLANTCD(.text)
					
				Case C_BpPop
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenBPCD(.text)	
							
				Case C_UnitPop
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenUnit(.text)	
				
				Case C_RetTypePop
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenRetType(.text)		
			End Select
		End If
    
	End With	
End Sub


'================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'================================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()

End Sub 

'================================================================================================================================
Sub txtRetFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtRetFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtRetFrDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtRetToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtRetToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtRetToDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtRetFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'================================================================================================================================
Sub txtRetToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'================================================================================================================================
Function BtnPreview()
    
    Dim strEbrFile
    Dim objName
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	dim var8
	dim var9
	dim var10
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)
	
	var1 = "0"
	var2 = "zzzz"
	var3 = "0"
	var4 = "zzzzzzzzzzzzzzzzzz"
	var5 = "0"
	var6 = "zzzzzzzzzz"		
	var7 = "1900-01-01" 
	var8 = "2999-12-31"
	
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_OrderNo
		
	var9 = Trim(frm1.vspdData1.text)
	var10 = "%"
		
	strUrl = strUrl & "FRBIZ_AREA_CD|" & var1
	strUrl = strUrl & "|TOBIZ_AREA_CD|" & var2 
	strUrl = strUrl & "|FRITEM_CD|" & var3 
	strUrl = strUrl & "|TOITEM_CD|" & var4
	strUrl = strUrl & "|FRBP_CD|" & var5 
	strUrl = strUrl & "|TOBP_CD|" & var6
	strUrl = strUrl & "|FRPLAN_DVRY_DT|" & var7 
	strUrl = strUrl & "|TOPLAN_DVRY_DT|" & var8
	strUrl = strUrl & "|PO_NO|" & var9 
	strUrl = strUrl & "|RET_FLG|" & var10

	strEbrFile = "mscm08oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'================================================================================================================================
Function BtnPrint()
	
	Dim strEbrFile
    Dim objName
	
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	dim var8
	dim var9
	dim var10
	Dim var11

	dim strUrl
	dim arrParam, arrField, arrHeader

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)	

	var1 = "0"
	var2 = "zzzz"
	var3 = "0"
	var4 = "zzzzzzzzzzzzzzzzzz"
	var5 = "0"
	var6 = "zzzzzzzzzz"		
	var7 = "1900-01-01" 
	var8 = "2999-12-31"
	
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Col = C_OrderNo
		
	var9 = Trim(frm1.vspdData1.text)
	var10 = "%"	
		
	strUrl = strUrl & "FRBIZ_AREA_CD|" & var1
	strUrl = strUrl & "|TOBIZ_AREA_CD|" & var2 
	strUrl = strUrl & "|FRITEM_CD|" & var3 
	strUrl = strUrl & "|TOITEM_CD|" & var4
	strUrl = strUrl & "|FRBP_CD|" & var5 
	strUrl = strUrl & "|TOBP_CD|" & var6
	strUrl = strUrl & "|FRPLAN_DVRY_DT|" & var7 
	strUrl = strUrl & "|TOPLAN_DVRY_DT|" & var8
	strUrl = strUrl & "|PO_NO|" & var9 
	strUrl = strUrl & "|RET_FLG|" & var10

	strEbrFile = "mscm08oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRprint(EBAction, objName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function


'================================================================================================================================
Function FncQuery()
 
    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True OR lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
	If IntRetCD = vbNo Then
	    Exit Function
	End If
    End If

    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True
   
End Function

'================================================================================================================================
Function FncNew() 
	On Error Resume Next	
	Dim IntRetCD 
    
    FncNew = False                                                  
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                          
    Call ggoOper.ClearField(Document, "2")                          
    Call ggoOper.ClearField(Document, "3")                          
    Call ggoOper.LockField(Document, "N")                              
    Call SetDefaultVal
    Call InitVariables													

    frm1.txtbpcd.focus
	Set gActiveElement = document.activeElement
    
    FncNew = True		
	
End Function

'================================================================================================================================
Function FncDelete() 
	On Error Resume Next   

    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                            '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")                         '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbDelete = False Then                                                      '☜: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement       
    If Err.number = 0 Then
       FncDelete = True                                                              '☜: Processing is OK
    End If   
	
End Function

'================================================================================================================================
Function FncSave()
    Dim IntRetCD 
         
    FncSave = False 
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData1
    If  ggoSpread.SSCheckChange = False And lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If

	Call DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True
    
End Function

'================================================================================================================================
Function FncCopy() 
        
    If frm1.vspdData1.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData1.focus
    Set gActiveElement = document.activeElement 
    frm1.vspdData1.EditMode = True
	    
    frm1.vspdData1.ReDraw = False
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.CopyRow
    frm1.vspdData1.ReDraw = True
    SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow

End Function

'================================================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'================================================================================================================================
Function FncCancel() 
    If frm1.vspdData1.MaxRows < 1 Then Exit Function	
    ggoSpread.EditUndo
    'Call initData(frm1.vspdData1.ActiveRow)
End Function

'================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
Dim IntRetCD
Dim imRow
Dim pvRow

Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
On Error Resume Next
	
    If Trim(Frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		Frm1.txtplantcd.focus
		Exit Function
    End If
	
	FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If
    
    With frm1
	.vspdData1.focus
	Set gActiveElement = document.activeElement 
	ggoSpread.Source = .vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
    SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1
    
   .vspddata1.Col  = C_DT
   .vspddata1.text = LocSvrDate
   .vspddata1.Col  = C_BPCD
   .vspddata1.text = .txtBPCD.VALUE
   .vspddata1.Col  = C_BPNM
   .vspddata1.text = .txtBPNM.VALUE
    
	.vspdData1.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement
	If Err.number = 0 Then FncInsertRow = True
End Function


'================================================================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData1.MaxRows < 1 Then Exit Function

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim iRow
	
	iRow = frm1.vspdData1.ActiveRow
	frm1.vspdData1.Row = iRow
	frm1.vspdData1.Col = C_STATUS2
	
	If frm1.vspdData1.Text <> "P" Then
		Call DisplayMsgBox("202411","X","X","X")
		Exit Function
	End If
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
	
End Function

'================================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'================================================================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'================================================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												
End Function

'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         
End Function

'================================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'================================================================================================================================
Function FncExit()

    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'================================================================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'================================================================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'================================================================================================================================
Sub MakeKeyStream(pOpt)

   		With frm1
			lgKeyStream =               Trim(.txtBpCd.VALUE)   & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtRetFrDt.TEXT) & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtRetToDt.TEXT) & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtBpCd2.VALUE)  & Parent.gColSep
			
			If .rdoflg1.checked = True Then 
				lgKeyStream = lgKeyStream & "P"  & Parent.gColSep
			ElseIf .rdoflg2.checked = True Then
				lgKeyStream = lgKeyStream & "O"  & Parent.gColSep	
			ElseIf .rdoflg3.checked = True Then
				lgKeyStream = lgKeyStream & "A"  & Parent.gColSep		
			ElseIf .rdoflg4.checked = True Then
				lgKeyStream = lgKeyStream & "E"  & Parent.gColSep			
			End If
			
			lgKeyStream = lgKeyStream & Trim(.txtPLANTCD.VALUE)  & Parent.gColSep
			
		End With
	   
End Sub

'================================================================================================================================
Function DbQuery() 
    
    Dim strVal
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Call MakeKeyStream("M")
    
	strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
 
    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'================================================================================================================================
Function DbQueryOk(ByVal LngMaxRow)

 	Dim lRow
 	Dim LngRow    
	
    Call ggoOper.LockField(Document, "Q")
    'ggoOper.SetReqAttr	frm1.txtdlvyno2, "Q"
    Call SetToolBar("11101111000111")

    
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call RestoreToolBar()
	End If
    Call initdata()
    Frm1.vspddata1.focus
	lgIntFlgMode = parent.OPMD_UMODE

End Function

'================================================================================================================================
Function DbQueryNotOk()	

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	DbDelete = False			                                                  '☜: Processing is NG

	Call LayerShowHide(1)

    Call MakeKeyStream("M")
		
    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0003                        '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                      '☜: Query Key

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
		
	Call RunMyBizASP(MyBizASP, strVal)                                            '☜: Run Biz logic
	
    If Err.number = 0 Then
       DbDelete = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'================================================================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
	exit Function
	end if
	
    strVal = ""
    strDel = ""
    lGrpCnt = 1
	
	With Frm1
    
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text

               Case  ggoSpread.InsertFlag                                      '☜: Insert
               
														  strVal = strVal & "C"  &  parent.gColSep					
														  strVal = strVal & lRow &  parent.gColSep
						.vspdData1.Col = C_DT		    : strVal = strVal & Trim(.vspdData1.Text)  & parent.gColSep	'2
						.vspdData1.Col = C_ITEMCD	    : strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'3
						.vspdData1.Col = C_QTY			: strVal = strVal & Trim(.vspdData1.value) & parent.gColSep	'4
						.vspdData1.Col = C_UNIT			: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'5
						.vspdData1.Col = C_PLANTCD		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'6
						.vspdData1.Col = C_BPCD			: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'7
						.vspdData1.Col = C_RETTYPE		: strVal = strVal & Trim(.vspdData1.Value) & parent.gRowSep	'8
						
						lGrpCnt = lGrpCnt + 1
				
				Case  ggoSpread.UpdateFlag                                      '☜: Update
               
														  strVal = strVal & "U"  &  parent.gColSep					
														  strVal = strVal & lRow &  parent.gColSep
						.vspdData1.Col = C_DT		    : strVal = strVal & Trim(.vspdData1.Text)  & parent.gColSep	'2
						.vspdData1.Col = C_ITEMCD	    : strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'3
						.vspdData1.Col = C_QTY			: strVal = strVal & Trim(.vspdData1.value) & parent.gColSep	'4
						.vspdData1.Col = C_UNIT			: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'5
						.vspdData1.Col = C_PLANTCD		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'6
						.vspdData1.Col = C_BPCD			: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'7
						.vspdData1.Col = C_RETTYPE		: strVal = strVal & Trim(.vspdData1.Value) & parent.gRowSep	'8
						
						lGrpCnt = lGrpCnt + 1
				
				
				
				Case  ggoSpread.DeleteFlag                                      '☜: Delete
               
														  strDel = strDel & "D"  &  parent.gColSep					
														  strDel = strDel & lRow &  parent.gColSep
						.vspdData1.Col = C_DT			: strDel = strDel & Trim(.vspdData1.Text)  & parent.gColSep	'2
						.vspdData1.Col = C_ITEMCD		: strDel = strDel & Trim(.vspdData1.Value) & parent.gColSep	'3
						.vspdData1.Col = C_PLANTCD		: strDel = strDel & Trim(.vspdData1.Value) & parent.gColSep	'4
						.vspdData1.Col = C_BPCD			: strDel = strDel & Trim(.vspdData1.Value) & parent.gRowSep	'5
						
						lGrpCnt = lGrpCnt + 1		
						
           End Select
       Next
		
	   .txtFlgMode.value	 = lgIntFlgMode
	   .txtMode.value        =  parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
    DbSave = True
    
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	Call InitVariables()

	Call SetToolbar("1111111111111111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call MainNew()	

    Set gActiveElement = document.ActiveElement   

End Sub


'================================================================================================================================
Function DbSaveOk()

	Call InitVariables
	ggoSpread.source = frm1.vspdData1
    frm1.vspdData1.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function

'================================================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData1.Col = iDx
              Frm1.vspdData1.Row = iRow
              Frm1.vspdData1.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'================================================================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'================================================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)
End Function

'==============================================================================================================================
'Function OpenFirmPORcpt()
Function OpenReqRef()

	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Then Exit Function
	
	'If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
	'	Call DisplayMsgBox("210033","X","X","X")
	'	Call SetToolBar("10000000000011")
	'	Exit Function
	'End If
	
	if Trim(frm1.txtbpcd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "업체", "X")
		frm1.txtbpcd.focus
    	Exit Function
    End if
	
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.txtbpcd.value)
	arrParam(1) = Trim(frm1.txtbpNm.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = ""
	arrParam(8) = ""
	arrParam(9) = ""
	arrParam(10)= ""
	arrParam(11)= ""
	arrParam(12)= ""

	iCalledAspName = AskPRAspName("MSCM1BRA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "MSCM1BRA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function					
	
	If strRet(0,0) = "" Then
		frm1.txtdlvyno.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		Call SetFirmPoRef(strRet)
	End If	
		
End Function
'==============================================================================================================================
Function SetFirmPoRef(strRet)

	Dim Index1, Count1, Row1
	Dim temp
		
	const		C_OrderNo_Ref		= 0
	const		C_OrderSeq_Ref		= 1
	const		C_ItemCode_Ref		= 2
	const		C_ItemName_Ref		= 3
	const		C_Spec_Ref			= 4
	const		C_OrderUnit_Ref		= 5
	const		C_DvryPlanDt_Ref	= 6
	const		C_DvryQty_Ref		= 7
	const		C_PlantCd_Ref		= 8
	const		C_PlantNm_Ref		= 9 
	const		C_SPLITSEQNO_Ref	= 10
    
	Count1 = Ubound(strRet,1)
	
	With frm1.vspdData1
		
		.Redraw = False
	
		Call fncinsertrow(Count1 + 1)
		
		For index1 = 0 to Count1
			
			Row1 = .ActiveRow + Index1
			
			Call .SetText(C_OrderNo,	Row1, strRet(index1,C_OrderNo_Ref))
			Call .SetText(C_OrderSeq,	Row1, strRet(index1,C_OrderSeq_Ref))
			Call .SetText(C_ItemCode,	Row1, strRet(index1,C_ItemCode_Ref))
			Call .SetText(C_ItemName,	Row1, strRet(index1,C_ItemName_Ref))
			Call .SetText(C_Spec,		Row1, strRet(index1,C_Spec_Ref))
			Call .SetText(C_OrderUnit,  Row1, strRet(index1,C_OrderUnit_Ref))
			Call .SetText(C_DvryPlanDt,	Row1, strRet(index1,C_DvryPlanDt_Ref)) 
			Call .SetText(C_DvryQty,	Row1, strRet(index1,C_DvryQty_Ref))
			Call .SetText(C_PlantCd,	Row1, strRet(index1,C_PlantCd_Ref))
			Call .SetText(C_PlantNm,	Row1, strRet(index1,C_PlantNm_Ref))
			Call .SetText(C_SPLITSEQNO,	Row1, strRet(index1,C_SPLITSEQNO_Ref))
			
		Next
	
		'Call LocalReFormatSpreadCellByCellByCurrency()
		'Call setReference()
		lgBlnFlgChgValue = True
		
		.ReDraw = True
		
	End with
	
End Function 

Sub txtDlvyTime_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDlvyTime2_Change()
	lgBlnFlgChgValue = True
End Sub