'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : U1118MA1
'*  4. Program Name         : 
'*  5. Program Desc         : 입고예정일등록 (Manage Planned Delivery Date)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/27
'*  8. Modified date(Last)  : 2004/08/12
'*  9. Modifier (First)     : NHG
'* 10. Modifier (Last)      : NHG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************

Const BIZ_PGM_ID			= "U1118MB1.asp"			'☆: List & Manage SCM Orders

Dim C_BPCd
Dim C_BPNm
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_DvryPlanDt
Dim C_DvryQty
'Dim C_M_Type
Dim C_D_Bp_Cd
Dim C_D_Bp_Pop
Dim C_D_Bp_Nm
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_SplitSeqNo
Dim C_DvryDt
Dim C_OrderQty
Dim C_OrderDt
Dim	C_RcptQty
Dim C_DLVYNO
Dim C_RetFlg
Dim	C_RcptFlg

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgIntGrpCount = 0
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1
End Sub

'================================================================================================================================
Sub InitSpreadComboBox()

End Sub

'================================================================================================================================
Sub InitData()

	Dim intRow
    Dim intIndex
    
	With frm1.vspdData
		.ReDraw = False
			
			For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_DLVYNO
				If .text <> ""  Then
				
					ggoSpread.SSSetProtected C_DvryPlanDt,	intRow, intRow
					ggoSpread.SSSetProtected C_DvryQty,	intRow, intRow
					ggoSpread.SSSetProtected C_D_BP_CD,	intRow, intRow
					ggoSpread.SSSetProtected C_D_BP_Pop,	intRow, intRow
			    End If
			Next
			
		.ReDraw = True	
	End With    
	
End Sub
'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtDvFrDt.text = UniConvDateAToB(UNIDateAdd ("D", 0, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtDvToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
End Sub

'================================================================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()    
    With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20070124", , Parent.gAllowDragDropSpread
		.vspdData.ReDraw = False
	
		.vspdData.MaxCols = C_RcptFlg + 1
		.vspdData.MaxRows = 0
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_BPCd,			"공급처"		,	8,,,10,2
		ggoSpread.SSSetEdit		C_BPNm,			"공급처명"		,	20
		ggoSpread.SSSetEdit		C_ItemCode,		"품목"			,	18,,,18,2
		ggoSpread.SSSetEdit		C_ItemName,		"품목명"		,	20
		ggoSpread.SSSetEdit		C_Spec,			"규격"			,	15
		ggoSpread.SSSetDate 	C_DvryPlanDt,	"입고예정일자"	,	12, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_DvryQty,		"입고예정수량"	,	12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_D_BP_CD,		"입고창고"		,	10,,,10, 2
		ggoSpread.SSSetButton	C_D_BP_Pop
		ggoSpread.SSSetEdit		C_D_BP_Nm,		"창고명"		,	20
		ggoSpread.SSSetEdit		C_OrderUnit,	"단위"			,	4,,,3,2
		ggoSpread.SSSetEdit		C_OrderNo,		"발주번호"		,	15
		ggoSpread.SSSetEdit		C_OrderSeq,		"순번"			,	4
		ggoSpread.SSSetDate 	C_DvryDt,		"납기일"		,	12, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C_OrderQty,		"발주량"		,	12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate 	C_OrderDt,		"발주일자"		,	12, 2, parent.gDateFormat		 
		ggoSpread.SSSetFloat	C_RcptQty,		"입고수량"		,	12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_DLVYNO,	    "납품명세서발행번호"    ,14
		ggoSpread.SSSetEdit		C_RetFlg,		"구분"    ,10,2
		ggoSpread.SSSetEdit		C_RcptFlg,		"입출고구분"    ,10,2
		
		Call ggoSpread.SSSetColHidden( C_SplitSeqNo, C_SplitSeqNo, True)
		Call ggoSpread.SSSetColHidden( C_DLVYNO, C_DLVYNO, True)
		Call ggoSpread.SSSetColHidden( .vspdData.MaxCols, .vspdData.MaxCols , True)
		
		.vspdData.ReDraw = True
   
		ggoSpread.SSSetSplit2(3)
    
    End With
    
    Call SetSpreadLock()
    
End Sub

'================================================================================================================================
Sub SetSpreadLock()

    With frm1
	ggoSpread.Source = .vspdData	
	.vspdData.ReDraw = False
	
   	ggoSpread.SpreadLock	 C_BPCd, -1, C_BPCd
	ggoSpread.SpreadLock	 C_BPNm, -1, C_BPNm
   	ggoSpread.SpreadLock	 C_ItemCode, -1, C_ItemCode
	ggoSpread.SpreadLock	 C_ItemName, -1, C_ItemName
	ggoSpread.SpreadLock	 C_Spec, -1, C_Spec
	ggoSpread.SSSetRequired  C_DvryPlanDt, -1
	ggoSpread.SSSetRequired  C_DvryQty, -1
	'ggoSpread.SSSetRequired  C_M_Type, -1
	ggoSpread.SSSetRequired  C_D_Bp_Cd, -1
	
	ggoSpread.SpreadLock	 C_D_Bp_Nm, -1, C_D_Bp_Nm
	ggoSpread.SpreadLock	 C_OrderUnit, -1, C_OrderUnit
	ggoSpread.SpreadLock	 C_OrderNo, -1, C_OrderNo
	ggoSpread.SpreadLock	 C_OrderSeq, -1, C_OrderSeq
	ggoSpread.SpreadLock	 C_SplitSeqNo, -1, C_SplitSeqNo
	ggoSpread.SpreadLock	 C_DvryDt, -1, C_DvryDt
	ggoSpread.SpreadLock	 C_OrderQty, -1, C_OrderQty
   	ggoSpread.SpreadLock	 C_OrderDt, -1, C_OrderDt
   	ggoSpread.SpreadLock	 C_RcptQty, -1, C_RcptQty
   	ggoSpread.SpreadLock	 C_RetFlg, -1, C_RetFlg
   	ggoSpread.SpreadLock	 C_RcptFlg, -1, C_RcptFlg
	.vspdData.ReDraw = True
    End With

End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData

    ggoSpread.SSSetProtected C_BPCd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BPNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCode,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemName,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec,		pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_DvryPlanDt,	pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_DvryQty,		pvStartRow, pvEndRow
    'ggoSpread.SSSetRequired  C_M_Type,		pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_D_Bp_Cd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_D_Bp_Nm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderUnit,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderNo,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderSeq,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SplitSeqNo,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DvryDt,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderQty,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderDt,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RcptQty,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RetFlg,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_RcptFlg,		pvStartRow, pvEndRow

    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
    
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()	

	C_BPCd			= 1
	C_BPNm			= 2
	C_ItemCode		= 3
	C_ItemName		= 4
	C_Spec			= 5
	C_DvryPlanDt	= 6
	C_DvryQty		= 7
	'C_M_Type		= 8
	C_D_Bp_Cd		= 8
	C_D_Bp_Pop		= 09
	C_D_Bp_Nm		= 10
	C_OrderUnit		= 11
	C_OrderNo		= 12
	C_OrderSeq		= 13
	C_SplitSeqNo	= 14
	C_DvryDt		= 15
	C_OrderQty		= 16
	C_OrderDt		= 17
	C_RcptQty		= 18
	C_DLVYNO		= 19
	C_RetFlg		= 20
	C_RcptFlg		= 21
	
End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_BPCd			= iCurColumnPos(1)
		C_BPNm			= iCurColumnPos(2)
		C_ItemCode		= iCurColumnPos(3)
		C_ItemName		= iCurColumnPos(4)
		C_Spec			= iCurColumnPos(5)
		C_DvryPlanDt	= iCurColumnPos(6)
		C_DvryQty		= iCurColumnPos(7)
		'C_M_Type		= iCurColumnPos(8)
		C_D_Bp_Cd		= iCurColumnPos(8)
		C_D_Bp_Pop		= iCurColumnPos(09)
		C_D_Bp_Nm		= iCurColumnPos(10)	
		C_OrderUnit		= iCurColumnPos(11)
		C_OrderNo		= iCurColumnPos(12)
		C_OrderSeq		= iCurColumnPos(13)
		C_SplitSeqNo	= iCurColumnPos(14)
		C_DvryDt		= iCurColumnPos(15)
		C_OrderQty		= iCurColumnPos(16)
		C_OrderDt		= iCurColumnPos(17)
		C_RcptQty		= iCurColumnPos(18)
		C_DLVYNO		= iCurColumnPos(19)
		C_RetFlg			= iCurColumnPos(20)
		C_RcptFlg			= iCurColumnPos(21)
		
 	End Select
 
End Sub

Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
	With frm1.vspdData
		 ggoSpread.Source = frm1.vspdData
		 .Row = Row
         .Col = Col
		If Row > 0 Then
			Select Case Col
				
				Case C_D_BP_POP
					.Col = Col - 1
			    	.Row = Row
					
					Call OPENBPCD2()
				
				
			End Select
		End If
    
	End With	
End Sub

'================================================================================================================================
Function OpenBpCd2()
	
	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	Dim iCalledAspName, IntRetCD

	IsOpenPop = True
	
		
	arrParam(0) = "납입처"										' 팝업 명칭 
	arrParam(1) = "B_STORAGE_LOCATION"								' TABLE 명칭 
	frm1.vspdData.Row  = frm1.vspdData.ActiveRow
	frm1.vspdData.Col  = C_D_BP_CD
	arrParam(2) = Trim(frm1.vspdData.value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "SL_TYPE = 'E' "'AND PLANT_CD = '" & frm1.txtPlantCd.value & "' "								' Where Condition
	arrParam(5) = "납입처"										' TextBox 명칭 
	
    arrField(0) = "SL_CD"										' Field명(0)
    arrField(1) = "SL_NM"										' Field명(1)
    
    arrHeader(0) = "납입처"										' Header명(0)
    arrHeader(1) = "납입처명"									' Header명(1)
    
	
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		
		Exit Function
	Else
		Frm1.vspddata.col = C_D_BP_CD
		frm1.vspddata.Value = arrRet(0)
		Frm1.vspddata.col = C_D_BP_NM
		frm1.vspddata.Value = arrRet(1)
		 ggoSpread.UpdateRow frm1.vspddata.activeRow
	End If	
End Function


'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'================================================================================================================================
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = "품목"							' 팝업 명칭 
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtItemCd.Value)		' Code Condition
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		' Name Condition
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = 'N' "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd='" & FilterVar(Trim(frm1.txtPlantCd.Value), "", "SNM") & "'"    
	End if
	arrParam(5) = "품목"							' TextBox 명칭 

    arrField(0) = "B_Item.Item_Cd"					' Field명(0)
    arrField(1) = "B_Item.Item_NM"					' Field명(1)
    arrField(2) = "B_Plant.Plant_Cd"				' Field명(2)
    arrField(3) = "B_Plant.Plant_NM"				' Field명(3)
    
    arrHeader(0) = "품목"							' Header명(0)
    arrHeader(1) = "품목명"							' Header명(1)
    arrHeader(2) = "공장"							' Header명(2)
    arrHeader(3) = "공장명"							' Header명(3)
    
	arrRet = window.showModalDialog("../m1/m1111pa1.asp", Array(window.parent,arrParam, arrField, arrHeader), _
		"dialogWidth=695px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"										' 팝업 명칭 
	arrParam(1) = "B_Biz_Partner"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBpCd.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	' Where Condition
	arrParam(5) = "공급처"										' TextBox 명칭 
	
    arrField(0) = "BP_CD"										' Field명(0)
    arrField(1) = "BP_NM"										' Field명(1)
    
    arrHeader(0) = "공급처"										' Header명(0)
    arrHeader(1) = "공급처명"									' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
		
	End If	
End Function

'================================================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
End Function

'================================================================================================================================
Function SetItemInfo(Byval arrRet)
    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtDvFrDt.Text
	arrParam(4) = frm1.txtDvToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function



'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
End Function

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub
 
'================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'================================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")
  	Else
  		Call SetPopupMenuItemInf("1001111111")
  	End If

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
	        If lgSortKey = 1 Then
        	    ggoSpread.SSSort Col
	            lgSortKey = 2
	        Else
        	    ggoSpread.SSSort Col, lgSortKey
	            lgSortKey = 1
        	End If
	End If

End Sub

'================================================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If

End Sub
 
'================================================================================================================================
Sub vspddata_KeyPress(index , KeyAscii )
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub


'================================================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'================================================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    On Error Resume Next                                                    '☜: Protect system from crashing
End Sub

'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If CheckRunningBizProcess = True Then Exit Sub
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If
    End if
    
End Sub



'================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
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
Sub txtDvFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvFrDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvFrDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDvToDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtDvToDt.Focus
    End If
End Sub

'================================================================================================================================
Sub txtDvFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtDvToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
Function FncQuery()
 
    Dim IntRetCD 
    
    ncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
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
End Function

'================================================================================================================================
Function FncDelete() 
	On Error Resume Next   
End Function

'================================================================================================================================
Function FncSave()
    Dim IntRetCD 
         
    FncSave = False 
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
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
        
    If frm1.vspdData.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData.focus
    Set gActiveElement = document.activeElement 
    frm1.vspdData.EditMode = True
	    
    frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    frm1.vspdData.ReDraw = True
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

End Function

'================================================================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'================================================================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function	
    ggoSpread.EditUndo
    Call initData(frm1.vspdData.ActiveRow)
End Function

'================================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
Dim IntRetCD
Dim imRow
Dim pvRow
	
On Error Resume Next
	
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
	.vspdData.focus
	Set gActiveElement = document.activeElement 
	ggoSpread.Source = .vspdData
	.vspdData.ReDraw = False
	ggoSpread.InsertRow .vspdData.ActiveRow, imRow
    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
	.vspdData.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement
	If Err.number = 0 Then FncInsertRow = True
End Function


'================================================================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData.MaxRows < 1 Then Exit Function

'--------------

	With frm1.vspdData
	.Col = C_DLVYNO
	.Row = .ActiveRow
		If .value <> ""  Then
			Call DisplayMsgBox("SCM014","X","X","X")
			Exit Function
		End If
	End With	
'----------

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
    
    ggoSpread.Source = frm1.vspdData							'⊙: Preset spreadsheet pointer 
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
Sub MakeKeyStream()

	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hDvFrDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hDvToDt.value)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.htrackingno.value)  & Parent.gColSep
		Else
			lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txttrackingno.VALUE)  & Parent.gColSep
							
			.hPlantCd.value		= .txtPlantCd.value
			.hItemCd.value		= .txtItemCd.value
			.hBPCd.value		= .txtBPCd.value
			.hDvFrDt.value		= .txtDvFrDt.Text
			.hDvToDt.value		= .txtDvToDt.Text
			.htrackingno.value	= .txttrackingno.value
			
		End If
	End With
	   
End Sub

'================================================================================================================================
Function DbQuery() 
    
    Dim strVal
    
    Err.Clear

    DbQuery = False
    
    Call LayerShowHide(1)
 
    Call MakeKeyStream()
    
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
    Call SetToolBar("11001011000111")

    '-----------------------
    'Reset variables area
    '-----------------------
    
    call initdata()
    
    lgIntFlgMode = parent.OPMD_UMODE
	Frm1.vspddata.focus
End Function

'================================================================================================================================
Function DbQueryNotOk()	

	Call SetToolBar("11001101001111")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE

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
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text

               Case  ggoSpread.UpdateFlag                                      '☜: Update
               
													  strVal = strVal & "U"  &  parent.gColSep					
													  strVal = strVal & lRow &  parent.gColSep
					.vspdData.Col = C_OrderNo	    : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep	'2
					.vspdData.Col = C_OrderSeq     	: strVal = strVal & Trim(.vspdData.Value)&  parent.gColSep	'3
					.vspdData.Col = C_SplitSeqNo   	: strVal = strVal & Trim(.vspdData.Value)&  parent.gColSep	'4
					.vspdData.Col = C_DvryPlanDt	: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep	'5
					
					.vspdData.Col = C_DvryQty	:
					If Trim(.vspdData.Value) > 0 Then 
						strVal = strVal & Trim(.vspdData.Value) & parent.gColSep	'6
					Else
						Call DisplayMsgBox("169918","X", "X", "X")
						Call LayerShowHide(0) 
						.Action = 0
						Exit Function
					End If
													  strVal = strVal & "N"					  & parent.gColSep
					.vspdData.Col = C_D_Bp_Cd     	: strVal = strVal & Trim(.vspdData.Value)&  parent.gRowSep	'8

                   
					lGrpCnt = lGrpCnt + 1

               Case  ggoSpread.DeleteFlag                                      '☜: Delete

													  strDel = strDel & "D"  &  parent.gColSep
													  strDel = strDel & lRow &  parent.gColSep
                    .vspdData.Col = C_OrderNo	    : strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep	'2
                    .vspdData.Col = C_OrderSeq     	: strDel = strDel & Trim(.vspdData.Value)&  parent.gColSep	'3
					.vspdData.Col = C_SplitSeqNo   	: strDel = strDel & Trim(.vspdData.Value)&  parent.gColSep	'4
                    .vspdData.Col = C_DvryPlanDt	: strDel = strDel & Trim(.vspdData.Text) &  parent.gColSep	'5
                    .vspdData.Col = C_DvryQty		: strDel = strDel & Trim(.vspdData.Value) & parent.gColSep	'6
                    .vspdData.Col = C_D_Bp_Cd     	: strDel = strDel & Trim(.vspdData.Value)&  parent.gRowSep	'8

                    lGrpCnt = lGrpCnt + 1

           End Select
       Next

       .txtMode.value        =  parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
    DbSave = True
    
End Function

'================================================================================================================================
Function DbSaveOk()

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
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
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
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
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function
