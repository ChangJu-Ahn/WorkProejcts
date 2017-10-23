
'==========================================================================================================
Const BIZ_PGM_QRY_ID		= "p2218mb1.asp"
Const BIZ_PGM_SAVE_ID		= "p2218mb2.asp"
Const BIZ_PLANT_ID			= "p2218mb3.asp"

Dim C_Select	
Dim C_ItemCode	
Dim C_ItemName	
Dim C_ItemSpec	
Dim C_TrackingNo
Dim C_PlndDt		
Dim C_PlndQty		
Dim C_Unit			
Dim C_MPSConfirmFlg	
Dim C_MRPConfirmFlg	
Dim C_MPSStatus		
Dim C_MPSNo			
Dim C_ProdEnv		
Dim C_MaxMrpQty		
Dim C_MinMrpQty		
Dim C_RondQty	
Dim C_OrderLt	
Dim C_CumulativeLt
Dim C_MPSOrigin	
Dim C_ItemGroupCd
Dim C_ItemGroupNm


Dim ihGridCnt
Dim intItemCnt
Dim IsOpenPop

Dim lsDTF
Dim lsPTF
Dim lsPH

Dim lgQryFlg
Dim lgButtonSelection

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_Select			= 1
    C_ItemCode			= 2
    C_ItemName			= 3
    C_ItemSpec			= 4
    C_TrackingNo		= 5
    C_PlndDt			= 6
    C_PlndQty			= 7
    C_Unit				= 8
    C_MPSConfirmFlg		= 9
    C_MRPConfirmFlg		= 10
    C_MPSStatus			= 11
    C_MPSNo				= 12
    C_ProdEnv			= 13
    C_MaxMrpQty			= 14
    C_MinMrpQty			= 15
    C_RondQty			= 16
    C_OrderLt			= 17
    C_CumulativeLt		= 18
    C_MPSOrigin			= 19
	C_ItemGroupCd		= 20
	C_ItemGroupNm		= 21
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgLngCurRows = 0
	lgSortKey    = 1
	lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
	
End Sub

'==========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'==========================================================================================================
Sub SetDefaultVal()
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Sub


'==========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
        
    Dim sList
    
    Call initSpreadPosVariables()    
    
    With frm1
      
    ggoSpread.Source = .vspdData
	ggoSpread.Spreadinit "V20030224",, parent.gAllowDragDropSpread  
	
	.vspdData.ReDraw = False
	
    .vspdData.MaxCols = C_ItemGroupNm + 1
    .vspdData.MaxRows = 0
	Call AppendNumberPlace("6", "6", "0")
	
    Call GetSpreadColumnPos("A")
       
	ggoSpread.SSSetCheck	C_Select,		"", 2,,,1    
    ggoSpread.SSSetEdit		C_ItemCode,		"품목", 18,,,18,2
    ggoSpread.SSSetEdit		C_ItemName,		"품목명", 25
    ggoSpread.SSSetEdit		C_ItemSpec,		"규격", 25
    ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25,,,25,2
    ggoSpread.SSSetDate 	C_PlndDt,		"계획일", 11, 2, gDateFormat
    ggoSpread.SSSetFloat	C_PlndQty,		"계획수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetEdit		C_Unit,			"단위", 8
    ggoSpread.SSSetEdit		C_MPSConfirmFlg, "MPS 확정여부", 10, 2
    ggoSpread.SSSetEdit		C_MRPConfirmFlg, "MRP 확정여부", 10, 2
    ggoSpread.SSSetEdit		C_MPSStatus,	"Status", 8
    ggoSpread.SSSetEdit		C_MPSNo,		"MPS No.", 18
    ggoSpread.SSSetEdit		C_ProdEnv,		"생산전략", 10
    ggoSpread.SSSetFloat	C_MaxMrpQty,	"Max MRP Qty", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_MinMrpQty,	"Min MRP Qty", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_RondQty,		"Round Qty", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_OrderLt,		"Order LT", 15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_CumulativeLt, "C_CUMULATIVELT", 15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetEdit		C_MPSOrigin,	"MPS Origin", 8
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"품목그룹",		15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"품목그룹명",	30
    
    ggoSpread.Source = .vspdData
    
    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
    Call ggoSpread.SSSetColHidden(C_MPSConfirmFlg, C_MPSConfirmFlg, True)
    Call ggoSpread.SSSetColHidden(C_MRPConfirmFlg, C_MRPConfirmFlg, True)
	Call ggoSpread.SSSetColHidden(C_MaxMrpQty, C_MaxMrpQty, True)
	Call ggoSpread.SSSetColHidden(C_MinMrpQty, C_MinMrpQty, True)
	Call ggoSpread.SSSetColHidden(C_RondQty, C_RondQty, True)
	Call ggoSpread.SSSetColHidden(C_OrderLt, C_OrderLt, True)
	Call ggoSpread.SSSetColHidden(C_CumulativeLt, C_CumulativeLt, True)
	Call ggoSpread.SSSetColHidden(C_MPSOrigin, C_MPSOrigin, True)
	Call ggoSpread.SSSetColHidden(C_ProdEnv, C_ProdEnv, True)

	ggoSpread.SSSetSplit2(3)
	
	.vspdData.ReDraw = True
    
    End With
    
    Call SetSpreadLock()
    
End Sub

'==========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================
Sub SetSpreadLock()
    With frm1
    
	.vspdData.ReDraw = False
	
    ggoSpread.SpreadLock -1,	-1
    ggoSpread.SpreadUnLock C_Select,	-1, C_Select

	.vspdData.ReDraw = True

    End With
End Sub

'==========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'==========================================================================================================
Sub InitComboBox()

    Call SetCombo(frm1.cboMPSStatus, "FM", "Firm")
    Call SetCombo(frm1.cboMPSStatus, "OP", "Open")
    Call SetCombo(frm1.cboMPSStatus, "PL", "Plan")
    
End Sub

'==========================================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Select			= iCurColumnPos(1)
			C_ItemCode			= iCurColumnPos(2)
			C_ItemName			= iCurColumnPos(3)
			C_ItemSpec			= iCurColumnPos(4)
			C_TrackingNo		= iCurColumnPos(5)
			C_PlndDt			= iCurColumnPos(6)
			C_PlndQty			= iCurColumnPos(7)
			C_Unit				= iCurColumnPos(8)    
			C_MPSConfirmFlg		= iCurColumnPos(9)
			C_MRPConfirmFlg		= iCurColumnPos(10)
			C_MPSStatus			= iCurColumnPos(11)
			C_MPSNo				= iCurColumnPos(12)
			C_ProdEnv			= iCurColumnPos(13)
			C_MaxMrpQty			= iCurColumnPos(14)    
			C_MinMrpQty			= iCurColumnPos(15)
			C_RondQty			= iCurColumnPos(16)
			C_OrderLt			= iCurColumnPos(17)
			C_CumulativeLt		= iCurColumnPos(18)
			C_MPSOrigin			= iCurColumnPos(19)
			C_ItemGroupCd		= iCurColumnPos(20)
			C_ItemGroupNm		= iCurColumnPos(21)
			
    End Select    

End Sub

'-------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
End Function

'-----------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval iWhere)
	Dim iCalledAspName, IntRetCD
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	
	If iWhere = 0 Then
		arrParam(1) = Trim(frm1.txtTrackingNo.value)
		arrParam(2) = Trim(frm1.txtItemCd.value)
	Else
		frm1.vspdData.Col =  C_TrackingNo
		frm1.vspdData.Row = iWhere
		arrParam(1) = Trim(frm1.vspdData.text)

		frm1.vspdData.Col = C_ItemCode
		frm1.vspdData.Row = iWhere
		arrParam(2) = Trim(frm1.vspdData.text)
	End If 
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet, iwhere)
	End If
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(10)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 							'ITEM_CD
	arrField(1) = 2 							'ITEM_NM
	arrField(2) = 4 							'BASIC_UNIT
	arrField(3) = 24							'ORDER_LT
	arrField(4) = 25							'CUMULATIVE_LT
	arrField(5) = 29							'MIN_MRP_QTY
	arrField(6) = 30							'MAX_MRP_QTY
	arrField(7) = 31							'ROND_QTY
	arrField(8) = 33							'MPS_FLAG
	arrField(9) = 21    						'Tracking Flag
	
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet)
	End If	

End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function
'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)

    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)
    frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement    

End Function
'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetTrackingNo(ByRef arrRet, Byval iwhere)

    With frm1
   
		If iWhere = "0" Then
			.txtTrackingNo.Value = arrRet(0)
			.txtTrackingNo.focus
			Set gActiveElement = document.activeElement
		Else
		   	Call .vspdData.SetText(C_TrackingNo,.vspdData.ActiveRow,arrRet(0))
		End If

	End With
	
End Function
'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(ByRef arrRet)
    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With
End Function
'===========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  BatchSelect()  -----------------------------------------
Function btnAutoSel_onClick()

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "전체선택"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "전체선택취소"
	End If

	Dim index,Count
	Dim strFlag
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Select
		
		strFlag = frm1.vspdData.Value

		If lgButtonSelection = "SELECT" Then 
			frm1.vspdData.Value = 1
			frm1.vspdData.Col = 0 
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData.Value = 0
			frm1.vspdData.Col = 0 
			frm1.vspdData.Text=""
		End if

	Next 
	
	frm1.vspdData.ReDraw = true

End Function

'=========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")
	Set gActiveSpdSheet = frm1.vspdData
	gMouseClickStatus = "SPC"	
	
	If Col < 0 Then
		Exit Sub
	End If
	
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
        Exit Sub
    End If

    With frm1.vspdData
	
		.Row = .ActiveRow

		frm1.txtMaxLotQty.Text = GetSpreadText(frm1.vspdData,C_MaxMrpQty,.ActiveRow,"X","X")
		frm1.txtMinLotQty.Text = GetSpreadText(frm1.vspdData,C_MinMrpQty,.ActiveRow,"X","X")
		frm1.txtRondQty.Text = GetSpreadText(frm1.vspdData,C_RondQty,.ActiveRow,"X","X")
		frm1.txtItemLT.Value = GetSpreadText(frm1.vspdData,C_OrderLt,.ActiveRow,"X","X")
		frm1.txtAccumLT.Value = GetSpreadText(frm1.vspdData,C_CumulativeLt,.ActiveRow,"X","X")
	
    End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
   If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)	Then
		If lgStrPrevKey1 <> "" Then
			Call DisableToolBar(parent.TBC_QUERY)
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
	
		If Row <= 0 Then Exit Sub

		If Col = C_Select Then
	
			If Buttondown = 1 Then
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SSDeleteFlag Row,Row
			End If

		End If
	
    End With
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPlndFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlndFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPlndFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPlndToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlndToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtPlndToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtPlndFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtPlndFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtPlndToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtPlndToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    lsDTF = frm1.txtDTF.Text
    lsPTF = frm1.txtDTF.Text
    lsPH   = frm1.txtDTF.Text

    Call ggoOper.ClearField(Document, "2")  
   
    frm1.txtDTF.Text = lsDTF
    frm1.txtPTF.Text = lsPTF
    frm1.txtDTF.Text = lsPH
    
    Call InitVariables		

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtPlndFromDt, frm1.txtPlndToDt)  = False Then		
		Exit Function
	End If

    If DbQuery = False Then
		Exit Function
	End If
       
    FncQuery = True
   
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True

End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()

	Dim ChkFlg
	
	If frm1.vspdData.MaxRows <= 0 Then
		Exit Function
	End If
	
	frm1.vspdData.Col = C_Select
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	If GetSpreadValue(frm1.vspdData,C_Select,frm1.vspdData.ActiveRow,"X","X") = "1" Then
		ChkFlg = True
	End If		
	 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    
    If ChkFlg = True Then
		Call SetSpreadValue(frm1.vspdData,C_Select,frm1.vspdData.ActiveRow,"1","X","X")
	End If		
	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then 
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim lRow
	Dim strStatus, strMRP,strMPS, strDate, strProdEnv
    
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	With frm1.vspdData

	.ReDraw = False

	If .MaxRows > 0 Then
		
		frm1.txtMaxLotQty.Text = GetSpreadText(frm1.vspdData,C_MaxMrpQty,1,"X","X")
		frm1.txtMinLotQty.Text = GetSpreadText(frm1.vspdData,C_MinMrpQty,1,"X","X")
		frm1.txtRondQty.Text = GetSpreadText(frm1.vspdData,C_RondQty,1,"X","X")
		frm1.txtItemLT.Value = GetSpreadText(frm1.vspdData,C_OrderLt,1,"X","X")
		frm1.txtAccumLT.Value = GetSpreadText(frm1.vspdData,C_CumulativeLt,1,"X","X")

	End If

	frm1.vspdData.ReDraw = True
	
    End With

    
    Set gActiveElement = document.ActiveElement   

End Sub
'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001 
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtPlndFromDt=" & Trim(frm1.hPlndFromDt.value)
		strVal = strVal & "&txtPlndToDt=" & Trim(frm1.hPlndToDt.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&cboMPSStatus=" & Trim(frm1.hMPSStatus.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtPlndFromDt=" & Trim(frm1.txtPlndFromDt.Text)
		strVal = strVal & "&txtPlndToDt=" & Trim(frm1.txtPlndToDt.Text)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&cboMPSStatus=" & Trim(frm1.cboMPSStatus.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	End If
	
    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Dim lRow
	Dim strStatus, strMRP,strMPS, strDate, strProdEnv
		
	Call SetToolBar("11001000000111")
    
    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")

	lgQryFlg = True
	
    With frm1.vspdData

	.ReDraw = False

	For lRow = LngMaxRow To .MaxRows

		strStatus = GetSpreadValue(frm1.vspdData,C_MPSStatus,lRow,"X","X")
		strMRP = GetSpreadValue(frm1.vspdData,C_MRPConfirmFlg,lRow,"X","X")
		strMPS = GetSpreadValue(frm1.vspdData,C_MPSConfirmFlg,lRow,"X","X")
		strProdEnv = GetSpreadValue(frm1.vspdData,C_ProdEnv,lRow,"X","X")
		strDate = GetSpreadText(frm1.vspdData,C_PlndDt,lRow,"X","X")

	Next
	
	If .MaxRows > 0 Then

		frm1.txtMaxLotQty.Text = GetSpreadText(frm1.vspdData,C_MaxMrpQty,1,"X","X")
		frm1.txtMinLotQty.Text = GetSpreadText(frm1.vspdData,C_MinMrpQty,1,"X","X")
		frm1.txtRondQty.Text = GetSpreadText(frm1.vspdData,C_RondQty,1,"X","X")
		frm1.txtItemLT.Value = GetSpreadText(frm1.vspdData,C_OrderLt,1,"X","X")
		frm1.txtAccumLT.Value = GetSpreadText(frm1.vspdData,C_CumulativeLt,1,"X","X")

	End If

	.ReDraw = True
	
    End With
	
	frm1.btnAutoSel.disabled = False
	
	frm1.vspdData.Focus
	
End Function
'========================================================================================
' Function Name : DbSave
' Function Desc : 
'========================================================================================
Function DBSave()

    Dim strVal
    Dim IntRows
    Dim ChgVal,strMPS
    Dim iRowSep
    Dim arrVal
    ReDim arrVal(0)
    
    iRowSep = parent.gRowSep
    
    Dim IntRetCD
    
    ggoSpread.Source = frm1.vspdData
    
	DbSave = False
	
    Call LayerShowHide(1)
   
	With frm1.vspdData

	For IntRows = 1 To .MaxRows
		
		strVal = ""
		
		ChgVal = GetSpreadValue(frm1.vspdData,C_Select,IntRows,"X","X")
		strMPS = GetSpreadValue(frm1.vspdData,C_MPSConfirmFlg,IntRows,"X","X")
			
		If GetSpreadText(frm1.vspdData,0,IntRows,"X","X") <> ggoSpread.InsertFlag _
			And UCase(Trim(ChgVal)) = "1" _
			And UCase(Trim(strMPS)) = "Y" Then
				
			strVal = Trim(GetSpreadText(frm1.vspdData,C_MPSNo,IntRows,"X","X")) & iRowSep
		            
		End If
		
		ReDim Preserve arrVal(IntRows)
		arrVal(IntRows) = strVal
		
	Next

	End With

	frm1.txtSpread.value = Join(arrVal,"")
		    
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	
	DbSave = True
	
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()	

	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	Call MainQuery()

End Function
