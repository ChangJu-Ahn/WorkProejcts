'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : 
'*  3. Program ID           : u2211qa1.vbs
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

Const BIZ_PGM_ID	= "u2211qb1.asp"			'☆: List & Manage SCM Orders

Dim C_OrderDt
Dim C_ItemCode
Dim C_ItemName
Dim C_Spec
Dim C_TrackingNo
Dim C_PlantCd
Dim C_PlantNm
Dim C_OrderUnit
Dim C_OrderNo
Dim C_OrderSeq
Dim C_OrderQty
Dim C_DvryDt
Dim	C_RcptQty
Dim	C_UnRcptQty
Dim	C_FirmDvryQty
Dim C_RemainQty
Dim C_DvryPlanDt
Dim C_DvryQty
'Dim C_SLYN
Dim C_SLCD
Dim C_SLPOP
Dim C_SLNM
Dim C_SPLITSEQNO

Dim	C_LotNo
Dim	C_LotSubNo
Dim	C_LotFlg

Dim C_Title
Dim C_DlvyPlanDt
Dim C_DlvyQty
'Dim	C_SerialNo
Dim C_RcptDt
Dim	C_ReceiptQty
Dim	C_RcptRemainQty

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
    
End Sub

'================================================================================================================================
Sub SetDefaultVal()

	Call SetBPCD()
	Call SetToolBar("11000000000111")								'⊙: 버튼 툴바 제어 
End Sub

'================================================================================================================================
Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call DisplayMsgBox("210033","X","X","X")
		Call ggoOper.SetReqAttr(frm1.txtDlvyNo,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDlvyNo2,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDlvyTIME,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTRANSCO,"Q")
		Call ggoOper.SetReqAttr(frm1.txtVEHICLENO,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDRIVER,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTELNO1,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTELNO2,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDlvyPLACE,"Q")
		Call ggoOper.SetReqAttr(frm1.txtREMARK,"Q")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
	    Call SetToolBar("11001011000011")								'⊙: 버튼 툴바 제어 
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
			ggoSpread.Spreadinit "V20050420", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False
	
			.vspdData1.MaxCols = C_SPLITSEQNO + 1
			.vspdData1.MaxRows = 0

			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_OrderNo,		"수주번호", 15
			ggoSpread.SSSetEdit		C_OrderSeq,		"행번"    ,  7
			ggoSpread.SSSetEdit		C_ItemCode,		"품목"    , 18,,,18,2
			ggoSpread.SSSetEdit		C_ItemName,		"품목명"  , 18
			ggoSpread.SSSetEdit		C_Spec,			"규격"    , 15
			ggoSpread.SSSetEdit		C_TRACKINGNO,	"Tracking No."   , 15
			ggoSpread.SSSetEdit		C_OrderUnit,	"단위"    ,  7,,,3,2
			ggoSpread.SSSetDate 	C_DvryPlanDt,	"납품예정일자",12, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_DvryQty,		"납품예정수량",12,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
            ggoSpread.SSSetEdit		C_PlantCd,		"납품처"  , 10
			ggoSpread.SSSetEdit		C_PlantNm,		"납품처명",	12
			ggoSpread.SSSetEdit		C_SPLITSEQNO,	"분할번호",	12
			
			Call ggoSpread.SSSetColHidden( C_SPLITSEQNO, C_SPLITSEQNO , True)
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
   			ggoSpread.SpreadLock	 C_OrderNo,   -1, C_OrderNo
			ggoSpread.SpreadLock	 C_OrderSeq,  -1, C_OrderSeq
			ggoSpread.SpreadLock	 C_ItemCode,  -1, C_ItemCode
			ggoSpread.SpreadLock	 C_ItemName,  -1, C_ItemName
			ggoSpread.SpreadLock	 C_Spec,      -1, C_Spec
			ggoSpread.SpreadLock	 C_OrderUnit, -1, C_OrderUnit
			ggoSpread.SpreadLock	 C_DvryPlanDt,-1, C_DvryPlanDt
			ggoSpread.SpreadLock	 C_DvryQty,   -1, C_DvryQty
			ggoSpread.SpreadLock	 C_PlantCd,   -1, C_PlantCd
			ggoSpread.SpreadLock	 C_PlantNm,   -1, C_PlantNm
			ggoSpread.SpreadLock	 C_SPLITSEQNO,-1, C_SPLITSEQNO
			
			.vspdData1.ReDraw = True
	
		End With
	End If

End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
    With frm1.vspdData1 
    
    .Redraw = False

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSSetProtected C_OrderNo,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderSeq,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemCode,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemName,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Spec,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderUnit,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DvryPlanDt,	pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_DvryQty,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantCd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlantNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SPLITSEQNO,	pvStartRow, pvEndRow
    
    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True
    
    .Redraw = True
    
    End With
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()	
	
		' Grid 1(vspdData1)
		C_OrderNo		= 1
		C_OrderSeq		= 2
		C_ItemCode		= 3
		C_ItemName		= 4
		C_Spec			= 5
		C_TrackingNo	= 6
		C_OrderUnit		= 7
		C_DvryPlanDt	= 8
		C_DvryQty		= 9
		C_PlantCd		= 10
		C_PlantNm		= 11
		C_SPLITSEQNO	= 12
	
End Sub
 
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case Ucase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_OrderNo		= iCurColumnPos(1)
			C_OrderSeq		= iCurColumnPos(2)
			C_ItemCode		= iCurColumnPos(3)
			C_ItemName		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_TrackingNo	= iCurColumnPos(6)
			C_OrderUnit		= iCurColumnPos(7)
			C_DvryPlanDt	= iCurColumnPos(8)
			C_DvryQty		= iCurColumnPos(9)
			C_PlantCd		= iCurColumnPos(10)
			C_PlantNm		= iCurColumnPos(11)
			C_SPLITSEQNO	= iCurColumnPos(12)
					
 	End Select
 
End Sub

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenDlvyNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenDlvyNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
			
	If IsOpenPop = True Or UCase(frm1.txtDlvyNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(Trim(frm1.txtbpcd.Value), "", "S"), lgF0) = False Then
		Call DisplayMsgBox("971012", "X", "공급처", "X")
		frm1.txtbpNM.VALUE = ""
		frm1.txtbpcd.focus
    	Exit Function
    Else
		lgF0 = replace(lgF0,chr(12),"")
		frm1.txtbpnm.value = replace(lgF0,chr(11),"")
	End If
	

	if Trim(frm1.txtbpcd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "업체", "X")
		frm1.txtbpcd.focus
    	Exit Function
    End if
				
	arrParam(0) = Trim(frm1.txtBpCd.value)

	iCalledAspName = AskPRAspName("U2122PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "U2122PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDlvyNo.value = strRet(0)
		frm1.txtDlvyNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
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
				
				Case C_SLPOP
					.Col = Col - 1
			    	.Row = Row
					
					Call OpenSLCD(.text)
					
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

'================================================================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtPoFrDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
		Call SetFocusToDocument("M") 
		frm1.txtPoToDt.focus
	End If
End Sub

'================================================================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'================================================================================================================================
Sub txtPoToDt_KeyDown(KeyCode, Shift)
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
	var11 = Trim(frm1.txtdlvyno2.VALUE)
		
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
	strUrl = strUrl & "|DLVY_NO|" & var11
	
	'strEbrFile = "U2118OA1"
	strEbrFile = "U1121MA1"
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
	var11 = Trim(frm1.txtdlvyno2.VALUE)
			
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
	strUrl = strUrl & "|DLVY_NO|" & var11	

	'strEbrFile = "U2118OA1"
	strEbrFile = "U1121MA1"
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

    frm1.txtdlvyno.focus
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
    Call initData(frm1.vspdData1.ActiveRow)
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
	.vspdData1.focus
	Set gActiveElement = document.activeElement 
	ggoSpread.Source = .vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
    SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1
	.vspdData1.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement
	If Err.number = 0 Then FncInsertRow = True
End Function


'================================================================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    
    If frm1.vspdData1.MaxRows < 1 Then Exit Function

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

	Dim strPoNo
	Dim strPoSeqNo

   Select Case pOpt
       Case "M"
			With frm1
				lgKeyStream =               Trim(.txtBpCd.VALUE)  & Parent.gColSep
				lgKeyStream = lgKeyStream & Trim(.txtDlvyNo.VALUE)  & Parent.gColSep
					
			End With

       Case "S"
			With frm1
				.vspdData1.Row = .vspdData1.ActiveRow
				.vspdData1.Col = C_OrderNo
				strPoNo = .vspdData1.text
				.vspdData1.Col = C_OrderSeq
				strPoSeqNo = .vspdData1.text
					
				lgKeyStream = lgKeyStream & UCase(Trim(strPoNo))  & Parent.gColSep
				lgKeyStream = lgKeyStream & UCase(Trim(strPoSeqNo))  & Parent.gColSep

			End With

	End Select
	   
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
    ggoOper.SetReqAttr	frm1.txtdlvyno2, "Q"
    Call SetToolBar("11100000000111")

    
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call RestoreToolBar()
	End If

    Call initdata()
	lgIntFlgMode = parent.OPMD_UMODE
	
End Function

'================================================================================================================================
Function DbQueryNotOk()	

'	Call SetToolBar("11000000000011")
'    '-----------------------
'    'Reset variables area
'    '-----------------------
'    lgIntFlgMode = parent.OPMD_CMODE

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
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim strDlvyNo
	
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
	exit Function
	end if
	
    strVal = ""
    strDel = ""
    lGrpCnt = 1

	if lgBlnFlgChgValue = True then
		'--형균 채번로직 
		Call CommonQueryRs( " ISNULL(DPNO,'DP' + CONVERT(VARCHAR(8),GETDATE(),112) + '001' ) " , " (SELECT (CASE WHEN SUBSTRING(MAX(DLVY_NO),3,8) <> CONVERT(VARCHAR(8),GETDATE(),112)   THEN 'DP' + CONVERT(VARCHAR(8),GETDATE(),112) + '0001' ELSE 'DP' + CONVERT(VARCHAR(8),GETDATE(),112) + RIGHT('000' + CONVERT(VARCHAR(3),CONVERT(NUMERIC,RIGHT(MAX(DLVY_NO),3),112) + 1,112),3) END) DPNO  FROM M_SCM_DLVY_PUR_RCPT WHERE LEFT(DLVY_NO,2) = 'DP')A ", " 1=1 ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Frm1.txtdlvyno2.value = "" Then
			Frm1.txtdlvyno2.value = Replace(lgF0,Chr(11),"")
			Frm1.txtdlvyno.value = Replace(lgF0,Chr(11),"")
		End If
		'---
	end iF
	
	With Frm1
    
       For lRow = 1 To .vspdData1.MaxRows
    
           .vspdData1.Row = lRow
           .vspdData1.Col = 0
        
           Select Case .vspdData1.Text

               Case  ggoSpread.InsertFlag                                      '☜: Update
               
														  strVal = strVal & "U"  &  parent.gColSep					
														  strVal = strVal & lRow &  parent.gColSep
						.vspdData1.Col = C_OrderNo	    : strVal = strVal & Trim(.vspdData1.value) & parent.gColSep	'2
						.vspdData1.Col = C_OrderSeq     : strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'3
						.vspdData1.Col = C_ItemCode		: strVal = strVal & Trim(.vspdData1.value) & parent.gColSep	'4
						.vspdData1.Col = C_ItemName		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'5
						.vspdData1.Col = C_Spec			: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'6
						.vspdData1.Col = C_OrderUnit	: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'7
						.vspdData1.Col = C_DvryPlanDt	: strVal = strVal & Trim(.vspdData1.text)  & parent.gColSep	'8
						.vspdData1.Col = C_DvryQty		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'9
						.vspdData1.Col = C_PlantCd		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'10
						.vspdData1.Col = C_PlantNm		: strVal = strVal & Trim(.vspdData1.Value) & parent.gColSep	'11
						.vspdData1.Col = C_SPLITSEQNO	: strVal = strVal & Trim(.vspdData1.value) & parent.gRowSep	'12
                   
						lGrpCnt = lGrpCnt + 1
				
				Case  ggoSpread.DeleteFlag                                      '☜: Delete
               
														  strDel = strDel & "D"  &  parent.gColSep					
														  strDel = strDel & lRow &  parent.gColSep
						.vspdData1.Col = C_OrderNo	    : strDel = strDel & Trim(.vspdData1.value) & parent.gColSep	'2
						.vspdData1.Col = C_OrderSeq     : strDel = strDel & Trim(.vspdData1.Value) & parent.gColSep	'3
						.vspdData1.Col = C_SPLITSEQNO	: strDel = strDel & Trim(.vspdData1.value) & parent.gRowSep	'4
                   
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
	
	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Function
	End If
	
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

	iCalledAspName = AskPRAspName("U2125RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "U2125RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If isEmpty(strRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.	
	
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