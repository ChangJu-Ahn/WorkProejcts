Const BIZ_PGM_QRY_ID    = "i1231mb1_ko441.asp"							
Const BIZ_PGM_DQRY_ID   = "i1231mb1_ko441.asp"							
Const BIZ_PGM_SAVE_ID   = "i1231mb2_ko441.asp"							
Const BIZ_PGM_DEL_ID    = "i1231mb3.asp"							
Const BIZ_PGM_LOOKUP_ID	= "i1231mb4.asp"							

Dim C_ItemCd 
Dim C_ItemPopup
Dim C_ItemNm 
Dim C_EntryQty 
Dim C_EntryUnit
Dim C_EntryUnitPopup
Dim C_ItemSpec 
Dim C_InvUnit 
Dim C_TrackingNo 
Dim C_TrackingNoPopUp 
Dim C_LotNo 
Dim C_LotSubNo 
Dim C_LotNoPopup 
Dim C_ProdOrdNo 
Dim C_ProdOrdNoPopUp
Dim C_ReqNo 
Dim C_SeqNo 
Dim C_SubSeqNo 
    

'******************************* Sub InitVariables() *******************************************
Sub InitVariables()
    lgIntFlgMode		= parent.OPMD_CMODE         
    lgBlnFlgChgValue	= False                 
    lgIntGrpCount		= 0                        
    lgStrPrevKey		= ""                        
    lgLngCurRows		= 0                         
    lgSortKey			= 1
    lgMovType			= ""
End Sub


'******************************* Sub InitSpreadPosVariables() ***********************************
Sub InitSpreadPosVariables()
	C_ItemCd			= 1									
	C_ItemPopup			= 2
	C_ItemNm			= 3
	C_EntryQty			= 4
	C_EntryUnit			= 5
	C_EntryUnitPopup	= 6
	C_ItemSpec			= 7
	C_InvUnit			= 8
	C_TrackingNo		= 9
	C_TrackingNoPopUp	= 10
	C_LotNo				= 11
	C_LotSubNo			= 12
	C_LotNoPopup		= 13
	C_ProdOrdNo			= 14
	C_ProdOrdNoPopUp	= 15
	C_ReqNo				= 16
	C_SeqNo				= 17
	C_SubSeqNo			= 18
End Sub


'******************************** Sub SetDefaultVal() *****************************************
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtDocumentDt.text = StartDate
	frm1.txtPostingDt.text  = StartDate
	Call ExtractDateFrom(Currentdate, Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.txtYear.Year = strYear
	frm1.txtDocumentNo1.focus
	lgBlnFlgChgValue = False	
	
	if frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtDocumentNo1.focus 
	Else
		frm1.txtPlantCd.focus 
	End If 
	Set gActiveElement = document.activeElement
		 
End Sub

'***************************** Sub InitSpreadSheet() ****************************************
Sub InitSpreadSheet()
        
    Call InitSpreadPosVariables()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030423", ,parent.gAllowDragDropSpread
        
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_SubSeqNo + 1	
		.MaxRows = 0
	    
	    Call GetSpreadColumnPos("A")
	    Call AppendNumberPlace("6", "3", "0")
		
		ggoSpread.SSSetEdit			C_ItemCd,				"품목",18, 0, -1, 18, 2		
		ggoSpread.SSSetButton 		C_ItemPopup	
		ggoSpread.MakePairsColumn	C_ItemCd,C_ItemPopup
			
		ggoSpread.SSSetEdit			C_ItemNm,				"품목명", 20, 0, -1, 50	
		ggoSpread.SSSetFloat		C_EntryQty,				"출고수량", 15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , , "Z"
		ggoSpread.SSSetEdit			C_EntryUnit,			"출고단위", 10, 0, -1, 3, 2
		ggoSpread.SSSetButton		C_EntryUnitPopup
		ggoSpread.MakePairsColumn	C_EntryUnit,C_EntryUnitPopup
		
		ggoSpread.SSSetEdit			C_ItemSpec,				"규격", 20, 0, -1, 50		
		ggoSpread.SSSetEdit			C_InvUnit,				"재고단위", 10, 0, -1, 3
		ggoSpread.SSSetEdit			C_TrackingNo,			"Tracking No",20, 0, -1, 25, 2
		ggoSpread.SSSetButton		C_TrackingNoPopup
		ggoSpread.MakePairsColumn	C_TrackingNo,C_TrackingNoPopup
		
		ggoSpread.SSSetEdit 		C_LotNo,				"LOT NO", 20, 0, -1, 25, 2
		ggoSpread.SSSetFloat		C_LotSubNo,				"순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetButton 		C_LotNoPopup
		ggoSpread.MakePairsColumn	C_LotNo,C_LotNoPopup

		ggoSpread.SSSetEdit			C_ProdOrdNo,			"제조오더번호", 18, 0, -1, 18, 2
		ggoSpread.SSSetButton		C_ProdOrdNoPopUp
		ggoSpread.MakePairsColumn	C_ProdOrdNo,C_ProdOrdNoPopUp

		ggoSpread.SSSetEdit			C_ReqNo,	"", 1, 0
		ggoSpread.SSSetEdit			C_SeqNo,    "PK1", 1, 0
		ggoSpread.SSSetEdit			C_SubSeqNo, "PK2", 1, 0
		Call ggoSpread.SSSetColHidden(C_ReqNo,    .MaxCols,     True)
		ggoSpread.SpreadLock -1, -1
		
		.ReDraw = true
		
		ggoSpread.SSSetSplit2(3)
	End With
    
End Sub

'********************************** Sub GetSpreadColumnPos(ByVal pvSpdNo) *******************************
Sub GetSpreadColumnPos(ByVal pvSpdNo)

	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
		      
		    ggoSpread.Source = frm1.vspdData
		      
		    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		      
            C_ItemCd          = iCurColumnPos(1)									
			C_ItemPopup       = iCurColumnPos(2)
			C_ItemNm          = iCurColumnPos(3)
			C_EntryQty        = iCurColumnPos(4)
			C_EntryUnit       = iCurColumnPos(5)
			C_EntryUnitPopup  = iCurColumnPos(6)
			C_ItemSpec        = iCurColumnPos(7)
			C_InvUnit         = iCurColumnPos(8)
			C_TrackingNo      = iCurColumnPos(9)
			C_TrackingNoPopUp = iCurColumnPos(10)
			C_LotNo           = iCurColumnPos(11)
			C_LotSubNo        = iCurColumnPos(12)
			C_LotNoPopup      = iCurColumnPos(13)
			C_ProdOrdNo       = iCurColumnPos(14)
			C_ProdOrdNoPopUp  = iCurColumnPos(15)
			C_ReqNo			  = iCurColumnPos(16)
			C_SeqNo           = iCurColumnPos(17)
			C_SubSeqNo        = iCurColumnPos(18)
	End Select

End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()	
End Sub


Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet
	Call ggoSpread.ReOrderingSpreadData
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.SSSetRequired		C_ItemCd,   pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_ItemNm,   pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_ItemSpec, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_InvUnit,  pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_EntryQty, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_EntryUnit,pvStartRow, pvEndRow
End Sub


'******************************** Function SheetFocus(lRow, lCol) ******************************************
Function SheetFocus(lRow, lCol)
	Call changeTabs(2)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")					
	ElseIf lgIntFlgMode = Parent.OPMD_UMODE Then
	    Call SetToolBar("11101011000111")					
	End If
	
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

'================================ Sub HideTextBox() =================================================
Sub HideTextBox()

	 txtWCCdTitle.style.display = "none"
	 frm1.txtWCCd.style.display = "none"
	 frm1.txtWCNm.style.display = "none"
	 frm1.btnWC.style.display  = "none"
	 
	 txtCostCenterTitle.style.display = "none"
	 frm1.txtCostCenter.style.display = "none"
	 frm1.txtCostCenterNm.style.display = "none"
	 frm1.btnCostCenter.style.display = "none" 
	
End Sub

'20080226::hanc***********************************************************
Function OpenMoveInvRef1()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)
	Dim strTempGlNo
	Dim strGlNo
	Dim strRefNo
	Dim strFrom

				
	Dim Param1, Param2, Param3, Param4, Param5, Param6 

	If IsOpenPop = True Then Exit Function
	
	


	If Trim(frm1.txtPlantCd.value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd.Focus
	    Set gActiveElement = document.activeElement
		Exit Function
	End If

	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtPlantNm.value)
	
'	If Trim(frm1.txtMovType.value) = "" then
'		
'		Call ClickTab1()
'		Msgbox "수불유형을 선택하십시오.",vbInformation, parent.gLogoName
''		MsgBox "수불유형을 선택하십시오."
'		frm1.txtMovType.Focus
'			'frm1.txtSLCd2.Focus
'			'Set gActiveElement = document.activeElement
'		Exit Function
'	End If

	Param3 = "OI"
	Param4 = Trim(frm1.txtMovType.value)
	
	
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I1311XA1_KO441")     '20080226::HANC
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1311RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    

	IsOpenPop = False
	
	If arrRet(0,0) = "" Then
		frm1.txtSLCd2.focus
		Exit Function
	Else
		Call SetMoveInvRef1(arrRet)
	End If
	
End Function
'20080226::hanc***********************************************************
Function SetMoveInvRef1(arrRet)
	Dim TempRow
	Dim intLoopCnt
	Dim intCnt
	Dim iRow
		
    frm1.txtMovType.value  =         arrRet(0, 7)   '20080312::HANC

	Call ClickTab2()
	
	With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData   
		.ReDraw = False 

		iLngStartRow = .MaxRows + 1            '☜: 현재까지의 MaxRows 
		iLngLoopCnt = Ubound(arrRet, 1)           '☜: Reference Popup에서 선택되어진 Row만큼 추가 

			TempRow = .MaxRows
			.MaxRows = TempRow + intLoopCnt ' + 1
			ggoSpread.SpreadUnLock	 C_ItemCd,		TempRow + 1,	C_ItemPopup,		.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryQty,	TempRow + 1,	C_EntryQty,			.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryUnit,	TempRow + 1,	C_EntryUnitPopup,	.MaxRows
			ggoSpread.SpreadUnLock	 C_TrackingNo,	TempRow + 1,	C_TrnsLotSubNo,		.MaxRows
			ggoSpread.SSSetRequired  C_ItemCd,		TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryQty,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryUnit,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemCd,  TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemPopup, TempRow + 1,	.MaxRows
			

		For iLngCnt = 0 to iLngLoopCnt - 1
			.MaxRows = .MaxRows + 1
			.Row = .MaxRows

			.Col = 0			:		.Text = ggoSpread.InsertFlag
			.Col = C_ItemCd     :		.text = arrRet(iLngCnt, 0)	
			.Col = C_ItemNm     :		.text = arrRet(iLngCnt, 1)	
			.Col = C_TrackingNo :		.text = arrRet(iLngCnt, 2)	
			.Col = C_EntryQty   :		.text = arrRet(iLngCnt, 3)	
			.Col = C_EntryUnit  :		.text = arrRet(iLngCnt, 4)	
			.Col = C_LotSubNo   :		.text = "0"			        
			.Col = C_LotNo      :		.text = "*"			        
			.Col = C_SeqNo      :		.text = arrRet(iLngCnt, 5)	        '20080227::HANC
			frm1.txtDocumentText.value  =       arrRet(iLngCnt, 6)	        '20080227::HANC
			frm1.txtMovType.value  =         arrRet(iLngCnt, 7)

'				Call .SetText(C_ItemCd,	iRow, arrRet(intCnt, 0))
'				Call .SetText(C_ItemNm, iRow, arrRet(intCnt, 1))
'				Call .SetText(C_TrackingNo,iRow, arrRet(intCnt, 2))
'				Call .SetText(C_EntryQty,  iRow, arrRet(intCnt, 3))
'				Call .SetText(C_EntryUnit, iRow, arrRet(intCnt, 4))
'				Call .SetText(C_LotSubNo,  iRow, "0")
'				Call .SetText(C_LotNo,     iRow, "*")
'				Call .SetText(C_TrnsLotNo, iRow, "*")

            
		Next

		Call SetSpreadColor(iLngStartRow, .MaxRows)

		.ReDraw = True    

	End With


'	With frm1.vspdData
'			.focus
'			ggoSpread.Source = frm1.vspdData			
'			
'			.ReDraw = False	
'			intLoopCnt = Ubound(arrRet, 1)
'			
'			TempRow = .MaxRows
'			.MaxRows = TempRow + intLoopCnt ' + 1
'			ggoSpread.SpreadUnLock	 C_ItemCd,		TempRow + 1,	C_ItemPopup,		.MaxRows
'			ggoSpread.SpreadUnLock	 C_EntryQty,	TempRow + 1,	C_EntryQty,			.MaxRows
'			ggoSpread.SpreadUnLock	 C_EntryUnit,	TempRow + 1,	C_EntryUnitPopup,	.MaxRows
'			ggoSpread.SpreadUnLock	 C_TrackingNo,	TempRow + 1,	C_TrnsLotSubNo,		.MaxRows
'			ggoSpread.SSSetRequired  C_ItemCd,		TempRow + 1,	.MaxRows
'			ggoSpread.SSSetRequired	 C_EntryQty,	TempRow + 1,	.MaxRows
'			ggoSpread.SSSetRequired	 C_EntryUnit,	TempRow + 1,	.MaxRows
'			ggoSpread.SSSetProtected C_TrnsItemCd,  TempRow + 1,	.MaxRows
'			ggoSpread.SSSetProtected C_TrnsItemPopup, TempRow + 1,	.MaxRows
'			
'			For intCnt = 0 to intLoopCnt
'				
'				iRow = TempRow + intCnt + 1
' 				Call .SetText(0, iRow, ggoSpread.InsertFlag)
'				Call .SetText(C_ItemCd,	iRow, arrRet(intCnt, 0))
'				Call .SetText(C_ItemNm, iRow, arrRet(intCnt, 1))
'				Call .SetText(C_TrackingNo,iRow, arrRet(intCnt, 2))
'				Call .SetText(C_EntryQty,  iRow, arrRet(intCnt, 3))
'				Call .SetText(C_EntryUnit, iRow, arrRet(intCnt, 4))
'				Call .SetText(C_LotSubNo,  iRow, "0")
'				Call .SetText(C_LotNo,     iRow, "*")
'				Call .SetText(C_TrnsLotNo, iRow, "*")
'					
'			Next							
'			.ReDraw = True
'
'		End With

End Function

'================================= ClickTab() ==============================================
Function ClickTab1()

    Dim strVal
    Call changeTabs(1)	
	If lgIntFlgMode = parent.OPMD_CMODE Then
	    Call SetToolBar("11101000000011")								    
	ElseIf lgIntFlgMode = parent.OPMD_UMODE Then
	    Call SetToolBar("11101000000111")							    
	End If
	If frm1.txtPlantCd.value <> "" Then
		frm1.txtDocumentNo1.focus
	Else
		frm1.txtPlantCd.focus
	End If
	 
End Function


Function ClickTab2()
		 
	Dim strVal
			
    Call changeTabs(2)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")							
	ElseIf lgIntFlgMode = Parent.OPMD_UMODE Then
	    Call SetToolBar("11101011000111")							
	End If	
	
	If frm1.txtMovType.value <> lgMovType Then
		ggoSpread.Source = frm1.vspdData   
		ggoSpread.ClearSpreadData
		lgMovType = frm1.txtMovType.value	
	End if	
End Function

'=================================== Function OpenPlant() ===========================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

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
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function



'====================================== OpenDocumentNo =======================================
Function OpenDocumentNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1, Param2, Param3, Param4
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")   
		frm1.txtPlantCd.Focus
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
    
	IsOpenPop = True

	Param1 = Trim(frm1.txtDocumentNo1.Value)
	Param2 = Trim(frm1.txtYear.Text)
	Param3 = "OI"
	Param4 = Trim(frm1.txtPlantCd.Value)
	
	iCalledAspName = AskPRAspName("I1111PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1111PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4), _
		"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocumentNo1.focus
		Exit Function
	Else
		Call SetDocumentNo(arrRet)
	End If	
	Set gActiveElement = document.activeElement
End Function



'================================ OpenSL() ===================================================
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd.Focus
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtSLcd.ClassName) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD=" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
	Set gActiveElement = document.activeElement   
End Function



'=================================== OpenWC() ================================================
Function OpenWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtWCCd.ClassName) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = Trim(frm1.txtWCCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "작업장"			
	
	arrField(0) = "WC_CD"	
	arrField(1) = "WC_NM"	
	
	arrHeader(0) = "작업장"		
	arrHeader(1) = "작업장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtWCCd.focus
		Exit Function
	Else
		Call SetWC(arrRet)
	End If	
	
End Function



'================================= OpenProdOrderNo() ==============================================
Function OpenProdOrderNo()
	Dim iCalledAspName
	Dim IntRetCD
	Dim strItem

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("169901","X", "X", "X")   
		IsOpenPop = False
		Call changeTabs(1) 
		frm1.txtPlantCd.Focus
		Exit Function
	End If
	
	If UCase(frm1.txtMovType.value) = "I97" Then
		If frm1.txtWCCd.value = "" Then
			Call DisplayMsgBox("221805","X", "X", "X")   
			IsOpenPop = False
			Call changeTabs(1) 
			frm1.txtWCCd.Focus
			Exit Function
		End If
	
		frm1.vspdData.Col = C_ItemCd
		strItem   = frm1.vspdData.Text

		If strItem = "" Then
			Call DisplayMsgBox("169915","X", "X", "X")
			IsOpenPop = False
			frm1.vspdData.Col = C_ItemCd
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Action = 0
			Exit Function
		End If
	End If
	
	Dim arrRet
	Dim arrParam(9)
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	frm1.vspdData.row = frm1.vspdData.ActiveRow
	frm1.vspdData.col = C_ProdOrdNo
	arrParam(5) = frm1.vspdData.Text
	arrParam(6) = ""
	arrParam(7) = ""
	
	If UCase(frm1.txtMovType.value) = "I97" Then
		arrParam(8) = strItem
		arrParam(9) = frm1.txtWCCd.value
		
		iCalledAspName = AskPRAspName("P4211PA1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"P4211PA1","x")
			IsOpenPop = False
			Exit Function
		End If
	Else
		iCalledAspName = AskPRAspName("P4111PA1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"P4111PA1","x")
			IsOpenPop = False
			Exit Function
		End If
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ProdOrdNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		If UCase(frm1.txtMovType.value) = "I97" Then
			Call SetProdOrderNo2(arrRet)
		Else
			Call SetProdOrderNo(arrRet)
		End If
	End If
	Set gActiveElement = document.activeElement	
End Function



'================================= OpenMovType() ==============================================
Function OpenMovType()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1, Param2, Param3
	
	If IsOpenPop = True Then Exit Function

	If UCase(frm1.txtMovType.ClassName) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True

	Param1 = Trim(frm1.txtMovType.Value)
	Param2 = Trim(frm1.txtMovTypeNm.Value)
	Param3 = "OI"

	iCalledAspName = AskPRAspName("I1411PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1411PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMovType.focus
		Exit Function
	Else
		Call SetMovType(arrRet)
	End If	
	
End Function



'==================================== OpenItem() ===========================================
Function OpenItem(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5),arrField(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901", "X", "X", "X")
		Call ClickTab1()
		Call SetFocusToDocument("M")  
        frm1.txtPlantCd.Focus
		Exit Function
	End If
	
	If Trim(frm1.txtSlCd.Value) = "" then 
		Call DisplayMsgBox("169902", "X", "X", "X")
		Call ClickTab1()
		Call SetFocusToDocument("M")  
        frm1.txtSlCd.Focus
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = Trim(frm1.txtSLCd.value)
	arrParam(3) = Trim(frm1.txtSLNm.value)
	arrParam(4) = strCode
	arrParam(5) = ""	
	
	arrField(0) = 1 'ITEM_CD					' Field명(0)
	arrField(1) = 2 'ITEM_NM					' Field명(1)
	arrField(2) = 3	'SPECIFICATION	
	arrField(3) = 4
	arrField(4) = 5
	arrField(5)	= 6
	arrField(6) = 7
	
	iCalledAspName = AskPRAspName("I1211PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1211PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetItem(arrRet)
	End If	
	Set gActiveElement = document.activeElement   
End Function



'=============================== OpenTrackingNo() ================================================
Function OpenTrackingNo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "TRACKINGNO"	
	arrParam(1) = "s_so_tracking"				
	
	frm1.vspdData.Col = C_TrackingNo
	arrParam(2)   = frm1.vspdData.Text		
	
	arrParam(3) = ""
	
	arrParam(4) = ""			
	arrParam(5) = "Tracking No"			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking_No"		
    arrHeader(1) = "품목"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_TrackingNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If	
End Function



'================================ OpenEntryUnit() =========================================
Function OpenEntryUnit(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위 팝업"						
	arrParam(1) = "B_UNIT_OF_MEASURE"						
	arrParam(2) = strCode						
	arrParam(3) = ""			 					
	arrParam(4) = ""					
	arrParam(5) = "단위"						
	
	arrField(0) = "UNIT"	
	arrField(1) = "UNIT_NM"	
	
	arrHeader(0) = "단위"		
	arrHeader(1) = "단위명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_EntryUnit,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetEntryUnit(arrRet)
	End If	
End Function



'======================================= OpenLotNo() ============================================
Function OpenLotNo(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	Dim Param4
	Dim Param5
	Dim Param6
	Dim Param7
	Dim Param8
	Dim Param9
	
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtSLCd.value)
	Param4 = Trim(frm1.txtPlantCd.value)	
	Param5 = "J"				  
	if Param1 = "" then
		Call DisplayMsgBox("169902", "X", "X", "X")   
		frm1.txtSLCd.Focus
		Set gActiveElement = document.activeElement   
		Exit Function
	End If
	ggoSpread.Source = frm1.vspdData    

	With frm1.vspdData	    
		If .MaxRows = 0 Then
			Exit Function
		else
			.Col = C_ItemCd
			.Row = .ActiveRow
			 Param2 = Trim(.Text )
			.Col = C_TrackingNo
			.Row = .ActiveRow
			 Param3 = Trim(.Text )
			.Col = C_LotNo
			.Row = .ActiveRow 
			 Param6 = Trim(.Text)
			.Col = C_EntryUnit
			.Row = .ActiveRow
			 Param9 = Trim(.Text)
		End If	
    	
		if Param2 = "" then
			Call DisplayMsgBox("169903", "X" ,"X", "X")   
			Call SetActiveCell(frm1.vspdData, C_ItemCd, .ActiveRow,"M","X","X")
			Exit Function
		End If
    End With
    
	iCalledAspName = AskPRAspName("I2212RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2212RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_LotNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
    	Call SetLotNo(arrRet)
	End If	
End Function



'====================================== OpenCostCd() ============================================
Function OpenCostCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtCostCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd.Focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "Cost Center 팝업"			
	arrParam(1) = "B_COST_CENTER A,B_PLANT B"
	arrParam(2) = Trim(frm1.txtCostCd.Value)		
	arrParam(3) = ""								
	arrParam(4) = "A.BIZ_AREA_CD = B.BIZ_AREA_CD AND B.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "Cost Center"					
	
	arrField(0) = "COST_CD"							
	arrField(1) = "COST_NM"							
    
	arrHeader(0) = "Cost Center"			    	
	arrHeader(1) = "Cost Center 명"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		Call SetCostCd(arrRet)
	End If	
    
End Function



'=================================== OpenPopupGL() ===============================================
Function OpenPopupGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)
	Dim strTempGlNo
	Dim strGlNo
	Dim strRefNo
	Dim strFrom
   
	If IsOpenPop = True Then Exit Function    
	
	if Trim(frm1.txtDocumentNo1.Value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169804","X", "X", "X")   
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement   
       Exit Function
    End If
	
	strRefNo	=	Trim(frm1.txtDocumentNo1.value) & "-" & Trim(frm1.txtYear.Year)  
	strFrom		=	"ufn_a_GetGlNo( " & FilterVar((strRefNo), "''" , "S") & " )"   
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", StrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
		If lgF0 <> "" then 
			strTempGlNo = Split(lgF0, Chr(11))
			strGlNo		= Split(lgF1, Chr(11))
		Else
			call DisplayMsgBox("205154","X","X","X")
			Exit Function
		End if 
		
		arrParam(0) = strGlNo(0)
		arrParam(1) = ""
	
	iCalledAspName = AskPRAspName("A5120RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5120RA1","x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
   
End Function


'===================================== OpenPopupGL2() ==============================================
Function OpenPopupGL2()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)	
	Dim strTempGlNo
	Dim strGlNo
	Dim strRefNo
	Dim StrFrom

	If IsOpenPop = True Then Exit Function    
	
	if Trim(frm1.txtDocumentNo1.Value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169804","X", "X", "X")   
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement   
       Exit Function
    End If
	
	strRefNo	=	Trim(frm1.txtDocumentNo1.Value) & "-" & Trim(frm1.txtYear.Year)
	StrFrom		=  " ufn_a_GetGlNo( " & FilterVar((strRefNo), "''" , "S") & " )"   
	
	Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", StrFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
		If lgF0 <> "" then 
			strTempGlNo = Split(lgF0, Chr(11))
			strGlNo		= Split(lgF1, Chr(11))
		else
			call DisplayMsgBox("205154","X","X","X")
			Exit Function
		end if
		
	arrParam(0) = strTempGlNo(0)
	arrParam(1) = ""

	iCalledAspName = AskPRAspName("A5130RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5130RA1","x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
   
End Function
 

'================================== Set() ===================================================
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus	
End Function


Function SetDocumentNo(byRef arrRet)
	frm1.txtDocumentNo1.Value    = arrRet(0)
	frm1.txtYear.Year            = arrRet(1)
	frm1.txtDocumentNo1.focus	
End Function


Function SetSL(byRef arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue = True
End Function


Function SetWC(byRef arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)
	frm1.txtWCCd.focus		
	lgBlnFlgChgValue = True
End Function


Function SetCostCd(byRef arrRet)
	frm1.txtCostCd.value = arrRet(0)
	frm1.txtCostNm.value = arrRet(1)
	frm1.txtCostCd.focus
	lgBlnFlgChgValue = True
End Function


Function SetMovType(byRef arrRet)
	frm1.txtMovType.Value    = arrRet(0)
	frm1.txtMovTypeNm.Value  = arrRet(1)
	Call WcCdChange()
	frm1.txtMovType.focus
	lgBlnFlgChgValue = True
End Function


Function SetItem(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_ItemCd, .ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm, .ActiveRow, arrRet(1))
		Call .SetText(C_ItemSpec, .ActiveRow, arrRet(2))
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(3))
		Call .SetText(C_InvUnit, .ActiveRow, arrRet(3))
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(4))
		Call .SetText(C_LotNo, .ActiveRow, arrRet(5))
		Call .SetText(C_LotSubNo, .ActiveRow, arrRet(6))
		Call vspdData_Change(C_ItemCd, .ActiveRow)		
		Call SetActiveCell(frm1.vspdData, C_EntryQty, .ActiveRow,"M","X","X")
	End With
End Function


Function SetTrackingNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_TrackingNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_LotNo,.ActiveRow,"M","X","X")		
	End With
End Function


Function SetEntryUnit(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_EntryUnit, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_TrackingNo,.ActiveRow,"M","X","X")				
	End With
End Function


Function SetLotNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(2))
		Call .SetText(C_LotNo, .ActiveRow, arrRet(3))
		Call .SetText(C_LotSubNo, .ActiveRow, arrRet(4))
		Call vspdData_Change(C_LotNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_ProdOrdNo,.ActiveRow,"M","X","X")		
	End With
End Function


Function SetProdOrderNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_ProdOrdNo, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_ProdOrdNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_ProdOrdNo,.ActiveRow,"M","X","X")		
	End With
End Function

Function SetProdOrderNo2(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_ProdOrdNo, .ActiveRow, arrRet(0))
		Call .SetText(C_ReqNo, .ActiveRow, arrRet(1))
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(2))
		.Col = C_EntryQty
		.Row = .ActiveRow
		If .Text = "" Then
			Call .SetText(C_EntryQty, .ActiveRow, arrRet(3))
		End If
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(4))
		Call vspdData_Change(C_ProdOrdNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_ProdOrdNo,.ActiveRow,"M","X","X")		
	End With
End Function


'======================================= SSColHiddenChange() ===========================================
Function SSColHiddenChange
	With frm1
		If UCase(.txtMovType.Value) = "R01" then
			Call ggoSpread.SSSetColHidden(C_LotNo, C_LotNoPopup, True)
		Else
			Call ggoSpread.SSSetColHidden(C_LotNo, C_LotNoPopup, False)
		End If
	End With
End Function

'==================================== User defined function ============================================
Sub txtMovType_OnChange()                    
	Call WcCdChange()	
End Sub


Sub WcCdChange()
	If UCase(frm1.txtMovType.value) = "I97" Then
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "N")
	Else 
		Call ggoOper.SetReqAttr(frm1.txtWcCd, "D")
	End if
End Sub


Sub txtDocumentDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt.Focus
    End If
End Sub


Sub txtDocumentDt_Change()
    lgBlnFlgChgValue = True
End Sub


Sub txtPostingDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtPostingDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtPostingDt.Focus
    End If
End Sub


Sub txtPostingDt_Change()
    lgBlnFlgChgValue = True
End Sub


Sub txtYear_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYear.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtYear.Focus
    End If
End Sub


Sub txtYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub



'========================================== vspdData_ function ========================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )

  	Dim DblEntryQty

	With frm1.vspdData 
		If Col = C_Entryqty then
		    .Col = C_EntryQty
		    If .Text = "" Then
		       	DblEntryQty = 0
		    Else
		       	DblEntryQty = UNICDbl(.Text)
		    End If
		End If
	End With
                
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	With Frm1.vspdData
		Select Case Col
		Case C_ItemCd
			.Col = Col
			.Row = Row		
			
			If 	CommonQueryRs(" A.item_nm, A.spec, A.basic_unit ", " B_ITEM A, B_ITEM_BY_PLANT B ", _
			    " A.item_cd = B.item_cd AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.item_cd = " & FilterVar(Frm1.vspdData.Text, "''", "S"), _
			    lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

				Call .SetText(C_ItemNm,   Row, "")
				Call .SetText(C_ItemSpec, Row, "")
				Call .SetText(C_EntryUnit,Row, "")
				Call .SetText(C_InvUnit,  Row, "")
				.focus
				Exit Sub
			End If
			
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			lgF2 = Split(lgF2, Chr(11))
			Call .SetText(C_ItemNm,   Row, lgF0(0))
			Call .SetText(C_ItemSpec, Row, lgF1(0))
			Call .SetText(C_EntryUnit,Row, lgF2(0))
			Call .SetText(C_InvUnit,  Row, lgF2(0))
			
		End Select
	End With
End Sub


Sub vspdData_Click(ByVal Col , ByVal Row )
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
	End If	

	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
	
	If Row = 0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortKey = 1 then
			ggoSpread.SSSort Col			
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	
			lgSortKey = 1
		End if
		Exit Sub
	End if
End Sub


Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 


Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 


Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 


Sub vspddata_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey <> "" Then							
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if
    
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_ItemPopUp Then
			.Col = C_ItemCd
			.Row = Row
			Call OpenItem(.Text)		
		
		Elseif Row > 0 And Col = C_LotNoPopUp Then
			.Col = C_LotNo
			.Row = Row
			Call OpenLotNo(.Text)
			
		Elseif Row > 0 And Col = C_TrackingNoPopUp Then
			.Col = C_TrackingNo
			.Row = Row
			Call OpenTrackingNo(.Text)
			
		Elseif Row > 0 And Col = C_EntryUnitPopUp Then
			.Col = C_EntryUnit
			.Row = Row
			Call OpenEntryUnit(.Text)
			
		Elseif Row > 0 And Col = C_ProdOrdNoPopUp Then
			.Col = C_ProdOrdNo
			.Row = Row
			Call OpenProdOrderNo()
			
		End If
	End With
End Sub



'======================================== FncQuery() =================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData                                               

    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					
		If IntRetCD = vbNo Then Exit Function
    End If
    
	If Not chkFieldByCell(frm1.txtPlantCd, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtDocumentNo1, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtYear, "A",1) Then Exit Function							

	Call ggoOper.ClearField(Document, "2")									
	Call ggoOper.LockField(Document, "N")
	
	frm1.txtDocumentDt.text = StartDate
	frm1.txtPostingDt.text  = StartDate
	Call InitVariables                                
	
    If 	CommonQueryRs(" A.MINOR_TYPE "," B_MINOR A, I_GOODS_MOVEMENT_HEADER  B ", _
					" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("OI", "''", "S") & " AND B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		If 	CommonQueryRs(" B.ITEM_DOCUMENT_NO "," I_GOODS_MOVEMENT_HEADER  B ", _
						"B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	    	Call DisplayMsgBox("160101","X","X","X")
		Else
			Call DisplayMsgBox("169955","X","X","X")
		End if
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement		
		Exit function
    End If
    lgF0 = Split(lgF0,Chr(11))
    if	Trim(lgF0(0)) <> "U" then
		Call DisplayMsgBox("169955","X","X","X")
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement
		Exit Function
    end if
    
    Call ClickTab1() 
    
    If DbQuery() = False Then Exit Function
		
    FncQuery = True																
   
End Function



'===================================== FncNew() ================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")         
		If IntRetCD = vbNo Then Exit Function
	End If
    
    ggoSpread.ClearSpreadData
	
    If gPageNo <> 1 Then
		Call changeTabs(1)
	End If
	
    Call ggoOper.ClearField(Document, "A")                                         
    Call ggoOper.LockField(Document, "N")                                         
    Call InitVariables                                                      
    Call SetDefaultVal
    Call SetToolBar("11101000000011")										
    
    FncNew = True                                                           

End Function


'===================================== FncDelete() ===============================================
Function FncDelete() 
    
    FncDelete = False                                                      
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
    FncDelete = True                                                        

End Function


'===================================== FncSave() ================================================
Function FncSave() 
	Dim IntRetCD 
	Dim idx
	Dim iRet
	Dim strYear2
	Dim strMonth2
	Dim strDay2
	Dim strCurrDt
    Dim strDocumentDt
    Dim strPostingDt
	
	FncSave = False                                                        
	
	Err.Clear                                                               
	On Error Resume Next   

	ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            
		Exit Function
	End If

	If Not chkField(Document, "2")  Then     	
       If gPageNo = 1 Then	              
          Call SetToolBar("11101000000011")
       Else	             
          Call SetToolBar("11101101000011")
       End If
       Exit Function
    End If 

	If UCase(Trim(frm1.txtMovType.Value)) = "I97" Then
		If Trim(frm1.txtWcCd.value) = "" then	
			iRet = DisplayMsgBox("970021", "X", frm1.txtWcCd.alt,"x")
			Exit Function
		End if
	End If	

	ggoSpread.Source = frm1.vspdData
	
	If Not ggoSpread.SSDefaultCheck Then Exit Function

	If frm1.vspdData.MaxRows <= 0  Then
		IntRetCD = DisplayMsgBox("122610","X", "X", "X")
		Exit Function
	End If

	If UCase(Trim(frm1.txtMovType.Value)) = "I97" Then
		for idx = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = idx
			frm1.vspdData.Col = C_ProdOrdNo

			If (frm1.vspdData.Text = "" or frm1.vspdData.Text = Null) Then
				Call DisplayMsgBox("169945","X", "X", "X")
				Call ClickTab2()
				frm1.vspdData.action = 0
				Exit Function
			End if
			
			frm1.vspdData.Col = C_ReqNo
			If (frm1.vspdData.Text = "" or frm1.vspdData.Text = Null) Then
				Call DisplayMsgBox("162053","X", "X", "X")
				Call ClickTab2()
				frm1.vspdData.Col = C_ProdOrdNo
				frm1.vspdData.action = 0
				Exit Function
			End If
		next
	End If		

    If 	CommonQueryRs(" A.MINOR_TYPE "," B_MINOR A, I_MOVETYPE_CONFIGURATION B ", _
		" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("OI", "''", "S") & " AND A.MINOR_CD = " & Trim(FilterVar(frm1.txtMovType.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

		Call DisplayMsgBox("169955","X","X","X")
		Call ClickTab1()
		frm1.txtMovType.Focus		
		Exit function
    End If

    lgF0 = Split(lgF0,Chr(11))

    if	Trim(lgF0(0)) <> "U" then
		Call DisplayMsgBox("169955","X","X","X")
		Call ClickTab1()
		frm1.txtMovType.Focus
		Exit Function
    end if

	Call ExtractDateFrom(frm1.txtDocumentDt.text, Parent.gDateFormat,Parent.gComDateType,strYear2,strMonth2,strDay2)

	frm1.txtDocumentDt.Year = strYear2
	
	If Trim(frm1.txtYear.Year) <> Trim(frm1.txtDocumentDt.Year) then
		Call DisplayMsgBox("169940","X","X","X")   
		Call ClickTab1()
		Call SetFocusToDocument("M")  
        frm1.txtYear.Focus
        Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
    strDocumentDt  = UniConvDateToYYYYMMDD(frm1.txtDocumentDt.text,Parent.gDateFormat,"")
    strCurrDt      = UniConvDateToYYYYMMDD(StartDate,parent.gDateFormat,"")
	strPostingDt   = UniConvDateToYYYYMMDD(frm1.txtPostingDt.text,Parent.gDateFormat,"")
	
    if strDocumentDt > strCurrDt then	
		Call DisplayMsgBox("169944","X","X","X")
		Call ClickTab1()
		Call SetFocusToDocument("M")
		frm1.txtDocumentDt.focus
		Set gActiveElement = document.activeElement
		Exit Function
    End if

	if strPostingDt > strCurrDt then	
		Call DisplayMsgBox("169944","X","X","X")
		Call ClickTab1()
		Call SetFocusToDocument("M")
		frm1.txtPostingDt.focus
		Set gActiveElement = document.activeElement
		Exit Function
    End if

	If DbSave() = False Then Exit Function
	
	FncSave = True                                                         
End Function



'====================================== FncCopy() =========================================================
Function FncCopy() 
    If frm1.vspdData.maxrows < 1 then exit function
	
	If gPageNo = 2 Then
    	frm1.vspdData.ReDraw = False
    	
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.CopyRow
        SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow
        
    	frm1.vspdData.ReDraw = True
    ElseIf gPageNo = 1 Then 
        Call ggoOper.ClearField(Document, "1")                        
    End If 
End Function



'======================================= FncPaste() =================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function



'======================================= FncCancel() ==================================================
Function FncCancel() 
    If frm1.vspdData.maxrows < 1 then exit function
	 ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  
End Function


'======================================== FncInsertRow() ================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
	Dim imRow
	Dim intRow
	
	On Error Resume Next
	
	FncInsertRow = False
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow ="" Then Exit Function
	End if
	
	If gPageNo <> 2 Then
		Call ClickTab2		
	End If
	
	With frm1.vspdData
	
		.focus
    	ggoSpread.Source = frm1.vspdData    		
    	.ReDraw = False
    	ggoSpread.InsertRow .ActiveRow,  imRow
    
    	SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
    
    	lgBlnFlgChgValue = True
    	
    	For intRow= .ActiveRow  to .ActiveRow +imRow-1
 			Call .SetText(C_TrackingNo,intRow, "*")
			Call .SetText(C_LotSubNo,  intRow, 0)
			Call .SetText(C_LotNo,     intRow, "*")
   		Next
		.ReDraw = True
    
    End With

End Function



'============================================= FncDeleteRow() =========================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 
	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function


Function FncPrint()     
	Call parent.FncPrint()
End Function


Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)											 
End Function


Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , False)                                                   
End Function


Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function


Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 11
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		frm1.vspdData.Col = iColumnLimit:
		frm1.vspdData.Row = 0
       iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function  
    End If  
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    
End Function


'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function


'============================== DbDeleteOk() ===========================================
Function DbDeleteOk()												
	Call MainNew()
End Function

 

'================================ DbQuery() ============================================
Function DbQuery() 
        
	DbQuery = False                                                        
	
	Dim strVal
	
	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY_ID &	"?txtMode="			& parent.UID_M0001					& _							
								"&txtPlantCd="      & Trim(frm1.txtPlantCd.value)		& _
								"&txtDocumentNo1=" & Trim(frm1.txtDocumentNo1.value)	& _				
								"&txtYear="			& Trim(frm1.txtYear.Year)			& _
								"&lgStrPrevKey="    & lgStrPrevKey
	
	Call RunMyBizASP(MyBizASP, strVal)							
				
	DbQuery = True                                                          

End Function



'================================== DbQueryOk() =========================================
Function DbQueryOk()													
	
    If gPageNo = 1 Then
		Call SetToolBar("11101000000111")
	Else
		Call SetToolBar("11101011000111")
	End if														
	
    lgIntFlgMode = parent.OPMD_UMODE													
    
    Call ggoOper.LockField(Document, "Q")										
	lgBlnFlgChgValue = False
	
	lgMovType = frm1.txtMovType.value
End Function


'===================================== DbSave() ============================================
Function DbSave() 

    
   	Dim IntRows 
	Dim strVal
	Dim iRowSep, iColSep
	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount

    Call LayerShowHide(1)
	
	iRowSep = Parent.gRowSep
	iColSep = Parent.gColSep

	DbSave = False                                                         
	
	On Error Resume Next                                                   
		
	With frm1
		.txtMode.value         = Parent.UID_M0002											
		.hYear.value           = .txtYear.Year
	End With

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
		
	
	With frm1.vspdData
	    
		For IntRows = 1 To .MaxRows
		
			.Row = IntRows
			.Col = 0

		
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					
					If .Text = ggoSpread.InsertFlag Then
						strVal = "C" & iColSep & IntRows & iColSep				
					Else
						strVal = "U" & iColSep & IntRows & iColSep				
					End If
							
					.Col = C_ItemCd	
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_TrackingNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotSubNo	
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_EntryQty	
					if  UNICDbl(.Text) = 0 then
						Call DisplayMsgBox("169918","X", "X", "X")
						Call LayerShowHide(0) 
						Call ClickTab2()
						.Action = 0
						exit function
					end if
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_EntryUnit		
					strVal = strVal & Trim(.Text) & iColSep	
					.Col = C_ProdOrdNo
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ReqNo      
					strVal = strVal & Trim(.Text) & iRowSep
					
				Case ggoSpread.DeleteFlag
		
					strVal = "D" & iColSep & IntRows & iColSep				
					.Col = C_SeqNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_SubSeqNo 	
					strVal = strVal & Trim(.Text) & iColSep				
					.Col = C_ItemCd
					strVal = strVal & Trim(.Text) & iColSep	
					.Col = C_ProdOrdNo
					strVal = strVal & Trim(.Text) & iRowSep
			End Select
			
			.Row = IntRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag, ggoSpread.DeleteFlag
					If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
								                            
						Set objTEXTAREA = document.createElement("TEXTAREA")
						objTEXTAREA.name = "txtCUSpread"
						objTEXTAREA.value = Join(iTmpCUBuffer,"")
						divTextArea.appendChild(objTEXTAREA)     
									 
						iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
						ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
						iTmpCUBufferCount = -1
						strCUTotalvalLen  = 0
											
					End If
								       
					iTmpCUBufferCount = iTmpCUBufferCount + 1
								      
					If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
						iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
						ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
					End If   
											
					iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
					strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End Select		

		Next
	End With
	
	If iTmpCUBufferCount > -1 Then 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	Call RemovedivTextArea								
	
	DbSave = True                                                          
    
End Function



'==================================== DbSaveOk() ===============================================
Function DbSaveOk()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData													
	lgBlnFlgChgValue = false
	
	Call ClickTab1()
	Call MainQuery()

End Function


