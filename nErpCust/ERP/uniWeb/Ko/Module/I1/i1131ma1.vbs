Const BIZ_PGM_QRY_ID  = "i1131mb1.asp"							
Const BIZ_PGM_DQRY_ID = "i1131mb1.asp"							
Const BIZ_PGM_SAVE_ID = "i1131mb2.asp"							
Const BIZ_PGM_DEL_ID  = "i1131mb3.asp"							

Dim C_ItemCd          
Dim C_ItemPopup       
Dim C_ItemNm          
Dim C_EntryAmount     
Dim C_EntryQty        
Dim C_EntryUnit       
Dim C_EntryUnitPopup  
Dim C_InvUnit         
Dim C_ItemSpec        
Dim C_TrackingNo      
Dim C_TrackingNoPopup 
Dim C_LotNo          
Dim C_LotSubNo       
Dim C_LotNoPopup     
Dim C_SeqNo          
Dim C_SubSeqNo       
    

'***********************************Sub InitVariables()***************************************
Sub InitVariables()
	
	lgIntFlgMode		= parent.OPMD_CMODE           
    lgBlnFlgChgValue	= False                   
    lgIntGrpCount		= 0                          
    lgStrPrevKey		= ""                          
    lgLngCurRows		= 0                           
    lgSortKey			= 1
    lgMovType			= ""  
End Sub


'**********************************Sub InitSpreadPosVariables()*******************************
Sub InitSpreadPosVariables()
	 C_ItemCd          = 1        
	 C_ItemPopup       = 2    
	 C_ItemNm          = 3    
	 C_EntryAmount     = 4    
	 C_EntryQty        = 5    
	 C_EntryUnit       = 6    
	 C_EntryUnitPopup  = 7  
	 C_InvUnit         = 8 
	 C_ItemSpec        = 9 
	 C_TrackingNo      = 10 
	 C_TrackingNoPopup = 11
	 C_LotNo           = 12
	 C_LotSubNo        = 13
	 C_LotNoPopup      = 14
	 C_SeqNo           = 15
	 C_SubSeqNo        = 16

End Sub


'*********************************Sub SetDefaultVal()****************************************
Sub SetDefaultVal()
	
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtDocumentDt.text = StartDate
	frm1.txtPostingDt.text  = StartDate
	
	Call ExtractDateFrom(Currentdate, Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.txtYear.Year = strYear 
	
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

'************************************ Sub InitSpreadSheet() ************************************
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030423", , parent.gAllowDragDropSpread
    
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_SubSeqNo + 1	
		.MaxRows = 0
    	    		
	    Call GetSpreadColumnPos("A")		
	    Call AppendNumberPlace("6", "3", "0")
		
		ggoSpread.SSSetEdit		  C_ItemCd,			"품목",15, 0, -1, 18, 2		
		ggoSpread.SSSetButton 	  C_ItemPopup		
		ggoSpread.MakePairsColumn C_ItemCd, C_ItemPopup
		
		ggoSpread.SSSetEdit		  C_ItemNm,			"품목명", 20, 0, -1, 50		
		ggoSpread.SSSetFloat      C_EntryAmount,	"입고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat      C_EntryQty,		"입고수량", 15, Parent.ggQtyNo,        ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		
		ggoSpread.SSSetEdit	      C_EntryUnit,		"입고단위", 10, 0, -1, 3, 2 
		ggoSpread.SSSetButton     C_EntryUnitPopup
		ggoSpread.MakePairsColumn C_EntryUnit, C_EntryUnitPopup
		
		ggoSpread.SSSetEdit		  C_InvUnit,		"재고단위", 10, 0, -1, 3 	
		ggoSpread.SSSetEdit		  C_ItemSpec,		"규격", 20, 0, -1, 50		
		ggoSpread.SSSetEdit		  C_TrackingNo,		"Tracking No", 20, 0, -1, 25, 2
		ggoSpread.SSSetButton 	  C_TrackingNoPopup	
		ggoSpread.MakePairsColumn C_TrackingNo, C_TrackingNoPopup
					
		ggoSpread.SSSetEdit 	C_LotNo,			"LOT NO",20, 0, -1, 25, 2
		ggoSpread.SSSetFloat	  C_LotSubNo,		"순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetButton 	  C_LotNoPopup
		ggoSpread.MakePairsColumn C_LotNo,C_LotNoPopup

		ggoSpread.SSSetEdit		  C_SeqNo,    "PK1", 1, 0
		ggoSpread.SSSetEdit		  C_SubSeqNo, "PK2", 1, 0
	    
		Call ggoSpread.SSSetColHidden(C_SeqNo,   .MaxCols,   True)	
		
		ggoSpread.SpreadLockWithOddEvenRowColor	
		
		.ReDraw = true
		
		ggoSpread.SSSetSplit2(3) 
		
	End With
	
End Sub


'************************************* Sub GetSpreadColumnPos(ByVal pvSpdNo) **********************************
Sub GetSpreadColumnPos(ByVal pvSpdNo)

Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
		      
		      ggoSpread.Source = frm1.vspdData
		      
		      Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		      
              C_ItemCd          =iCurColumnPos(1)
 			  C_ItemPopup       =iCurColumnPos(2)
 			  C_ItemNm          =iCurColumnPos(3)
			  C_EntryAmount     =iCurColumnPos(4)
			  C_EntryQty        =iCurColumnPos(5)
			  C_EntryUnit       =iCurColumnPos(6)
			  C_EntryUnitPopup  =iCurColumnPos(7)
			  C_InvUnit         =iCurColumnPos(8)
			  C_ItemSpec        =iCurColumnPos(9)
			  C_TrackingNo      =iCurColumnPos(10)
			  C_TrackingNoPopup =iCurColumnPos(11)
			  C_LotNo           =iCurColumnPos(12)
			  C_LotSubNo        =iCurColumnPos(13)
			  C_LotNoPopup      =iCurColumnPos(14)
			  C_SeqNo           =iCurColumnPos(15)
		      C_SubSeqNo        =iCurColumnPos(16)
		     
	End Select

End Sub


'******************************** Sub PopSaveSpreadColumnInf() ********************************************
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()	
End Sub


'********************************** Sub PopRestoreSpreadColumnInf() ****************************************
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet
	Call ggoSpread.ReOrderingSpreadData
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor( ByVal pvStartRow, ByVal pvEndRow)
    
   	With frm1
   	
   		If 	CommonQueryRs("REFERENCE","B_CONFIGURATION", " MAJOR_CD = " & FilterVar("I0001", "''", "S") & "" _
   																		& " AND SEQ_NO = " & 97 _
   																		& " AND MINOR_CD = " & FilterVar(UCase(.txtMovType.value), "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			StrCompany = lgF0(0)
		End If
		
		ggoSpread.SSSetRequired	 C_ItemCd,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm,   pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_InvUnit,  pvStartRow, pvEndRow
		
		If UCase(StrCompany) <> UCase(parent.gCompany) Then  
			ggoSpread.SSSetProtected C_EntryAmount, pvStartRow,                pvEndRow
		ElseIf UCase(StrCompany) = UCase(parent.gCompany) Then
			ggoSpread.SpreadUnLock   C_EntryAmount, pvStartRow, C_EntryAmount, pvEndRow 
			ggoSpread.SSSetRequired  C_EntryAmount, pvStartRow,                pvEndRow
		End If
			
		If UCase(.txtMovType.value) = "R90" or UCase(.txtMovType.value) = "R89" Then 
			ggoSpread.SSSetProtected C_EntryQty,       pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_EntryUnit,      pvStartRow, pvEndRow
			ggoSpread.SSSetProtected C_EntryUnitPopup, pvStartRow, pvEndRow
		ElseIf UCase(.txtMovType.value) <> "R90" and UCase(.txtMovType.value) <> "R89" Then
		    ggoSpread.SpreadUnLock	 C_EntryQty,       pvStartRow, C_EntryQty,       pvEndRow
		    ggoSpread.SSSetRequired	 C_EntryQty,       pvStartRow,                   pvEndRow
		    ggoSpread.SpreadUnLock   C_EntryUnit,      pvStartRow, C_EntryUnit,      pvEndRow
		    ggoSpread.SSSetRequired	 C_EntryUnit,      pvStartRow,                   pvEndRow
		    ggoSpread.SpreadUnLock   C_EntryUnitPopup, pvStartRow, C_EntryUnitPopup, pvEndRow
		    ggoSpread.SSSetProtected C_InvUnit,        pvStartRow,                   pvEndRow
		    ggoSpread.SSSetProtected C_ItemSpec,       pvStartRow,                   pvEndRow
		End if
	End With
End Sub


'**************************************Function SheetFocus(lRow, lCol)**********************
Function SheetFocus(lRow, lCol)
	Call changeTabs(2)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")						
	ElseIf lgIntFlgMode = Parent.OPMD_UMODE Then
	    Call SetToolBar("11101011000111")						
	End If
	
	frm1.vspdData.focus
	frm1.vspdData.Row		= lRow
	frm1.vspdData.Col		= lCol
	frm1.vspdData.Action	= 0
	frm1.vspdData.SelStart	= 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

'*************************Function ClickTab1()**********************************************
Function ClickTab1()

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
	
	gSelframeFlg=1
	
End Function


'***********************Function ClickTab2()************************************************
Function ClickTab2()
	Dim IRow
	
	if (frm1.txtMovType.value = "" and lgIntFlgMode = parent.OPMD_CMODE ) then
	    Call DisplayMsgBox("169904","X", "X", "X")   
	    frm1.txtMovType.Focus
	    Exit function
	End if
	
	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1
	
    Call changeTabs(2)
		
	If lgIntFlgMode = parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")						
	ElseIf lgIntFlgMode = parent.OPMD_UMODE Then
	    Call SetToolBar("11101011000111")					
	End If
	
	If frm1.txtMovType.value <> lgMovType Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		lgMovType = frm1.txtMovType.value	
	End if
	
	gSelframeFlg=2
	Set gActiveElement = document.activeElement
End Function

'**************************** Function OpenPlant() ***********************************8
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


'************************************** Function OpenDocumentNo() **************************************
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
	Param3 = "OR"
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
	
End Function


'************************************** Function OpenSL() ****************************************
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd.Focus
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtSLCd.ClassName) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLcd.Value)
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
		frm1.txtSLcd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
	
End Function


'************************************** Function OpenMovType() ********************************************
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
	Param3 = "OR"

	iCalledAspName = AskPRAspName("I1411PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1411PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3), _
		"dialogWidth=465px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMovType.focus
		Exit Function
	Else
		Call SetMovType(arrRet)
	End If	
	
End Function



'************************************* Function OpenItem(Byval strCode) ***********************************
Function OpenItem(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(5), arrField(6)

	if UCase(frm1.txtMovType.value) <> "R90" and UCase(frm1.txtMovType.value) <> "R89" then 	
		Dim arrRet
		
		If Trim(frm1.txtPlantCd.Value) = "" then
			Call DisplayMsgBox("169901","X","X","X")   
			Call ClickTab1()
			Call SetFocusToDocument("M")  
			frm1.txtPlantCd.Focus
			Exit Function
		End If
	
		If IsOpenPop = True Then Exit Function	

		IsOpenPop = True
	
		arrParam(0) = Trim(frm1.txtPlantCd.value)	
		arrParam(1) = strCode						
		arrParam(2) = ""           
		arrParam(3) = ""			
	
		arrField(0) = 1
		arrField(1) = 2
		arrField(2) = 3
		arrField(3) = 4
		arrField(4) = 4

		iCalledAspName = AskPRAspName("B1B11PA3")				
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","x")
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
	Else
		Dim arrRet2
	
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
	    
		arrRet2 = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
		IsOpenPop = False
	
		If arrRet2(0) = "" Then
			Call SetActiveCell(frm1.vspdData,C_ItemCd,frm1.vspdData.ActiveRow,"M","X","X")
			Exit Function
		Else
			Call SetItem2(arrRet2)
		End If	
	End if
	Set gActiveElement = document.activeElement   
End Function



'************************************** Function OpenTrackingNo(Byval strCode) **********************************
Function OpenTrackingNo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strItemCd

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(4) = ""							
	arrParam(5) = "Major코드"			

	arrParam(0) = "TRACKINGNO"				
	arrParam(1) = "s_so_tracking"				
	
	With frm1.vspdData
		.Col = C_TrackingNo
		arrParam(2) = .Text		 
		arrParam(3) = ""		
		.Col = C_ItemCd
		strItemCd   = .Text		
	End With
	arrParam(4) = " "
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


'*************************************** Function OpenEntryUnit(Byval strCode) ********************************
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



'*********************************** Function OpenLotNo(Byval strCode) *************************************
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
			Call SetActiveCell(frm1.vspdData,C_ItemCd,.ActiveRow,"M","X","X")
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
	Set gActiveElement = document.activeElement	
End Function



'************************************ Function OpenPopupGL() *********************************************
Function OpenPopupGL()
	Dim iCalledAspName
	Dim IntRetCD

	If GetSetupMod(parent.gSetupMod, "a") <> "Y" Then 
       Call DisplayMsgBox("169934","X", "X", "X")
       frm1.txtPlantCd.focus
       Exit Function
	End if 
	
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
       Exit Function
    End If
	   	
	strRefNo = Trim(frm1.txtDocumentNo1.value) & "-" & Trim(frm1.txtYear.Year)  
	strFrom	 = "ufn_a_GetGlNo( " & FilterVar((strRefNo), "''" , "S") & " )"
		
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
		
	IsOpenPop = True

	iCalledAspName = AskPRAspName("A5120RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5120RA1","x")
		IsOpenPop = False
		Exit Function
	End If
   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	Set gActiveElement = document.activeElement
	
End Function



'********************************* Function OpenPopupGL2() ************************************
Function OpenPopupGL2()
	Dim iCalledAspName
	Dim IntRetCD
	
	If GetSetupMod(parent.gSetupMod, "a") <> "Y" Then
       Call DisplayMsgBox("169934","X", "X", "X")
       frm1.txtPlantCd.focus
       Exit Function
	End if
     
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
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("A5130RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A5130RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	   
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	Set gActiveElement = document.activeElement   
End Function



'*************************************** Function OpenCostCd() *****************************************
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


'********************************* Function Set() **************************************************
Function SetPlant(byRef arrRet)
	frm1.txtPlantCd.Value   = arrRet(0)		
	frm1.txtPlantNm.Value   = arrRet(1)
	frm1.txtPlantCd.focus
End Function

Function SetDocumentNo(byval arrRet)
	frm1.txtDocumentNo1.Value   = arrRet(0)	
	frm1.txtYear.Year           = arrRet(1)
	frm1.txtDocumentNo1.focus
End Function

Function SetSL(byval arrRet)
	frm1.txtSLCd.Value		= arrRet(0)		
	frm1.txtSLNm.Value		= arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue		= True
End Function

Function SetMovType(byval arrRet)
	frm1.txtMovType.Value		= arrRet(0)
	frm1.txtMovTypeNm.Value		= arrRet(1)
	lgBlnFlgChgValue			= True
	frm1.txtMovType.focus
End Function

Function SetCostCd(byval arrRet)
	frm1.txtCostCd.value	= arrRet(0)
	frm1.txtCostNm.value	= arrRet(1)
	frm1.txtCostCd.focus
	lgBlnFlgChgValue		= True
End Function

Function SetItem(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_ItemCd,	   .ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm,	   .ActiveRow, arrRet(1))
		Call .SetText(C_ItemSpec,  .ActiveRow, arrRet(2))
		Call .SetText(C_InvUnit,   .ActiveRow, arrRet(3))
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(4))
		Call vspdData_Change(C_ItemCd, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_EntryQty,.ActiveRow,"M","X","X")		
	End With
End Function

Function SetItem2(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_ItemCd,	   .ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm,	   .ActiveRow, arrRet(1))
		Call .SetText(C_ItemSpec,  .ActiveRow, arrRet(2))
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(3))
		Call .SetText(C_InvUnit,   .ActiveRow, arrRet(3))
		Call .SetText(C_TrackingNo,.ActiveRow, arrRet(4))
		Call .SetText(C_LotNo,     .ActiveRow, arrRet(5))
		Call .SetText(C_LotSubNo,  .ActiveRow, arrRet(6))
		Call vspdData_Change(C_ItemCd, .ActiveRow)		
		Call SetActiveCell(frm1.vspdData,C_EntryAmount,.ActiveRow,"M","X","X")	
	End With
End Function

Function SetTrackingNo(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo,.ActiveRow, arrRet(0))
		Call vspdData_Change(C_TrackingNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_LotNo,.ActiveRow,"M","X","X")		
	End With
End Function

Function SetEntryUnit(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_EntryUnit, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_TrackingNo,.ActiveRow,"M","X","X")		
	End With
End Function

Function SetLotNo(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(2))
		Call .SetText(C_LotNo,		.ActiveRow, arrRet(3))
		Call .SetText(C_LotSubNo,	.ActiveRow, arrRet(4))
		Call vspdData_Change(C_LotNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_LotNo,.ActiveRow,"M","X","X")		
	End With
End Function

'*********************************** Sub txtDocumentDt_DblClick(Button)  *********************************
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



'************************************ Sub vspdData_EditChange() ************************************
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



'*************************************** Sub vspdData_Change() ***************************************
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	With frm1.vspdData	
	
	If Col = C_ItemCd Then
		.Col = Col
		.Row = Row		
		
		If 	CommonQueryRs(" A.item_nm, A.spec, A.basic_unit ", " B_ITEM A, B_ITEM_BY_PLANT B ", _
		    " A.item_cd = B.item_cd AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.item_cd = " & FilterVar(.Text, "''", "S"), _
		    lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			Call .SetText(C_ItemNm,		Row, "")
			Call .SetText(C_ItemSpec,	Row, "")
			Call .SetText(C_EntryUnit,	Row, "")
			Call .SetText(C_InvUnit,	Row, "")
			.focus
			Exit Sub
		End If
		
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		Call .SetText(C_ItemNm,		Row, lgF0(0))
		Call .SetText(C_ItemSpec,	Row, lgF1(0))
		Call .SetText(C_EntryUnit,	Row, lgF2(0))
		Call .SetText(C_InvUnit,	Row, lgF2(0))
	End If
	End With
End Sub



'********************************* Sub vspdData_Click() **************************************
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



'*************************************** Sub vspdData_ScriptDragDropBlock() **************************
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


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
    If CheckRunningBizProcess = True Then Exit Sub
    
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call DisableToolBar(parent.TBC_QUERY)
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
		
		Elseif Row > 0 And Col = C_TrackingNoPopup Then
			.Col = C_TrackingNo
			.Row = Row
			Call OpenTrackingNo(.Text)
		
		Elseif Row > 0 And Col = C_EntryUnitPopUp Then
			.Col = C_EntryUnit
			.Row = Row
			Call OpenEntryUnit(.Text)
		
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			
		If IntRetCD =  vbNo Then Exit Function
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
		" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("OR", "''", "S") & " AND B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

		If 	CommonQueryRs(" B.ITEM_DOCUMENT_NO "," I_GOODS_MOVEMENT_HEADER  B ", _
		"B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
    
     		Call DisplayMsgBox("160101","X","X","X")
		Else
			Call DisplayMsgBox("169954","X","X","X")
		End if
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement		
		Exit function
    End If
    lgF0 = Split(lgF0,Chr(11))
    if	Trim(lgF0(0)) <> "U" then
		Call DisplayMsgBox("169954","X","X","X")
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement
		Exit Function
    end if
    
    Call ClickTab1()

	If DBQuery = False Then Exit Function 
	
	FncQuery = True		
   
End Function



'************************************ Function FncNew()  ********************************************
Function FncNew() 
	Dim IntRetCD 

	FncNew = False                                                       
	
	ggoSpread.Source = frm1.vspdData

	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")   
		If IntRetCD =  vbNo Then Exit Function
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
	
	lgBlnFlgChgValue = False

	FncNew = True                                                       

End Function



'*********************************** Function FncDelete() *********************************************
Function FncDelete() 
	If lgIntFlgMode <> Parent.OPMD_UMODE Then                          
	    Call DisplayMsgBox("900002", "X", "X", "X")                    
		Exit Function
	End If
End Function


'*********************************** Function FncSave() ***********************************************
Function FncSave() 
	Dim IntRetCD 
	Dim strYear2
	Dim strMonth2
	Dim strDay2
    Dim strCurrDt
    Dim strDocumentDt
    Dim strPostingDt

	FncSave = False                                                       
	
	Err.Clear                                                             

	ggoSpread.Source = frm1.vspdData
		
	If lgBlnFlgChgValue = False AND ggoSpread.SSCheckChange = False  Then
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

	ggoSpread.Source = frm1.vspdData
	
	If ggoSpread.SSDefaultCheck = False Then Exit Function                              

	If frm1.vspdData.MaxRows <= 0  Then
		IntRetCD = DisplayMsgBox("122610", "X", "X", "X")
		Exit Function
	End If
	
	If 	CommonQueryRs(" A.MINOR_TYPE "," B_MINOR A, I_MOVETYPE_CONFIGURATION B ", _
		" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("OR", "''", "S") & " AND A.MINOR_CD = " & Trim(FilterVar(frm1.txtMovType.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false Then
	
		Call DisplayMsgBox("169954","X","X","X")
		Call ClickTab1()
		frm1.txtMovType.Focus		
		Exit function
    End If

    lgF0 = Split(lgF0,Chr(11))
    if	Trim(lgF0(0)) <> "U" then
		Call DisplayMsgBox("169954","X","X","X")
		Call ClickTab1()
		frm1.txtMovType.Focus
		Exit Function
    end if
	
	
	Call ExtractDateFrom(frm1.txtDocumentDt.text, Parent.gDateFormat, Parent.gComDateType,strYear2,strMonth2,strDay2)

	frm1.txtDocumentDt.Year = strYear2
	
	If Trim(frm1.txtYear.Year) <> Trim(frm1.txtDocumentDt.Year) then
		Call DisplayMsgBox("169940","X","X","X")  
		Call ClickTab1()
		Call SetFocusToDocument("M")  
        frm1.txtYear.Focus
        Set gActiveElement = document.activeElement
		Exit Function
	End If

    strDocumentDt  = UniConvDateToYYYYMMDD(frm1.txtDocumentDt.text, Parent.gDateFormat,"")
    strCurrDt      = UniConvDateToYYYYMMDD(StartDate, parent.gDateFormat, "")
	strPostingDt   = UniConvDateToYYYYMMDD(frm1.txtPostingDt.text, Parent.gDateFormat,"")
	
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
 
	If DBSave = False Then Exit Function 
		
	FncSave = True                                                          
End Function



'***************************** Function FncCopy() ****************************************************
Function FncCopy() 
    If frm1.vspdData.maxrows < 1 then exit function
	
	If gPageNo = 2 Then
    	frm1.vspdData.ReDraw = False
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.CopyRow
        SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    	frm1.vspdData.ReDraw = True
    ElseIf gPageNo = 1 Then 
        Call ggoOper.ClearField(Document, "1")                               
    End If 
    
    Set gActiveElement = document.activeElement
    
End Function


'******************************* Function FncPaste() ************************************************
Function FncPaste() 
	ggoSpread.SpreadPaste
End Function


'******************************** Function FncCancel() **********************************************
Function FncCancel()
     If frm1.vspdData.maxrows < 1 then exit function
	 ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo                                                
End Function


'********************************* Function FncInsertRow(ByVal pvRowCnt) ******************************
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
    	ggoSpread.InsertRow  .ActiveRow,  imRow
    	SetSpreadColor .ActiveRow, .ActiveRow + imRow -1
    	
    	lgBlnFlgChgValue = True
    	
    	For intRow= .ActiveRow  to .ActiveRow +imRow-1
  			Call .SetText(C_TrackingNo, intRow, "*")
  			Call .SetText(C_LotSubNo, intRow, 0)
  			Call .SetText(C_LotNo, intRow, "*")
		Next
    	.ReDraw = True
    End With
End Function


'********************************* Function FncDeleteRow() *********************************************
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
		If IntRetCD =  vbNo Then Exit Function
    End If
    FncExit = True
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


Function DbDeleteOk()											
	Call MainNew()
End Function


'************************************* Function DbQuery() *******************************************
Function DbQuery() 

	Dim strVal
	Err.Clear
	
   	DbQuery = False                                                       
   	
   	LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID &	"?txtMode="        & Parent.UID_M0001					& _
								"&txtPlantCd="     & Trim(frm1.txtPlantCd.value)		& _
								"&txtDocumentNo1=" & Trim(frm1.txtDocumentNo1.value)	& _	
								"&txtYear="        & Trim(frm1.txtYear.Year)			& _
								"&lgStrPrevKey="   & lgStrPrevKey
	
	Call RunMyBizASP(MyBizASP, strVal)									
	
	DbQuery = True                                                      
    
End Function


'****************************** Function DbQueryOk() *********************************************
Function DbQueryOk()													

	If gPageNo = 1 Then
		Call SetToolBar("11101000000111")
	Else
		Call SetToolBar("11101011000111")
	End if													

	lgIntFlgMode = Parent.OPMD_UMODE									
	
	Call ggoOper.LockField(Document, "Q")								
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor
    
    Set gActiveElement = document.activeElement 

	lgBlnFlgChgValue = False	
	lgMovType = frm1.txtMovType.value
	
End Function


'************************************ Function DbSave() ****************************************
Function DbSave() 

   	Call LayerShowHide(1)  
	
	Dim IntRows 

	Dim strVal
	Dim IRow
	Dim iRowSep, iColSep
	
	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount
	
	iRowSep = Parent.gRowSep
	iColSep = Parent.gColSep

	DbSave = False                                                      
	
	On Error Resume Next                                                
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
	
	With frm1
		.txtMode.value         = Parent.UID_M0002						
		.txtFlgMode.value      = lgIntFlgMode							
		.hYear.value           = .txtYear.Year
	End With
	
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
						
					if UCase(frm1.txtMovType.value) <> "R90" and UCase(frm1.txtMovType.value) <> "R89"  and UCase(frm1.txtMovType.value) <> "RX1" then		  
				 		if  UNICDbl(.Text) = 0 then
							Call DisplayMsgBox("169918","X", "X", "X")
							Call LayerShowHide(0) 
							Call ClickTab2()
							.Action = 0
							exit function
						end if
					End if	
						
					strVal = strVal & Trim(.Text) & iColSep

					.Col = C_EntryAmount	
					if  UCase(frm1.txtMovType.value) = "R91" or UCase(frm1.txtMovType.value) = "R90" or UCase(frm1.txtMovType.value) = "R89" then
						if uniCdbl(.Text) = 0 then
							Call DisplayMsgBox("169939","X", "X", "X")
							Call LayerShowHide(0) 
							Call ClickTab2()
							exit function
						end if	  
					end if

					strVal = strVal & Trim(.Text) & iColSep
				
					.Col = C_EntryUnit		
					strVal = strVal & Trim(.Text) & iRowSep		
					
				Case ggoSpread.DeleteFlag
				
					strVal = "D" & iColSep & IntRows & iColSep			
			
					.Col = C_SeqNo		
					strVal = strVal & Trim(.Text) & iColSep
					    
					.Col = C_SubSeqNo 	
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


'********************************* Function DbSaveOk()	******************************************
Function DbSaveOk()												
   	lgBlnFlgChgValue = false   	
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	Call ClickTab1()
	Call MainQuery()
	
End Function


'******************************** Function FncSplitColumn() **************************************
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 9
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		frm1.vspdData.Col = iColumnLimit:
		frm1.vspdData.Row = 0
       iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    Frm1.vspdData.Action = 0    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function
