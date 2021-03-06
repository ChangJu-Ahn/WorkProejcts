Const BIZ_PGM_QRY_ID    = "i1311mb1_ko441.asp"					
Const BIZ_PGM_DQRY_ID   = "i1311mb1_ko441.asp"					
Const BIZ_PGM_SAVE_ID   = "i1311mb2_ko441.asp"					
Const BIZ_PGM_DEL_ID    = "i1311mb3_ko441.asp"					
Const BIZ_PGM_LOOKUP_ID	= "i1311mb4_ko441.asp"					

Const GoodQty = "양품"							 
Const BadQty  = "불량품"
Const InspQty = "검사품"
Const TrnsQty = "이동품"							

Dim C_ItemCd          									
Dim C_ItemPopup       
Dim C_ItemNm          
Dim C_TrnsItemCd      
Dim C_TrnsItemPopup   
Dim C_EntryQty        
Dim C_EntryUnit       
Dim C_EntryUnitPopup  
Dim C_ItemSpec        
Dim C_InvUnit         
Dim C_TrackingNo      
Dim C_TrackingNoPopup 

'2008-05-19 6:46오후 :: hanc
Dim C_EXT2_CD   				
Dim C_BP_CD    				
Dim C_BP_NM

Dim C_LotNo           
Dim C_LotSubNo        
Dim C_LotNoPopup      
Dim C_TrnsLotNo       
Dim C_TrnsLotSubNo
Dim C_ItemStatus      
Dim C_TrnsPlantCd     
Dim C_TrnsSLCd        
Dim C_TrnsTrackingNo  
Dim C_SeqNo         


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
	C_ItemCd          = 1
	C_ItemPopup       = 2
	C_ItemNm          = 3 
	C_TrnsItemCd      = 4   
	C_TrnsItemPopup   = 5
	C_EntryQty        = 6
	C_EntryUnit       = 7
	C_EntryUnitPopup  = 8 
	C_ItemSpec        = 9
	C_InvUnit         = 10
	C_TrackingNo      = 11
	C_TrackingNoPopup = 12 

'	C_LotNo           = 13
'	C_LotSubNo        = 14
'	C_LotNoPopup      = 15
'	C_TrnsLotNo       = 16
'	C_TrnsLotSubNo    = 17
'	C_ItemStatus      = 18
'	C_TrnsPlantCd     = 19
'	C_TrnsSLCd        = 20
'	C_TrnsTrackingNo  = 21
'	C_SeqNo           = 22

	C_LotNo           = 13
	C_LotSubNo        = 14
	C_LotNoPopup      = 15
	C_TrnsLotNo       = 16
	C_TrnsLotSubNo    = 17
	C_ItemStatus      = 18
	C_EXT2_CD         = 19
	C_BP_CD           = 20
	C_BP_NM           = 21
	C_TrnsPlantCd     = 22
	C_TrnsSLCd        = 23
	C_TrnsTrackingNo  = 24
	C_SeqNo           = 25
End Sub


'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtDocumentDt.text = StartDate
	frm1.txtPostingDt.text  = StartDate
	
	Call ExtractDateFrom(Currentdate, Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtYear.Year	    = strYear
	lgBlnFlgChgValue = False  
	if frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
	End if
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd1.value = UCase(Parent.gPlant)
		frm1.txtPlantNm1.value = Parent.gPlantNm
		frm1.txtDocumentNo1.focus 
	Else
		frm1.txtPlantCd1.focus 
	End If	
	Set gActiveElement = document.activeElement
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
'========================================================================================
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20080429", ,parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_SeqNo + 1	
		.MaxRows = 0
		 
	    Call GetSpreadColumnPos("A")
	    Call AppendNumberPlace("6", "3", "0")
	    
		ggoSpread.SSSetEdit		  C_ItemCd,			"품목",			18, 0, -1, 18, 2		
		ggoSpread.SSSetButton 	  C_ItemPopup, -1	 
		ggoSpread.MakePairsColumn C_ItemCd, C_ItemPopup
		ggoSpread.SSSetEdit		  C_ItemNm,			"품목명",		20, 0, -1, 50		
		
		ggoSpread.SSSetEdit		  C_TrnsItemCd,		"변경품목",		18, 0, -1, 18, 2
		ggoSpread.SSSetButton 	  C_TrnsItemPopup		
		ggoSpread.MakePairsColumn C_TrnsItemCd, C_TrnsItemPopup		
		
		ggoSpread.SSSetFloat      C_EntryQty,		"이동수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetEdit		  C_EntryUnit,		"이동단위",		10, 0, -1, 3, 2
		ggoSpread.SSSetButton 	  C_EntryUnitPopup
		ggoSpread.MakePairsColumn C_EntryUnit, C_EntryUnitPopup
		
		ggoSpread.SSSetEdit		  C_ItemSpec,		"규격",			20, 0, -1, 50		
		ggoSpread.SSSetEdit		  C_InvUnit,		"단위",			10, 0, -1, 3
		ggoSpread.SSSetEdit		  C_TrackingNo,		"Tracking No",	20, 0, -1, 25, 2
		ggoSpread.SSSetButton	  C_TrackingNoPopup
		ggoSpread.MakePairsColumn C_TrackingNo, C_TrackingNoPopup

		ggoSpread.SSSetEdit 	  C_LotNo,			"LOT NO",		20, 0, -1, 25, 2
		ggoSpread.SSSetFloat	  C_LotSubNo,		"순번",			8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetButton 	  C_LotNoPopup
		ggoSpread.MakePairsColumn C_LotNo, C_LotNoPopup
		
		ggoSpread.SSSetEdit 	  C_TrnsLotNo,		"이동Lot No",	20, 0, -1, 25, 2
		ggoSpread.SSSetFloat	  C_TrnsLotSubNo,	"이동순번",		10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"		
		ggoSpread.MakePairsColumn C_TrnsLotNo, C_TrnsLotSubNo

		ggoSpread.SSSetEdit 	  C_ItemStatus,     "품목상태",	20, 0, -1, 25, 2

    	ggoSpread.SSSetEdit 	  C_EXT2_CD,			"STOCK TYPE",		20, 0, -1, 25, 2
		ggoSpread.SSSetEdit 	  C_BP_CD  ,			"공급처",		20, 0, -1, 25, 2
		ggoSpread.SSSetEdit 	  C_BP_NM  ,			"공급처명",		20, 0, -1, 25, 2

		ggoSpread.SSSetEdit 	  C_TrnsPlantCd,    "TrnsPlantCd",		100, 0
		ggoSpread.SSSetEdit 	  C_TrnsSLCd,       "TrnsSLCd",			100, 0
		ggoSpread.SSSetEdit 	  C_TrnsTrackingNo, "TrnsTrackingNo",	1, 0
		ggoSpread.SSSetEdit 	  C_SeqNo,          "순번",			1, 0

		Call ggoSpread.SSSetColHidden(C_ItemStatus, C_ItemStatus, True)
		Call ggoSpread.SSSetColHidden(C_TrnsPlantCd, C_TrnsPlantCd, True)
		Call ggoSpread.SSSetColHidden(C_TrnsSLCd, C_TrnsSLCd, True)
		Call ggoSpread.SSSetColHidden(C_TrnsTrackingNo, C_TrnsTrackingNo, True)
		Call ggoSpread.SSSetColHidden(C_SeqNo, C_SeqNo, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLock -1, -1
		
		.ReDraw = true
		
		ggoSpread.SSSetSplit2(3)
	End With
	
End Sub

'================================ GetSpreadColumnPos() ============================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
		      
			ggoSpread.Source = frm1.vspdData
		      
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
					      
			C_ItemCd          =iCurColumnPos(1)
			C_ItemPopup       =iCurColumnPos(2)
			C_ItemNm          =iCurColumnPos(3)
			C_TrnsItemCd      =iCurColumnPos(4)
			C_TrnsItemPopup   =iCurColumnPos(5)
			C_EntryQty        =iCurColumnPos(6)
			C_EntryUnit       =iCurColumnPos(7)
			C_EntryUnitPopup  =iCurColumnPos(8)
			C_ItemSpec        =iCurColumnPos(9)
			C_InvUnit         =iCurColumnPos(10)
			C_TrackingNo      =iCurColumnPos(11)
			C_TrackingNoPopup =iCurColumnPos(12)
		
'			C_LotNo           =iCurColumnPos(13)
'			C_LotSubNo        =iCurColumnPos(14)
'			C_LotNoPopup      =iCurColumnPos(15)
'			C_TrnsLotNo       =iCurColumnPos(16)
'			C_TrnsLotSubNo    =iCurColumnPos(17)
'			C_ItemStatus      =iCurColumnPos(18)
'			C_TrnsPlantCd     =iCurColumnPos(19)
'			C_TrnsSLCd        =iCurColumnPos(20)
'			C_TrnsTrackingNo  =iCurColumnPos(21)
'			C_SeqNo           =iCurColumnPos(22)
		      
			C_LotNo           =iCurColumnPos(13)
			C_LotSubNo        =iCurColumnPos(14)
			C_LotNoPopup      =iCurColumnPos(15)
			C_TrnsLotNo       =iCurColumnPos(16)
			C_TrnsLotSubNo    =iCurColumnPos(17)
			C_ItemStatus      =iCurColumnPos(18)
			C_EXT2_CD         =iCurColumnPos(19)
			C_BP_CD           =iCurColumnPos(20)
			C_BP_NM           =iCurColumnPos(21)
			C_TrnsPlantCd     =iCurColumnPos(22)
			C_TrnsSLCd        =iCurColumnPos(23)
			C_TrnsTrackingNo  =iCurColumnPos(24)
			C_SeqNo           =iCurColumnPos(25)
		      
	End Select
End Sub

'===================================== PopSaveSpreadColumnInf() ==================================
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
    
	If UCase(frm1.hGuiControlFlag3.value) = "Y" Then
		ggoSpread.SSSetRequired	 C_TrnsItemCd,    pvStartRow, pvEndRow
	Else
	    ggoSpread.SSSetProtected C_TrnsItemCd,    pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_TrnsItemPopup, pvStartRow, pvEndRow
	End if

	ggoSpread.SSSetRequired	 C_ItemCd,    pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm,    pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemSpec,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_InvUnit,   pvStartRow, pvEndRow
	ggoSpread.SSSetRequired	 C_EntryQty,  pvStartRow, pvEndRow
	ggoSpread.SSSetRequired	 C_EntryUnit, pvStartRow, pvEndRow

    '2008-05-23 11:54오전 :: hanc
	ggoSpread.SSSetProtected	 C_EXT2_CD, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	 C_BP_CD, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	 C_BP_NM, pvStartRow, pvEndRow


End Sub

'================================== 2.2.5 SheetFocus() ==================================================
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


'==========================================  2.3.1 Tab Click 처리  =================================================
Function ClickTab1()
	
    Call changeTabs(1)	

	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101000000011")		    			
	ElseIf lgIntFlgMode = Parent.OPMD_UMODE Then
	    Call SetToolBar("11101000000111")						
	End If	
	
	If frm1.txtPlantCd1.value <> "" Then
		frm1.txtDocumentNo1.focus
	Else
		frm1.txtPlantCd1.focus
	End If 
	
End Function


Function ClickTab2()
	Dim strVal
	ClickTab2	=	True

	If Trim(frm1.txtPlantCd1.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd1.Focus
		Set gActiveElement = document.activeElement
		ClickTab2	= False
		Exit Function
	End If
	
	if (frm1.txtMovType.value = "" and lgIntFlgMode = Parent.OPMD_CMODE ) then
	    Call DisplayMsgBox("169904","X", "X", "X")    
	    frm1.txtMovType.Focus
	    Set gActiveElement = document.activeElement
	    ClickTab2	= False
	    Exit function
	End if	

    Call changeTabs(2)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")						
	Else
	    Call SetToolBar("11101011000111")						
	End If		

	Call LayerShowHide(1)  

	strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtMovType=" & Trim(frm1.txtMovType.value)    	
	
	Call RunMyBizASP(MyBizASP, strVal)
	Call LayerShowHide(0)  


	If frm1.txtMovType.value <> lgMovType Then
		If lgIntFlgMode = Parent.OPMD_CMODE Then 
			ggoSpread.Source = frm1.vspdData
			ggoSpread.ClearSpreadData
			frm1.txtPlantCd2.value = ""
			frm1.txtPlantNm2.value = ""
			frm1.txtTrackingNo.value = ""
			frm1.txtSLCd2.value = ""
			frm1.txtSLNm2.value = ""
			frm1.txtCostCd2.value = ""
			frm1.txtCostNm2.value = ""
		End If
		lgMovType = frm1.txtMovType.value	
	End if		
	
	
End Function


'20080306::hanc --------------------------------------------------------------------------------------
'ClickTab2_1 만든이유 : 자재불출의뢰정보 참조팝업 시 수불유형 선택하지 않고 팝업창 띄우기 위함.
'-----------------------------------------------------------------------------------------------------
Function ClickTab2_1()
	Dim strVal
	ClickTab2_1	=	True

	If Trim(frm1.txtPlantCd1.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd1.Focus
		Set gActiveElement = document.activeElement
		ClickTab2_1	= False
		Exit Function
	End If
	
	if (frm1.txtMovType.value = "" and lgIntFlgMode = Parent.OPMD_CMODE ) then
	    Call DisplayMsgBox("169904","X", "X", "X")    
	    frm1.txtMovType.Focus
	    Set gActiveElement = document.activeElement
	    ClickTab2_1	= False
	    Exit function
	End if	
    Call changeTabs(2)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
	    Call SetToolBar("11101101001011")						
	Else
	    Call SetToolBar("11101011000111")						
	End If		

	Call LayerShowHide(1)  

	strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtMovType=" & Trim(frm1.txtMovType.value)    	
	
	Call RunMyBizASP(MyBizASP, strVal)
	Call LayerShowHide(0)  


	If frm1.txtMovType.value <> lgMovType Then
		If lgIntFlgMode = Parent.OPMD_CMODE Then 
			ggoSpread.Source = frm1.vspdData
			ggoSpread.ClearSpreadData
			frm1.txtPlantCd2.value = ""
			frm1.txtPlantNm2.value = ""
			frm1.txtTrackingNo.value = ""
			frm1.txtSLCd2.value = ""
			frm1.txtSLNm2.value = ""
			frm1.txtCostCd2.value = ""
			frm1.txtCostNm2.value = ""
		End If
		lgMovType = frm1.txtMovType.value	
	End if		
	
	
End Function

'------------------------------------------  OpenPlant1()  -------------------------------------------------
Function OpenPlant1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd1.Value)
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
		frm1.txtPlantCd1.focus
		Exit Function
	Else
		Call SetPlant1(arrRet)
	End If	
	
End Function

' '------------------------------------------  OpenPlant2()  -------------------------------------------------
Function OpenPlant2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

    If UCase(frm1.txtPlantCd2.ClassName) = UCase(Parent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd2.Value)
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
		frm1.txtPlantCd2.focus
		Exit Function
	Else
		Call SetPlant2(arrRet)
	End If	
	
End Function

' '------------------------------------------  OpenDocumentNo1()  --------------------------------------------
Function OpenDocumentNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1, Param2, Param3, Param4
	
	If Trim(frm1.txtPlantCd1.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")    
		frm1.txtPlantCd1.Focus
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Param1 = Trim(frm1.txtDocumentNo1.Value)
	Param2 = Trim(frm1.txtYear.Text)
	Param3 = "ST"
	Param4 = Trim(frm1.txtPlantCd1.Value)

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

' '------------------------------------------  OpenSL1()  -------------------------------------------------
Function OpenSL1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd1.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd1.Focus
		Exit Function	
	End if

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtSLCd1.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd1.Value)
	arrParam(3) = ""	
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd1.Value, "''", "S")		
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd1.focus
		Exit Function
	Else
		Call SetSL1(arrRet)
	End If	
	
End Function

' '------------------------------------------  OpenSL2()  -------------------------------------------------
Function OpenSL2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd2.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd2.Focus
		Exit Function
	End if

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtSLCd2.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd2.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd2.Value, "''", "S")		
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd2.focus
		Exit Function
	Else
		Call SetSL2(arrRet)
	End If	
	
End Function

' '------------------------------------------  OpenMovType()  -------------------------------------------------
Function OpenMovType()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1, Param2, Param3
	
	If IsOpenPop = True Then Exit Function

	If UCase(frm1.txtMovType.ClassName) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True

	Param1 = Trim(frm1.txtMovType.Value)
	Param2 = Trim(frm1.txtMovTypeNm.Value)
	Param3 = "ST"

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


' '------------------------------------------  OpenItem1()  -------------------------------------------------
Function OpenItem1(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5),arrField(14)
	
	If Trim(frm1.txtPlantCd1.Value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169901","X", "X", "X")   
		Call SetFocusToDocument("M")  
		frm1.txtPlantCd1.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If Trim(frm1.txtSlCd1.Value) = "" then
		Call ClickTab1() 
		Call DisplayMsgBox("169902", "X", "X", "X")   
		Call SetFocusToDocument("M")  
		frm1.txtSlCd1.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd1.value)	
	arrParam(1) = Trim(frm1.txtPlantNm1.value)
	arrParam(2) = Trim(frm1.txtSLCd1.value)
	arrParam(3) = Trim(frm1.txtSLNm1.value)
	arrParam(4) = strCode
	arrParam(5) = ""	
	
	arrField(0) = 1 'ITEM_CD					' Field명(0)
	arrField(1) = 2 'ITEM_NM					' Field명(1)
	arrField(2) = 3	'SPECIFICATION	
	arrField(3) = 4
	arrField(4) = 5
	arrField(5)	= 6
	arrField(6) = 7
	
	arrField(7) = 8
	arrField(8) = 9
	arrField(9) = 10
	arrField(10) = 11
	arrField(11) = 12
	arrField(12) = 13
	arrField(13) = 14
	arrField(14) = 15
	
	iCalledAspName = AskPRAspName("I1211PA1_ko441")     '2008-06-12 6:35오후 :: hanc
'	iCalledAspName = AskPRAspName("I1211PA1")     '2008-06-12 6:35오후 :: hanc
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
		Call SetItem1(arrRet)
	End If	

End Function

' '------------------------------------------  OpenItem2()  -------------------------------------------------
Function OpenItem2(Byval strCode)
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd1.Value) = "" then
		Call ClickTab1() 
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd1.Focus
	    Set gActiveElement = document.activeElement
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd1.value)	
	arrParam(1) = Trim(strCode)	                
	arrParam(2) = ""							
	arrParam(3) = ""							
	
	arrField(0) = 1 
    arrField(1) = 2 
    arrField(2) = 10
    arrField(3) = 6 

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
		Call SetActiveCell(frm1.vspdData,C_TrnsItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetItem2(arrRet)
	End If	

End Function

' '------------------------------------------  OpenLotNo()  -------------------------------------------------
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

	Param1 = Trim(frm1.txtSLCd1.value)
	Param4 = Trim(frm1.txtPlantCd1.value)
	Param5 = "J"				 
	if Param1 = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169902","X", "X", "X")    
		frm1.txtSLCd1.Focus
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
			 Param3 = Trim(.Text)
			.Col = C_LotNo
			.Row = .ActiveRow 
			 Param6 = Trim(.Text)
			.Col = C_EntryUnit
			.Row = .ActiveRow
			 Param9 = Trim(.Text)
		End If	
    	
		if Param2 = "" then
			Call DisplayMsgBox("169903","X", "X", "X")    
			Call SetActiveCell(frm1.vspdData,C_ItemCd,.ActiveRow,"M","X","X")
			Exit Function    	
		End If
    End With
	
	if Param3 = "" Then
	    Param3 = "*"
	End if		
		
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("I2212RA1_ko441")     '2008-05-19 3:50오후 :: hanc
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2212RA1","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6,Param7,Param8,Param9), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_LotNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
    	Call SetLotNo(arrRet)
	End If	
End Function

' '------------------------------------------  OpenTrackingNo1()  -------------------------------------------------
Function OpenTrackingNo1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    
	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.ClassName)= UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "TRACKINGNO"	
	arrParam(1) = "s_so_tracking"				
	
	arrParam(2) = Trim(frm1.txtTrackingNo.Value)
	
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
		frm1.txtTrackingNo.Value = arrRet(0)
	End If	
End Function

'------------------------------------------  OpenTrackingNo2()  -------------------------------------------------
Function OpenTrackingNo2(Byval strCode)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

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
		Call SetActiveCell(frm1.vspdData,C_TrnsTrackingNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If	
End Function

' '------------------------------------------  OpenEntryUnit()  -------------------------------------------------
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

' '------------------------------------------  OpenPopupGL()  -------------------------------------------------
Function OpenPopupGL()
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
	
End Function

'------------------------------------------  OpenPopupGL2()  -------------------------------------------------
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
   
End Function

'20080226::hanc***********************************************************
Function OpenMoveInvRef1()
	Dim iCalledAspName
	Dim IntRetCD

	If GetSetupMod(Parent.gSetupMod, "p") <> "Y" Then
    	Call DisplayMsgBox("169936","X", "X", "X")
		Exit Function
	End if
				
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6 

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd1.value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd1.Focus
	    Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	Param1 = Trim(frm1.txtPlantCd1.value)
	Param2 = Trim(frm1.txtPlantNm1.value)
'20080312::hanc	If Trim(frm1.txtSLCd1.value) = "" then
'20080312::hanc		Call ClickTab1()   
'20080312::hanc		Call DisplayMsgBox("169902","X","X","X")
'20080312::hanc		frm1.txtSLCd1.Focus
'20080312::hanc	    Set gActiveElement = document.activeElement
'20080312::hanc		Exit Function
'20080312::hanc	End If
	
	Param3 = "ST"
	Param4 = Trim(frm1.txtMovType.value)
	
	
'20080312::hanc	If Trim(frm1.txtSLCd2.value) = "" then
'20080312::hanc		
'20080312::hanc		If ClickTab2_1() Then                       '20080306::hanc::ClickTab2_1 만든이유 : 자재불출의뢰정보 참조팝업 시 수불유형 선택하지 않고 팝업창 띄우기 위함.
'20080312::hanc			Call DisplayMsgBox("169937","X","X","X")
'20080312::hanc			'frm1.txtSLCd2.Focus
'20080312::hanc			'Set gActiveElement = document.activeElement
'20080312::hanc		End if
'20080312::hanc	    Exit Function
'20080312::hanc	End If
	
	Param5 = Trim(frm1.txtSLCd2.value)
	Param6 = Trim(frm1.txtSLNm2.value)
	
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
			.Col = C_TrnsLotSubNo   :		.text = arrRet(iLngCnt, 5)	     
			.Col = C_LotNo      :		.text = "*"			        
			.Col = C_TrnsLotNo  :		.text = "*"			
			frm1.txtDocumentText.value  =         arrRet(iLngCnt, 6)
			frm1.txtMovType.value  =         arrRet(iLngCnt, 7)

            '2008-04-23 1:31오후 :: hanc
			ggoSpread.SpreadUnLock	 C_LotNo,	    .Row,	C_LotNo,		    .Row
			ggoSpread.SpreadUnLock	 C_LotNoPopup,	    .Row,	C_LotNoPopup,		    .Row
			ggoSpread.SpreadUnLock	 C_TrnsLotNo,	.Row,	C_TrnsLotNo,		.Row
			ggoSpread.SpreadUnLock	 C_EntryQty,	.Row,	C_EntryQty,		.Row   '2008-04-29 6:00오후 :: hanc
			

            '2008-04-23 1:18오후 :: hanc
'			Call .SetText(C_LotSubNo,       .Row, 0)
'			Call .SetText(C_LotNo,          .Row, "*")
'			Call .SetText(C_TrnsLotNo,      .Row, "*")
'			Call .SetText(C_TrnsLotSubNo,   .Row, 0)

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


 '------------------------------------------  OpenMoveInvRef()  -------------------------------------------------
Function OpenMoveInvRef()
	Dim iCalledAspName
	Dim IntRetCD

	If GetSetupMod(Parent.gSetupMod, "p") <> "Y" Then
    	Call DisplayMsgBox("169936","X", "X", "X")
		Exit Function
	End if
				
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6 

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd1.value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd1.Focus
	    Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	Param1 = Trim(frm1.txtPlantCd1.value)
	Param2 = Trim(frm1.txtPlantNm1.value)
	If Trim(frm1.txtSLCd1.value) = "" then
		Call ClickTab1()   
		Call DisplayMsgBox("169902","X","X","X")
		frm1.txtSLCd1.Focus
	    Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	Param3 = Trim(frm1.txtSlCd1.value)
	Param4 = Trim(frm1.txtSLNm1.value)
	
	If Trim(frm1.txtSLCd2.value) = "" then
		
		If ClickTab2() Then
			Call DisplayMsgBox("169937","X","X","X")
			'frm1.txtSLCd2.Focus
			'Set gActiveElement = document.activeElement
		End if
	    Exit Function
	End If
	
	Param5 = Trim(frm1.txtSLCd2.value)
	Param6 = Trim(frm1.txtSLNm2.value)
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I1311RA1")
	
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
		Call SetMoveInvRef(arrRet)
	End If
	
End Function

'------------------------------------------  OpenCostCd1()  -------------------------------------------------
  Function OpenCostCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtCostCd1.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	If Trim(frm1.txtPlantCd1.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd1.Focus
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "Cost Center 팝업"			
	arrParam(1) = "B_COST_CENTER A,B_PLANT B"
	arrParam(2) = Trim(frm1.txtCostCd1.Value)		
	arrParam(3) = ""								
	arrParam(4) = "A.BIZ_AREA_CD = B.BIZ_AREA_CD AND B.PLANT_CD =" & FilterVar(frm1.txtPlantCd1.Value, "''", "S")
	arrParam(5) = "Cost Center"					
	
	arrField(0) = "COST_CD"							
	arrField(1) = "COST_NM"							
    
	arrHeader(0) = "Cost Center"			    	
	arrHeader(1) = "Cost Center 명"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd1.focus
		Exit Function
	Else
		Call SetCostCd1(arrRet)
	End If	
End Function

'------------------------------------------  OpenCostCd2()  -------------------------------------------------
  Function OpenCostCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtCostCd2.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
	
	If Trim(frm1.txtPlantCd2.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")  
	    frm1.txtPlantCd2.Focus
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "Cost Center 팝업"			
	arrParam(1) = "B_COST_CENTER A,B_PLANT B"
	arrParam(2) = Trim(frm1.txtCostCd2.Value)		
	arrParam(3) = ""								
	arrParam(4) = "A.BIZ_AREA_CD = B.BIZ_AREA_CD AND B.PLANT_CD =" & FilterVar(frm1.txtPlantCd2.Value, "''", "S")
	arrParam(5) = "Cost Center"					
	
	arrField(0) = "COST_CD"							
	arrField(1) = "COST_NM"							
    
	arrHeader(0) = "Cost Center"			    	
	arrHeader(1) = "Cost Center 명"				

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtCostCd2.focus
		Exit Function
	Else
		Call SetCostCd2(arrRet)
	End If	
    
End Function

' '------------------------------------------  SetPlant1()  --------------------------------------------------
Function SetPlant1(byRef arrRet)
	frm1.txtPlantCd1.Value    = arrRet(0)		
	frm1.txtPlantNm1.Value    = arrRet(1)
	frm1.txtPlantCd1.focus	
End Function

' '------------------------------------------  SetPlant2()  --------------------------------------------------
Function SetPlant2(byRef arrRet)

	if (arrRet(0) = frm1.txtPlantCd1.value) then 	
	 Call DisplayMsgBox("169905","X", "X", "X")
	 frm1.txtPlantCd2.focus    
	 Exit Function
	end if

	frm1.txtPlantCd2.Value    = arrRet(0)		
	frm1.txtPlantNm2.Value    = arrRet(1)
	frm1.txtPlantCd2.focus		
End Function

' '------------------------------------------  SetDocumentNo()  --------------------------------------------------
Function SetDocumentNo(byRef arrRet)
	frm1.txtDocumentNo1.Value    = arrRet(0)
	frm1.txtYear.Year            = arrRet(1)
	
	frm1.txtDocumentNo1.focus
End Function

' '------------------------------------------  SetSL1()  --------------------------------------------------
Function SetSL1(byRef arrRet)
	frm1.txtSLCd1.Value    = arrRet(0)		
	frm1.txtSLNm1.Value    = arrRet(1)
	frm1.txtSLCd1.focus
	lgBlnFlgChgValue = True
End Function

' '------------------------------------------  SetSL2()  --------------------------------------------------
Function SetSL2(byRef arrRet)
	if (arrRet(0) = frm1.txtSLCd1.value) then
	Call DisplayMsgBox("169906","X", "X", "X")
	frm1.txtSLCd2.focus   
	Exit Function
	end if
	frm1.txtSLCd2.Value    = arrRet(0)		
	frm1.txtSLNm2.Value    = arrRet(1)
	frm1.txtSLCd2.focus
	lgBlnFlgChgValue = True
End Function

' '------------------------------------------  SetMovType()  --------------------------------------------------
Function SetMovType(byRef arrRet)
	frm1.txtMovType.Value      = arrRet(0)
	frm1.txtMovTypeNm.Value    = arrRet(1)
	frm1.txtMovType.focus
	lgBlnFlgChgValue = True
End Function

' '------------------------------------------  SetItem1()  --------------------------------------------------
Function SetItem1(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_ItemCd,		.ActiveRow, arrRet(0))
		Call .SetText(C_ItemNm,		.ActiveRow, arrRet(1))
		Call .SetText(C_ItemSpec,	.ActiveRow, arrRet(2))
		Call .SetText(C_EntryUnit,	.ActiveRow, arrRet(3))
		Call .SetText(C_InvUnit,    .ActiveRow, arrRet(3))
		Call .SetText(C_TrackingNo,	.ActiveRow, arrRet(4))
		Call .SetText(C_LotNo,		.ActiveRow, arrRet(5))
		Call .SetText(C_LotSubNo,	.ActiveRow, arrRet(6))

		Call .SetText(C_EXT2_CD,	.ActiveRow, arrRet(12))
		Call .SetText(C_BP_CD,	.ActiveRow, arrRet(13))
		Call .SetText(C_BP_NM,	.ActiveRow, arrRet(14))
		
		Call vspdData_Change(C_ItemCd, .ActiveRow)		 
		Call SetActiveCell(frm1.vspdData,C_EntryQty,.ActiveRow,"M","X","X")
	End With
End Function

' '------------------------------------------  SetItem2()  --------------------------------------------------
Function SetItem2(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrnsItemCd,.ActiveRow, arrRet(0))
		Call vspdData_Change(C_TrnsItemCd, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_EntryQty,.ActiveRow,"M","X","X")		
	End With
End Function


' '------------------------------------------  SetLotNo()  --------------------------------------------------
Function SetLotNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(2))
		Call .SetText(C_LotNo,		.ActiveRow, arrRet(3))
		Call .SetText(C_LotSubNo,	.ActiveRow, arrRet(4))
		Call .SetText(C_EXT2_CD,	.ActiveRow, arrRet(5))
		Call .SetText(C_BP_CD,	.ActiveRow, arrRet(6))
		Call .SetText(C_BP_NM,	.ActiveRow, arrRet(7))
		Call vspdData_Change(C_LotNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_LotNo,.ActiveRow,"M","X","X")		 
	End With
End Function


'------------------------------------------  SetTrackingNo()  --------------------------------------------------
Function SetTrackingNo(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_TrackingNo, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_TrackingNo, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_LotNo,.ActiveRow,"M","X","X")		 
	End With
End Function

' '------------------------------------------  SetEntryUnit()  --------------------------------------------------
Function SetEntryUnit(byRef arrRet)
	With frm1.vspdData
		Call .SetText(C_EntryUnit, .ActiveRow, arrRet(0))
		Call vspdData_Change(C_EntryUnit, .ActiveRow)
		Call SetActiveCell(frm1.vspdData,C_TrackingNo,.ActiveRow,"M","X","X")		
	End With
End Function

'------------------------------------------  SetCostCd()  --------------------------------------------------
Function SetCostCd1(byRef arrRet)
	frm1.txtCostCd1.value = arrRet(0)
	frm1.txtCostNm1.value = arrRet(1)
	frm1.txtCostCd1.focus
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCostCd()  --------------------------------------------------
Function SetCostCd2(byRef arrRet)
	frm1.txtCostCd2.value = arrRet(0)
	frm1.txtCostNm2.value = arrRet(1)
	frm1.txtCostCd2.focus
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'   Event Name : txtDocumentDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt_Change()
'=======================================================================================================
Sub txtDocumentDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPostingDt_DblClick(Button)
'=======================================================================================================
Sub txtPostingDt_DblClick(Button) 
    If Button = 1 Then
        frm1.txtPostingDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtPostingDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPostingDt_Change()
'=======================================================================================================
Sub txtPostingDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'=======================================================================================================
Sub txtYear_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYear.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtYear.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'=======================================================================================================
Sub  txtYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtSLCd2_onChange()
'=======================================================================================================
Sub txtSLCd2_onChange()
	if frm1.txtSLCd2.value = frm1.txtSLCd1.value then
		Call DisplayMsgBox("169906","X", "X", "X")    
		frm1.txtSLCd2.value = ""
		frm1.txtSLNm2.value = ""
		Call SetFocusToDocument("M")  
		frm1.txtSLCd2.focus
		Set gActiveElement = document.activeElement 
	end if
	
End Sub

'=======================================================================================================
'   Event Name : txtItemCd2_onChange()
'=======================================================================================================
Sub txtItemCd2_onChange()
	if frm1.txtItemCd2.value = frm1.txtItemCd.value then
		Call DisplayMsgBox("169917","X", "X", "X")  
		frm1.txtItemCd2.value = ""
		frm1.txtItemNm2.value = ""
		Call SetFocusToDocument("M")  
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement 		 
	end if
End Sub

'=======================================================================================================
'   Event Name : txtPlantCd2_onChange()
'=======================================================================================================
Sub txtPlantCd2_onChange()
	If Trim(frm1.hGuiControlFlag2.value) = "Y" Then
		if (Trim(UCase(frm1.txtPlantCd2.value)) = Trim(UCase(frm1.txtPlantCd1.value))) then 	
			Call DisplayMsgBox("169905","X", "X", "X")  
			frm1.txtPlantCd2.value = ""
			frm1.txtPlantNm2.value = ""
			Call SetFocusToDocument("M")  
			frm1.txtPlantCd2.focus
			Set gActiveElement = document.activeElement 		
			Exit Sub
		end if
	End if
End Sub

'==========================================================================================
'   Event Name :vspddata_EditChange
'==========================================================================================
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

'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	With Frm1.vspdData
		Select Case Col
		Case C_ItemCd
			.Col = Col
			.Row = Row		
			
			If 	CommonQueryRs(" A.item_nm, A.spec, A.basic_unit ", " B_ITEM A, B_ITEM_BY_PLANT B ", _
			    " A.item_cd = B.item_cd AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd1.Value, "''", "S") & " AND A.item_cd = " & FilterVar(Frm1.vspdData.Text, "''", "S"), _
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

'==========================================================================================
'   Event Name : vspdData_Click
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("0101111111")
	End If	

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 then
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

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    


Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_ItemPopUp Then
			.Col = C_ItemCd
			.Row = Row
			Call OpenItem1(.Text)

		Elseif Row > 0 And Col = C_TrnsItemPopup Then
			.Col = C_TrnsItemCd
			.Row = Row
			Call OpenItem2(.Text)

		Elseif Row > 0 And Col = C_TrackingNoPopUp Then
			.Col = C_TrackingNo
			.Row = Row
			Call OpenTrackingNo2(.Text)
			
		Elseif Row > 0 And Col = C_EntryUnitPopUp Then
			.Col = C_EntryUnit
			.Row = Row
			Call OpenEntryUnit(.Text)
			
		Elseif Row > 0 And Col = C_LotNoPopUp Then
			.Col = C_LotNo
			.Row = Row			
			Call OpenLotNo(.Text)
		End If
	
	End With
End Sub


'========================================================================================
' Function Name : FncQuery
'========================================================================================
 Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                  
    
    Err.Clear                                                         
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then Exit Function
	End If

	If Not chkFieldByCell(frm1.txtPlantCd1, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtDocumentNo1, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtYear, "A",1) Then Exit Function

	Call ggoOper.ClearField(Document, "2")										
	Call ggoOper.LockField(Document, "N")  
	                                     
 	frm1.txtDocumentDt.text = StartDate
	frm1.txtPostingDt.text  = StartDate
	Call InitVariables
   
    If 	CommonQueryRs(" A.MINOR_TYPE "," B_MINOR A, I_GOODS_MOVEMENT_HEADER  B ", _
					" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("ST", "''", "S") & " AND B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		If 	CommonQueryRs(" B.ITEM_DOCUMENT_NO "," I_GOODS_MOVEMENT_HEADER  B ", _
						"B.ITEM_DOCUMENT_NO = " & Trim(FilterVar(frm1.txtDocumentNo1.value," ","S")), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
         	Call DisplayMsgBox("160101","X","X","X")
		Else
			Call DisplayMsgBox("169956","X","X","X")
		End if
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement		
		Exit function
    End If
    lgF0 = Split(lgF0,Chr(11))
    if	Trim(lgF0(0)) <> "U" then
		Call DisplayMsgBox("169956","X","X","X")
		frm1.txtDocumentNo1.Focus
		Set gActiveElement = document.activeElement
		Exit Function
    end if
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call ClickTab1()
   
    If DbQuery() = False Then Exit Function

    FncQuery = True															
   
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                               
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")      
		If IntRetCD = vbNo Then Exit Function
	End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    ggoSpread.ClearSpreadData
    
    Call changeTabs(1)	
	
    Call ggoOper.ClearField(Document, "A") 
    Call ggoOper.LockField(Document, "N")  
    Call InitVariables                     
    Call SetDefaultVal
    Call SetToolBar("1110100000011")		
    
    FncNew = True                           

End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
    
    FncDelete = False                                          
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                    
        Call DisplayMsgBox("900002","X", "X", "X")               
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete() = False Then Exit Function
    
    FncDelete = True                                           

End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================

Function FncSave() 
	Dim IntRetCD 

	FncSave = False                                            
		
	Err.Clear                                                  
	On Error Resume Next                                       
	
	'-----------------------
	'Precheck area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If lgBlnFlgChgValue = False  And ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001","X", "X", "X")                    
		Exit Function
	End If
	
	'-----------------------
	'Check content area
	'-----------------------	
    If Not chkField(Document, "2")  Then    
 	   If gPageNo = 1 Then	              
          Call SetToolBar("11101000000011")
	   Else	             
          Call SetToolBar("11101101000011")
       End If
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function
    	
	If frm1.vspdData.MaxRows <= 0  Then
		IntRetCD = DisplayMsgBox("122610","X", "X", "X")
		Exit Function
	End If
	
	'-----------------------
	'Check MINOR CODE
	'-----------------------
    If 	CommonQueryRs(" A.MINOR_TYPE "," B_MINOR A, I_MOVETYPE_CONFIGURATION B ", _
		" A.MINOR_CD = B.MOV_TYPE AND A.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " and B.TRNS_TYPE =" & FilterVar("ST", "''", "S") & " AND A.MINOR_CD = " & Trim(FilterVar(frm1.txtMovType.value," ","S")), _
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

	'-----------------------
	'Save function call area
	'-----------------------
	Dim strYear2
	Dim strMonth2
	Dim strDay2

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

	'-----------------------
	'Check DocumentDt 
	'-----------------------
    Dim strCurrDt
    Dim strDocumentDt
    Dim strPostingDt
    
    strDocumentDt  = UniConvDateToYYYYMMDD(frm1.txtDocumentDt.text,Parent.gDateFormat,"")
    strCurrDt      = UniConvDateToYYYYMMDD(startdate, parent.gDateFormat,"")
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

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 
    With frm1.vspdData
		If .maxrows < 1 then exit function

		.ReDraw = False
		If gPageNo = 2 Then
			
		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.CopyRow
		    SetSpreadColor .ActiveRow, .ActiveRow
		    
		ElseIf gPageNo = 1 Then 
		    Call ggoOper.ClearField(Document, "1")                      
		End If 
		.ReDraw = True

    End With
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPaste
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel()
    If frm1.vspdData.maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                            
End Function

'========================================================================================
' Function Name : FncInsertRow
'========================================================================================
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
			Call .SetText(C_TrnsLotNo,     intRow, "*")
			Call .SetText(C_TrnsLotSubNo,     intRow, 0)
   		Next
		.ReDraw = True
    End With	

    Set gActiveElement = document.activeElement

    If Err.number = 0 Then
		FncInsertRow = True
	End If

End Function

'========================================================================================
' Function Name : FncDeleteRow
'========================================================================================
Function FncDeleteRow() 
	
	Dim lDelRows 
	Dim lTempRows 
	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncExcel
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)								
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , False)                                       
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 10
    
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

'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'========================================================================================
' Function Name : DbDelete
'========================================================================================
Function DbDelete() 

    DbDelete = False											
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & Parent.UID_M0003
    strVal = strVal         & "&txtApNo=" & Trim(frm1.txtApNo.value)	
    
	Call RunMyBizASP(MyBizASP, strVal)									
    
    DbDelete = True                                                     

End Function

'========================================================================================
' Function Name : DbDeleteOk
'========================================================================================
Function DbDeleteOk()											
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    
    Call LayerShowHide(1)  
    DbQuery = False                                      
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID &	"?txtMode="        & Parent.UID_M0001					& _				
								"&txtPlantCd="     & Trim(frm1.txtPlantCd1.value)		& _	
								"&txtDocumentNo1=" & Trim(frm1.txtDocumentNo1.value)	& _
								"&txtYear="        & Trim(frm1.txtYear.Year)			& _			
								"&lgStrPrevKey="   & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbQuery = True                                                      
    
End Function

'========================================================================================
' Function Name : MovTypeDbQueryOk()
'========================================================================================
Function MovTypeDbQueryOk()												
    if lgIntFlgMode = Parent.OPMD_UMODE Then		
        Call ggoOper.LockField(Document, "Q")        
        Call SetToolBar("11101011000111")			
    end if   
    Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : DbQueryOk 
'========================================================================================
Function DbQueryOk()										
    '-----------------------
    'Reset variables area
    '-----------------------  
    If gPageNo = 1 Then
		Call SetToolBar("11101000000111")				
	Else
		Call SetToolBar("11101011000111")
	End If
		
    lgIntFlgMode = Parent.OPMD_UMODE					
	
    Call ggoOper.LockField(Document, "Q")    
	lgBlnFlgChgValue = False
   
End Function

'========================================================================================
' Function Name : OpenSubCtctRef
'========================================================================================
Function OpenSubCtctRef()
		Dim arrRet
		Dim Param1 
		Dim Param2 
		Dim Param3 
		Dim Param4
		Dim Param5
		Dim Param6
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call DisplayMsgBox("K21022","X","X","X")
		Exit Function
	End If

	If Trim(frm1.txtPlantCd1.value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169901","X","X","X")  
		frm1.txtPlantCd1.focus
		Exit Function
	End If
	
	Param1 = Trim(frm1.txtPlantCd1.value)
	Param2 = Trim(frm1.txtPlantNm1.value)
	
	If Trim(frm1.txtSLCd1.value) = "" then
		Call ClickTab1()
		Call DisplayMsgBox("169902","X","X","X")  
		frm1.txtSLCd1.focus
		Exit Function
	End If
	
	Param3 = Trim(frm1.txtSlCd1.value)
	Param4 = Trim(frm1.txtSLNm1.value)
	
	If Trim(frm1.txtSLCd2.value) = "" then
		If  ClickTab2()	then 
			Call DisplayMsgBox("169937","X","X","X")  
			'frm1.txtSLCd2.focus
			'Set gActiveElement = document.activeElement
		End if
		Exit Function	
	End If
	
	Param5 = Trim(frm1.txtSLCd2.value)
	Param6 = Trim(frm1.txtSLNm2.value)
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("I1321RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1321RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	
	If arrRet(0,0) = "" Then
		frm1.txtSLCd2.focus 
		Exit Function
	Else
		Call SetSubCtctRef(arrRet)
	End If
End Function


Function SetSubCtctRef(arrRet)
		
	Dim TempRow
	Dim intLoopCnt
	Dim intCnt
	Dim iRow
		
	Call ClickTab2()
	
	With frm1.vspdData
			.focus
			ggoSpread.Source = frm1.vspdData			
			
			.ReDraw = False	
			intLoopCnt = Ubound(arrRet, 1)									
			
			TempRow = .MaxRows

			.MaxRows = TempRow + intLoopCnt + 1

			ggoSpread.SpreadUnLock	 C_ItemCd,		TempRow + 1,	C_ItemPopup,		.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryQty,	TempRow + 1,	C_EntryQty,			.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryUnit,	TempRow + 1,	C_EntryUnitPopup,	.MaxRows
			ggoSpread.SpreadUnLock	 C_TrackingNo,	TempRow + 1,	C_TrnsLotSubNo,		.MaxRows
			ggoSpread.SSSetRequired  C_ItemCd,		TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryQty,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryUnit,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemCd,  TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemPopup, TempRow + 1,	.MaxRows
			
			For intCnt = 0 to intLoopCnt
				
				iRow = TempRow + intCnt + 1
					
 				Call .SetText(0, iRow, ggoSpread.InsertFlag)
				Call .SetText(C_ItemCd,	iRow, arrRet(intCnt, 0))
				Call .SetText(C_ItemNm, iRow, arrRet(intCnt, 1))
				Call .SetText(C_TrackingNo,iRow, arrRet(intCnt, 2))
				Call .SetText(C_EntryQty,  iRow, arrRet(intCnt, 3))
				Call .SetText(C_EntryUnit, iRow, arrRet(intCnt, 4))
				Call .SetText(C_InvUnit,   iRow, arrRet(intCnt, 9))
				Call .SetText(C_ItemSpec,  iRow, arrRet(intCnt, 10))
				Call .SetText(C_LotSubNo,  iRow, "0")
				Call .SetText(C_LotNo,     iRow, "*")
				Call .SetText(C_TrnsLotNo, iRow, "*")
					
			Next							
			.ReDraw = True

		End With

End Function

Function SetMoveInvRef(arrRet)
		
	Dim TempRow
	Dim intLoopCnt
	Dim intCnt
	Dim iRow
		
	Call ClickTab2()
	
	With frm1.vspdData
			.focus
			ggoSpread.Source = frm1.vspdData			
			
			.ReDraw = False	
			intLoopCnt = Ubound(arrRet, 1)									
			
			TempRow = .MaxRows
			.MaxRows = TempRow + intLoopCnt + 1

			ggoSpread.SpreadUnLock	 C_ItemCd,		TempRow + 1,	C_ItemPopup,		.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryQty,	TempRow + 1,	C_EntryQty,			.MaxRows
			ggoSpread.SpreadUnLock	 C_EntryUnit,	TempRow + 1,	C_EntryUnitPopup,	.MaxRows
			ggoSpread.SpreadUnLock	 C_TrackingNo,	TempRow + 1,	C_TrnsLotSubNo,		.MaxRows
			ggoSpread.SSSetRequired  C_ItemCd,		TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryQty,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetRequired	 C_EntryUnit,	TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemCd,  TempRow + 1,	.MaxRows
			ggoSpread.SSSetProtected C_TrnsItemPopup, TempRow + 1,	.MaxRows
			
			For intCnt = 0 to intLoopCnt
				
				iRow = TempRow + intCnt + 1
					
 				Call .SetText(0, iRow, ggoSpread.InsertFlag)
				Call .SetText(C_ItemCd,	iRow, arrRet(intCnt, 0))
				Call .SetText(C_ItemNm, iRow, arrRet(intCnt, 1))
				Call .SetText(C_TrackingNo,iRow, arrRet(intCnt, 2))
				Call .SetText(C_EntryQty,  iRow, arrRet(intCnt, 3))
				Call .SetText(C_EntryUnit, iRow, arrRet(intCnt, 4))
				Call .SetText(C_LotSubNo,  iRow, "0")
				Call .SetText(C_LotNo,     iRow, "*")
				Call .SetText(C_TrnsLotNo, iRow, "*")
					
			Next							
			.ReDraw = True

		End With

End Function


'=============================== DbSave() ==================================================
Function DbSave() 
	Dim IntRows 
	Dim strVal
	Dim RowItem
	Dim iRowSep, iColSep

	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount
		
	iRowSep = Parent.gRowSep
	iColSep = Parent.gColSep				
	
	DbSave = False                                                       
	
	Call LayerShowHide(1)        
	
    Err.Clear		
	On Error Resume Next                                             
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
	
	With frm1
		.txtMode.value         = Parent.UID_M0002						
		.hYear.value           = frm1.txtYear.Year
	End With

	'-----------------------
	'Data manipulate area
	'-----------------------
	
	With frm1.vspdData
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
		
			Select Case Trim(.Text)
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					
					If .Text = ggoSpread.InsertFlag Then
						strVal = "C" & iColSep & IntRows & iColSep			
					Else
						strVal = "U" & iColSep & IntRows & iColSep			
					End If
		
					.Col = C_ItemCd	
					RowItem = Trim(.Text)
					strVal = strVal & RowItem & iColSep
					.Col = C_TrackingNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_LotSubNo	
					strVal = strVal & Trim(.Text) & iColSep
					
					.Col = C_EntryQty	
					If  UNICDbl(.Text) = 0 then
					    Call DisplayMsgBox("169918","X", "X", "X")
						Call LayerShowHide(0) 
						Call changeTabs(2)
						.Action = 0
					    Call SetToolBar("11101101001011")					
						Exit function
                    End if
					strVal = strVal & Trim(.Text) & iColSep					
					
					.Col = C_EntryUnit	
					strVal = strVal & Trim(.Text) & iColSep
					
					.Col = C_TrnsItemCd	
					if Trim(.Text) = "" then
						strVal = strVal & RowItem & iColSep				
					Else
						strVal = strVal & Trim(.Text) & iColSep				
					end if

					.Col = C_TrnsPlantCd		
					strVal = strVal & Trim(frm1.txtPlantCd2.value) & iColSep	
					.Col = C_TrnsSLCd	
					strVal = strVal & Trim(frm1.txtSLCd2.value) & iColSep	
					.Col = C_TrnsLotNo	
					strVal = strVal & Trim(.Text) & iColSep	
					.Col = C_TrnsLotSubNo	
					strVal = strVal & Trim(.Text) & iColSep	 		
					.Col = C_TrnsTrackingNo	
					strVal = strVal & Trim(frm1.txtTrackingNo.value) & iRowSep  
					
				
				Case ggoSpread.DeleteFlag

					strVal = "D" & iColSep & IntRows & iColSep	
			
					.Col = C_SeqNo		
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ItemCd		
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

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = false
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData 	
	Call ClickTab1()
	Call MainQuery()
End Function


