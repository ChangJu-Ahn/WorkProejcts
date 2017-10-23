
'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "i2211mb1.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_ItemCode            
Dim C_ItemName    
Dim C_ItemUnit    
Dim C_ItemSpec    
Dim C_Location    
Dim C_TrackingNo  
Dim C_GoodQty     
Dim C_BadQty      
Dim C_InspQty     
Dim C_TrnsQty     
Dim C_SchRcptQty  
Dim C_SchIssueQty 
Dim C_PrevGoodQty 
Dim C_PrevBadQty  
Dim C_PrevInspQty 
Dim C_PrevTrnsQty
Dim C_AllocationQty
Dim C_PickingQty

 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop          
 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

 '==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                 
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                        
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	If frm1.txtPlant_Cd.value = "" Then
		frm1.txtPlant_Nm.value = ""
	End if 
 
	If Parent.gPlant <> "" Then
		frm1.txtPlant_Cd.value = UCase(Parent.gPlant)
		frm1.txtPlant_Nm.value = Parent.gPlantNm
		frm1.txtSL_Cd.focus 
	Else
		frm1.txtPlant_Cd.focus 
	End If
 
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
	
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_PickingQty+1        
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		     
		ggoSpread.SSSetEdit C_ItemCode, "품목", 18
		ggoSpread.SSSetEdit C_ItemName, "품목명", 25
		ggoSpread.SSSetEdit C_ItemUnit, "단위", 8,2
		ggoSpread.SSSetEdit C_ItemSpec, "규격", 17 
		ggoSpread.SSSetEdit C_Location, "Location", 20
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20
		ggoSpread.SSSetFloat C_GoodQty, "양품재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec    
		ggoSpread.SSSetFloat C_BadQty, "불량재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_InspQty, "검사중수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
		ggoSpread.SSSetFloat C_TrnsQty, "이동중수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
		ggoSpread.SSSetFloat C_SchRcptQty, "입고예정량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
		ggoSpread.SSSetFloat C_SchIssueQty, "출고예정량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_PrevGoodQty, "전월양품재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_PrevBadQty, "전월불량재고량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_PrevInspQty, "전월검사중수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_PrevTrnsQty, "전월이동중수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_AllocationQty, "재고할당량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_PickingQty, "PICKING수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
		'ggoSpread.MakePairsColumn()
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		.ReDraw = true
		
		Call SetSpreadLock 
		ggoSpread.SSSetSplit2(2)
	End With
    
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode		= 1
	C_ItemName		= 2
	C_ItemUnit		= 3
	C_ItemSpec		= 4
	C_Location		= 5
	C_TrackingNo	= 6
	C_GoodQty		= 7
	C_BadQty		= 8
	C_InspQty		= 9
	C_TrnsQty		= 10
	C_SchRcptQty	= 11
	C_SchIssueQty	= 12
	C_PrevGoodQty	= 13
	C_PrevBadQty	= 14
	C_PrevInspQty	= 15
	C_PrevTrnsQty	= 16
	C_AllocationQty	= 17
	C_PickingQty	= 18
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode		= iCurColumnPos(1)
		C_ItemName		= iCurColumnPos(2)
		C_ItemUnit		= iCurColumnPos(3)
		C_ItemSpec		= iCurColumnPos(4)
		C_Location		= iCurColumnPos(5)
		C_TrackingNo	= iCurColumnPos(6)
		C_GoodQty		= iCurColumnPos(7)
		C_BadQty		= iCurColumnPos(8)
		C_InspQty		= iCurColumnPos(9)
		C_TrnsQty		= iCurColumnPos(10)
		C_SchRcptQty	= iCurColumnPos(11)
		C_SchIssueQty	= iCurColumnPos(12)
		C_PrevGoodQty	= iCurColumnPos(13)
		C_PrevBadQty	= iCurColumnPos(14)
		C_PrevInspQty	= iCurColumnPos(15)
		C_PrevTrnsQty	= iCurColumnPos(16)
		C_AllocationQty	= iCurColumnPos(17)	
		C_PickingQty	= iCurColumnPos(18)		
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
       ggoSpread.Source = frm1.vspdData
       ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================== 2.4.2 Open???()  =============================================
' Name : Open???()
' Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'      ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant1()  -------------------------------------------------
' Name : OpenPlant1()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlantCode()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 

 IsOpenPop = True

 arrParam(0) = "공장팝업" 
 arrParam(1) = "B_PLANT"    
 arrParam(2) = Trim(frm1.txtPlant_Cd.Value)
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
	frm1.txtPlant_Cd.focus 
	Exit Function
 Else
	frm1.txtPlant_Cd.Value    = arrRet(0)  
	frm1.txtPlant_Nm.Value    = arrRet(1)  
	frm1.txtPlant_Cd.focus 
 End If  
End Function


 '------------------------------------------  OpenSLCode()  -------------------------------------------------
' Name : OpenSLCode()
' Description : Storage Location Display PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenSLCode()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)
 
 If Trim(frm1.txtPlant_Cd.Value) = "" then 
  Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
  frm1.txtPlant_Cd.focus
  Exit Function
 End if

 '-----------------------
 'Check Plant CODE  '공장코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("125000","X","X","X")
  frm1.txtPlant_Nm.value = ""
  frm1.txtPlant_Cd.focus
  Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtPlant_Nm.value = lgF0(0)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "창고조회팝업"   
 arrParam(1) = "B_STORAGE_LOCATION"  
 arrParam(2) = Trim(frm1.txtSL_Cd.value)  
 arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S")  ' Where Condition    
 arrParam(5) = "창고"    ' TextBox 명칭 
 
 arrField(0) = "SL_CD"     
 arrField(1) = "SL_NM"     
    
 arrHeader(0) = "창고"   
 arrHeader(1) = "창고명"    
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
	frm1.txtSL_Cd.focus 
	Exit Function
 Else
	frm1.txtSL_Cd.value = arrRet(0) 
	frm1.txtSL_Nm.value	= arrRet(1)   
	frm1.txtSL_Cd.focus  
 End If 
 
End Function

 '------------------------------------------  OpenItemcode()  -------------------------------------------------
' Name : OpenItemCode()
' Description OPen Item Code Reference
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCode()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlant_Cd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlant_Cd.focus
		Exit Function
	End if

	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlant_Nm.value = ""
		frm1.txtPlant_Cd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlant_Nm.value = lgF0(0)
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = Trim(frm1.txtPlant_Cd.value) 
	arrParam(1) = Trim(frm1.txtItem_Cd.Value)      
	arrParam(2) = ""       
	arrParam(3) = ""       
	 
	arrField(0) = 1  
	arrField(1) = 2  
	arrField(2) = 9  
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
		frm1.txtItem_Cd.focus 
		Exit Function
	Else
		frm1.txtItem_Cd.value	= arrRet(0) 
		frm1.txtItem_Nm.value	= arrRet(1)   
		frm1.txtItem_Cd.focus
	End If 
End Function


'------------------------------------------  OpenEntryUnit()  -------------------------------------------------
' Name : OpenEntryUnit()
' Description : Entry Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenEntryUnit()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "단위 팝업"      
 arrParam(1) = "B_UNIT_OF_MEASURE"      
 arrParam(2) = Trim(frm1.txtinvUnit.Value)      ' Code Condition 
 arrParam(3) = ""        ' Name Condition 
 arrParam(4) = ""     ' Where Condition 
 arrParam(5) = "단위"      
 
 arrField(0) = "UNIT" 
 arrField(1) = "UNIT_NM" 
 
 arrHeader(0) = "단위"  
 arrHeader(1) = "단위명"  
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtinvunit.focus
	Exit Function
 Else
	frm1.txtinvunit.value = arrRet(0) 
	frm1.txtunit_Nm.value = arrRet(1)   
	frm1.txtinvunit.focus
 End If 
End Function


 '------------------------------------------  OpenOnhandDtlRef()  -------------------------------------------------
' Name : OpenOnhandDtlRefCode()
' Description : OnahndStock detail Reference
'--------------------------------------------------------------------------------------------------------- 

Function OpenOnhandDtlRef()
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

	Param1 = Trim(frm1.txtSL_Cd.value)
	Param7 = Trim(frm1.txtSL_Nm.value)
	Param4 = Trim(frm1.txtPlant_Cd.value)
	Param5 = "I"     
	 
	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(Param4, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlant_Nm.value = ""
		frm1.txtPlant_Cd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlant_Nm.value = lgF0(0)

	'-----------------------
	'Check Plant CODE  '창고코드가 있는 지 체크 
	'-----------------------
	If Param1 = "" then
		Call DisplayMsgBox("169902","X", "X", "X")   
		frm1.txtSL_Cd.focus 
		Exit Function
	Else
		If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S") & " AND SL_CD = " & FilterVar(Param1, "''", "S"), _
						  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		    
			If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(Param1, "''", "S"), _
							  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

				Call DisplayMsgBox("125700","X","X","X")
			Else
				Call DisplayMsgBox("169922","X","X","X")
			End If
			frm1.txtSL_Nm.value = ""
			frm1.txtSL_Cd.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtSL_Nm.value = lgF0(0)
	End If

	ggoSpread.Source = frm1.vspdData    
	Param2 = ""
	With frm1.vspdData     
		If .MaxRows = 0 Then
			Call DisplayMsgBox("169903","X", "X", "X")   
			frm1.txtItem_Cd.focus 
			Exit Function
		else
			.Col = C_ItemCode
			.Row = .ActiveRow
			Param2 = Trim(.Text )
			.Col = C_ItemName
			.Row = .ActiveRow
			Param8 = Trim(.Text )
			.Col = C_TrackingNo
			.Row = .ActiveRow
			Param3 = Trim(.Text )
			.Col = C_ItemUnit
			.Row = .ActiveRow
			Param9 = Trim(.Text)
		End If 
	End With
	     
	If Param2 = "" then
		Call DisplayMsgBox("169903","X", "X", "X")   
		frm1.txtItem_Cd.focus
		Exit Function
	Else
		If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " ITEM_CD= " & FilterVar(Param2, "''", "S") & " AND PLANT_CD = " & FilterVar(Param4, "''", "S"), _
						  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("122700","X","X","X")
			Exit Function
		End If
	End If

	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2212RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2212RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent , Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")      
	     
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlant_Cd.focus 
		Exit Function
	End If 
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
     
	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then 
		If lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" Then       
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End If
	End if 
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
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

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
   
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 

    FncQuery = False                                                      
    
    Err.Clear                                                             

    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables              
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPlant_Cd, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtSL_Cd, "A",1) Then Exit Function
    
    
 '-----------------------
 'Check Plant CODE  '공장코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("125000","X","X","X")
  frm1.txtPlant_Nm.value = ""
  frm1.txtPlant_Cd.focus
  Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtPlant_Nm.value = lgF0(0)

 '-----------------------
 'Check Plant CODE  '창고코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S") & " AND SL_CD = " & FilterVar(frm1.txtSL_Cd.value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSL_Cd.value, "''", "S"), _
   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

   Call DisplayMsgBox("125700","X","X","X")
  Else
   Call DisplayMsgBox("125710","X","X","X")
  End If
  frm1.txtSL_Nm.value = ""
  frm1.txtSL_Cd.focus
  Exit function
 End If
    
    lgF0 = Split(lgF0,Chr(11))
    frm1.txtSL_Nm.value = lgF0(0)

 '-----------------------
  'Check ItemCD CODE     '공장코드별 품목코드가 있는 지 체크 
 '-----------------------
 frm1.txtItem_Nm.value = ""
 If frm1.txtItem_Cd.value <> "" Then
  If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItem_Cd.value, "''", "S"), _
   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
  
   lgF0 = Split(lgF0,Chr(11))
   frm1.txtItem_Nm.value = lgF0(0)
  End If
 End If
 
 '-----------------------
  'Check invunit CODE     '재고단위코드가 있는 지 체크 
 '-----------------------
 If frm1.txtinvunit.value <> "" Then
  If  CommonQueryRs(" UNIT_NM "," B_UNIT_OF_MEASURE ", " UNIT = " & FilterVar(frm1.txtinvunit.Value, "''", "S"), _
   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
    
   Call DisplayMsgBox("124000","X","X","X")
   frm1.txtunit_Nm.value = ""
   frm1.txtinvunit.focus
   Exit function
  End If 
  lgF0 = Split(lgF0,Chr(11))
  frm1.txtunit_Nm.value = lgF0(0)  
 End If

    '-----------------------
    'Query function call area
    '-----------------------
     Call SetToolbar("11000000000111")         
    If DbQuery = False Then              
		Exit Function
	End if
       
    FncQuery = True                
    Set gActiveElement = document.activeElement
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
    Call parent.FncExport(Parent.C_MULTI)         
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , True)        
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit() 
    FncExit = True
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

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
' 설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    'Show Processing Bar
    Call LayerShowHide(1)  
    DbQuery = False
    
    Err.Clear                                                           

    Dim strVal
    Dim strValid
    Dim strQtyFlag
    
    
    If frm1.RadioOutputType.rdoCase1.Checked Then
		strValid = "Y"
	Else
		strValid = "N"
	End if
	
	If frm1.RadioOutputType2.rdoCase1.Checked Then
		strQtyFlag = "Y"
	Else
		strQtyFlag = "N"
	End if
    
    With frm1
    if lgIntFlgMode = Parent.OPMD_UMODE Then
     
    strVal = BIZ_PGM_ID & "?txtSL_Cd="        & Trim(.hSL_Cd.value)			& _   
						 "&txtItem_Cd="		  & Trim(.hItem_Cd.value)		& _
						 "&txtQryFrItem_Cd="  & Trim(.txtItem_Cd.value)		& _
						 "&txtPlant_Cd="      & Trim(.hPlant_Cd.value)		& _
						 "&txtCheck="		  & strValid					& _
						 "&txtQtyCheck="	  & strQtyFlag					& _
						 "&lgStrPrevKey1="    & lgStrPrevKey1				& _
						 "&lgStrPrevKey2="    & lgStrPrevKey2			
     
    Else
    
    strVal = BIZ_PGM_ID & "?txtSL_Cd="        & Trim(.txtSL_Cd.value)		& _   
						 "&txtItem_Cd="       & Trim(.txtItem_Cd.value)		& _
						 "&txtQryFrItem_Cd="  & Trim(.txtItem_Cd.value)		& _
						 "&txtinvunit="       & Trim(.txtinvunit.value)		& _
						 "&txtPlant_Cd="      & Trim(.txtPlant_Cd.value)	& _
						 "&txtCheck="		  & strValid					& _
						 "&txtQtyCheck="	  & strQtyFlag					& _	
						 "&lgStrPrevKey1="    & lgStrPrevKey1				& _
						 "&lgStrPrevKey2="    & lgStrPrevKey2
    
    End if
    
    Call RunMyBizASP(MyBizASP,strVal)         
        
    End With
    
    DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()             
 
    '-----------------------
    'Reset variables area    
    '-----------------------
    
    Call ggoOper.LockField(Document, "Q")   
    frm1.vspdData.focus 
    
End Function
