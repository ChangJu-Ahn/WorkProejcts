<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Monthly Inventory
'*  3. Program ID           : I2241ma1.asp
'*  4. Program Name         : 이월재고현황 
'*  5. Program Desc         : 
'*  6. Comproxy List        :                        
'                             
'*  7. Modified date(First) : 2000/04/18
'*  8. Modified date(Last)  : 2005/08/09
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/18 : ..........
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Const BIZ_PGM_ID = "i2241mb1.asp"            '☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Dim	C_ItemCode       
Dim	C_ItemName
Dim	C_Spec
'************* Tracking No Addition LSW 2005-04-08 ************************************
Dim C_TrackingNo
Dim	C_ItemUnit
Dim	C_BaseQty
Dim	C_BaseAmount
Dim	C_EntryQty
Dim	C_EntryAmount
Dim	C_OrderQty
Dim	C_OrderAmount 
Dim	C_Qty 
Dim	C_Amount 

 '==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
 
Dim IsOpenPop

Dim lgStrPrevToKey
Dim DbChkFlag        
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntGrpCount	= 0                           
    lgIntFlgMode	= Parent.OPMD_CMODE                        
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevToKey  = 1
    lgLngCurRows	= 0                            
    
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	If frm1.txtPlantCd.value = "" Then frm1.txtPlantNm.value = ""
 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		Call txtPlantCd_LostFocus

		frm1.txtItemAcct.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20081006", , Parent.gAllowDragDropSpread
	
	With frm1.vspdData
		.ReDraw = False
		.MaxCols = C_Amount+1       
		.MaxRows = 0
		  
		Call GetSpreadColumnPos("A")
		  
		ggoSpread.SSSetEdit C_ItemCode, "품목", 18
		ggoSpread.SSSetEdit C_ItemName, "품목명", 25
		ggoSpread.SSSetEdit C_Spec, "규격", 20
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No.", 20
		ggoSpread.SSSetEdit C_ItemUnit, "단위", 10,2
		ggoSpread.SSSetFloat C_BaseQty, "기초수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_BaseAmount, "기초금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_EntryQty, "입고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_EntryAmount, "입고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_OrderQty, "출고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_OrderAmount, "출고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Qty, "재고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_Amount, "재고금액", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec

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
	C_Spec			= 3
	C_TrackingNo	= 4
	C_ItemUnit		= 5
	C_BaseQty		= 6
	C_BaseAmount	= 7
	C_EntryQty		= 8
	C_EntryAmount	= 9
	C_OrderQty		= 10
	C_OrderAmount	= 11
	C_Qty			= 12
	C_Amount		= 13
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

		C_ItemCode    = iCurColumnPos(1)
		C_ItemName    = iCurColumnPos(2)
		C_Spec        = iCurColumnPos(3)
		C_TrackingNo  = iCurColumnPos(4)
		C_ItemUnit    = iCurColumnPos(5)
		C_BaseQty     = iCurColumnPos(6)
		C_BaseAmount  = iCurColumnPos(7)
		C_EntryQty    = iCurColumnPos(8)
		C_EntryAmount = iCurColumnPos(9)
		C_OrderQty    = iCurColumnPos(10)
		C_OrderAmount = iCurColumnPos(11)
		C_Qty         = iCurColumnPos(12)
		C_Amount      = iCurColumnPos(13)
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
 With frm1
  .vspdData.ReDraw = False    
  ggoSpread.SpreadLockWithOddEvenRowColor()
  .vspdData.ReDraw = True
 End With
End Sub


 '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
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
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtPlantCd.focus
	Exit Function
 Else
	Call SetPlant(arrRet)
 End If 
 
End Function

 '------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenItem()
' Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	'공장코드가 있는 지 체크 
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus  
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	  
	'품목계정코드가 있는지 체크 
	If Trim(frm1.txtItemAcct.Value) <> "" then
		If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & Trim(FilterVar(frm1.txtItemAcct.Value," ","S")), _
						  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("169952",vbOKOnly, "x", "x")
			frm1.txtItemAcctNm.value = ""
			frm1.txtItemAcct.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtItemAcctNm.value = lgF0(0)
	End If
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = Trim(frm1.txtPlantCd.value) 
	arrParam(1) = Trim(frm1.txtItemCd.Value)     
	arrParam(2) = ""
	arrParam(3) = Trim(frm1.txtItemAcct.Value)      
	 
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
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItem(arrRet)
	End If 
End Function

 '------------------------------------------  OpenItemAcct()  --------------------------------------------------
' Name : OpenItemAcct()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 
 IsOpenPop = True

 arrParam(0) = "품목계정 팝업"    
 arrParam(1) = "B_MINOR"     
 arrParam(2) = Trim(frm1.txtItemAcct.Value)     
 arrParam(3) = ""      
 arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""    
 arrParam(5) = "품목계정"   
 
 arrField(0) = "MINOR_CD"   
 arrField(1) = "MINOR_NM"   
 
 arrHeader(0) = "품목계정"     
 arrHeader(1) = "품목계정명"     
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtItemAcct.focus
	Exit Function
 Else
	Call SetItemAcct(arrRet)
 End If 
 
End Function

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Tracking No."	
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackingNo.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "Tracking No."			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking_No"		
    arrHeader(1) = "품목"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
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
	 
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtPlantCd.value)
	if Param1 = "" then
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If
	 
	Param2 = Trim(frm1.txtPlantNm.Value)
	Param3 = Trim(frm1.txtYyyyMm.Text)
	 
	ggoSpread.Source = frm1.vspdData    
	Param4 = ""
	With frm1.vspdData
		If .MaxRows = 0 Then
			Call DisplayMsgBox("169903","X", "X", "X")
			frm1.txtItemCd.focus    
			Exit Function
		else
			.Row = .ActiveRow
			.Col = C_ItemCode
			Param4 = Trim(.Text)
			.Col = C_ItemName
			Param5 = Trim(.Text)
			.Col = C_Spec
			Param6 = ""
			.Col = C_ItemUnit
			Param7 = Trim(.Text )
		End If 
	End With
	     
	if Param4 = "" then
		Call DisplayMsgBox("169903","X", "X", "X") 
		frm1.txtItemCd.focus 
		Exit Function
	End If

	Param8 = frm1.txtYyyyMm.UserDefinedFormat
	 
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2241RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2241RA1","x")
		IsOpenPop = False
		Exit Function
	End If
 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")  
	     
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.focus
		Exit Function
	End If 
End Function

 '------------------------------------------  SetPlant()  --------------------------------------------------
' Name : SetPlant()
' Description : OpenPlant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
 frm1.txtPlantCd.Value    = arrRet(0)  
 frm1.txtPlantNm.Value    = arrRet(1)
 frm1.txtPlantCd.focus  
End Function

 '------------------------------------------  SetItemCode()  --------------------------------------------------
' Name : SetItemCode()
' Description : ItemCode Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItem(byRef arrRet)

 frm1.txtItemCd.value = arrRet(0) 
 frm1.txtItemNm.value = arrRet(1)   
 frm1.txtItemCd.focus
End Function

 '------------------------------------------  SetItemAcct()  --------------------------------------------------
' Name : SetItemAcct()
' Description : ItemAcct Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byRef arrRet)
 frm1.txtItemAcct.Value		= arrRet(0)
 frm1.txtItemAcctNm.Value	= arrRet(1)
 frm1.txtItemAcct.focus
End Function

 '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    
 Call LoadInfTB19029              
 Call ggoOper.LockField(Document, "N")                                 
 Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
 Call ggoOper.FormatDate(frm1.txtYyyyMm, Parent.gDateFormat, "2")

 Call InitSpreadSheet 
 Call InitVariables                                                   
 Call SetDefaultVal

 Call SetToolbar("11000000000011")
 
End Sub


'=======================================================================================================
'   Event Name : txtYyyyMm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtYyyyMm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyyMm.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtYyyyMm.Focus
    End If
End Sub


Function txtYyyyMm_KeyPress(KeyAscii)
 If KeyAscii = 13 Then
  Call MainQuery()
 End If
End Function


'=======================================================================================================
'   Event Name : txtPlantCd_LostFocus()
'   Event Desc : 공장명과 최종마감년월을 찾는다.
'=======================================================================================================
Sub txtPlantCd_LostFocus()
    Dim strYear
    Dim strMonth
    Dim strDay
 
	If frm1.txtPlantCd.value <> "" Then
		If  CommonQueryRs(" PLANT_NM, CONVERT(CHAR(10), INV_CLS_DT, 21) "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			frm1.txtPlantNm.Value  = ""
			frm1.txtYyyyMm.text  = ""
			Exit Sub
		End If

		lgF0 = Split(lgF0,Chr(11))
		lgF1 = Split(lgF1,Chr(11))
  
		frm1.txtPlantNm.Value = lgF0(0)
         
	Else
		frm1.txtPlantNm.Value  = ""
		frm1.txtYyyyMm.text  = ""
	End If
End Sub

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
	
	If DbChkFlag = True Then
		Exit Sub
	End If
	
	'----------  Coding part  -------------------------------------------------------------   
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End if
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
	If Not chkField(Document, "1") Then         
		Exit Function
	End If

 '-----------------------
 '공장코드가 있는 지 체크 
 '----------------------- 
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		Call DisplayMsgBox("125000","X", "X", "X")    
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus  
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
 '----------------------- 
 '품목계정코드가 있는지 체크 
 '----------------------- 
	If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & Trim(FilterVar(frm1.txtItemAcct.Value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
		Call DisplayMsgBox("169952",vbOKOnly, "x", "x")   
		frm1.txtItemAcctNm.value = ""
		frm1.txtItemAcct.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtItemAcctNm.value = lgF0(0)
 
 '-----------------------
 '품목코드가 있는 지 체크 
 '-----------------------
	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
  
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemNm.value = lgF0(0)
		End If
	End If

 '-----------------------
 'Query function call area
 '-----------------------
	lgIntFlgMode = Parent.OPMD_CMODE
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
 call parent.FncExport(Parent.C_MULTI)
    On Error Resume Next                                                  
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

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim LngLastRow      
	Dim LngMaxRow       
	'Dim StrNextKey      
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strVal

	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
 
	DbQuery = False
 
	Call LayerShowHide(1)
 
	Err.Clear    
 
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtPlantCd="	& .hPlantCd.value		& _
								"&txtYyyy="			& .hYyyy.value			& _
								"&txtMm="			& .hMm.value			& _
								"&txtItemAcct="		& .hItemAcct.value		& _
								"&txtTrackingNo="	& .hTrackingNo.value	& _
								"&lgStrPrevToKey="	& lgStrPrevToKey		& _
								"&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtPlantCd="	& Trim(.txtPlantCd.value)		& _	
								"&txtYyyy="			& strYear						& _
								"&txtMm="			& strMonth						& _
								"&txtItemAcct="		& Trim(.txtItemAcct.value)		& _
								"&txtQryFrItemCd="	& Trim(.txtItemCd.value)		& _
								"&txtItemCd="		& Trim(.txtItemCd.value)		& _
								"&txtTrackingNo="	& Trim(.txtTrackingNo.value)	& _
								"&lgStrPrevToKey="	& lgStrPrevToKey				& _
								"&txtMaxRows="		& .vspdData.MaxRows
		End If 
     
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

'========================================================================================
' Function Name : ViewHidden
' Function Desc : Show Detail Field
'========================================================================================
Function ViewHidden(StrMnuID, MnuCount, StrImageSize )
    Dim ii

    For ii = 1 To MnuCount
        If document.all(StrMnuID & ii).style.display = "" Then 
           document.all(StrMnuID & ii).style.display = "none"
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/Smallplus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigPlus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
			End Select		
        Else
           document.all(StrMnuID & ii).style.display = ""
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/SmallMinus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigMinus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
			End Select
        End If
    Next    

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%> >
			<TR>
				<TD <%=HEIGHT_TYPE_00%> >
				</TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_10%> >
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
									<TR>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>이월재고현황</font></TD>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
									</TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenOnhandDtlRef()">수불상세</A></TD>     
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD <%=HEIGHT_TYPE_02%> >
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
										<TR>
											<TD CLASS="TD5" NOWRAP>공장</TD>      
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="12XXXU" ALT = "공장" onBlur="vbscript:txtPlantCd_LostFocus()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=27 MAXLENGTH=30 tag="14"></TD>    
											<TD CLASS="TD5" NOWRAP>연월</TD>
											<TD CLASS="TD6" NOWRAP>
											<script language =javascript src='./js/i2241ma1_fpDateTime1_txtYyyyMm.js'></script>
											</TD>       
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>품목계정</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=6 MAXLENGTH=2 tag="12XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=27 MAXLENGTH=20 tag="14"></TD>
											<TD CLASS="TD5" NOWRAP>품목</TD>     
											<TD CLASS="TD6" NOWRAP >
												<TABLE CELLSPACING=0 CELLPADDING= 0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT = "품목" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=30 tag="14"></TD>
														</TD>
														<TD WIDTH="*">
															&nbsp;
														</TD>
														<TD WIDTH="20" STYLE="TEXT-ALIGN: RIGHT" ><IMG SRC="../../../CShared/image/BigPlus.gif" Style="CURSOR: hand" ALT="DetailCondition" ALIGN= "TOP" ID = "IMG_DetailCondition" NAME="pop1" ONCLICK= 'vbscript:viewHidden "DetailCondition" ,1, 3' ></IMG>
														</TD>
													</TR>
												</TABLE>
											</TD>		
										</TR>
										<TR ID="DetailCondition1" style="display: none">
											<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()">
											</TD>
											<TD CLASS="TD5" NOWRAP></TD>
											<TD CLASS="TD6" NOWRAP></TD>
										</TR>
										<TR>
											<TD <%=HEIGHT_TYPE_03%> >
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=* valign=top>
									<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
										<TR>
											<TD HEIGHT="100%">
											<script language =javascript src='./js/i2241ma1_I619976311_vspdData.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_01%> >
					</TD>
				</TR>
				<TR HEIGHT=20 >
					<TD>
						<TABLE <%=LR_SPACE_TYPE_30%> >
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=<%=BizSize%>>
						<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
						</IFRAME>
					</TD>
				</TR>
			</TABLE>
			<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
			<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hYyyy" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hMM" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hcomcfg" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24" TABINDEX="-1">
		</FORM>
		<DIV ID="MousePT" NAME="MousePT">
			<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
		</DIV>
	</BODY>
</HTML>

