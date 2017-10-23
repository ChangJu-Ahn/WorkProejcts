<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Material Valuation
'*  3. Program ID           : I2321ma1.asp
'*  4. Program Name         : 표준단가수정 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'
'*  7. Modified date(First) : 2006/05/09
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : Update Standard Price
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                           
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

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
Const BIZ_PGM_QRY_ID = "i2321mb1.asp"
Const BIZ_PGM_SAVE_ID = "i2321mb2.asp"        

'==========================================  1.2.1 Global 상수 선언  ======================================
Dim IsOpenPop 
Dim lgOldRow

Dim	C_ItemCd           
Dim	C_ItemNm     
Dim	C_Spec
Dim C_TrackingNo        
Dim	C_ItemUnit     
Dim	C_Qty          
Dim	C_Amount       
Dim	C_Price        
Dim	C_PrevQty      
Dim	C_PrevAmount   
Dim	C_PrevPrice    

'***********************************Sub InitVariables()***************************************
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgIntGrpCount     = 0						
	lgBlnFlgChgValue  = False
	lgStrPrevKeyIndex = ""                      
	lgLngCurRows      = 0
    lgOldRow		  = 0
	
End Sub

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End if 
 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
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
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<%Call LoadBNumericFormatA("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
' Name : InitComboBox()	
'========================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboProcurType, lgF0, lgF1, Chr(11))
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20060509", , Parent.gAllowDragDropSpread
	
	With frm1.vspdData
		.ReDraw = false
		
		.MaxCols = C_PrevPrice+1         
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")		
				  
		ggoSpread.SSSetEdit		C_ItemCd,		"품목",			18
		ggoSpread.SSSetEdit		C_ItemNm,		"품목명",		20
		ggoSpread.SSSetEdit		C_Spec,			"규격",			20
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.",	20
		ggoSpread.SSSetEdit		C_ItemUnit,		"단위",			8,2
		ggoSpread.SSSetFloat	C_Qty,			"현재고수량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_Amount,		"현재고금액",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_Price,		"단가",			12, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 
		ggoSpread.SSSetFloat	C_PrevQty,		"전월재고수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PrevAmount,	"전월재고금액", 12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PrevPrice,	"전월단가",		12, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 
		  
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		.ReDraw = true
		ggoSpread.SpreadLockWithOddEvenRowColor
		
		ggoSpread.SpreadLock -1, -1
		ggoSpread.SpreadUnLock C_Price, -1, C_Price
		ggoSpread.SpreadUnLock C_PrevPrice, -1, C_PrevPrice
		ggoSpread.SSSetRequired  C_Price, -1, -1
		ggoSpread.SSSetRequired  C_PrevPrice, -1, -1
		ggoSpread.SSSetSplit2(2)
	End With

End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCd		= 1
	C_ItemNm		= 2
	C_Spec			= 3
	C_TrackingNo	= 4
	C_ItemUnit		= 5
	C_Qty			= 6
	C_Amount		= 7
	C_Price			= 8
	C_PrevQty		= 9
	C_PrevAmount	= 10
	C_PrevPrice		= 11
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

		C_ItemCd       = iCurColumnPos(1)
		C_ItemNm	   = iCurColumnPos(2)
		C_Spec         = iCurColumnPos(3)
		C_TrackingNo   = iCurColumnPos(4)
		C_ItemUnit     = iCurColumnPos(5)
		C_Qty          = iCurColumnPos(6)
		C_Amount       = iCurColumnPos(7)
		C_Price        = iCurColumnPos(8)
		C_PrevQty      = iCurColumnPos(9)
		C_PrevAmount   = iCurColumnPos(10)
		C_PrevPrice    = iCurColumnPos(11)
	End Select

End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
Sub SetSpreadColor( ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.SSSetProtected C_ItemCd,   pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm,   pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TrackingNo,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemUnit,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Qty,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Amount,  pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_Price,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PrevQty,  pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PrevAmount,  pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_PrevPrice,  pvStartRow, pvEndRow
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
		frm1.txtPlantCd.Value    = arrRet(0)  
		frm1.txtPlantNm.Value    = arrRet(1) 
		frm1.txtPlantCd.focus 
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
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	 
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus 
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	 
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
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA1","x")
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
		frm1.txtItemCd.value = arrRet(0) 
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
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
 arrHeader(1) = "계정명"      
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtItemAcct.focus
	Exit Function
 Else
	frm1.txtItemAcct.Value		= arrRet(0)
	frm1.txtItemAcctNm.Value	= arrRet(1)
	frm1.txtItemAcct.focus
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

'------------------------------------------  OpenItemGroup()  --------------------------------------------------
' Name : OpenItemGroup()
' Description : Item Group Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
 
	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"    
	arrParam(1) = "B_ITEM_GROUP"      
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value) 
	arrParam(3) = ""       
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "  
	arrParam(5) = "품목그룹"   
 
	arrField(0) = "ITEM_GROUP_CD"      
	arrField(1) = "ITEM_GROUP_NM"      
 
	arrHeader(0) = "품목그룹"     
	arrHeader(1) = "품목그룹명"      
 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtItemGroupCd.focus
		Exit Function
	Else
		frm1.txtItemGroupCd.Value	= arrRet(0)
		frm1.txtItemGroupNm.Value	= arrRet(1)
		frm1.txtItemGroupCd.focus
	End If 
 
End Function


'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
 
    Call LoadInfTB19029 
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call InitSpreadSheet
    Call InitVariables
    
     '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
	Call SetToolbar("11000000000011")         
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End if
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then Exit Sub
		
		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then
				If DbQuery = False Then	Exit Sub
			End if
		End if
	
	End With
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

'================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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
	Dim IntRetCD
	FncQuery = False                                                      
 
	Err.Clear                                                             
 
	'-----------------------
	'Erase contents area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData

    Call InitVariables
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
	   
	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
	 
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemNm.value = lgF0(0)
		End If
	End If
	
	frm1.txtItemGroupNm.value = ""
	If frm1.txtItemGroupCd.value <> "" Then
		If  CommonQueryRs(" ITEM_GROUP_NM "," B_ITEM_GROUP ", " ITEM_GROUP_CD = " & FilterVar(frm1.txtItemGroupCd.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
	 
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItemGroupNm.value = lgF0(0)
		End If
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
' Function Name : FncSave
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncSave()
	Dim IntRetCD 
         
    FncSave = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit function
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


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery()
	Err.Clear												                   
	    
    DbQuery = False											                   
	
    Call LayerShowHide(1)

    Dim strVal
    Dim strFlag
    
    If frm1.RadioOutputType2.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
	
	
	With frm1
		
		'if lgIntFlgMode = Parent.OPMD_UMODE Then
		'	strVal = BIZ_PGM_QRY_ID		& "?txtPlantCd="		& .hPlantCd.value				& _
		'								"&txtItemAcct="			& .hItemAcct.value				& _
		'								"&txtItemCd="			& .hItemCd.value				& _
		'								"&txtTrackingNo="		& .hTrackingNo.value			& _
		'								"&txtItemGrpCd="		& .hItemGroupCd.value			& _
		'								"&cboProcType="			& .hProcureType.value			& _
		'								"&txtFlag="				& strFlag						& _
		'								"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
		'								"&txtMaxRows="			& .vspdData.MaxRows

		'Else
			strVal = BIZ_PGM_QRY_ID		&"?txtPlantCd="			& Trim(.txtPlantCd.value)		& _
										"&txtItemAcct="			& Trim(.txtItemAcct.value)		& _
										"&txtItemCd="			& Trim(.txtItemCd.value)		& _
										"&txtTrackingNo="		& Trim(.txtTrackingNo.value)	& _
										"&txtItemGrpCd="		& Trim(.txtItemGroupCd.value)	& _
										"&cboProcType="			& Trim(.cboProcurType.value)	& _
										"&txtFlag="				& strFlag						& _
										"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
										"&txtMaxRows="			& .vspdData.MaxRows
		'End if
		
	End With								
	

    Call RunMyBizASP(MyBizASP, strVal)						                    

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
frm1.vspdData.focus
lgIntFlgMode = Parent.OPMD_UMODE 

Call SetToolbar("11001000000111")


End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave()
	Dim lRow
    Dim lGrpCnt
    Dim strVal
    
    DbSave = False
    
    If LayerShowHide(1) = False Then
		Exit Function
    End If
    
    strVal = ""
    lGrpCnt = 1
    
    With frm1
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0
			
			Select Case .vspdData.Text
				Case ggoSpread.UpdateFlag
				
													strVal = strVal & "U"  &  parent.gColSep
													strVal = strVal & lRow &  parent.gColSep
				.vspdData.Col = C_ItemCd		:	strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
				.vspdData.Col = C_TrackingNo	:	strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep				
				.vspdData.Col = C_Price			:	strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
				.vspdData.Col = C_PrevPrice		:	strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
				
				lGrpCnt = lGrpCnt + 1
			End Select
			
		Next
		
		.txtMaxRows.value     = lGrpCnt-1
		.txtSpread.value      = strVal
		
    End With

    Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	DbSave = True	
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()

	Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call MainQuery()

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
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>표준단가수정</font></TD>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
									</TR>
								</TABLE>
							</TD>
							<TD WIDTH=*>&nbsp;</TD>
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
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4" tag="12XXXU" ALT = "공장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=30 tag="14"></TD>    
											<TD CLASS="TD5" NOWRAP>품목계정</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=6 MAXLENGTH=2 tag="12XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=27 MAXLENGTH=20 tag="14"></TD>      
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>품목</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=30 tag="14"></TD>
											<TD CLASS="TD5" NOWRAP>Tracking No.</TD>     
											<TD CLASS="TD6" NOWRAP >
												<TABLE CELLSPACING=0 CELLPADDING= 0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 tag="11XXXU" ALT = "Tracking No." ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>
														</TD>
														<TD WIDTH="*">
															&nbsp;
														</TD>
														<TD WIDTH="20" STYLE="TEXT-ALIGN: RIGHT" ><IMG SRC="../../../CShared/image/BigPlus.gif" Style="CURSOR: hand" ALT="DetailCondition" ALIGN= "TOP" ID = "IMG_DetailCondition" NAME="pop1" ONCLICK= 'vbscript:viewHidden "DetailCondition" ,2, 3' ></IMG>
														</TD>
													</TR>
												</TABLE>
											</TD>		
										</TR>
										<TR ID="DetailCondition1" style="display: none">
											<TD CLASS="TD5" NOWRAP>품목그룹</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
											<TD CLASS="TD5" NOWRAP>조달구분</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboProcurType" ALT="조달구분" STYLE="Width: 160px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
										</TR>
										<TR ID="DetailCondition2" style="display: none">
											<TD CLASS="TD5" NOWRAP>수량유무</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">수량있음</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">전품목</LABEL>
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
											<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1>
											<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
											</OBJECT>
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
			<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hProcureType" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
		</FORM>
		<DIV ID="MousePT" NAME="MousePT">
			<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
		</DIV>
	</BODY>
</HTML>

