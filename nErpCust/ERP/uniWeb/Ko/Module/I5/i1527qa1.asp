<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory List onhand stock
'*  2. Function Name          : 
'*  3. Program ID             : I1527qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 재고현황조회(VMI)
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/02/18
'*  8. Modified date(Last)    : 2003/04/28
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : Ahn Jung Je
'* 11. Comment                :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit       

Const BIZ_PGM_ID = "I1527qb1.asp"

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

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgUserFlag 

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                 
    lgIntGrpCount = 0                           
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            
End Sub

'==========================================  2.2.1 SetDefaultVal() ========================================
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
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
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
		ggoSpread.SSSetFloat C_PickingQty, "Picking 수량", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = true
		
		ggoSpread.SSSetSplit2(2)
	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
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

 '------------------------------------------  OpenPlant1()  -------------------------------------------------
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
		Call SetPlant(arrRet)
	End If  
End Function

 '------------------------------------------  OpenSLCode()  -------------------------------------------------
Function OpenSLCode()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If Trim(frm1.txtPlant_Cd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")    
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
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S") 
	arrParam(5) = "창고"   
	 
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
		Call SetSLCode(arrRet)
	End If 
End Function

'------------------------------------------  OpenItemcode()  -------------------------------------------------
Function OpenItemCode()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	If Trim(frm1.txtPlant_Cd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")   
		Exit Function
	End if

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
		Call SetItem(arrRet)
	End If 
End Function

'------------------------------------------  OpenEntryUnit()  -------------------------------------------------
Function OpenEntryUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위 팝업"      
	arrParam(1) = "B_UNIT_OF_MEASURE"  
	arrParam(2) = Trim(frm1.txtinvUnit.Value) 
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
		frm1.txtinvunit.focus
		Exit Function
	Else
		Call SetEntryUnit(arrRet)
	End If 
End Function


'------------------------------------------  OpenOnhandDtlRef()  -------------------------------------------------
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
	 

	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(Param4, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlant_Nm.value = ""
		frm1.txtPlant_Cd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlant_Nm.value = lgF0(0)

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
		frm1.txtSL_Cd.focus 
		Exit Function
	End If 
End Function

'------------------------------------------  SetPlant1()  --------------------------------------------------
Function SetPlant(byRef arrRet)
	frm1.txtPlant_Cd.Value    = arrRet(0)  
	frm1.txtPlant_Nm.Value    = arrRet(1)  
	frm1.txtPlant_Cd.focus 
End Function

'------------------------------------------  SetSLCode()  --------------------------------------------------
Function SetSLCode(byRef arrRet)
	frm1.txtSL_Cd.value = arrRet(0) 
	frm1.txtSL_Nm.value = arrRet(1)   
	frm1.txtSL_Cd.focus 
End Function

'------------------------------------------  SetItemCode()  --------------------------------------------------
Function SetItem(byRef arrRet)
	frm1.txtItem_Cd.value = arrRet(0) 
	frm1.txtItem_Nm.value = arrRet(1)   
	frm1.txtItem_Cd.focus
End Function

'------------------------------------------  SetEntryUnit()  --------------------------------------------------
Function SetEntryUnit(byRef arrRet)
	frm1.txtinvunit.value = arrRet(0) 
	frm1.txtunit_Nm.value = arrRet(1)   
	frm1.txtinvunit.focus
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
     If OldLeft <> NewLeft Then Exit Sub
     If CheckRunningBizProcess = True Then Exit Sub
    
	 If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" Then       '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End if
	 End if 
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029             
    Call ggoOper.LockField(Document, "N")                                     
    Call InitSpreadSheet    
    Call InitVariables                                                    
    Call SetDefaultVal
    Call SetToolbar("11000000000011")        
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
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
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                    
    Err.Clear
    
    If Not chkField(Document, "1") Then Exit Function                                                           

	If Trim(frm1.txtPlant_Cd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		frm1.txtPlant_Nm.Value = ""
		frm1.txtPlant_Cd.focus
		Exit function
	Else
		If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlant_Cd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlant_Nm.value = ""
			frm1.txtPlant_Cd.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtPlant_Nm.value = lgF0(0)
	End If
	
 	If Trim(frm1.txtSL_Cd.Value) = "" Then
		Call DisplayMsgBox("169902","X","X","X")
		frm1.txtSL_Nm.Value = ""
		frm1.txtSL_Cd.focus
		Exit function
	Else
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
	End If    

	frm1.txtItem_Nm.value = ""
	If frm1.txtItem_Cd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItem_Cd.value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			lgF0 = Split(lgF0,Chr(11))
			frm1.txtItem_Nm.value = lgF0(0)
		End If
	End If
 
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
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData         
    Call InitVariables             
    '-----------------------
    'Query function call area
    '-----------------------
    Call SetToolbar("11000000000111")         
	If DbQuery = False Then Exit Function
    Set gActiveElement = document.activeElement   
    FncQuery = True                
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)         
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , True)      
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit() 
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

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
			strVal = BIZ_PGM_ID &	"?txtSL_Cd="        & Trim(.hSL_Cd.value)		& _  
									"&txtItem_Cd="      & Trim(.hItem_Cd.value)		& _
									"&txtQryFrItem_Cd=" & Trim(.txtItem_Cd.value)	& _
									"&txtPlant_Cd="     & Trim(.hPlant_Cd.value)	& _
									"&txtCheck="		& strValid					& _
									"&txtQtyCheck="		& strQtyFlag				& _
									"&lgStrPrevKey1="   & lgStrPrevKey1				& _
									"&lgStrPrevKey2="   & lgStrPrevKey2				& _
									"&txthUserFlag="	& "V"
		Else
			strVal = BIZ_PGM_ID &	"?txtSL_Cd="        & Trim(.txtSL_Cd.value)		& _ 
									"&txtItem_Cd="      & Trim(.txtItem_Cd.value)	& _
									"&txtQryFrItem_Cd=" & Trim(.txtItem_Cd.value)	& _
									"&txtinvunit="      & Trim(.txtinvunit.value)	& _
									"&txtPlant_Cd="     & Trim(.txtPlant_Cd.value)	& _
									"&txtCheck="		& strValid					& _
									"&txtQtyCheck="		& strQtyFlag				& _
									"&lgStrPrevKey1="   & lgStrPrevKey1				& _
									"&lgStrPrevKey2="   & lgStrPrevKey2				& _
									"&txthUserFlag="	& "V"
		End if
    
		Call RunMyBizASP(MyBizASP,strVal)        
        
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()           
    '-----------------------
    'Reset variables area    
    '-----------------------
    frm1.vspdData.focus 
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고현황조회(VMI품목)</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOnhandDtlRef()">재고상세정보</A></TD>     
					<TD WIDTH=10>
					</TD>
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
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>      
									<TD CLASS="TD6" NOWRAP >
									<input NAME="txtPlant_Cd" TYPE="Text" MAXLENGTH="4" tag="12XXXU" ALT = "공장" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode"  align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPlantCode()">&nbsp;<input NAME="txtPlant_Nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14N"></td>    
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP >
									<input NAME="txtSL_Cd" TYPE="Text" MAXLENGTH="7" tag="12XXXU" ALT = "창고" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenSLCode()">&nbsp;<input NAME="txtSL_Nm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14N"></td>    
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>      
									<TD CLASS="TD6" NOWRAP >
									<input NAME="txtItem_Cd" TYPE="Text" MAXLENGTH="18" tag="11NXXU" ALT = "품목" size="15"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenItemCode()">&nbsp;<input NAME="txtItem_Nm" TYPE="Text" MAXLENGTH="40" tag="14N"></td>     
									<TD CLASS="TD5" NOWRAP>재고단위</td>      
									<TD CLASS="TD6" NOWRAP >
									<input NAME="txtinvunit" TYPE="Text" MAXLENGTH="3" tag="11NXXU" ALT = "재고단위" size="8"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgSaleOrgCode" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenEntryUnit()">&nbsp;<input NAME="txtunit_Nm" TYPE="Text" MAXLENGTH="40" SIZE=20 tag="14N"></td>     
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목유효일체크</TD>
									<TD CLASS="TD6">
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">예</LABEL>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">아니오</LABEL>
									</TD>
									<TD CLASS="TD5" NOWRAP>양품수량유무</TD>      
									<TD CLASS="TD6" NOWRAP >
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">수량있음</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">전품목</LABEL>
									</TD>						                           
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
									<script language =javascript src='./js/i1527qa1_OBJECT1_vspdData.js'></script></TD>
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
	<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
	<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hPlant_Cd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hSL_Cd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hItem_Cd" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txthUserFlag" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

