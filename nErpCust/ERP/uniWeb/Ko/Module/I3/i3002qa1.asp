<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i3002qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 재고현황 상세조회 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/07/02
'*  8. Modified date(Last)    : 2003/07/02
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

Const BIZ_PGM_QRY_ID	= "i3002qb1.asp"

Dim C_SlCd
Dim C_SlNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_Unit
Dim C_OnhandQty
Dim C_GoodQty
Dim C_BadQty
Dim C_InspQty
Dim C_TrnsQty
Dim C_PrevOnhandQty
Dim C_PrevGoodQty
Dim C_PrevBadQty
Dim C_PrevInspQty
Dim C_PrevTrnsQty
Dim C_PikingQty
Dim C_LastRcptDt
Dim C_LastIssueDt
Dim C_ExpiaryDt
Dim C_LastPhyDt
Dim C_PriceFlag
Dim C_Price

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop 
Dim lgOldRow


'==========================================  2.1.1 InitVariables()  =====================================
Sub InitVariables()
	lgIntFlgMode		= parent.OPMD_CMODE
	lgIntGrpCount		= 0	
	lgBlnFlgChgValue	= False
	lgLngCurRows		= 0
    lgOldRow			= 0
    lgStrPrevKeyIndex	= ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemAcct.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement

End Sub


'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "QA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030702", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
			
		.MaxCols = C_Price+1					
		.MaxRows = 0
				
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
				
		ggoSpread.SSSetEdit     C_SlCd,				"창고",				10
		ggoSpread.SSSetEdit     C_SlNm,				"창고명",			20
		ggoSpread.SSSetEdit     C_ItemCd,			"품목",				18
		ggoSpread.SSSetEdit		C_ItemNm,			"품목명",			20
		ggoSpread.SSSetEdit		C_Spec,				"규격",				20
		ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.",		20
		ggoSpread.SSSetEdit 	C_LotNo,			"Lot No.",			12
		ggoSpread.SSSetFloat	C_LotSubNo,			"Lot순번",			10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit		C_Unit,				"단위",				10
		ggoSpread.SSSetFloat    C_OnhandQty,		"현재고수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_GoodQty,			"양품재고수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_BadQty,			"불량재고수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_InspQty,			"검사중수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_TrnsQty,			"이동중수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PrevOnhandQty,	"전월재고수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PrevGoodQty,		"전월양품재고량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PrevBadQty,		"전월불량재고량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PrevInspQty,		"전월검사중수량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PrevTrnsQty,		"전월이동중수량",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_PikingQty,		"PICKING수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetDate		C_LastRcptDt,		"최종입고일",		12, 2,Parent.gDateFormat
		ggoSpread.SSSetDate		C_LastIssueDt,		"최종출고일",		12, 2,Parent.gDateFormat
		ggoSpread.SSSetDate		C_ExpiaryDt,		"실효일",			12, 2,Parent.gDateFormat
		ggoSpread.SSSetDate		C_LastPhyDt,		"최종실사일",		12, 2,Parent.gDateFormat
		ggoSpread.SSSetEdit		C_PriceFlag,		"단가구분",			15
		ggoSpread.SSSetFloat    C_Price,			"단가",				15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		 		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
	End With
		
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_SlCd			= 1
	C_SlNm			= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_TrackingNo	= 6
	C_LotNo			= 7
	C_LotSubNo		= 8
	C_Unit			= 9
	C_OnhandQty		= 10
	C_GoodQty		= 11
	C_BadQty		= 12
	C_InspQty		= 13
	C_TrnsQty		= 14
	C_PrevOnhandQty	= 15
	C_PrevGoodQty	= 16
	C_PrevBadQty	= 17
	C_PrevInspQty	= 18
	C_PrevTrnsQty	= 19
	C_PikingQty		= 20
	C_LastRcptDt	= 21
	C_LastIssueDt	= 22
	C_ExpiaryDt		= 23
	C_LastPhyDt		= 24
	C_PriceFlag		= 25
	C_Price			= 26
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
		C_SlCd				= iCurColumnPos(1)
		C_SlNm				= iCurColumnPos(2)
		C_ItemCd			= iCurColumnPos(3)
		C_ItemNm			= iCurColumnPos(4)
		C_Spec				= iCurColumnPos(5)
		C_TrackingNo		= iCurColumnPos(6)
		C_LotNo				= iCurColumnPos(7)
		C_LotSubNo			= iCurColumnPos(8)
		C_Unit				= iCurColumnPos(9)
		C_OnhandQty			= iCurColumnPos(10)
		C_GoodQty			= iCurColumnPos(11)
		C_BadQty			= iCurColumnPos(12)
		C_InspQty			= iCurColumnPos(13)
		C_TrnsQty			= iCurColumnPos(14)
		C_PrevOnhandQty		= iCurColumnPos(15)
		C_PrevGoodQty		= iCurColumnPos(16)
		C_PrevBadQty		= iCurColumnPos(17)
		C_PrevInspQty		= iCurColumnPos(18)
		C_PrevTrnsQty		= iCurColumnPos(19)
		C_PikingQty			= iCurColumnPos(20)
		C_LastRcptDt		= iCurColumnPos(21)
		C_LastIssueDt		= iCurColumnPos(22)
		C_ExpiaryDt			= iCurColumnPos(23)
		C_LastPhyDt			= iCurColumnPos(24)
		C_PriceFlag			= iCurColumnPos(25)
		C_Price				= iCurColumnPos(26)
	
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
		frm1.txtPlantCd.Value    = arrRet(0)  
		frm1.txtPlantNm.Value    = arrRet(1)  
		frm1.txtPlantCd.focus
	End If  
End Function


'------------------------------------------  OpenItemAcct()  --------------------------------------------------
'	Name : OpenItemAcct() 
'	Description : Item Account Popup
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
	
	arrHeader(0) = "품목계정코드"	
	arrHeader(1) = "품목계정명"		
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemAcct.focus
		Exit Function
	Else
		frm1.txtItemAcct.Value	    =	arrRet(0)
		frm1.txtItemAcctNm.Value	=	arrRet(1)
		frm1.txtItemAcct.focus
	End If	
	
End Function

 '------------------------------------------  OpenSlCd()  -------------------------------------------------
' Name : OpenSlCd()
' Description : Storage Location Display PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenSlCd()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)
 
 If Trim(frm1.txtPlantCd.Value) = "" then 
	Call DisplayMsgBox("169901","X", "X", "X")    
	frm1.txtPlantCd.focus
	Exit Function
 End if

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "창고조회팝업"   
 arrParam(1) = "B_STORAGE_LOCATION"  
 arrParam(2) = Trim(frm1.txtSlCd.value)  
 arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") 
 arrParam(5) = "창고"    
 
 arrField(0) = "SL_CD"     
 arrField(1) = "SL_NM"     
    
 arrHeader(0) = "창고"   
 arrHeader(1) = "창고명"    
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
	frm1.txtSlCd.focus 
	Exit Function
 Else
	frm1.txtSlCd.value	= arrRet(0) 
	frm1.txtSlNm.value	= arrRet(1)   
	frm1.txtSlCd.focus  
 End If 
 
End Function

'------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If


	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목 팝업"
	arrParam(1) = "B_ITEM_BY_PLANT P,B_ITEM I"	
	arrParam(2) = Trim(frm1.txtItemCd.Value)	
	arrParam(3) = ""							
	arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "품목"			
	
	arrField(0) = "I.ITEM_CD"
	arrField(1) = "I.ITEM_NM"
	
	arrHeader(0) = "품목코드"
	arrHeader(1) = "품목명"	
	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	=arrRet(0)
		frm1.txtItemNm.Value	=arrRet(1)
		frm1.txtItemCd.focus
	End If	
	
End Function


'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("11000000000011")
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


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub

		if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKeyIndex <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
				If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
				End if
		End If
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
 		Exit Sub
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData.MaxRows = 0 Then Exit Sub
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                    
    Err.Clear
    
    If Not chkField(Document, "1") Then Exit Function                                                           

	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	Else
		If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.value = ""
			frm1.txtPlantCd.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtPlantNm.value = lgF0(0)
	End If
	
	If Trim(frm1.txtItemAcct.Value) = "" Then
		Call DisplayMsgBox("169953","X","X","X")
		frm1.txtItemAcctNm.Value = ""
		frm1.txtItemAcct.focus
		Exit function
	Else
		If  CommonQueryRs(" MINOR_NM "," B_MINOR ", "MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & FilterVar(frm1.txtItemAcct.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("169952","X","X","X")
			frm1.txtItemAcctNm.value = ""
			frm1.txtItemAcct.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtItemAcctNm.value = lgF0(0)
	End If
	
	frm1.txtSlNm.value = ""  
	If frm1.txtSlCd.value <> "" Then
		If  CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSlCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		 
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtSlNm.value = lgF0(0)
		End If  
	End If
	

	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtItemNm.value = lgF0(0)
		End if
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
    Call parent.FncExport(parent.C_MULTI)			
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , True) 
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

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	
    Dim strVal
    
    
	On error resume next
    Err.Clear        
    Call LayerShowHide(1)
    
    DbQuery = False
    With frm1    
		strVal = BIZ_PGM_QRY_ID &	"?txtMode="				& Parent.UID_M0001				& _
									"&txtPlantCd="			& Trim(.txtPlantCd.value)		& _
									"&txtItemAcct="			& Trim(.txtItemAcct.value)		& _
									"&txtSlCd="				& Trim(.txtSlCd.value)			& _
									"&txtItemCd="			& Trim(.txtItemCd.value)		& _
									"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
									"&txtMaxRows="			& .vspdData.MaxRows
		Call RunMyBizASP(MyBizASP, strVal)
    End With
    DbQuery = True
                             

End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()		
	frm1.vspdData.focus 
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고현황상세조회</font></TD>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14">	
								</TD>
								<TD CLASS="TD5" NOWRAP>품목계정</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemAcct" SIZE="8" MAXLENGTH="2" tag="12XXXU" ALT="품목계정"  ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAcct()" >&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT Name="txtSlCd" SIZE="8" MAXLENGTH="7"  ALT="창고" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnSlCd align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSlCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" size="30" tag="14">
								</TD>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCd" SIZE="18" MAXLENGTH="18" ALT="품목" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top  TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="14">
								</TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=*>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
								<script language =javascript src='./js/i3002qa1_OBJECT1_vspdData.js'></script>
							</TD>
						</TR>											
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
	
	
	
	
	
